"""재시도/이미지 페이로드 처리를 포함한 OpenAI 클라이언트 래퍼.

(레거시 대비 변경점)
- 설치된 openai 라이브러리가 지원하는 경우 JSON 모드(response_format)를 선택적으로 사용.
- temperature를 선택적으로 설정(일부 모델은 이 파라미터를 거부함).
- 촘촘한 표가 과도하게 압축되지 않도록 이미지 용량 한도를 약간 더 크게 사용.
- 호환성을 위해 max_completion_tokens ↔ max_tokens를 상황에 따라 폴백.

참고
- openai가 설치되지 않은 환경에서도 writer/config가 동작할 수 있도록 openai는 지연 import한다.
"""

from __future__ import annotations

import base64
import io
import time
from dataclasses import dataclass
from typing import Any, Dict, List, Optional, Sequence, Tuple

from PIL import Image


# 교차표가 조밀해 과도한 압축 시 품질이 크게 떨어짐: 중간 수준 유지
MAX_IMAGE_SIZE_KB = 3500


@dataclass
class OpenAIOptions:
    model: str
    max_completion_tokens: int = 16000
    max_retries: int = 3
    # 참고: gpt-5 계열/mini는 temperature 파라미터 자체를 거부하는 경우가 있습니다.
    # 기본값은 None(미전송)으로 두고, 필요할 때만 명시적으로 지정하세요.
    temperature: Optional[float] = None
    json_mode: bool = True


def _compress_image(
    image: Image.Image,
    max_size_kb: int = MAX_IMAGE_SIZE_KB,
    quality: int = 85,
) -> Tuple[bytes, str]:
    """(바이트, 포맷) 반환. 포맷은 PNG 또는 JPEG."""

    buf = io.BytesIO()
    image.save(buf, format="PNG")
    size_kb = len(buf.getvalue()) / 1024
    if size_kb <= max_size_kb:
        return buf.getvalue(), "PNG"

    im = image
    if im.mode == "RGBA":
        im = im.convert("RGB")

    # JPEG 품질 단계별 시도
    for q in (quality, 80, 75, 70, 65, 60, 55, 50, 45):
        buf = io.BytesIO()
        im.save(buf, format="JPEG", quality=q, optimize=True)
        if len(buf.getvalue()) / 1024 <= max_size_kb:
            return buf.getvalue(), "JPEG"

    # 그래도 크면 가장 낮은 품질 결과 반환
    return buf.getvalue(), "JPEG"


def _image_to_content(image: Image.Image) -> dict:
    img_bytes, fmt = _compress_image(image)
    base64_image = base64.b64encode(img_bytes).decode("utf-8")
    mime_type = "image/png" if fmt == "PNG" else "image/jpeg"
    return {
        "type": "image_url",
        "image_url": {"url": f"data:{mime_type};base64,{base64_image}"},
    }


class OpenAIChatClient:
    def __init__(self, api_key: str) -> None:
        try:
            import openai  # type: ignore
        except Exception as e:
            raise ImportError(
                "openai 패키지가 설치되어 있지 않습니다.\n"
                "  pip install openai\n"
                "를 실행한 뒤 다시 시도하세요."
            ) from e

        self._openai = openai
        self._client = openai.OpenAI(api_key=api_key)
        # 모델별로 특정 파라미터(예: temperature)가 거부되는 경우가 있어
        # 한 번 감지되면 같은 런에서 재시도/중복 호출을 줄이기 위해 캐시합니다.
        self._no_temperature_models: set[str] = set()

    def chat(self, prompt: str, images: Sequence[Image.Image], options: OpenAIOptions) -> str:
        content: List[dict] = [{"type": "text", "text": prompt}]
        for im in images:
            content.append(_image_to_content(im))

        openai = self._openai
        last_error: Exception | None = None

        for attempt in range(int(options.max_retries or 1)):
            try:
                kwargs: Dict[str, Any] = {
                    "model": options.model,
                    "messages": [{"role": "user", "content": content}],
                    # 참고: 일부 모델(예: gpt-5-mini)은 temperature 파라미터를 허용하지 않습니다.
                    # 아래에서 필요 시(지원되는 경우)만 넣고, 미지원이면 자동으로 제거 후 재시도합니다.
                }

                # 명시적으로 요청된 경우에만 temperature를 설정
                # (gpt-5 계열 등 일부 모델은 기본값만 허용하며 파라미터를 보내면 400 발생)
                if options.temperature is not None and options.model not in self._no_temperature_models:
                    kwargs["temperature"] = float(options.temperature)

                # max_completion_tokens 우선, 없으면 max_tokens로 폴백
                kwargs["max_completion_tokens"] = int(options.max_completion_tokens)

                if options.json_mode:
                    kwargs["response_format"] = {"type": "json_object"}

                def _do_create(call_kwargs: Dict[str, Any]):
                    """호환성을 고려해 ChatCompletions 호출."""
                    try:
                        return self._client.chat.completions.create(**call_kwargs)
                    except TypeError:
                        # 호환 경로(구버전 openai용)
                        call_kwargs = dict(call_kwargs)
                        call_kwargs.pop("response_format", None)
                        mct = call_kwargs.pop("max_completion_tokens", None)
                        if mct is not None:
                            call_kwargs["max_tokens"] = mct
                        return self._client.chat.completions.create(**call_kwargs)

                try:
                    resp = _do_create(kwargs)
                except openai.BadRequestError as e:
                    # 일부 모델이 특정 파라미터(temperature 등)를 거부할 수 있음
                    # 감지되면 해당 파라미터를 제거하고 한 번 재시도
                    body = getattr(e, "body", None) or {}
                    err = (body.get("error") or {}) if isinstance(body, dict) else {}
                    param = err.get("param")
                    code = err.get("code")
                    msg = str(e)

                    if (
                        (param == "temperature")
                        or ("temperature" in msg and "Only the default" in msg)
                        or (code == "unsupported_value" and "temperature" in msg)
                    ):
                        self._no_temperature_models.add(options.model)
                        kwargs.pop("temperature", None)
                        resp = _do_create(kwargs)
                    else:
                        raise

                return resp.choices[0].message.content

            except openai.RateLimitError as e:
                last_error = e
                time.sleep((2**attempt) * 2)
            except (openai.APITimeoutError, openai.APIConnectionError) as e:
                last_error = e
                time.sleep((2**attempt) * 1)
            except openai.APIError as e:
                last_error = e
                if attempt < int(options.max_retries or 1) - 1:
                    time.sleep(1)

        raise last_error or RuntimeError("OpenAI call failed")
