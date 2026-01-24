from __future__ import annotations

import base64
from dataclasses import dataclass
import json
from pathlib import Path
from typing import Any

from dotenv import load_dotenv
from openai import OpenAI

from .pricing import estimate_cost_usd


@dataclass
class LLMResult:
    text: str
    input_tokens: int
    output_tokens: int
    cost_usd: float
    raw: dict[str, Any]


def _png_to_data_url(png_path: Path) -> str:
    b = png_path.read_bytes()
    b64 = base64.b64encode(b).decode("ascii")
    return f"data:image/png;base64,{b64}"


class OpenAIResponsesClient:
    def __init__(self) -> None:
        load_dotenv()
        self.client = OpenAI()

    def ask_text(self, *, model: str, question: str, context_text: str) -> LLMResult:
        """
        Responses API: text-only
        """
        resp = self.client.responses.create(
            model=model,
            input=[
                {
                    "role": "user",
                    "content": [
                        {
                            "type": "input_text",
                            "text": "You are a strict JSON extraction engine. Output JSON only.",
                        },
                        {"type": "input_text", "text": f"[QUESTION]\n{question}"},
                        {"type": "input_text", "text": f"[CONTEXT]\n{context_text}"},
                    ],
                }
            ],
        )

        text = resp.output_text  # SDK helper
        usage = getattr(resp, "usage", None) or {}
        in_tok = int(usage.get("input_tokens", 0))
        out_tok = int(usage.get("output_tokens", 0))
        cost = estimate_cost_usd(model, in_tok, out_tok)

        raw = json.loads(resp.model_dump_json())
        return LLMResult(
            text=text,
            input_tokens=in_tok,
            output_tokens=out_tok,
            cost_usd=cost,
            raw=raw,
        )

    def ask_images(
        self, *, model: str, question: str, image_paths: list[Path]
    ) -> LLMResult:
        """
        Responses API: image + text
        """
        content: list[dict[str, Any]] = [
            {
                "type": "input_text",
                "text": "You are a strict JSON extraction engine. Output JSON only.",
            },
            {"type": "input_text", "text": f"[QUESTION]\n{question}"},
        ]
        for p in image_paths:
            content.append({"type": "input_image", "image_url": _png_to_data_url(p)})

        resp = self.client.responses.create(
            model=model,
            input=[{"role": "user", "content": content}],
        )

        text = resp.output_text
        usage = getattr(resp, "usage", None) or {}
        in_tok = int(usage.get("input_tokens", 0))
        out_tok = int(usage.get("output_tokens", 0))
        cost = estimate_cost_usd(model, in_tok, out_tok)

        raw = json.loads(resp.model_dump_json())
        return LLMResult(
            text=text,
            input_tokens=in_tok,
            output_tokens=out_tok,
            cost_usd=cost,
            raw=raw,
        )
