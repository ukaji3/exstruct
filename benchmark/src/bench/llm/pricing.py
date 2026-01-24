from __future__ import annotations

# gpt-4o pricing: per 1M tokens
# Input $2.50 / 1M, Output $10.00 / 1M (cached input not used here)
# Source: model compare page
# https://platform.openai.com/docs/models/compare?model=gpt-4o
# (You will cite this in README/report; code keeps constants.)
GPT4O_INPUT_PER_1M = 2.50
GPT4O_OUTPUT_PER_1M = 10.00


def estimate_cost_usd(model: str, input_tokens: int, output_tokens: int) -> float:
    if model != "gpt-4o":
        # ベンチでは統一前提。拡張するならここをテーブル化。
        raise ValueError(f"Unsupported model for cost table: {model}")

    return (input_tokens / 1_000_000) * GPT4O_INPUT_PER_1M + (
        output_tokens / 1_000_000
    ) * GPT4O_OUTPUT_PER_1M
