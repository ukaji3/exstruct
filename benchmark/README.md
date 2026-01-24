# ExStruct Benchmark

This benchmark compares methods for answering questions about Excel documents using GPT-4o:

- exstruct
- openpyxl
- pdf (xlsx->pdf->text)
- html (xlsx->html->table text)
- image_vlm (xlsx->pdf->png -> GPT-4o vision)

## Requirements

- Python 3.11+
- LibreOffice (`soffice` in PATH)
- OPENAI_API_KEY in `.env`

## Setup

```bash
cd benchmark
cp .env.example .env
pip install -e ..  # install exstruct from repo root
pip install -e .
```

## Run

```bash
make all
```

Outputs:

- outputs/extracted/\* : extracted context (text or images)
- outputs/prompts/\*.jsonl
- outputs/responses/\*.jsonl
- outputs/results/results.csv
- outputs/results/report.md

## Notes:

- GPT-4o Responses API supports text and image inputs. See docs:
  - [https://platform.openai.com/docs/api-reference/responses](https://platform.openai.com/docs/api-reference/responses)
  - [https://platform.openai.com/docs/guides/images-vision](https://platform.openai.com/docs/guides/images-vision)
- Pricing for gpt-4o used in cost estimation:
  - https://platform.openai.com/docs/models/compare?model=gpt-4o
