from __future__ import annotations

from pathlib import Path

ROOT = Path(__file__).resolve().parents[2]  # benchmark/
DATA_DIR = ROOT / "data"
RAW_DIR = DATA_DIR / "raw"
TRUTH_DIR = DATA_DIR / "truth"

OUT_DIR = ROOT / "outputs"
EXTRACTED_DIR = OUT_DIR / "extracted"
PROMPTS_DIR = OUT_DIR / "prompts"
RESPONSES_DIR = OUT_DIR / "responses"
RESULTS_DIR = OUT_DIR / "results"
