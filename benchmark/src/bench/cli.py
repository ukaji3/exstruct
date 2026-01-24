from __future__ import annotations

import json
from pathlib import Path
from typing import Any

from rich import print
from rich.console import Console
import typer

from .eval.exact_match import exact_match
from .eval.normalize import normalize_json_text
from .eval.report import write_results_csv
from .llm.openai_client import OpenAIResponsesClient
from .manifest import Case, load_manifest
from .paths import DATA_DIR, EXTRACTED_DIR, PROMPTS_DIR, RESPONSES_DIR, RESULTS_DIR
from .pipeline.common import ensure_dir, sha256_text, write_json
from .pipeline.exstruct_adapter import extract_exstruct
from .pipeline.html_text import html_to_text, xlsx_to_html
from .pipeline.image_render import xlsx_to_pngs_via_pdf
from .pipeline.openpyxl_pandas import extract_openpyxl
from .pipeline.pdf_text import pdf_to_text, xlsx_to_pdf

app = typer.Typer(add_completion=False)
console = Console()

METHODS_TEXT = ["exstruct", "openpyxl", "pdf", "html"]
METHODS_ALL = ["exstruct", "openpyxl", "pdf", "html", "image_vlm"]


def _manifest_path() -> Path:
    return DATA_DIR / "manifest.json"


def _select_cases(manifest_cases: list[Case], case: str) -> list[Case]:
    if case == "all":
        return manifest_cases
    ids = {c.strip() for c in case.split(",")}
    return [c for c in manifest_cases if c.id in ids]


def _select_methods(method: str) -> list[str]:
    if method == "all":
        return METHODS_ALL
    return [m.strip() for m in method.split(",")]


@app.command()
def extract(case: str = "all", method: str = "all") -> None:
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    methods = _select_methods(method)

    for c in cases:
        xlsx = Path(c.xlsx)
        console.rule(f"EXTRACT {c.id} ({xlsx.name})")

        if "exstruct" in methods:
            out_txt = EXTRACTED_DIR / "exstruct" / f"{c.id}.txt"
            extract_exstruct(xlsx, out_txt, c.sheet_scope)
            print(f"[green]exstruct -> {out_txt}[/green]")

        if "openpyxl" in methods:
            out_txt = EXTRACTED_DIR / "openpyxl" / f"{c.id}.txt"
            extract_openpyxl(xlsx, out_txt, c.sheet_scope)
            print(f"[green]openpyxl -> {out_txt}[/green]")

        if "pdf" in methods:
            out_pdf = EXTRACTED_DIR / "pdf" / f"{c.id}.pdf"
            out_txt = EXTRACTED_DIR / "pdf" / f"{c.id}.txt"
            xlsx_to_pdf(xlsx, out_pdf)
            pdf_to_text(out_pdf, out_txt)
            print(f"[green]pdf -> {out_txt}[/green]")

        if "html" in methods:
            out_html = EXTRACTED_DIR / "html" / f"{c.id}.html"
            out_txt = EXTRACTED_DIR / "html" / f"{c.id}.txt"
            xlsx_to_html(xlsx, out_html)
            html_to_text(out_html, out_txt)
            print(f"[green]html -> {out_txt}[/green]")

        if "image_vlm" in methods:
            out_dir = EXTRACTED_DIR / "image_vlm" / c.id
            pngs = xlsx_to_pngs_via_pdf(
                xlsx, out_dir, dpi=c.render.dpi, max_pages=c.render.max_pages
            )
            write_json(out_dir / "images.json", {"images": [str(p) for p in pngs]})
            print(f"[green]image_vlm -> {len(pngs)} png(s) in {out_dir}[/green]")


@app.command()
def ask(case: str = "all", method: str = "all", model: str = "gpt-4o") -> None:
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    methods = _select_methods(method)

    client = OpenAIResponsesClient()
    ensure_dir(PROMPTS_DIR)
    ensure_dir(RESPONSES_DIR)

    for c in cases:
        console.rule(f"ASK {c.id}")
        q = c.question

        for m in methods:
            prompt_rec: dict[str, Any] = {
                "case_id": c.id,
                "method": m,
                "model": model,
                "question": q,
            }
            resp_rec: dict[str, Any] = {"case_id": c.id, "method": m, "model": model}

            if m == "image_vlm":
                img_dir = EXTRACTED_DIR / "image_vlm" / c.id
                imgs = json.loads(
                    (img_dir / "images.json").read_text(encoding="utf-8")
                )["images"]
                img_paths = [Path(p) for p in imgs]
                prompt_hash = sha256_text(
                    q + "|" + "|".join([p.name for p in img_paths])
                )
                prompt_rec["prompt_hash"] = prompt_hash
                prompt_rec["images"] = [p.name for p in img_paths]

                res = client.ask_images(model=model, question=q, image_paths=img_paths)

            else:
                txt_path = EXTRACTED_DIR / m / f"{c.id}.txt"
                context = txt_path.read_text(encoding="utf-8")
                prompt_hash = sha256_text(q + "|" + context)
                prompt_rec["prompt_hash"] = prompt_hash

                res = client.ask_text(model=model, question=q, context_text=context)

            # save prompt/response
            prompt_file = PROMPTS_DIR / f"{c.id}.jsonl"
            resp_file = RESPONSES_DIR / f"{c.id}.jsonl"

            prompt_rec_line = json.dumps(prompt_rec, ensure_ascii=False)
            resp_rec.update(
                {
                    "prompt_hash": prompt_hash,
                    "text": res.text,
                    "input_tokens": res.input_tokens,
                    "output_tokens": res.output_tokens,
                    "cost_usd": res.cost_usd,
                    "raw": res.raw,
                }
            )
            resp_rec_line = json.dumps(resp_rec, ensure_ascii=False)

            with prompt_file.open("a", encoding="utf-8") as f:
                f.write(prompt_rec_line + "\n")
            with resp_file.open("a", encoding="utf-8") as f:
                f.write(resp_rec_line + "\n")

            print(
                f"[cyan]{c.id} {m}[/cyan] tokens(in/out)={res.input_tokens}/{res.output_tokens} cost=${res.cost_usd:.6f}"
            )


@app.command()
def eval(case: str = "all", method: str = "all") -> None:
    mf = load_manifest(_manifest_path())
    cases = _select_cases(mf.cases, case)
    methods = _select_methods(method)

    rows: list[dict[str, Any]] = []

    for c in cases:
        truth = json.loads(Path(c.truth).read_text(encoding="utf-8"))
        resp_file = RESPONSES_DIR / f"{c.id}.jsonl"
        if not resp_file.exists():
            print(f"[yellow]skip: no responses for {c.id}[/yellow]")
            continue

        # 最新の各method結果を採用（同じmethodが複数行ある場合、最後の行が最新）
        latest: dict[str, dict[str, Any]] = {}
        for line in resp_file.read_text(encoding="utf-8").splitlines():
            rec = json.loads(line)
            if rec["method"] in methods:
                latest[rec["method"]] = rec

        for m, rec in latest.items():
            ok = False
            pred_obj = None
            err = None
            try:
                pred_obj = normalize_json_text(rec["text"])
                ok = exact_match(pred_obj, truth)
            except Exception as e:
                err = str(e)

            rows.append(
                {
                    "case_id": c.id,
                    "type": c.type,
                    "method": m,
                    "model": rec.get("model"),
                    "ok": ok,
                    "input_tokens": rec.get("input_tokens", 0),
                    "output_tokens": rec.get("output_tokens", 0),
                    "cost_usd": rec.get("cost_usd", 0.0),
                    "error": err,
                }
            )

    out_csv = RESULTS_DIR / "results.csv"
    write_results_csv(rows, out_csv)
    print(f"[green]Wrote {out_csv} ({len(rows)} rows)[/green]")


@app.command()
def report() -> None:
    """
    雑に Markdown レポートを作る（必要なら後で強化）
    """
    csv_path = RESULTS_DIR / "results.csv"
    if not csv_path.exists():
        raise typer.Exit(code=1)

    import pandas as pd

    df = pd.read_csv(csv_path)
    # 集計: method別の正解率/平均トークン/平均コスト
    g = (
        df.groupby("method")
        .agg(
            acc=("ok", "mean"),
            avg_in=("input_tokens", "mean"),
            avg_cost=("cost_usd", "mean"),
            n=("ok", "count"),
        )
        .reset_index()
    )

    md_lines = []
    md_lines.append("# Benchmark Report")
    md_lines.append("")
    md_lines.append("## Summary by method")
    md_lines.append("")
    md_lines.append(g.to_markdown(index=False))
    md_lines.append("")
    out_md = RESULTS_DIR / "report.md"
    out_md.write_text("\n".join(md_lines), encoding="utf-8")
    print(f"[green]Wrote {out_md}[/green]")
