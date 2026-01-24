from __future__ import annotations

from pathlib import Path
import subprocess

import fitz  # PyMuPDF

from .common import ensure_dir, write_text


def xlsx_to_pdf(xlsx_path: Path, out_pdf: Path) -> None:
    ensure_dir(out_pdf.parent)
    # LibreOffice headless convert
    # soffice --headless --convert-to pdf --outdir <dir> <xlsx>
    cmd = [
        "soffice",
        "--headless",
        "--nologo",
        "--nolockcheck",
        "--convert-to",
        "pdf",
        "--outdir",
        str(out_pdf.parent),
        str(xlsx_path),
    ]
    subprocess.run(cmd, check=True)
    produced = out_pdf.parent / (xlsx_path.stem + ".pdf")
    produced.replace(out_pdf)


def pdf_to_text(pdf_path: Path, out_txt: Path) -> None:
    doc = fitz.open(pdf_path)
    parts: list[str] = []
    for i in range(doc.page_count):
        page = doc.load_page(i)
        parts.append(f"\n# PAGE {i + 1}")
        parts.append(page.get_text("text"))
    text = "\n".join(parts).strip()

    lines: list[str] = []
    lines.append("[DOC_META]")
    lines.append(f"source={pdf_path.name}")
    lines.append("method=pdf_text")
    lines.append("")
    lines.append("[CONTENT]")
    lines.append(text)

    write_text(out_txt, "\n".join(lines).strip() + "\n")
