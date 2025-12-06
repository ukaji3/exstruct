import subprocess
import sys
from pathlib import Path

import pytest

import xlwings as xw
from openpyxl import Workbook
from importlib import util


def _excel_available() -> bool:
    try:
        app = xw.App(add_book=False, visible=False)
        app.quit()
        return True
    except Exception:
        return False


def _pypdfium_available() -> bool:
    try:
        import pypdfium2  # noqa: F401

        return True
    except Exception:
        return False


def _toon_available() -> bool:
    try:
        import toon  # noqa: F401

        return True
    except Exception:
        return False


def _prepare_sample_excel(tmp_path: Path) -> Path:
    """
    Prepare a minimal Excel workbook for CLI tests.
    - If repo sample exists, copy it.
    - Otherwise, generate a tiny workbook with openpyxl.
    """
    sample = Path("sample") / "sample.xlsx"
    dest = tmp_path / "sample.xlsx"
    if sample.exists():
        import shutil

        shutil.copy(sample, dest)
        return dest

    wb = Workbook()
    ws = wb.active
    ws.title = "Sheet1"
    ws.append(["A", "B"])
    ws.append([1, 2])
    wb.save(dest)
    wb.close()
    return dest


def test_CLIでjson出力が成功する(tmp_path: Path) -> None:
    xlsx = _prepare_sample_excel(tmp_path)
    out_json = tmp_path / "out.json"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(xlsx), "-o", str(out_json)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert out_json.exists()
    # stdout may be empty when writing to a file; ensure no errors surfaced
    assert "Error" not in result.stdout


def test_CLIでyamlやtoon指定は未サポート(tmp_path: Path) -> None:
    xlsx = _prepare_sample_excel(tmp_path)
    out_yaml = tmp_path / "out.yaml"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(xlsx), "-o", str(out_yaml), "-f", "yaml"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if util.find_spec("yaml") is not None:
        assert result.returncode == 0
        assert out_yaml.exists()
    else:
        assert result.returncode != 0
        assert "pyyaml" in result.stdout or "pyyaml" in result.stderr

    out_toon = tmp_path / "out.toon"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(xlsx), "-o", str(out_toon), "-f", "toon"]
    result = subprocess.run(cmd, capture_output=True, text=True)
    if _toon_available():
        assert result.returncode == 0
        assert out_toon.exists()
    else:
        assert result.returncode != 0
        assert "TOON export requires python-toon" in result.stdout


@pytest.mark.skipif(
    not _excel_available() or not _pypdfium_available(),
    reason="Excel COM or pypdfium2 unavailable; skipping PDF/PNG export tests.",
)
def test_CLIでpdfと画像が出力される(tmp_path: Path) -> None:
    xlsx = _prepare_sample_excel(tmp_path)
    out_json = tmp_path / "out.json"
    cmd = [
        sys.executable,
        "-m",
        "exstruct.cli.main",
        str(xlsx),
        "-o",
        str(out_json),
        "--pdf",
        "--image",
    ]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    pdf_path = out_json.with_suffix(".pdf")
    images_dir = out_json.parent / f"{out_json.stem}_images"
    assert pdf_path.exists()
    assert images_dir.exists()
    assert any(images_dir.glob("*.png"))


def test_CLIで無効ファイルは安全終了する(tmp_path: Path) -> None:
    bad_path = tmp_path / "nope.xlsx"
    out_json = tmp_path / "out.json"
    cmd = [sys.executable, "-m", "exstruct.cli.main", str(bad_path), "-o", str(out_json)]
    result = subprocess.run(cmd, capture_output=True, text=True)
    assert result.returncode == 0
    assert "not found" in (result.stdout + result.stderr).lower()
