"""Tests for LibreOffice runtime gating in ``tests/conftest.py``."""

from __future__ import annotations

import importlib.util
from pathlib import Path
from types import SimpleNamespace

from _pytest.monkeypatch import MonkeyPatch
import pytest


def test_has_libreoffice_runtime_returns_false_on_expected_probe_failure(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that the smoke gate rejects incompatible bridge runtimes."""

    module_path = Path(__file__).with_name("conftest.py")
    spec = importlib.util.spec_from_file_location("tests_runtime_conftest", module_path)
    if spec is None or spec.loader is None:
        raise RuntimeError("Failed to load tests/conftest.py")
    conftest = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(conftest)

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")
    probe_error = type("ProbeUnavailableError", (RuntimeError,), {})

    def _raise_probe_failure(_path: Path) -> Path:
        """Simulate an explicit probe compatibility failure."""

        raise probe_error("bridge probe failed")

    monkeypatch.setattr(
        conftest,
        "_load_libreoffice_runtime_module",
        lambda: SimpleNamespace(
            LibreOfficeUnavailableError=probe_error,
            _which_soffice=lambda: soffice_path,
            _resolve_python_path=_raise_probe_failure,
        ),
    )
    monkeypatch.delenv("EXSTRUCT_LIBREOFFICE_PATH", raising=False)
    conftest._has_libreoffice_runtime.cache_clear()

    assert conftest._has_libreoffice_runtime() is False

    conftest._has_libreoffice_runtime.cache_clear()


def test_has_libreoffice_runtime_raises_on_unexpected_probe_failure(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that unexpected runtime probe errors fail the test setup loudly."""

    module_path = Path(__file__).with_name("conftest.py")
    spec = importlib.util.spec_from_file_location(
        "tests_runtime_conftest_unexpected", module_path
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("Failed to load tests/conftest.py")
    conftest = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(conftest)

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")

    monkeypatch.setattr(
        conftest,
        "_load_libreoffice_runtime_module",
        lambda: SimpleNamespace(
            LibreOfficeUnavailableError=RuntimeError,
            _which_soffice=lambda: soffice_path,
            _resolve_python_path=lambda _path: (_ for _ in ()).throw(
                ValueError("unexpected probe regression")
            ),
        ),
    )
    monkeypatch.delenv("EXSTRUCT_LIBREOFFICE_PATH", raising=False)
    conftest._has_libreoffice_runtime.cache_clear()

    with pytest.raises(ValueError, match="unexpected probe regression"):
        conftest._has_libreoffice_runtime()

    conftest._has_libreoffice_runtime.cache_clear()


def test_libreoffice_skip_reason_fails_fast_when_forced(
    tmp_path: Path, monkeypatch: MonkeyPatch
) -> None:
    """Verify that required LibreOffice smoke fails instead of skipping."""

    module_path = Path(__file__).with_name("conftest.py")
    spec = importlib.util.spec_from_file_location(
        "tests_runtime_conftest_force", module_path
    )
    if spec is None or spec.loader is None:
        raise RuntimeError("Failed to load tests/conftest.py")
    conftest = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(conftest)

    soffice_path = tmp_path / "soffice"
    soffice_path.write_text("", encoding="utf-8")
    monkeypatch.setattr(
        conftest,
        "_load_libreoffice_runtime_module",
        lambda: SimpleNamespace(
            LibreOfficeUnavailableError=RuntimeError,
            _which_soffice=lambda: soffice_path,
            _resolve_python_path=lambda _path: (_ for _ in ()).throw(
                RuntimeError("bridge probe failed")
            ),
        ),
    )
    monkeypatch.setenv("RUN_LIBREOFFICE_SMOKE", "1")
    monkeypatch.setenv("FORCE_LIBREOFFICE_SMOKE", "1")
    conftest.__dict__["RUN_LIBREOFFICE_SMOKE"] = True
    conftest.__dict__["FORCE_LIBREOFFICE_SMOKE"] = True
    conftest._has_libreoffice_runtime.cache_clear()

    with pytest.raises(RuntimeError, match="FORCE_LIBREOFFICE_SMOKE=1"):
        conftest._libreoffice_skip_reason()

    conftest._has_libreoffice_runtime.cache_clear()
