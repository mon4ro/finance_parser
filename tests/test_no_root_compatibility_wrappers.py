import importlib
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_root_compatibility_wrappers_are_absent():
    assert not (PROJECT_ROOT / "settings.py").exists()
    assert not (PROJECT_ROOT / "common.py").exists()
    assert not (PROJECT_ROOT / "transaction_parser.py").exists()
    assert not (PROJECT_ROOT / "transaction_normaliser.py").exists()
    assert not (PROJECT_ROOT / "transaction_categoriser.py").exists()


def test_settings_imports_from_package():
    module = importlib.import_module("finance_parser.settings")
    assert hasattr(module, "AppSettings")
    assert hasattr(module, "DEFAULT_EXAMPLE_SETTINGS_PATH")
