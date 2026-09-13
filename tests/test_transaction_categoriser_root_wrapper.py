import importlib.util
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[1]


def test_legacy_budgeting_root_wrappers_are_removed():
    assert not (PROJECT_ROOT / "transaction_parser.py").exists()
    assert not (PROJECT_ROOT / "transaction_normaliser.py").exists()
    assert not (PROJECT_ROOT / "transaction_categoriser.py").exists()


def test_root_pipeline_wrapper_exists_and_points_to_package():
    path = PROJECT_ROOT / "run_budgeting_pipeline.py"
    assert path.exists()

    source = path.read_text(encoding="utf-8")
    assert "finance_parser.budgeting.run_budgeting_pipeline" in source
    assert "from finance_parser.budgeting.run_budgeting_pipeline import main" in source


def test_root_pipeline_wrapper_imports():
    path = PROJECT_ROOT / "run_budgeting_pipeline.py"
    spec = importlib.util.spec_from_file_location("run_budgeting_pipeline", path)
    assert spec is not None
    assert spec.loader is not None

    module = importlib.util.module_from_spec(spec)
    spec.loader.exec_module(module)

    assert hasattr(module, "main")
