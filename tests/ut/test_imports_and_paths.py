import importlib
import py_compile
from pathlib import Path


PROJECT_ROOT = Path(__file__).resolve().parents[2]


def test_core_modules_import():
    modules = [
        "finance_parser.budgeting.transaction_parser",
        "finance_parser.budgeting.transaction_normaliser",
        "finance_parser.budgeting.transaction_categoriser",
        "finance_parser.budgeting.run_budgeting_pipeline",
        "finance_parser.investments.investment_parser",
    ]

    for module_name in modules:
        importlib.import_module(module_name)


def test_root_budgeting_pipeline_wrapper_compiles():
    py_compile.compile(str(PROJECT_ROOT / "run_budgeting_pipeline.py"), doraise=True)


def test_legacy_budgeting_root_wrappers_are_absent():
    assert not (PROJECT_ROOT / "transaction_parser.py").exists()
    assert not (PROJECT_ROOT / "transaction_normaliser.py").exists()
    assert not (PROJECT_ROOT / "transaction_categoriser.py").exists()


def test_expected_budgeting_package_scripts_exist():
    assert (PROJECT_ROOT / "finance_parser" / "budgeting" / "transaction_parser.py").exists()
    assert (PROJECT_ROOT / "finance_parser" / "budgeting" / "transaction_normaliser.py").exists()
    assert (PROJECT_ROOT / "finance_parser" / "budgeting" / "transaction_categoriser.py").exists()
    assert (PROJECT_ROOT / "finance_parser" / "budgeting" / "run_budgeting_pipeline.py").exists()
