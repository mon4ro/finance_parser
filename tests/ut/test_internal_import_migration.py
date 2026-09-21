from finance_parser.utilities.migrate_internal_package_imports import REPLACEMENTS, TARGET_FILES

def test_internal_import_migration_targets_finance_parser_files_only():
    assert TARGET_FILES
    assert all(path.startswith("finance_parser/") for path in TARGET_FILES)

def test_internal_import_migration_replaces_old_shim_imports():
    old_imports = [old for replacements in REPLACEMENTS.values() for old, _ in replacements]
    assert "from common import (" in old_imports
    assert "from settings import get_settings" in old_imports
    assert "from parsers import op, norwegian, nordea, spankki" in old_imports
    assert "from investments.investment_common import INVESTMENT_RAW_COLUMNS" in old_imports
