# Utility scripts

Canonical utility scripts live in:

```text
utilities/
```

The older root-level copies should be removed.

## Commands

Cleanup/source-account review:

```bash
python utilities/cleanup_budgeting_source_accounts.py --dry-run
python utilities/cleanup_budgeting_source_accounts.py
```

Refresh bank raw IDs:

```bash
python utilities/refresh_budgeting_bank_raw_ids.py --dry-run
python utilities/refresh_budgeting_bank_raw_ids.py
```

Review/cleanup suspicious UnifiedTransactions duplicates:

```bash
python utilities/review_cleanup_unified_duplicates.py --dry-run
python utilities/review_cleanup_unified_duplicates.py
```

Local artifact cleanup:

```bash
python utilities/cleanup_local_artifacts.py --dry-run
python utilities/cleanup_local_artifacts.py
```

## Help

All utility scripts should support:

```bash
python utilities/<script>.py --help
```

## Safety

For workbook-mutating utilities, always run `--dry-run` first.

The local artifact cleanup utility removes only generated/local junk:

```text
.DS_Store
__pycache__/
*.pyc
.pytest_cache/
```

It does not remove personal input/output/config/rule files.
