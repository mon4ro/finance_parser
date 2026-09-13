# Local cleanup guide

These files/folders are local artifacts and should not be committed:

```text
.DS_Store
.pytest_cache/
__pycache__/
*.pyc
input/
output/
config/settings.yaml
rules/budgeting/TransactionRules.xlsx
rules/investments/InstrumentMaster.xlsx
```

## Safe cleanup command

Dry run:

```bash
python utilities/cleanup_local_artifacts.py --dry-run
```

Real cleanup:

```bash
python utilities/cleanup_local_artifacts.py
```

The helper removes only typical local/generated artifacts:

```text
.DS_Store
__pycache__/
*.pyc
.pytest_cache/
```

It does not remove:

```text
input/
output/
config/settings.yaml
personal rule workbooks
```

## If artifacts are already tracked by Git

Use:

```bash
git rm --cached -r .pytest_cache
git rm --cached -r '**/__pycache__'
git rm --cached '**/.DS_Store'
```

Exact shell quoting can differ by shell. Check with:

```bash
git status
```

before committing.
