# Sharing and local personal files

The project should be split into generic code/templates and user-specific local files.

## Safe to commit/share

```text
settings.py
config/settings.example.yaml
rules/budgeting/TransactionRules.template.xlsx
rules/investments/InstrumentMaster.template.xlsx
parsers/
investments/
tests/
docs/
```

## Keep local / do not share

```text
config/settings.yaml
rules/budgeting/TransactionRules.xlsx
rules/investments/InstrumentMaster.xlsx
input/
output/
```

These may contain personal account labels, merchant names, ownership rules,
source files, parsed transactions, or portfolio data.

## First-time setup for another user

```bash
cp config/settings.example.yaml config/settings.yaml
cp rules/budgeting/TransactionRules.template.xlsx rules/budgeting/TransactionRules.xlsx
cp rules/investments/InstrumentMaster.template.xlsx rules/investments/InstrumentMaster.xlsx
```

Then edit the copied files.

## Git note

If a personal file has already been committed, adding it to `.gitignore` is not enough.
It must also be removed from tracking with:

```bash
git rm --cached config/settings.yaml
git rm --cached rules/budgeting/TransactionRules.xlsx
git rm --cached rules/investments/InstrumentMaster.xlsx
```

Keep the files locally; `--cached` only removes them from Git tracking.
