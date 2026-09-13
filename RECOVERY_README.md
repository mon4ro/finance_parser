# Personal finance project recovery bundle

This bundle reconstructs the latest known project files from the patches and
uploaded files available in the ChatGPT session.

## What is included

Budgeting parser:
- transaction_parser.py
- common.py
- parsers/op.py
- parsers/nordea.py
- parsers/norwegian.py
- parsers/spankki.py
- parsers/__init__.py

Investment parser:
- investment_parser.py
- investments/
- investments/parsers/nordnet.py
- investments/parsers/seligson.py
- rules/investments/InstrumentMaster.xlsx

Folder structure:
- output/budgeting/
- output/investments/
- rules/budgeting/
- rules/investments/
- docs/patch_notes/

## Expected default outputs

Budgeting:
```bash
python transaction_parser.py --input path/to/bank_exports
```
writes to:
```text
output/budgeting/ParsedTransactions.xlsx
```

Investments:
```bash
python investment_parser.py --input path/to/investment_exports
```
writes to:
```text
output/investments/ParsedInvestments.xlsx
```

## Important warning

Do not copy a patch by replacing whole folders such as `rules/`, `output/`,
`docs/`, or `investments/` unless the patch is explicitly a full recovery
bundle like this one.

Safer method:
1. Extract the zip into a temporary folder.
2. Inspect the file list.
3. Copy individual files over, or use `rsync`/`ditto` merge behaviour.
4. Commit or backup before applying future patches.

## Suggested immediate Git recovery habit

After restoring this bundle and confirming scripts run:

```bash
git status
git add .
git commit -m "Recover parser project state"
```

Then before applying any future patch:

```bash
git status
git add .
git commit -m "Checkpoint before applying generated patch"
```

## Caveat

This recovery bundle reflects the latest state known in the ChatGPT session.
If you had local manual edits not uploaded or described here, they cannot be
reconstructed automatically.
