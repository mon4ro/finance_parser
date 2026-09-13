# Finance Parser

A modular personal finance parser for bank transactions and investment transactions.

The project keeps parsing, normalisation, categorisation, and local manual editing concerns separate.

## Setup

Copy the public example settings and edit the private local file:

```bash
cp config/settings.example.yaml config/settings.yaml
```

`config/settings.yaml` is for your private local settings and should not be committed.

## Budgeting workflow

Run the full budgeting pipeline:

```bash
python run_budgeting_pipeline.py
```

Run a full dry-run against a temporary copied workbook:

```bash
python run_budgeting_pipeline.py --dry-run
```

Keep the temporary dry-run workbook for inspection:

```bash
python run_budgeting_pipeline.py --dry-run --keep-temp
```

Run individual budgeting steps:

```bash
python -m finance_parser.budgeting.transaction_parser
python -m finance_parser.budgeting.transaction_normaliser
python -m finance_parser.budgeting.transaction_categoriser
```

Useful individual dry-runs:

```bash
python -m finance_parser.budgeting.transaction_parser --dry-run
python -m finance_parser.budgeting.transaction_normaliser --dry-run
python -m finance_parser.budgeting.transaction_categoriser --dry-run
```

## Investment workflow

Run the investment parser:

```bash
python investment_parser.py
```

## Tests

Run the unit tests:

```bash
python -m pytest
```

## Repository layout

```text
config/
docs/
finance_parser/
input/
output/
rules/
tests/
tools/
run_budgeting_pipeline.py
investment_parser.py
```
