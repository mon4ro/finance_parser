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

Run the full investment pipeline (parse, fetch prices, fetch FX rates, build positions,
check coverage, build the monthly value rollup):

```bash
python run_investment_pipeline.py
```

Run a full dry-run against temporary copied workbooks:

```bash
python run_investment_pipeline.py --dry-run
```

Skip the network-bound stages (price/FX fetch) for offline dev/debug iteration on
positions, coverage, or rollup logic - uses whatever price/FX data already exists:

```bash
python run_investment_pipeline.py --skip-pricing --skip-fx
```

Run individual investment steps:

```bash
python -m finance_parser.investments.investment_parser
python -m finance_parser.investments.fetch_instrument_prices
python -m finance_parser.investments.fetch_fx_rates
python -m finance_parser.investments.build_portfolio_positions
python -m finance_parser.investments.check_instrument_coverage
python -m finance_parser.investments.build_monthly_position_value
```

Each script has its own `--help`; use `--force` on the pipeline (or on
`check_instrument_coverage` directly) to proceed past incomplete coverage instead of
halting.

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
run_investment_pipeline.py
```
