# Finance Parser

A modular personal finance parser for bank transactions and investment transactions.

The project keeps parsing, normalisation, categorisation, and local manual editing concerns separate.

## Setup

Run the interactive setup wizard to generate `config/settings.yaml` without hand-editing
YAML - it asks what data you'll import (budgeting/investments/both), which of the
supported banks/brokers you use, and the account-level facts each one needs (owner
labels, OP filename prefixes, portfolio numbers, etc.), then shows you the exact result
and asks for confirmation before writing anything. Re-running it later only adds/updates
what you answer - it never drops or overwrites unrelated existing settings:

```bash
python setup_wizard.py
```

Try it risk-free first with `--dry-run` - it copies your real `config/settings.yaml`
and rule workbooks (if they exist) into a disposable temporary sandbox, runs the whole
wizard against that instead, and never touches the real files:

```bash
python setup_wizard.py --dry-run
python setup_wizard.py --dry-run --keep-temp  # inspect the sandbox output afterward
```

Or copy the example settings and edit the private local file by hand:

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
python -m finance_parser.budgeting.enrich_dividend_income
python -m finance_parser.budgeting.transaction_categoriser
```

Useful individual dry-runs:

```bash
python -m finance_parser.budgeting.transaction_parser --dry-run
python -m finance_parser.budgeting.transaction_normaliser --dry-run
python -m finance_parser.budgeting.enrich_dividend_income --dry-run
python -m finance_parser.budgeting.transaction_categoriser --dry-run
```

The full pipeline (`run_budgeting_pipeline.py`) also runs two dividend-income
stages between the normaliser and categoriser: an extra parser pass that
turns the investment side's own `DividendHistory.xlsx` into real income rows
for Nordnet/EVLI-held instruments (which never reach a bank transaction at
all - the cash sits in the broker's own balance instead), and
`enrich_dividend_income.py`, which cross-references OP-held instruments'
dividends against their real bank deposit by date and exact amount. Both
auto-skip with a message if `DividendHistory.xlsx` doesn't exist yet (i.e.
the investment pipeline hasn't been run) - pass `--skip-dividends` to
disable them explicitly.

The full pipeline also fetches a monthly snapshot of each actively-held fund's top-10
constituent holdings (`FundHoldingsSnapshot.xlsx`) - self-throttled to run once per
calendar month per fund, for future portfolio look-through/concentration analysis (e.g.
"how much of my total investments are really tied to Nokia, across direct stock plus
fund exposure"). Uses an unofficial Yahoo Finance endpoint, meaningfully more fragile
than the price-fetch API - a fund with no data, or a failed fetch, is recorded as 100%
`UNKNOWN` rather than silently dropped. There's no historical holdings data available
from Yahoo, so a fund's first-ever fetch also backfills 60 months of assumed-constant
history, tagged `BACKFILLED_ASSUMED_CONSTANT` to keep it distinguishable from a real
observed snapshot (`LIVE_FETCH`). Pass `--skip-fund-holdings` to disable it.

Once `FundHoldingsSnapshot.xlsx` exists, run the look-through exposure calculator
on demand (not part of the main pipeline - nothing downstream depends on it) to
see real total exposure per company, combining direct stock holdings with
indirect exposure via funds that also hold the same company:

```bash
python -m finance_parser.investments.build_look_through_exposure
```

Writes `output/investments/LookThroughExposure.xlsx` (a `LookThroughByCompany`
summary sheet plus a `LookThroughDetail` sheet showing which source instrument
contributed how much to each total). A stock held both directly and inside a
fund is merged into one combined total, matched by the shared Yahoo ticker
(`PriceSymbol`) - not by name, since a fund's constituent listing and this
project's own instrument naming don't always spell a company the same way.

## Investment workflow

Run the full investment pipeline (parse, fetch prices, fetch FX rates, build positions,
fetch fund holdings snapshots, build dividend history, check coverage, build the monthly
value rollup):

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
python -m finance_parser.investments.fetch_fund_holdings
python -m finance_parser.investments.build_dividend_history
python -m finance_parser.investments.check_instrument_coverage
python -m finance_parser.investments.build_monthly_position_value
```

Each script has its own `--help`; use `--force` on the pipeline (or on
`check_instrument_coverage` directly) to proceed past incomplete coverage instead of
halting. Dividend history also does a best-effort, read-only enrichment from the
budgeting side's own OP dividend notices (local-currency amount/exchange rate, e.g.
Telia's SEK dividends) - use `--no-local-currency` to skip it.

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
