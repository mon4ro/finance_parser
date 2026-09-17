# Workflow

## Normal budgeting workflow

1. Put new bank export files into:

   ```text
   input/budgeting/
   ```

2. Run the full budgeting pipeline:

   ```bash
   python run_budgeting_pipeline.py
   ```

3. Review:

   ```text
   output/budgeting/ParsedTransactions.xlsx
   ```

The pipeline runs:

```text
1. transaction parser
2. investment dividend parser (Nordnet/EVLI real dividend income, from the
   investment side's DividendHistory.xlsx - auto-skipped if it doesn't exist)
3. transaction normaliser
4. dividend income enrichment (OP-held instrument dividends, cross-referenced
   against their real bank deposit - also auto-skipped if DividendHistory.xlsx
   doesn't exist)
5. transaction categoriser
```

Pass `--skip-dividends` to disable steps 2 and 4 explicitly.

## Budgeting dry-run workflow

Use dry-run before a real import:

```bash
python run_budgeting_pipeline.py --dry-run
```

This runs the real pipeline against a copied temporary workbook, so the real workbook is not modified.

To inspect the simulated workbook:

```bash
python run_budgeting_pipeline.py --dry-run --keep-temp
```

## Individual budgeting steps

The individual budgeting scripts live in the package and can be run directly when needed:

```bash
python -m finance_parser.budgeting.transaction_parser
python -m finance_parser.budgeting.transaction_normaliser
python -m finance_parser.budgeting.enrich_dividend_income
python -m finance_parser.budgeting.transaction_categoriser
```

Individual dry-run commands:

```bash
python -m finance_parser.budgeting.transaction_parser --dry-run
python -m finance_parser.budgeting.transaction_normaliser --dry-run
python -m finance_parser.budgeting.enrich_dividend_income --dry-run
python -m finance_parser.budgeting.transaction_categoriser --dry-run
```

## Normal investment workflow

1. Put new broker export files into:

   ```text
   input/investments/
   ```

2. Run the full investment pipeline:

   ```bash
   python run_investment_pipeline.py
   ```

3. Review:

   ```text
   output/investments/ParsedInvestments.xlsx
   output/investments/PortfolioPositions.xlsx
   output/investments/MonthlyPositionValue.xlsx
   output/investments/DividendHistory.xlsx
   ```

The pipeline runs, in order:

```text
1. investment parser
2. instrument price fetch (Yahoo)
3. FX rate fetch
4. portfolio positions
5. dividend history
6. instrument coverage check (gate)
7. monthly position value rollup
```

Step 6 halts the pipeline if an actively-held instrument is missing classification, a
price symbol, or a recent price - pass `--force` to proceed anyway with incomplete
net worth data.

## Investment dry-run workflow

```bash
python run_investment_pipeline.py --dry-run
python run_investment_pipeline.py --dry-run --keep-temp
```

Same real-workbook-untouched guarantee as the budgeting dry-run.

For offline dev/debug iteration on positions, coverage, or rollup logic without
hitting the network:

```bash
python run_investment_pipeline.py --skip-pricing --skip-fx
```

## Individual investment steps

```bash
python -m finance_parser.investments.investment_parser
python -m finance_parser.investments.fetch_instrument_prices
python -m finance_parser.investments.fetch_fx_rates
python -m finance_parser.investments.build_portfolio_positions
python -m finance_parser.investments.build_dividend_history
python -m finance_parser.investments.check_instrument_coverage
python -m finance_parser.investments.build_monthly_position_value
```

## ChangeLog

Workbook-writing steps append a `ChangeLog` entry so imports and automatic updates can be audited later.

## Manual split rows

Manual split rows are Excel-side manual categorisation rows.

Parent transaction:

```text
Include = NO
TransactionType = Bank
```

Child rows:

```text
TransactionType = Split
Include = YES
UnifiedID = parent UnifiedID + -S01/-S02/...
RawID = same as parent RawID
```

Split rows are manually categorised and should be preserved by the workflow.

If needed, split-specific IDs can be introduced like:

```text
original-id-S01
original-id-S02
```
