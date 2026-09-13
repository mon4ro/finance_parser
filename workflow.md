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
2. transaction normaliser
3. transaction categoriser
```

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
python -m finance_parser.budgeting.transaction_categoriser
```

Individual dry-run commands:

```bash
python -m finance_parser.budgeting.transaction_parser --dry-run
python -m finance_parser.budgeting.transaction_normaliser --dry-run
python -m finance_parser.budgeting.transaction_categoriser --dry-run
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
