# Transaction categoriser

Budgeting processing order:

```text
parse transactions
normalise transactions
categorise transactions
```

Run:

```bash
python -m finance_parser.budgeting.transaction_categoriser --dry-run
python -m finance_parser.budgeting.transaction_categoriser
```

The real implementation lives in:

```text
finance_parser/budgeting/transaction_categoriser.py
```

The root-level `transaction_categoriser.py` is a compatibility command wrapper.

## Safety rule

The categoriser only processes `UnifiedTransactions` rows where `Comments` or `Comment` is blank. Any manual comment acts as a bypass switch.

The categoriser changes only fields in `UnifiedTransactions` and appends an audit row to `ChangeLog` on wet runs. It does not touch raw bank sheets or `ImportLog`.

## Rule sheets

Rules are read from:

```text
rules/budgeting/TransactionRules.xlsx
```

Required sheets:

```text
CategoryRules
OwnershipRules
```

`CategoryRules` can set:

```text
Supercategory
Category
Subcategory
Owner
Comments
Tag, if the UnifiedTransactions sheet has a Tag column
```

`OwnershipRules` can set or clear:

```text
Owner
```

## Overwrite modes

```text
BLANK_ONLY  set only blank target cells
FORCE       overwrite existing target values
RULE_ONLY   overwrite while Comments is blank; used for rule-managed refreshes
```


## Eligibility gate

Categorisation is applied only to rows where the `Comments` column is blank.
If the workbook has `Comment` instead of `Comments`, that column is used.

This is the manual bypass mechanism:

```text
Comments blank     -> categoriser may apply matching rules
Comments not blank -> categoriser skips the row completely
```

`OverwriteMode` only controls what happens **after** a row is eligible and a
rule matches.

Recommended interpretation:

```text
BLANK_ONLY = set only blank target cells
FORCE      = overwrite target cells on eligible rows
RULE_ONLY  = rule-managed refresh semantics for eligible rows; it is not
             a different Comments trigger
```

Because all categoriser processing is gated by blank `Comments`, `RULE_ONLY`
should not be read as "categorise only when Comments is non-blank".
Non-blank `Comments` always means skip.
