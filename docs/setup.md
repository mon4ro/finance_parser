# Setup

## 1. Install dependencies

Runtime dependencies:

```bash
python -m pip install -r requirements.txt
```

Test dependencies:

```bash
python -m pip install -r requirements-dev.txt
```

## 2. Create local configuration

Copy the example settings file:

```bash
cp config/settings.example.yaml config/settings.yaml
```

Edit:

```text
config/settings.yaml
```

for your own account labels and policy defaults.

Examples of user-specific settings:

```yaml
budgeting:
  source_account_inference:
    OP:
      filename_prefixes:
        FAMILY: "FAMILY"
    NORDEA:
      fixed_source_account: "PERSON_B"
    SPANKKI:
      fixed_source_account: "GROCERY_ACCOUNT"
      default_include: "NO"
```

## 3. Create local rule files

Copy the templates:

```bash
cp rules/budgeting/TransactionRules.template.xlsx rules/budgeting/TransactionRules.xlsx
cp rules/investments/InstrumentMaster.template.xlsx rules/investments/InstrumentMaster.xlsx
```

Then edit the copied files.

## 4. Add input files

Budgeting exports:

```text
input/budgeting/
```

Investment exports:

```text
input/investments/
```

## 5. Run the pipelines

Budgeting:

```bash
python run_budgeting_pipeline.py
```

Investments:

```bash
python run_investment_pipeline.py
```

Both support `--dry-run` (runs against temporary copied workbooks; the real ones are
untouched). Each individual stage can also be run directly, e.g.
`python -m finance_parser.budgeting.transaction_parser` or
`python -m finance_parser.investments.investment_parser` - see the main `README.md`
for the full list and other flags.

## 6. Run tests

```bash
python -m pytest
```

## Git hygiene

These local files should normally be ignored by Git:

```text
config/settings.yaml
rules/budgeting/TransactionRules.xlsx
rules/investments/InstrumentMaster.xlsx
input/
output/
```

If they were already committed earlier, remove them from Git tracking without deleting your local copies:

```bash
git rm --cached config/settings.yaml
git rm --cached rules/budgeting/TransactionRules.xlsx
git rm --cached rules/investments/InstrumentMaster.xlsx
```
