# BudgetTools Excel macros

This folder contains VBA tooling for manual editing of `ParsedTransactions.xlsx`.

The macros should **not** live inside `ParsedTransactions.xlsx`, because that file is generated/rebuilt by Python and remains `.xlsx`.

Recommended setup:

```text
WorkbookA = output/budgeting/ParsedTransactions.xlsx
WorkbookB = BudgetTools.xlsm
```

`BudgetTools.xlsm` stores the macros and edits the currently active `ParsedTransactions.xlsx`.

## Create BudgetTools.xlsm

1. Open Excel.
2. Create a blank workbook.
3. Save it as:

   ```text
   BudgetTools.xlsm
   ```

4. Press `Alt + F11` to open the VBA editor.
5. In the VBA editor, choose:

   ```text
   File > Import File...
   ```

6. Import:

   ```text
   tools/excel/BudgetTools.bas
   ```

7. Save `BudgetTools.xlsm`.

You can store the `.xlsm` locally, or in this folder if you want it versioned. The important point is that `ParsedTransactions.xlsx` remains macro-free.

## Use: Split selected transaction

1. Open `ParsedTransactions.xlsx`.
2. Open `BudgetTools.xlsm`.
3. Activate `ParsedTransactions.xlsx`.
4. Go to sheet:

   ```text
   UnifiedTransactions
   ```

5. Select any cell on the parent transaction row.
6. Run macro:

   ```text
   SplitSelectedTransaction
   ```

7. Enter the number of split rows.
8. The macro:
   - sets parent `Include = NO`
   - inserts split rows below the parent
   - copies source/date/receiver fields from the parent
   - sets child `TransactionType = Split`
   - sets child `Include = YES`
   - creates unique child `UnifiedID` values like `parent-S01`
   - keeps child `RawID` the same as the parent
   - clears child `Amount`, `Supercategory`, `Category`, `Subcategory`, and `Tag`
   - sets child `Comments = Split from <parent UnifiedID>`

Then fill the split rows manually.

## Use: Validate split sum

Select either the parent row or a child split row and run:

```text
ValidateSelectedSplitGroup
```

The macro compares:

```text
sum(child Amounts) == parent Amount
```

and reports whether the split is balanced.

## Data model

Parent row:

```text
TransactionType = Bank
Include = NO
UnifiedID = original unique row id
RawID = original bank/raw id
Comments = Split into N rows
```

Child rows:

```text
TransactionType = Split
Include = YES
UnifiedID = parent UnifiedID + -S01/-S02/...
RawID = same as parent RawID
Comments = Split from parent UnifiedID
```

`UnifiedID` stays unique. `RawID` stays shared for easy grouping.
