# finance_parser

Personal household finance pipeline. Parses bank/broker exports into two
Excel workbooks (`ParsedTransactions.xlsx`, `ParsedInvestments.xlsx`) that
feed a separate, hand-built Master Budget workbook via Excel formulas.

## Architecture: three layers, never collapsed

The budgeting pipeline (and the investment pipeline, in miniature) is three
strictly separated layers. Each layer is a different script/module, and each
runs on the previous layer's *output*, never inline:

1. **Raw** — `finance_parser/budgeting/parsers/{op,nordea,norwegian,spankki,cash,investment_dividends}.py`
   Bank-specific column mapping only: read the bank's native export, emit one
   row per `RAW_COLUMNS` (RawID, dates, amount, RawReceiver, reference,
   etc.). **A parser module must never set `Include`, `Owner`,
   `Supercategory`, `Category`, or `Subcategory`.** If you're adding or
   editing a parser and find yourself wanting to special-case a receiver name
   or an account's default owner in there, stop — that belongs in
   `config/settings.yaml` (account-level facts) or `TransactionRules.xlsx`
   (content-based rules), not in parser code. `cash.py` and
   `investment_dividends.py` aren't bank exports (a manually maintained
   ledger and a cross-pipeline read of the investment side's
   `DividendHistory.xlsx`, respectively) but follow the exact same contract
   and flow through the exact same three layers - **with one narrow,
   deliberate exception each to the "never set `Owner`" rule**: `cash.py`
   sets `Owner` directly from `CashEntries.xlsx`'s own `Owner` column, and
   `investment_dividends.py` sets it directly from `DividendHistory.xlsx`'s
   `PortfolioOwner`. Both are cases where the source data itself states
   whose transaction this is as a real recorded fact, not an inference from
   account naming - unlike every other parser, where `Owner` truly is
   inferred (via `SourceAccount` and `SOURCE_TO_DEFAULT_OWNER`) rather than
   a fact present in the source file. `investment_dividends.py`'s
   `SourceAccount` is the real broker the dividend was paid into (`NORDNET`
   or `EVLI`) - it used to carry `PortfolioOwner` instead, as a workaround
   to borrow the `SourceAccount`->`Owner` inference machinery without
   setting `Owner` directly, but that made it look like a real bank account
   matching the budgeting side's own personal account names (each real
   household member's own account), when this money never touched a bank
   account at all (fixed 2026-09-23).

2. **Unified (normalised)** — `raw_to_unified_rows()` in
   `finance_parser/common.py`, invoked by `transaction_parser.py`.
   Converts a raw row into a `UnifiedTransactions` row. This is the one
   place where `Include` and `Owner` get non-blank defaults *before* any
   rules run — but only from **account-level settings**
   (`default_include_for_source_bank`, `SOURCE_TO_DEFAULT_OWNER`), never
   from receiver text or transaction content. `Supercategory`/`Category`/
   `Subcategory` stay blank here, always — they're the categoriser's job.
   `transaction_normaliser.py` then runs against `TransactionRules.xlsx`
   (the `TransactionRules` sheet) to fill `NormalizedReceiver` and,
   occasionally, `Include`.

3. **Classified** — `transaction_categoriser.py`, driven by the
   `CategoryRules` and `OwnershipRules` sheets in `TransactionRules.xlsx`.
   This is the only place `Supercategory`/`Category`/`Subcategory` should
   be set from *content-matching* rules today. A row is only touched if its
   `Review/Notes` cell is blank — see below. `enrich_dividend_income.py` is
   a narrow, deliberate exception: it also sets these three fields, but via
   structured cross-reference against the investment pipeline's own
   `DividendHistory.xlsx` (matched by date + exact amount) rather than
   receiver-text rules — for the one case (OP-held instrument dividends)
   where the real receiver text is useless for categorisation (it's just
   the account holder's own name, an OP export quirk). Still respects
   `Review/Notes` and never overwrites an already-filled cell.

Pipeline order (`run_budgeting_pipeline.py`): **parser → investment
dividend parser → normaliser → dividend income enrichment → categoriser**
(the two dividend stages auto-skip with a message if the investment
pipeline hasn't been run yet, i.e. `DividendHistory.xlsx` doesn't exist).
Investments mirror the core three layers with `apply_instrument_master()`
in `investment_common.py` playing the categoriser's role, driven by
`InstrumentMaster.xlsx`.

**Why the separation matters:** `Owner` and `Include` are personal/
deployment-specific facts (which account belongs to which person; whether a
given account is a real spending account or a buffer/pass-through account)
that would leak household-specific assumptions into the shared,
content-based rule engine if they lived in `TransactionRules.xlsx`. Keeping
them in `config/settings.yaml` (gitignored) means the rule sheets and parser
code stay portable across a different household's account setup — only the
settings file and rule sheet contents are personal.

**Known latent inconsistency, not yet a bug:** the `TransactionRules` sheet
schema and `transaction_normaliser.py` code both support `SetOwner`,
`SetSupercategory`, `SetCategory`, `SetSubcategory` columns — but as of this
writing, zero rows in the live `TransactionRules.xlsx` sheet populate them
(306 rows use `SetNormalizedReceiver`, 13 use `SetInclude`). All real
categorisation happens via `CategoryRules`/`OwnershipRules`, the
categoriser, and (for dividend income specifically)
`enrich_dividend_income.py`'s cross-reference. Don't start populating those
normaliser-sheet columns without deciding whether that's intentionally
reclaiming the categoriser's job — two rule sheets quietly doing the same
thing is a maintenance trap.

## Include / Owner / ReviewStatus semantics

- **`Include`** — whether a transaction counts toward budget totals.
  Defaulted per source bank/account in `config/settings.yaml`
  (`budgeting.default_include`, `budgeting.source_account_inference.<BANK>.
  default_include`). Example: the S-Pankki buffer account defaults to
  `NO` because its balance is reconciled separately (see below), not
  because its individual transactions are irrelevant.

- **`Owner`** — which household member a transaction belongs to. Defaulted
  from `SOURCE_TO_DEFAULT_OWNER` in `common.py` (keyed by substrings in
  `SourceAccount`, e.g. "joint"/"shared" → `Shared`) and can be overridden
  by `OwnershipRules` in the categoriser. Personal-mapping concern — belongs
  in settings/rules, never hardcoded assumptions in parser or common code
  beyond the existing generic keyword map.

- **`Review/Notes`** — the categoriser's protection gate. A row is only
  eligible for automatic categorisation while this cell is blank; once a
  human writes anything in it, the categoriser leaves the row alone (unless
  explicitly overridden with `--recategorise --all`, which is a destructive,
  double-confirmed operation). This is how manual splits/overrides (e.g. the
  S-Pankki grocery reimbursement split, done via an Excel macro) survive
  repeated pipeline runs.

- **`ReviewStatus`** — defined in the `UnifiedTransactions` schema and
  initialised to `""` when a row is created, but **no script currently
  reads or writes it afterward**. It is effectively a manual-only column
  today (distinct from `Review/Notes`, which is load-bearing). Don't assume
  it drives any pipeline behavior unless/until code is added to use it.

## Business rule: no "Cars" category

The category taxonomy deliberately has no vehicle-ownership category.
Transportation is split by mode/purpose instead: `Personal transportation`
and `Public transportation and taxis` (see `CategoryRules` in
`TransactionRules.xlsx`). Don't introduce a `Cars` category when adding
rules — route car-related spend into the existing transportation categories
(or another existing category, e.g. `Housing` for a car loan) instead of
creating a new one.

## `ParsedTransactions.xlsx` regeneration risk

`transaction_normaliser.py` and `transaction_categoriser.py` both default to
**fresh-rebuild write mode**: they read the existing workbook as
values-only, then write a brand-new `.xlsx` package from scratch via
`replace_with_fresh_workbook()` (Excel Table objects are deliberately
disabled — real-workbook testing showed Excel repair warnings otherwise).
This is not an in-place edit; it's a new binary file replacing the old one.

**Consequence:** any external workbook (e.g. a hand-built Master Budget
workbook) that references `ParsedTransactions.xlsx` via `XLOOKUP`/formulas
pointing at the file, a sheet's Table object, or anything more structural
than plain cell ranges, can break on the next pipeline run — the file is a
new object, not a patched one. Before changing anything about workbook
writing (write mode, sheet layout, column order), check whether it could
break formula references in a workbook outside this repo. When in doubt,
treat any external `.xlsx` that reads from this repo's outputs as read-only
and out of scope for changes here — this repo only owns the parser →
normaliser → categoriser chain and its own outputs
(`output/budgeting/ParsedTransactions.xlsx`,
`output/investments/ParsedInvestments.xlsx`).

A legacy `--write-mode openpyxl-mutating` exists for the categoriser but is
flagged in its own help text as legacy/prone to Excel repair warnings — not
a safe alternative for avoiding the regeneration risk.

## No personal or real financial data in committed files

This repo is public. Nothing that identifies a person, or that reveals a
real number from this household's actual finances, may appear in any file
that gets `git add`ed — code, tests, fixtures, docs, comments, commit
messages. Two distinct categories, both forbidden:

- **Personal info:** real names (household members, anyone else), emails,
  phone numbers, addresses, real account numbers. Use generic placeholders
  instead: `PERSON_A`/`PERSON_B`/`SHARED` for people (matching
  `SOURCE_TO_DEFAULT_OWNER` in `common.py`), `HOUSEHOLD`/`PERSONAL`/`CHILD`
  for account labels (matching `settings.example.yaml`).
- **Real transaction data:** an actual amount, date, quantity, price, or
  rate copied from a real run or a real input file — even with no name
  attached, a real number ties a real event to this household's real
  finances. Every number in code comments, docstrings, and test fixtures
  must be invented, never copied from real output. This class of leak is
  more central to a *finance* tool than the personal-info case above and
  needs the same level of care: it's easy to paste an actual observed value
  into a comment mid-explanation ("confirmed against the real row above:
  ...") without registering that it's real household data. When writing a
  comment or docstring that explains real data shape/behavior, use the same
  synthetic numbers already in the corresponding test fixture (or invent new
  ones) — never the actual values observed while testing against real data,
  even to illustrate a genuinely real relationship (e.g. "eurAmount is
  fee-inclusive, confirmed against real data" is fine to say; pasting the
  real row that proved it is not).

Real names/data belong only in gitignored files (`config/settings.yaml`,
`rules/*.xlsx`, `output/*.xlsx`, `docs/patch_notes/`) or in the terminal/chat
transcript with the user — never in a file about to be `git add`ed.

**Watch for indirect leaks, not just literal typing:**
- Copying a real value into a **test fixture** while writing a test for
  settings-driven behavior (happened once: a real name ended up in
  `tests/ut/test_investment_portfolio_owner.py` as a stand-in owner value).
- Pasting a **real observed value into an explanatory code comment** while
  documenting how a parser handles real data (happened once, in a parser's
  header comment — don't do this again).
- Adding an example value to `config/settings.example.yaml` under a key
  (e.g. a broker name) that also exists as a real key the user's own
  gitignored `config/settings.yaml` doesn't happen to override — the
  example value silently deep-merges into their real settings and can then
  leak into real output data. Prefer an empty placeholder (`{}`) over a
  real-looking example value for any settings key where this is possible.

Before committing anything new, re-read every comment/docstring touched in
the diff and ask "is this number invented, or did I copy it from a real
run/file?" — not just grep for known real names. Do this on every commit,
not just at initial repo setup.
