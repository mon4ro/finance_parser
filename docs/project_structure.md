# Project structure

All live logic lives under `finance_parser/`: `settings.py`, `common.py`,
`budgeting/` (parser/normaliser/categoriser stages + `parsers/{op,nordea,
norwegian,spankki,cash,investment_dividends}.py`), `investments/`
(parser + pipeline stages + `parsers/{nordnet,seligson,evli,op_investment,
coinmotion,nordea}.py`), and `utilities/` (maintenance scripts, fresh-rebuild
workbook writer, etc.).

For the architectural rules - the raw/normalised/classified layer
separation, `Include`/`Owner`/`ReviewStatus` semantics, the no-`Cars`
category rule, and the `ParsedTransactions.xlsx` regeneration risk - see
`CLAUDE.md` at the repo root. That document is the current source of truth;
this one just describes where things live on disk.

## Root layout

```text
finance_parser/
├─ README.md                    # getting-started, per-pipeline workflow commands
├─ CLAUDE.md                    # architectural rules (source of truth)
├─ TASKLIST.md                  # local planning only, gitignored, not committed
├─ requirements.txt / requirements-dev.txt
├─ pytest.ini
│
├─ config/
│  ├─ settings.example.yaml
│  └─ settings.yaml              # local, gitignored
│
├─ input/                        # local, gitignored
│  ├─ budgeting/
│  └─ investments/
│
├─ output/                       # local, gitignored
│  ├─ budgeting/
│  └─ investments/
│
├─ rules/
│  ├─ budgeting/
│  │  ├─ TransactionRules.template.xlsx
│  │  └─ TransactionRules.xlsx    # local, gitignored
│  └─ investments/
│     ├─ InstrumentMaster.template.xlsx
│     └─ InstrumentMaster.xlsx    # local, gitignored
│
├─ finance_parser/                # the one importable package - all real logic
│  ├─ settings.py
│  ├─ common.py
│  ├─ setup_wizard.py
│  ├─ budgeting/
│  │  ├─ transaction_parser.py
│  │  ├─ transaction_normaliser.py
│  │  ├─ transaction_categoriser.py
│  │  ├─ run_budgeting_pipeline.py
│  │  ├─ seed_account_balance.py
│  │  ├─ find_duplicate_transactions.py
│  │  ├─ sync_category_taxonomy.py
│  │  └─ parsers/{op,nordea,norwegian,spankki,cash,investment_dividends}.py
│  ├─ investments/
│  │  ├─ investment_parser.py
│  │  ├─ investment_common.py
│  │  ├─ run_investment_pipeline.py
│  │  └─ parsers/{nordnet,seligson,evli,op_investment,coinmotion,nordea}.py
│  └─ utilities/                  # fresh_workbook_writer.py + maintenance scripts
│
├─ run_budgeting_pipeline.py      # thin root wrapper -> finance_parser.budgeting... (core, run often)
├─ run_investment_pipeline.py     # thin root wrapper -> finance_parser.investments... (core, run often)
├─ setup_wizard.py                # thin root wrapper (core - first-run onboarding)
│
├─ personal/                      # gitignored - real-data-bearing personal scripts
├─ tools/
│  ├─ excel/                      # BudgetTools.xlsm + decompiled source
│  ├─ seed_account_balance.py     # thin wrapper (occasional/maintenance - run once per account)
│  └─ find_duplicate_transactions.py  # thin wrapper (occasional/maintenance - spot-check tool)
│
├─ tests/
│  ├─ fixtures/
│  └─ test_*.py
│
└─ docs/
   ├─ setup.md
   ├─ workflow.md
   ├─ sharing.md
   ├─ utilities.md
   ├─ transaction_categoriser.md
   ├─ project_structure.md        # this file
   └─ patch_notes/                # gitignored - historical dev log, kept locally only
```

## Root wrapper scripts vs. tools/

`run_budgeting_pipeline.py`, `run_investment_pipeline.py`, and
`setup_wizard.py` are thin wrappers at the repo root (each just imports and
calls `main()` from the real module under `finance_parser/`) - kept there
deliberately for `python run_budgeting_pipeline.py`-style convenience.
These are the core, frequently-run commands (the two pipelines, plus
first-run onboarding).

`tools/seed_account_balance.py` and `tools/find_duplicate_transactions.py`
are the same kind of thin wrapper, but deliberately kept out of root -
they're occasional/maintenance tools (seed an account once, spot-check
duplicates now and then), not part of the everyday pipeline loop. Both
still just call `main()` from the real module under `finance_parser/
budgeting/` - moving them changed nothing about the underlying logic.

**Individual pipeline *stage* scripts are deliberately NOT allowed at
root** - `transaction_parser.py`, `transaction_normaliser.py`,
`transaction_categoriser.py`, `common.py`, and `settings.py` must not exist
at the repo root; run those via `python -m finance_parser.budgeting.<name>`
instead. Enforced by `tests/test_no_root_compatibility_wrappers.py`.

**2026-09-21 cleanup**: removed the old root-level `parsers/`, `investments/`,
and `investment_parser.py` re-export shims (thin `from finance_parser...
import *` files kept alive only because 3 tests imported from them
directly - not part of any real user-facing surface). Those tests now
import from `finance_parser.budgeting.parsers` / `finance_parser.
investments.parsers` directly. Also removed `RECOVERY_README.md` (an
obsolete pre-package, pre-git ChatGPT-patch recovery doc) and a stale
duplicate root `workflow.md` (superseded by the actively maintained
`docs/workflow.md`). Same day, moved `seed_account_balance.py` and
`find_duplicate_transactions.py` from root into `tools/` (see above).

## Local-only, gitignored (not clutter - by design)

- `config/settings.yaml`, `input/`, `output/`, `rules/budgeting/
  TransactionRules.xlsx`, `rules/investments/InstrumentMaster.xlsx` - real
  personal/financial data, never committed.
- `personal/` - real-data-bearing personal scripts/tools, deliberately kept
  outside the shared package.
- `TASKLIST.md` - local planning/backlog, not a commit-worthy artifact.
- `docs/patch_notes/` - historical per-change dev log, kept locally for
  reference, not published.
