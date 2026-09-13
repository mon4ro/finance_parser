# Project structure harmonisation plan

This document defines the target structure for the finance parser project.

The current project still works, but the structure grew organically:

```text
common.py                       # root-level shared code
parsers/                        # budgeting parsers
investments/investment_common.py
investments/parsers/            # investment parsers
utilities/                      # some maintenance scripts
root-level utility scripts       # older duplicate utility locations
src/                            # older/alternate architecture
```

The target is to move toward one importable package and thin user-facing wrappers.

## Current state (as of 2026-09-12)

Most of this plan was carried out, but the end state diverges from the target below in a few ways worth knowing before touching folder structure.

**Matches the plan:**
- All live logic sits under `finance_parser/`: `settings.py`, `common.py`, `budgeting/` (`transaction_parser.py`, `transaction_normaliser.py`, `transaction_categoriser.py`, `run_budgeting_pipeline.py`, `parsers/{op,nordea,norwegian,spankki}.py`), `investments/` (`investment_parser.py`, `investment_common.py`, `parsers/{nordnet,seligson,evli,op_investment}.py`), and `utilities/` (7 maintenance scripts).
- Root-level `common.py`, `settings.py`, `transaction_parser.py`, `transaction_normaliser.py`, and `transaction_categoriser.py` are gone, as planned — enforced by `tests/test_no_root_compatibility_wrappers.py` and `tests/test_imports_and_paths.py`.
- The `src/` folder (the "older/alternate architecture" from Principle 3) no longer exists.

**Diverges from the plan:**
- No `scripts/` folder was ever created. Principle 4's "user-facing wrappers may live in `scripts/utilities/`" didn't happen — utilities are run directly as package modules (`python -m finance_parser.utilities.<name>`), with no thin-wrapper layer.
- `run_budgeting_pipeline.py`, `parsers/`, `investments/`, and `investment_parser.py` remain at the repo root as thin compatibility wrappers rather than moving under a `scripts/` directory. This was a later, deliberate decision (see `docs/patch_notes/2026-06-08_2024_remove_legacy_budgeting_root_wrappers.md`) — but that cleanup only removed the three root-level *budgeting stage* scripts; `parsers/`, `investments/`, and `investment_parser.py` were intentionally kept.
- **Those root shims are not obsolete.** `tests/test_budgeting_parser_fixtures.py`, `tests/test_bank_raw_hashes.py`, and `tests/test_investment_parser_fixtures.py` import directly from the root `parsers`/`investments` packages. They're part of the tested public surface, not safe-to-delete leftovers.
- `finance_parser/budgeting/transaction_categoriser.py` exists and is a core pipeline stage (parser → normaliser → categoriser) but isn't mentioned anywhere in the target structure or the migration commits below — it was added after this plan was written.
- This working directory is not currently a git repository, so the "Proposed migration commits" section below is historical narrative, not an actionable commit sequence today.
- `.DS_Store` files are still present throughout the tree, so "Harmonisation Commit 2: remove obsolete local artifacts" was never fully completed (or didn't stick).

For the current architectural rules — the raw/normalised/classified layer separation, `Include`/`Owner`/`ReviewStatus` semantics, the no-`Cars`-category rule, and the `ParsedTransactions.xlsx` regeneration risk — see `CLAUDE.md` at the repo root. That document reflects current behavior; this document is retained as historical planning context.

## Target structure

```text
finance_parser/
├─ README.md
├─ requirements.txt
├─ requirements-dev.txt
├─ pytest.ini
│
├─ config/
│  ├─ settings.example.yaml
│  └─ settings.yaml                      # local, gitignored
│
├─ input/                                # local, gitignored
│  ├─ budgeting/
│  └─ investments/
│
├─ output/                               # local, gitignored
│  ├─ budgeting/
│  └─ investments/
│
├─ rules/
│  ├─ budgeting/
│  │  ├─ TransactionRules.template.xlsx
│  │  └─ TransactionRules.xlsx            # local, gitignored
│  └─ investments/
│     ├─ InstrumentMaster.template.xlsx
│     └─ InstrumentMaster.xlsx            # local, gitignored
│
├─ finance_parser/
│  ├─ __init__.py
│  ├─ settings.py
│  ├─ common.py
│  │
│  ├─ budgeting/
│  │  ├─ __init__.py
│  │  ├─ transaction_parser.py
│  │  ├─ transaction_normaliser.py
│  │  └─ parsers/
│  │     ├─ __init__.py
│  │     ├─ op.py
│  │     ├─ norwegian.py
│  │     ├─ nordea.py
│  │     └─ spankki.py
│  │
│  ├─ investments/
│  │  ├─ __init__.py
│  │  ├─ investment_parser.py
│  │  ├─ investment_common.py
│  │  └─ parsers/
│  │     ├─ __init__.py
│  │     ├─ nordnet.py
│  │     ├─ seligson.py
│  │     ├─ evli.py
│  │     └─ op_investment.py
│  │
│  └─ utilities/
│     ├─ __init__.py
│     ├─ cleanup_budgeting_source_accounts.py
│     ├─ refresh_budgeting_bank_raw_ids.py
│     └─ review_cleanup_unified_duplicates.py
│
├─ scripts/
│  ├─ transaction_parser.py               # thin wrapper
│  ├─ transaction_normaliser.py           # thin wrapper
│  ├─ investment_parser.py                # thin wrapper
│  └─ utilities/
│     ├─ cleanup_budgeting_source_accounts.py
│     ├─ refresh_budgeting_bank_raw_ids.py
│     └─ review_cleanup_unified_duplicates.py
│
├─ tests/
│  ├─ fixtures/
│  └─ test_*.py
│
└─ docs/
   ├─ setup.md
   ├─ workflow.md
   ├─ sharing.md
   ├─ project_structure.md
   └─ patch_notes/
```

## Principles

### 1. Generic code under package

Actual parser logic should eventually live under:

```text
finance_parser/
```

Root-level scripts should become thin wrappers only.

### 2. Budgeting and investment structure should be parallel

Budgeting:

```text
finance_parser/budgeting/
finance_parser/budgeting/parsers/
```

Investments:

```text
finance_parser/investments/
finance_parser/investments/parsers/
```

### 3. No competing architecture folders

The current `src/` folder appears to be an older/alternate architecture. It should not remain alongside the final package structure unless its contents are intentionally migrated.

### 4. Utilities should have one canonical location

Maintenance logic should eventually live in:

```text
finance_parser/utilities/
```

User-facing wrappers may live in:

```text
scripts/utilities/
```

Root-level utility duplicates should be removed once wrappers exist.

### 5. Local/private data stays outside commits

These should normally be ignored:

```text
config/settings.yaml
input/
output/
rules/budgeting/TransactionRules.xlsx
rules/investments/InstrumentMaster.xlsx
```

## Proposed migration commits

### Harmonisation Commit 1: hygiene and plan

- Add this structure plan.
- Add local artifact cleanup helper.
- Do not move parser imports yet.
- Do not delete personal rule/input/output files.

### Harmonisation Commit 2: remove obsolete local artifacts

- Remove `.DS_Store`, `__pycache__`, `.pytest_cache` from working tree.
- Confirm `.gitignore`.

### Harmonisation Commit 3: package skeleton

- Add `finance_parser/` package directories.
- Add thin wrappers but do not move heavy logic yet.

### Harmonisation Commit 4: move budgeting code

- Move `common.py`, `transaction_parser.py`, `transaction_normaliser.py`, and `parsers/`.
- Update imports and tests.

### Harmonisation Commit 5: move investment code

- Move `investment_parser.py`, `investments/investment_common.py`, and `investments/parsers/`.
- Update imports and tests.

### Harmonisation Commit 6: move utilities

- Move maintenance scripts into package utilities.
- Add script wrappers.
- Remove duplicate root-level utilities.

### Harmonisation Commit 7: cleanup and docs pass

- Remove obsolete `src/` if fully replaced.
- Update README/setup/workflow docs.
- Run full test suite.
