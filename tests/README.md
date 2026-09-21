# Tests

Run from the project root:

```bash
python -m pytest
```

These tests are intentionally small and fast. They protect the most fragile
behaviours: import paths, hash stability, receiver blank defaults, source
account canonicalisation, and normaliser dynamic rules.

## Layout

```text
tests/
├─ README.md
└─ ut/                 # unit tests - all of them today (see below)
   ├─ fixtures/
   └─ test_*.py
```

Every test here is unit-test-shaped: fast, self-contained, no real network
calls or real workbook I/O, synthetic fixtures only. If a genuinely
system/component-level test is ever added (e.g. a real end-to-end pipeline
run against real-shaped external systems), give it its own `tests/sct/`
sibling rather than mixing it into `ut/` - don't create that folder ahead
of actually having such a test.
