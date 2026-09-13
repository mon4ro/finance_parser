# Tests

Run from the project root:

```bash
python -m pytest
```

These tests are intentionally small and fast. They protect the most fragile
behaviours: import paths, hash stability, receiver blank defaults, source
account canonicalisation, and normaliser dynamic rules.
