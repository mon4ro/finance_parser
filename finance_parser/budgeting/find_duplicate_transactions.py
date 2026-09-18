from __future__ import annotations

import argparse
from pathlib import Path

import pandas as pd


PROJECT_ROOT = Path(__file__).resolve().parents[2]
DEFAULT_BUDGETING_WORKBOOK = PROJECT_ROOT / "output" / "budgeting" / "ParsedTransactions.xlsx"

# A group of real RawTransactions rows sharing all four of these is worth a
# second look - a real, distinct transaction essentially never repeats its
# exact account/date/amount/receiver combination by pure coincidence more
# than very occasionally (e.g. two identical-fare transit tickets bought the
# same day), so groups this small are cheap to review by hand.
CONTENT_KEY_COLUMNS = ["SourceAccount", "BookingDate", "Amount", "RawReceiver"]


def load_raw_transactions(budgeting_workbook: Path) -> pd.DataFrame:
    df = pd.read_excel(budgeting_workbook, sheet_name="RawTransactions", dtype=object, engine="openpyxl")
    df.columns = [str(c).strip() for c in df.columns]
    return df


# Some banks store an empty reference as a literal placeholder rather than
# a blank cell (OP: "ref="; S-Pankki: "-") - without stripping/rejecting
# these, two unrelated blank references would match each other and falsely
# count as "the same reference number", flooding results with unrelated
# real transactions that simply both happen to have no reference at all.
_BLANK_REFERENCE_PREFIXES = ("ref=",)
_BLANK_REFERENCE_VALUES = {"-"}

# A "reference" shorter than this (after stripping known blank-markers)
# isn't distinctive enough to trust as proof of a real duplicate - a short
# code could coincidentally repeat across genuinely separate transactions.
_MIN_MEANINGFUL_REFERENCE_MAGNITUDE = 1e7

# How many leading significant digits must agree for two long numeric
# references to count as "the same number, just imprecisely formatted".
# Empirically calibrated against real data: a genuine float64/text-display
# round-trip of the same 20-digit reference still agrees to at least this
# many significant digits, while two genuinely different real reference
# numbers that happen to share a long common prefix (e.g. both encoding the
# same date) diverge well before it - see test_find_duplicate_transactions.py.
_SIGNIFICANT_DIGITS_FOR_MATCH = 14


def _core_reference_text(value: str) -> str:
    value = value.strip()
    for prefix in _BLANK_REFERENCE_PREFIXES:
        if value.startswith(prefix):
            value = value[len(prefix):]
    value = value.strip()
    if value in _BLANK_REFERENCE_VALUES:
        return ""
    return value


def _references_look_like_the_same_number(a: str, b: str) -> bool:
    """
    True if two reference strings are either identical (after stripping a
    known blank-marker), or are the same underlying long number just read/
    formatted differently - one read as exact text, the other lossily
    coerced through float64 (and sometimes re-displayed with even fewer
    significant digits, e.g. Excel's own scientific-notation text) at some
    point upstream (see normalise_reference_text() in common.py for why
    this happens and why it can't always be prevented at read time).
    Deliberately conservative: a blank reference on either side is never a
    match, and the two values must agree to _SIGNIFICANT_DIGITS_FOR_MATCH
    leading digits - calibrated against real data so two genuinely
    different long references are never mistaken for a match just because
    they share a prefix.
    """
    core_a, core_b = _core_reference_text(a), _core_reference_text(b)
    if not core_a or not core_b:
        return False
    if core_a == core_b:
        return True
    try:
        value_a, value_b = float(core_a), float(core_b)
    except ValueError:
        return False
    if abs(value_a) < _MIN_MEANINGFUL_REFERENCE_MAGNITUDE or abs(value_b) < _MIN_MEANINGFUL_REFERENCE_MAGNITUDE:
        return False
    fmt = f"{{:.{_SIGNIFICANT_DIGITS_FOR_MATCH - 1}e}}"
    return fmt.format(value_a) == fmt.format(value_b)


def find_duplicate_groups(raw_transactions: pd.DataFrame) -> pd.DataFrame:
    """
    Groups RawTransactions rows sharing CONTENT_KEY_COLUMNS, then tags each
    group's confidence:

    - HIGH: the group's references also match (exactly, or as the same
      underlying number, see _references_look_like_the_same_number) - very
      likely the exact same real transaction imported twice under two
      different RawIDs, defeating the usual RawID-based de-duplication.
    - INFORMATIONAL: the references genuinely differ - very likely separate
      real transactions that happen to share account/date/amount/receiver
      (e.g. two identical-fare tickets bought the same day). Reported for
      visibility, not flagged as a probable bug.

    Read-only - never modifies raw_transactions or any file.
    """
    working = raw_transactions.copy()
    for col in CONTENT_KEY_COLUMNS:
        if col not in working.columns:
            working[col] = ""
    working["_Reference"] = working["Reference"].astype(str) if "Reference" in working.columns else ""

    dupe_mask = working.duplicated(subset=CONTENT_KEY_COLUMNS, keep=False)
    groups = working[dupe_mask].copy()
    if len(groups) == 0:
        groups["Confidence"] = pd.Series(dtype=object)
        return groups

    confidence_by_index: dict = {}
    for _, group in groups.groupby(CONTENT_KEY_COLUMNS):
        refs = list(group["_Reference"])
        confidence = "INFORMATIONAL"
        for i in range(len(refs)):
            for j in range(i + 1, len(refs)):
                if _references_look_like_the_same_number(refs[i], refs[j]):
                    confidence = "HIGH"
                    break
            if confidence == "HIGH":
                break
        for idx in group.index:
            confidence_by_index[idx] = confidence

    groups["Confidence"] = groups.index.map(confidence_by_index)
    return groups.sort_values(
        ["Confidence"] + CONTENT_KEY_COLUMNS,
        ascending=[False] + [True] * len(CONTENT_KEY_COLUMNS),
    )


def build_arg_parser() -> argparse.ArgumentParser:
    parser = argparse.ArgumentParser(
        description="Read-only report of possibly-duplicate RawTransactions rows (same "
                    "SourceAccount+BookingDate+Amount+RawReceiver). Never modifies any file - "
                    "review the HIGH-confidence group manually before deciding what to do about it."
    )
    parser.add_argument("--budgeting-workbook", default=str(DEFAULT_BUDGETING_WORKBOOK))
    return parser


def main() -> None:
    args = build_arg_parser().parse_args()
    workbook = Path(args.budgeting_workbook).expanduser().resolve()

    raw = load_raw_transactions(workbook)
    groups = find_duplicate_groups(raw)

    if len(groups) == 0:
        print("No possibly-duplicate transactions found.")
        return

    high = groups[groups["Confidence"] == "HIGH"]
    info = groups[groups["Confidence"] == "INFORMATIONAL"]
    report_cols = [c for c in ["RawID", "SourceAccount", "BookingDate", "Amount", "RawReceiver", "Reference", "ImportedAt"] if c in groups.columns]

    print(f"Found {len(groups)} rows in possibly-duplicate groups.")
    print()
    if len(high):
        print(f"HIGH confidence ({len(high)} rows) - same account/date/amount/receiver AND the same")
        print("underlying reference number, just formatted differently. Very likely a real double-import:")
        print(high[report_cols].to_string(index=False))
        print()
    if len(info):
        print(f"INFORMATIONAL ({len(info)} rows) - same account/date/amount/receiver but a DIFFERENT")
        print("reference - most likely genuinely separate real transactions. Reported for visibility only:")
        print(info[report_cols].to_string(index=False))


if __name__ == "__main__":
    main()
