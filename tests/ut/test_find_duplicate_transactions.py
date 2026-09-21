import pandas as pd

from finance_parser.budgeting.find_duplicate_transactions import (
    find_duplicate_groups,
    _references_look_like_the_same_number,
)


def _rows(records):
    return pd.DataFrame(records)


def test_no_duplicates_returns_empty():
    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -10, "RawReceiver": "Shop A", "Reference": "111"},
        {"RawID": "A-2", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-02", "Amount": -20, "RawReceiver": "Shop B", "Reference": "222"},
    ])

    result = find_duplicate_groups(df)

    assert len(result) == 0


def test_exact_matching_reference_is_high_confidence():
    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -50, "RawReceiver": "Loan Co", "Reference": "99887766554433221100"},
        {"RawID": "A-2", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -50, "RawReceiver": "Loan Co", "Reference": "99887766554433221100"},
    ])

    result = find_duplicate_groups(df)

    assert len(result) == 2
    assert set(result["Confidence"]) == {"HIGH"}


def test_float_precision_loss_reference_still_flagged_high_confidence():
    """
    Simulates a real bug found this session: a long reference read once as
    exact text and once lossily coerced through float64 (classic Excel/
    pandas dtype-inference quirk on a long numeric-looking cell) hashes
    differently and defeats RawID-based de-duplication, but the two
    references are still recognizably the "same" number within float64's
    own precision - close enough that this tool should still flag it.
    """
    exact = "99887766554433221100"
    lossy = str(int(float(exact)))
    assert exact != lossy  # sanity: precision was genuinely lost by the round-trip

    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -50, "RawReceiver": "Loan Co", "Reference": exact},
        {"RawID": "A-2", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -50, "RawReceiver": "Loan Co", "Reference": lossy},
    ])

    result = find_duplicate_groups(df)

    assert len(result) == 2
    assert set(result["Confidence"]) == {"HIGH"}


def test_different_reference_is_informational_only():
    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Transit Co", "Reference": "111"},
        {"RawID": "A-2", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Transit Co", "Reference": "222"},
    ])

    result = find_duplicate_groups(df)

    assert len(result) == 2
    assert set(result["Confidence"]) == {"INFORMATIONAL"}


def test_different_account_is_not_grouped_together():
    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Transit Co", "Reference": "111"},
        {"RawID": "B-1", "SourceAccount": "OTHER", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Transit Co", "Reference": "111"},
    ])

    result = find_duplicate_groups(df)

    assert len(result) == 0


def test_references_look_like_the_same_number_rejects_genuinely_different_values():
    assert _references_look_like_the_same_number("111", "222") is False
    assert _references_look_like_the_same_number("", "111") is False
    assert _references_look_like_the_same_number("abc", "111") is False
    assert _references_look_like_the_same_number("111", "111") is True


def test_blank_op_style_reference_never_counts_as_a_match():
    """
    Real bug found and fixed: OP stores an empty reference as the literal
    text "ref=" rather than a blank cell. Two unrelated transactions that
    both simply have no reference at all must never be treated as "the same
    reference" just because the placeholder text happens to be identical -
    this originally flooded results with unrelated real transactions.
    """
    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Coffee Shop", "Reference": "ref="},
        {"RawID": "A-2", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Coffee Shop", "Reference": "ref="},
    ])

    result = find_duplicate_groups(df)

    assert len(result) == 2
    assert set(result["Confidence"]) == {"INFORMATIONAL"}
    assert _references_look_like_the_same_number("ref=", "ref=") is False


def test_blank_spankki_style_reference_never_counts_as_a_match():
    """Same class of bug as the OP "ref=" case, but S-Pankki uses a plain
    "-" as its own blank-reference placeholder."""
    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Grocery Store", "Reference": "-"},
        {"RawID": "A-2", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Grocery Store", "Reference": "-"},
    ])

    result = find_duplicate_groups(df)

    assert set(result["Confidence"]) == {"INFORMATIONAL"}
    assert _references_look_like_the_same_number("-", "-") is False


def test_numerically_close_but_genuinely_different_long_references_are_not_a_match():
    """
    Real false positive found and fixed: two genuinely separate real
    transactions (confirmed distinct by their own separate ArchiveID) can
    still have long reference numbers that are numerically close (e.g. both
    encoding the same date, or issued in the same batch) without being a
    float64 round-trip of each other at all. A naive relative-tolerance
    check on numbers this large is fooled by this; comparing significant
    digits at a calibrated precision is not.
    """
    assert _references_look_like_the_same_number("11223344556677889900", "11223344556699887766") is False
    assert _references_look_like_the_same_number("20260101123456789012", "20260101123499998877") is False


def test_short_reference_below_magnitude_threshold_is_not_high_confidence():
    """A short numeric-looking reference isn't distinctive enough to trust
    as proof of a real duplicate on its own."""
    df = _rows([
        {"RawID": "A-1", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Transit Co", "Reference": "42"},
        {"RawID": "A-2", "SourceAccount": "HOUSEHOLD", "BookingDate": "2026-01-01", "Amount": -3, "RawReceiver": "Transit Co", "Reference": "43"},
    ])

    result = find_duplicate_groups(df)

    assert set(result["Confidence"]) == {"INFORMATIONAL"}
