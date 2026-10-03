import pandas as pd

from finance_parser.common import find_overlapping_import_windows


def _log(rows):
    defaults = {"Status": "Imported"}
    if not rows:
        return pd.DataFrame(columns=["SourceBank", "SourceFile", "Status"])
    return pd.DataFrame([{**defaults, **row} for row in rows])


_EMPTY = _log([])


def test_overlapping_date_ranges_within_one_batch_are_flagged():
    """
    Real motivating case (2026-10-01): two real Norwegian exports covering
    the same days, parsed in the same batch, described the identical
    transaction with different Text AND Type - existing dedup (keyed
    partly on receiver text) missed it entirely, silently double-counting
    the amount. This can't prevent that on its own (date+amount alone
    isn't a safe merge key - see the no-false-merge rationale elsewhere),
    but it can flag the overlap for a human to check.
    """
    new_log = _log([
        {"SourceBank": "NORWEGIAN", "SourceFile": "Tili_NORWEGIAN_20260720-20260820.xlsx"},
        {"SourceBank": "NORWEGIAN", "SourceFile": "Card_NORWEGIAN_Statement20260801-20260930.xlsx"},
    ])

    overlaps = find_overlapping_import_windows(_EMPTY, new_log)

    assert len(overlaps) == 1
    account, window_a, window_b = overlaps[0]
    assert account == "NORWEGIAN"
    names = {window_a[0], window_b[0]}
    assert names == {"Tili_NORWEGIAN_20260720-20260820.xlsx", "Card_NORWEGIAN_Statement20260801-20260930.xlsx"}


def test_new_file_overlapping_an_existing_imported_file_is_flagged():
    """The overlap doesn't have to be within one batch - a new file
    overlapping something imported in an earlier, separate session is
    just as real a risk."""
    existing_log = _log([
        {"SourceBank": "NORWEGIAN", "SourceFile": "Tili_NORWEGIAN_20260720-20260820.xlsx"},
    ])
    new_log = _log([
        {"SourceBank": "NORWEGIAN", "SourceFile": "Card_NORWEGIAN_Statement20260801-20260930.xlsx"},
    ])

    overlaps = find_overlapping_import_windows(existing_log, new_log)

    assert len(overlaps) == 1


def test_two_existing_files_overlapping_each_other_are_never_flagged():
    """
    Real bug this guards against: a purely historical overlap (e.g. a
    multi-year backfill file whose range happens to intersect an old
    month-by-month file, both imported years ago) was already dealt with
    at the time and isn't actionable now - comparing existing_log entries
    against each other would re-flag it on every future run forever,
    burying the one new warning that actually matters in noise. This ran
    into exactly this on the real household workbook: one account's
    multi-year backfill file overlaps several old month-by-month files,
    none of it a real, current problem.
    """
    existing_log = _log([
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20220101-20260518.xlsx"},
        {"SourceBank": "OP", "SourceFile": "PERSON_A_tapahtumat20260503-20260603.csv"},
    ])

    assert find_overlapping_import_windows(existing_log, _EMPTY) == []


def test_file_still_in_the_input_folder_is_not_treated_as_new_just_because_it_got_reparsed():
    """
    Real bug this guards against, found on the actual household workbook:
    old backfill files are never removed from input/budgeting/, so they
    get re-parsed (and so appear in new_log) on every single pipeline
    run, indistinguishable from a genuinely fresh file by new_log alone.
    Two such files whose overlap was already there (and already settled)
    years ago must NOT be re-flagged just because today's run re-parsed
    them both again - they're in new_log, but their filenames already
    exist in existing_log from prior runs, so neither is genuinely new.
    """
    existing_log = _log([
        {"SourceBank": "NORWEGIAN", "SourceFile": "NORWEGIAN_Statement_20220101-20241231.xlsx"},
        {"SourceBank": "NORWEGIAN", "SourceFile": "NORWEGIAN_Statement-20250101-20260522.xlsx"},
    ])
    # Same two files, re-parsed again this run (as they always are, since
    # nothing removes them from the input folder).
    new_log = _log([
        {"SourceBank": "NORWEGIAN", "SourceFile": "NORWEGIAN_Statement_20220101-20241231.xlsx"},
        {"SourceBank": "NORWEGIAN", "SourceFile": "NORWEGIAN_Statement-20250101-20260522.xlsx"},
    ])

    assert find_overlapping_import_windows(existing_log, new_log) == []


def test_same_file_logged_multiple_times_in_existing_log_is_not_duplicated():
    """
    Real bug this guards against: ImportLog records one row per past
    import RUN, so the same file can legitimately appear several times
    across separate historical runs - that must be deduped down to one
    window per file, never multiplied into repeated warnings.
    """
    existing_log = _log([
        {"SourceBank": "NORWEGIAN", "SourceFile": "Tili_NORWEGIAN_20260720-20260820.xlsx"},
        {"SourceBank": "NORWEGIAN", "SourceFile": "Tili_NORWEGIAN_20260720-20260820.xlsx"},
        {"SourceBank": "NORWEGIAN", "SourceFile": "Tili_NORWEGIAN_20260720-20260820.xlsx"},
    ])
    new_log = _log([
        {"SourceBank": "NORWEGIAN", "SourceFile": "Card_NORWEGIAN_Statement20260801-20260930.xlsx"},
    ])

    overlaps = find_overlapping_import_windows(existing_log, new_log)

    assert len(overlaps) == 1


def test_sequential_non_overlapping_files_are_not_flagged():
    new_log = _log([
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260601-20260721.csv"},
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260722-20260819.csv"},
    ])

    assert find_overlapping_import_windows(_EMPTY, new_log) == []


def test_touching_at_a_single_shared_boundary_day_is_not_flagged():
    """20260722 appears as both the end of one file and the start of the
    next - this household's own real, deliberate re-export convention
    (confirmed in the real input files), not a bug. Flagging every
    routine import's normal boundary touch would bury the real signal in
    noise, so only a genuine multi-day overlap should warn."""
    new_log = _log([
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260601-20260722.csv"},
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260722-20260819.csv"},
    ])

    assert find_overlapping_import_windows(_EMPTY, new_log) == []


def test_two_day_overlap_is_flagged():
    """One day beyond the normal boundary touch is enough to mark a real
    overlap - this is the threshold itself, not just the Norwegian case's
    much larger 19-day overlap."""
    new_log = _log([
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260601-20260723.csv"},
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260722-20260819.csv"},
    ])

    overlaps = find_overlapping_import_windows(_EMPTY, new_log)
    assert len(overlaps) == 1


def test_different_accounts_are_never_compared_against_each_other():
    new_log = _log([
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260601-20260819.csv"},
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_B_tapahtumat20260601-20260819.csv"},
    ])

    assert find_overlapping_import_windows(_EMPTY, new_log) == []


def test_skipped_or_format_drift_rows_are_ignored():
    new_log = _log([
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260601-20260819.csv", "Status": "Skipped: format drift"},
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260701-20260901.csv", "Status": "Skipped: unsupported file"},
    ])

    assert find_overlapping_import_windows(_EMPTY, new_log) == []


def test_cash_and_investment_dividend_are_never_flagged():
    new_log = _log([
        {"SourceBank": "CASH", "SourceFile": "CashEntries_20260601-20260819.xlsx"},
        {"SourceBank": "CASH", "SourceFile": "CashEntries_20260701-20260901.xlsx"},
        {"SourceBank": "INVESTMENT_DIVIDEND", "SourceFile": "DividendHistory_20260601-20260819.xlsx"},
        {"SourceBank": "INVESTMENT_DIVIDEND", "SourceFile": "DividendHistory_20260701-20260901.xlsx"},
    ])

    assert find_overlapping_import_windows(_EMPTY, new_log) == []


def test_filename_with_no_derivable_date_is_never_flagged():
    new_log = _log([
        {"SourceBank": "SPANKKI", "SourceFile": "S-PANKKI_export.csv"},
        {"SourceBank": "SPANKKI", "SourceFile": "S-PANKKI_export-2.csv"},
    ])

    assert find_overlapping_import_windows(_EMPTY, new_log) == []


def test_same_filename_twice_in_the_same_batch_is_not_flagged_as_overlapping_itself():
    new_log = _log([
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260601-20260819.csv"},
        {"SourceBank": "OP", "SourceFile": "Tili_PERSON_A_tapahtumat20260601-20260819.csv"},
    ])

    assert find_overlapping_import_windows(_EMPTY, new_log) == []
