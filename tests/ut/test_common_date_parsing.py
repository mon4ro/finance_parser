from finance_parser.common import format_date


def test_format_date_handles_bare_iso_date():
    assert format_date("2026-05-12") == "2026-05-12"


def test_format_date_handles_finnish_dotted_date():
    assert format_date("12.5.2026") == "2026-05-12"


def test_format_date_handles_full_iso8601_timestamp_with_timezone():
    """
    Real bug: Coinmotion's export uses full ISO 8601 timestamps
    ("2026-02-07T11:17:44+02:00"), which only matched the ISO-date branch's
    re.fullmatch (bare yyyy-mm-dd only) - so it fell through to the final
    dayfirst=True pandas parse, silently swapping month/day
    (2026-02-07 -> misread as 2 July instead of 7 February).
    """
    assert format_date("2026-02-07T11:17:44+02:00") == "2026-02-07"


def test_format_date_handles_iso8601_timestamp_without_timezone():
    assert format_date("2026-11-03T00:00:00") == "2026-11-03"


def test_format_date_blank_returns_empty_string():
    assert format_date("") == ""
    assert format_date(None) == ""


def test_format_date_handles_slash_separated_iso_date():
    """
    Real bug: a Nordea export started using "/" instead of "-" for its
    year-first date column ("2026/09/15" instead of "2026-09-15"). Only
    "-" was matched by the ISO-date branch, so these fell through to the
    dayfirst=True fallback, which swapped day and month whenever both were
    <=12 (pandas itself warns about this exact case) - e.g. "2026/05/12"
    (12 May) silently became 5 December. Confirmed against the real
    export: 77 of 191 rows in one real Nordea import were affected, 12 of
    them landing in a future month/day. Unambiguous cases (day > 12) look
    fine as a distraction - they only "worked" because the wrong day/month
    order happened to be impossible, not because parsing was correct.
    """
    assert format_date("2026/05/12") == "2026-05-12"
    assert format_date("2026/09/15") == "2026-09-15"
    assert format_date("2026/01/09") == "2026-01-09"


def test_format_date_still_treats_short_slashed_date_as_day_first():
    """
    A slash date with a 1-2 digit leading token (not a 4-digit year) must
    still go through the existing day-first euro branch, not the
    year-first ISO branch - "5/12/2026" is 5 December, not treated as
    a (nonsensical) 4-digit-year match.
    """
    assert format_date("5/12/2026") == "2026-12-05"
