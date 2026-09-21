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
