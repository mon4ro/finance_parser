from datetime import date

from finance_parser.common import extract_query_date_range_from_filename, extract_query_end_date_from_filename


_TODAY = date(2026, 10, 1)


def test_date_range_filename_returns_the_later_date():
    assert extract_query_end_date_from_filename(
        "Tili_CHILD_tapahtumat20260703-20260722.csv", today=_TODAY
    ) == date(2026, 7, 22)


def test_disambiguation_suffix_is_not_mistaken_for_part_of_the_date():
    """
    Real filename shape: a trailing "-N" disambiguates a same-day repeat
    export (e.g. re-downloading the same range twice) - it must not get
    pulled into the date itself.
    """
    assert extract_query_end_date_from_filename(
        "Tili_PERSON_A_tapahtumat20260719-20260819-3.csv", today=_TODAY
    ) == date(2026, 8, 19)


def test_hyphenated_account_name_does_not_confuse_the_date_range():
    """S-PANKKI's own hyphen must not be mistaken for a date separator."""
    assert extract_query_end_date_from_filename(
        "Tili_S-PANKKI_20260601-20260819.csv", today=_TODAY
    ) == date(2026, 8, 19)


def test_multi_year_backfill_range_is_handled():
    assert extract_query_end_date_from_filename(
        "Tili_PERSON_A_tapahtumat20220101-20260518.xlsx", today=_TODAY
    ) == date(2026, 5, 18)


def test_month_only_filename_resolves_to_the_last_day_of_that_month():
    assert extract_query_end_date_from_filename(
        "NORWEGIAN_Statement-202606.xlsx", today=_TODAY
    ) == date(2026, 6, 30)


def test_filename_with_no_date_at_all_returns_none():
    assert extract_query_end_date_from_filename("S-PANKKI_export.csv", today=_TODAY) is None


def test_typo_producing_an_impossible_calendar_date_returns_none_not_a_crash():
    """Real risk this guards against: a digit typo in a renamed filename
    (e.g. month 13) must degrade to "no information", never raise and
    never silently produce a wrong date."""
    assert extract_query_end_date_from_filename(
        "Tili_PERSON_A_tapahtumat20260719-20261332.csv", today=_TODAY
    ) is None


def test_implausibly_future_date_is_rejected_even_though_its_a_valid_calendar_date():
    """A transposed digit can easily produce a syntactically valid date
    that's still obviously wrong (here: years into the future) - this must
    not be trusted just because date() didn't raise."""
    assert extract_query_end_date_from_filename(
        "Tili_PERSON_A_tapahtumat20260719-20960819.csv", today=_TODAY
    ) is None


def test_implausibly_old_date_is_rejected():
    assert extract_query_end_date_from_filename(
        "Tili_PERSON_A_tapahtumat19500101-19500201.csv", today=_TODAY
    ) is None


def test_same_day_as_today_is_accepted():
    assert extract_query_end_date_from_filename(
        "Tili_PERSON_A_tapahtumat20260901-20261001.csv", today=_TODAY
    ) == _TODAY


def test_multiple_date_ranges_in_one_filename_uses_the_latest():
    """Defensive: if a filename somehow carries more than one range-like
    pattern, always trust the latest one, never the first found."""
    assert extract_query_end_date_from_filename(
        "merged_20260101-20260201_and_20260201-20260301.csv", today=_TODAY
    ) == date(2026, 3, 1)


def test_range_extractor_returns_matching_start_and_end():
    assert extract_query_date_range_from_filename(
        "Tili_PERSON_A_tapahtumat20260601-20260722.csv", today=_TODAY
    ) == (date(2026, 6, 1), date(2026, 7, 22))


def test_range_extractor_single_date_is_a_zero_width_range():
    assert extract_query_date_range_from_filename(
        "Tili_S-PANKKI_20260819.csv", today=_TODAY
    ) == (date(2026, 8, 19), date(2026, 8, 19))


def test_range_extractor_month_only_spans_the_whole_month():
    assert extract_query_date_range_from_filename(
        "NORWEGIAN_Statement-202606.xlsx", today=_TODAY
    ) == (date(2026, 6, 1), date(2026, 6, 30))


def test_range_extractor_rejects_a_reversed_range():
    """A start-after-end range is definitely a garbled filename, not real
    coverage - must degrade to None, not silently swap the two."""
    assert extract_query_date_range_from_filename(
        "Tili_PERSON_A_tapahtumat20260722-20260601.csv", today=_TODAY
    ) is None


def test_range_extractor_multiple_ranges_picks_the_pair_with_the_latest_end():
    """Must pick the END's own matching START, not mix start from one pair
    with end from a different one."""
    assert extract_query_date_range_from_filename(
        "merged_20260101-20260201_and_20260215-20260301.csv", today=_TODAY
    ) == (date(2026, 2, 15), date(2026, 3, 1))


def test_range_extractor_no_date_returns_none():
    assert extract_query_date_range_from_filename("S-PANKKI_export.csv", today=_TODAY) is None
