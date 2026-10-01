import pandas as pd

from finance_parser.budgeting.transaction_parser import disambiguate_duplicate_raw_ids


def test_no_collisions_leaves_every_raw_id_untouched():
    df = pd.DataFrame({"RawID": ["A", "B", "C"], "Amount": [-1, -2, -3]})
    out = disambiguate_duplicate_raw_ids(df)
    assert list(out["RawID"]) == ["A", "B", "C"]


def test_colliding_raw_ids_get_a_stable_deterministic_suffix():
    """
    Real motivating case: two genuinely separate airline charges (two
    seats, billed separately) hash to the identical RawID because
    NORWEGIAN's export has no field rich enough to tell them apart on its
    own. The first occurrence keeps the original RawID unchanged (so it
    still matches whatever was already imported before this fix existed);
    later occurrences get a stable suffix instead of being silently
    dropped.
    """
    df = pd.DataFrame({"RawID": ["NWG-X", "NWG-X", "NWG-X"], "Amount": [-89.90, -89.90, -89.90]})
    out = disambiguate_duplicate_raw_ids(df)
    assert list(out["RawID"]) == ["NWG-X", "NWG-X-dup2", "NWG-X-dup3"]
    assert out["RawID"].is_unique


def test_reparsing_the_same_file_reproduces_identical_suffixes():
    """
    Stability requirement: a genuine re-import of the exact same file
    parses rows in the same order, so the same occurrence -> same suffix -
    letting the existing "already imported" RawID match still work
    correctly on a real re-import, not just on first import.
    """
    df = pd.DataFrame({"RawID": ["NWG-X", "NWG-X"], "Amount": [-20.0, -20.0]})
    first = disambiguate_duplicate_raw_ids(df)
    second = disambiguate_duplicate_raw_ids(df.copy())
    assert list(first["RawID"]) == list(second["RawID"])


def test_multiple_distinct_colliding_groups_are_each_disambiguated_independently():
    df = pd.DataFrame({
        "RawID": ["A", "A", "B", "B", "B", "C"],
        "Amount": [-1, -1, -2, -2, -2, -3],
    })
    out = disambiguate_duplicate_raw_ids(df)
    assert list(out["RawID"]) == ["A", "A-dup2", "B", "B-dup2", "B-dup3", "C"]
    assert out["RawID"].is_unique


def test_does_not_mutate_the_input_frame():
    df = pd.DataFrame({"RawID": ["A", "A"], "Amount": [-1, -1]})
    disambiguate_duplicate_raw_ids(df)
    assert list(df["RawID"]) == ["A", "A"]
