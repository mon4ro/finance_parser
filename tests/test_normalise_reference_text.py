from finance_parser.common import normalise_reference_text


def test_plain_string_reference_passes_through():
    assert normalise_reference_text("REF-12345") == "REF-12345"


def test_none_and_blank_return_empty_string():
    assert normalise_reference_text(None) == ""
    assert normalise_reference_text("") == ""
    assert normalise_reference_text(float("nan")) == ""


def test_whole_number_float_formats_as_plain_integer_not_scientific_notation():
    """
    Real bug found and fixed: a long bank reference number read as float64
    (e.g. by openpyxl, if the source spreadsheet stored that cell as a
    number rather than text) renders via plain str() in lossy scientific
    notation - and since Reference feeds directly into RawID hashing, the
    same real transaction can hash differently between two imports
    depending on which dtype pandas happened to infer that time, silently
    defeating de-duplication.
    """
    value = 99887766554433.0
    result = normalise_reference_text(value)

    assert result == "99887766554433"
    assert "e" not in result.lower()


def test_non_integer_float_falls_back_to_normal_text_handling():
    assert normalise_reference_text(3.14) == "3.14"


def test_int_input_unaffected():
    assert normalise_reference_text(12345) == "12345"
