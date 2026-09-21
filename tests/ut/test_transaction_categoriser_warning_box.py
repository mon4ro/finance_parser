from finance_parser.budgeting import transaction_categoriser as categoriser


def test_recategorise_warning_uses_terminal_box_and_colour_helpers(capsys, monkeypatch):
    monkeypatch.setattr(categoriser, "supports_ansi_colour", lambda: False)

    categoriser.print_recategorise_warning("only_automatic")
    output = capsys.readouterr().out

    assert "DESTRUCTIVE RECATEGORISATION" in output
    assert "!==" in output
    assert "Manual categorisation values can be overwritten." in output
    assert "BIG RED WARNING" not in output


def test_recategorise_warning_supports_ansi_colour(monkeypatch):
    monkeypatch.setattr(categoriser, "supports_ansi_colour", lambda: True)

    assert categoriser.colour_text("test", "1;31") == "\033[1;31mtest\033[0m"
