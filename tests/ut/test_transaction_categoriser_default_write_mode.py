from finance_parser.budgeting.transaction_categoriser import build_arg_parser


def test_fresh_rebuild_is_default_write_mode():
    parser = build_arg_parser()
    args = parser.parse_args([])

    assert args.write_mode == "fresh-rebuild"
