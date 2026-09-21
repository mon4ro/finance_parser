import pandas as pd

from finance_parser.budgeting.parsers import op, norwegian, spankki


def test_spankki_bank_raw_hash_ignores_source_file():
    row = pd.Series({
        "Arkistointitunnus": "20260601990000900000",
        "Maksupäivä": "31.5.2026",
        "Kirjauspäivä": "1.6.2026",
        "Summa": "-4,50",
        "Saajan nimi": "S-MARKET",
        "Maksaja": "PERSON X",
        "Tapahtumalaji": "KORTTIOSTO",
    })

    assert spankki.make_bank_raw_export_id(row, "file_a.csv") == spankki.make_bank_raw_export_id(row, "file_b.csv")


def test_op_bank_raw_hash_ignores_source_file():
    row = pd.Series({
        "Arkistointitunnus": "ABC123",
        "Arvopäivä": "1.6.2026",
        "Kirjauspäivä": "1.6.2026",
        "Määrä EUROA": "-10,00",
        "Saaja/Maksaja": "PRISMA",
        "Viite": "",
        "Viesti": "",
        "Laji": "KORTTIOSTO",
    })

    assert op.make_bank_raw_export_id(row, "old.csv") == op.make_bank_raw_export_id(row, "new.csv")


def test_norwegian_bank_raw_hash_ignores_source_file():
    row = pd.Series({
        "TransactionDate": "2026-06-01",
        "BookDate": "2026-06-02",
        "ValueDate": "2026-06-02",
        "Amount": "-12,34",
        "Text": "K-MARKET",
        "Type": "Purchase",
        "Currency Amount": "",
        "Currency": "EUR",
    })

    assert norwegian.make_bank_raw_export_id(row, "old.xlsx") == norwegian.make_bank_raw_export_id(row, "new.xlsx")
