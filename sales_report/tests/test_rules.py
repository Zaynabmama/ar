from sales_report.logic import build


def line(inv, term, code, amt):
    return {"Invoice No": inv, "Pay Terms": term, "Cust Code": code, "Cust Name": "X",
            "Customer Region": "UAE", "Adjusted Sales": amt}


def test_rules():
    terms = {"ADV-TRF-ADVANCE PAYMENT": "Adv", "(blank)": "Open"}
    r = build([
        line("INV-1", "ADV-TRF-ADVANCE PAYMENT", "CD4A0001", 10),  # affiliate wins
        line("INV-2", "ADV-TRF-ADVANCE PAYMENT", "CD1A0001", 5),
        line("INV-2", "ADV-TRF-ADVANCE PAYMENT", "CD1A0001", 5),   # lines summed per invoice
        line("INV-3", "NEW-TERM", "CD1A0001", 7),                  # unknown -> Open + flagged
        line("SRCL-1", "", "CD1A0001", -3),                         # credit note
        line("SR-2", "", "CD1A0001", 0),                            # zero SR not in SR view
    ], terms)
    s = r["split"]
    assert [x[0] for x in s["Aff"]] == ["INV-1"]
    assert s["Adv"][0][5] == 10
    assert [x[0] for x in s["Open"]] == ["INV-3", "SRCL-1", "SR-2"]
    assert list(r["new_terms"]) == ["NEW-TERM"]
    assert [x[0] for x in r["sr"]] == ["SRCL-1"]
