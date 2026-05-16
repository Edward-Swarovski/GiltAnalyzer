from app import formulas


def test_effective_price_formula() -> None:
    assert formulas.effective_price(2) == (
        '=IFERROR(IF(INDEX(Inputs!D:D,MATCH(A2,Inputs!A:A,0))<>"",'
        'INDEX(Inputs!D:D,MATCH(A2,Inputs!A:A,0)),'
        'INDEX(\'Market Data\'!E:E,MATCH(A2,\'Market Data\'!A:A,0))),"")'
    )


def test_cgt_formula_is_zero() -> None:
    assert formulas.cgt_on_conventional_gilt_capital_gain(2) == "=0"


def test_net_gain_formula_references_tax_and_cgt_columns() -> None:
    assert formulas.approx_net_cash_gain(2, "M") == "=L2-M2-P2"



def test_nominal_amount_formula_falls_back_to_settings() -> None:
    formula = formulas.nominal_amount(2)
    assert "Settings!B:B" in formula
    assert 'MATCH("DefaultNominalAmount",Settings!A:A,0)' in formula
    assert "Inputs!C:C" in formula


def test_effective_yield_formula() -> None:
    assert formulas.effective_yield(2) == (
        '=IFERROR(IF(INDEX(Inputs!E:E,MATCH(A2,Inputs!A:A,0))<>"",'
        'INDEX(Inputs!E:E,MATCH(A2,Inputs!A:A,0)),'
        'INDEX(\'Market Data\'!F:F,MATCH(A2,\'Market Data\'!A:A,0))),"")'
    )
