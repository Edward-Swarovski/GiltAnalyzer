def effective_price(row: int) -> str:
    return (
        f'=IFERROR(IF(INDEX(Inputs!D:D,MATCH(A{row},Inputs!A:A,0))<>"",'
        f'INDEX(Inputs!D:D,MATCH(A{row},Inputs!A:A,0)),'
        f'INDEX(\'Market Data\'!E:E,MATCH(A{row},\'Market Data\'!A:A,0))),"")'
    )


def effective_yield(row: int) -> str:
    return (
        f'=IFERROR(IF(INDEX(Inputs!E:E,MATCH(A{row},Inputs!A:A,0))<>"",'
        f'INDEX(Inputs!E:E,MATCH(A{row},Inputs!A:A,0)),'
        f'INDEX(\'Market Data\'!F:F,MATCH(A{row},\'Market Data\'!A:A,0))),"")'
    )


def nominal_amount(row: int) -> str:
    # Falls back to Settings!B2 (DefaultNominalAmount) when Inputs col C is blank.
    return (
        f'=IFERROR('
        f'IF(INDEX(Inputs!C:C,MATCH(A{row},Inputs!A:A,0))<>"",'
        f'INDEX(Inputs!C:C,MATCH(A{row},Inputs!A:A,0)),'
        f'INDEX(Settings!B:B,MATCH("DefaultNominalAmount",Settings!A:A,0))),'
        f'"")'
    )


def years_to_maturity(row: int) -> str:
    return (
        f'=IFERROR(YEARFRAC('
        f'INDEX(\'Market Data\'!G:G,MATCH(A{row},\'Market Data\'!A:A,0)),'
        f'C{row},1),"")'
    )


def annual_coupon_cash(row: int) -> str:
    return f"=H{row}*D{row}/100"


def capital_uplift_to_par(row: int) -> str:
    return f"=H{row}*(100-E{row})/100"


def approx_gross_cash_gain_to_maturity(row: int) -> str:
    return f"=J{row}*I{row}+K{row}"


def coupon_tax(row: int, tax_rate: str) -> str:
    return f"=J{row}*I{row}*{tax_rate}"


def cgt_on_conventional_gilt_capital_gain(_: int) -> str:
    return "=0"


def approx_net_cash_gain(row: int, tax_column: str) -> str:
    return f"=L{row}-{tax_column}{row}-P{row}"


def annual_net_gain(row: int, net_column: str) -> str:
    return f"=IFERROR({net_column}{row}/I{row},\"\")"
