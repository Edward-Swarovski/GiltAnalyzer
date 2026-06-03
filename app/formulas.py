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
    # L(col12) = Nominal(J) * Coupon%(D) / 100
    return f"=J{row}*D{row}/100"


def capital_uplift_to_par(row: int) -> str:
    # M(col13) = Nominal(J) * (100 - EffPrice(E)) / 100
    return f"=J{row}*(100-E{row})/100"


def approx_gross_cash_gain_to_maturity(row: int) -> str:
    # N(col14) = AnnualCoupon(L) * Years(K) + CapitalUplift(M)
    return f"=L{row}*K{row}+M{row}"


def coupon_tax(row: int, tax_rate: str) -> str:
    # O/P/Q = AnnualCoupon(L) * Years(K) * rate
    return f"=L{row}*K{row}*{tax_rate}"


def approx_net_cash_gain(row: int, tax_column: str) -> str:
    # R/S/T = GrossCashGain(N) - CouponTax(O/P/Q)
    return f"=N{row}-{tax_column}{row}"


def annual_net_gain(row: int, net_column: str) -> str:
    # U/V/W = NetGain(R/S/T) / Years(K)
    return f"=IFERROR({net_column}{row}/K{row},\"\")"
