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


def accrued_interest(row: int) -> str:
    # Accrued Interest = (Nominal × Coupon% / 100 / 2) × (days since last coupon / days in period)
    # Coupon dates derived from maturity date (col C): same day, coupon months = MONTH(C) and MONTH(C)±6
    # cm1 = first coupon month, cm2 = second coupon month (6 months apart)
    # Last coupon = most recent of the 4 candidate dates (cm1/cm2 × this year/last year) that is <= TODAY()
    # Next coupon = earliest candidate that is > TODAY()
    c = f"C{row}"
    j = f"J{row}"
    d = f"D{row}"

    # Coupon months
    cm1 = f"MONTH({c})"
    cm2 = f"IF(MONTH({c})<=6,MONTH({c})+6,MONTH({c})-6)"
    dy = f"DAY({c})"

    # Four candidate coupon dates: cm1 this year, cm1 last year, cm2 this year, cm2 last year
    def cdate(month_expr: str, year_offset: int) -> str:
        y = f"YEAR(TODAY()){'+' + str(year_offset) if year_offset >= 0 else str(year_offset)}"
        return f"DATE({y},{month_expr},{dy})"

    c1_ty = cdate(cm1, 0)   # cm1 this year
    c1_ly = cdate(cm1, -1)  # cm1 last year
    c2_ty = cdate(cm2, 0)   # cm2 this year
    c2_ly = cdate(cm2, -1)  # cm2 last year

    # Last coupon = MAX of candidates that are <= TODAY()
    last = (
        f"MAX("
        f"IF({c1_ty}<=TODAY(),{c1_ty},{c1_ly}),"
        f"IF({c2_ty}<=TODAY(),{c2_ty},{c2_ly})"
        f")"
    )

    # Next coupon = MIN of candidates that are > TODAY()
    nxt = (
        f"MIN("
        f"IF({c1_ty}>TODAY(),{c1_ty},DATE(YEAR(TODAY())+1,{cm1},{dy})),"
        f"IF({c2_ty}>TODAY(),{c2_ty},DATE(YEAR(TODAY())+1,{cm2},{dy}))"
        f")"
    )

    # Accrued = (J * D/100/2) * (TODAY()-last) / (next-last)
    return (
        f"=IFERROR("
        f"({j}*{d}/100/2)*((TODAY()-({last}))/(({nxt})-({last}))),"
        f'"")'
    )
