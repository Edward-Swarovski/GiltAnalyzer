# UK Gilt Knowledge Reference

## What is a Gilt?

A gilt is a bond issued by the UK government (via the Debt Management Office, DMO). The name comes from the original certificates having gilded edges. Gilts are considered one of the safest investments available because the UK government has never defaulted on its debt.

When you buy a gilt, you are lending money to the government. In return:
- You receive fixed coupon payments twice a year
- You receive £100 per £100 nominal back when the gilt matures

---

## Key Terms

### Nominal Amount (Face Value)
The £ amount printed on the gilt — what the government promises to repay at maturity. Coupons are calculated as a percentage of this amount, not the price you paid.

### Clean Price
The market price of the gilt, expressed per £100 nominal, **excluding** accrued interest. This is what you see quoted on broker screens and in market data (e.g. DividendData). A price below 100 means the gilt trades at a discount to par; above 100 is a premium.

### Dirty Price (Full Price)
What you actually pay at settlement:
> Dirty Price = Clean Price + Accrued Interest

When buying a gilt between coupon dates, you compensate the seller for the coupon that has been accruing since the last payment date. This is automatic at settlement.

### Par (100)
The redemption value. All conventional gilts redeem at exactly £100 per £100 nominal at maturity, regardless of what price you paid.

### Coupon Rate
The annual interest rate, expressed as a percentage of nominal. A 4% gilt on £10,000 nominal pays £400/year, split as £200 every 6 months.

- Coupons are **fixed** — they never change over the life of the gilt
- Coupons do **not** compound — each payment is the same cash amount
- Coupons are paid **semi-annually** (twice a year) for all UK conventional gilts

### Running Yield (Current Yield)
The annual coupon cash as a percentage of the price you actually pay:
> Running Yield = Coupon % ÷ Clean Price × 100

Example: 1.5% coupon gilt priced at £91 → Running Yield = 1.5 ÷ 91 × 100 = **1.65%**

This is the most honest measure of your annual income relative to what you invest. It does not account for the capital gain/loss to maturity.

### Yield to Maturity (YTM)
The total annualised return if you buy today and hold to maturity, assuming you reinvest coupons at the same rate. It accounts for:
- Coupon income
- Capital gain (if bought below £100) or loss (if bought above £100)
- Time remaining to maturity

YTM moves **inversely** to price: if the price rises, the yield falls, and vice versa.

**Important:** YTM implicitly assumes coupon reinvestment at the same rate. In practice:
- If you spend the coupons (take as income), your actual return is slightly lower
- For short-dated gilts (≤3 years), the difference is small
- For long-dated gilts (20–30 years), the reinvestment assumption matters significantly

### Accrued Interest
The portion of the next coupon that has built up since the last payment date. When you sell before a coupon date, the buyer pays you this automatically in the dirty price. When you buy, you pay the seller their accrued interest.

### Capital Gain / Capital Uplift
The difference between £100 par and your purchase price, received at maturity:
> Capital Uplift = Nominal × (100 − Clean Price) ÷ 100

If you bought at £91, you receive £9 per £100 nominal extra at maturity. For UK **conventional gilts**, this capital gain is **exempt from Capital Gains Tax (CGT)** regardless of the gain size — a significant advantage for higher-rate taxpayers.

---

## Conventional vs Index-Linked Gilts

This tool covers **conventional gilts** only.

| | Conventional | Index-Linked |
|---|---|---|
| Coupon | Fixed £ amount | Rises with RPI inflation |
| Redemption | Fixed £100 | Rises with RPI inflation |
| CGT | Exempt | Exempt |
| Income tax | Coupon taxable | Coupon taxable |
| Best for | Certainty of return | Inflation protection |

---

## Tax Treatment (UK)

### Income Tax on Coupons
Coupon payments are taxed as **income** in the tax year received:
- Basic rate (20%): pay 20% of each coupon
- Higher rate (40%): pay 40% of each coupon
- Additional rate (45%): pay 45% of each coupon

If held in an ISA or SIPP, coupons are received free of income tax.

### Capital Gains Tax
**Zero.** Capital gains on conventional gilts are fully exempt from CGT for UK individuals. This makes gilts particularly attractive compared to equities or corporate bonds for higher-rate taxpayers — you can lock in a known capital gain at zero tax cost.

### Accrued Interest (Bond Washing Rules)
HMRC has anti-avoidance rules around buying gilts just before a coupon and selling just after. In practice, for a genuine buy-and-hold investor this is rarely relevant.

---

## Pricing Mechanics

### Why Gilt Prices Move
Gilt prices move primarily because **interest rates change**. If the Bank of England raises rates, newly issued gilts offer higher yields, making existing lower-coupon gilts less attractive — their prices fall until their yields match the market. The relationship is:

> Price ↑ → Yield ↓  
> Price ↓ → Yield ↑

### Duration (Interest Rate Sensitivity)
Longer-dated gilts are more sensitive to interest rate changes than short-dated gilts. A 30-year gilt might fall 20% in price from a 1% rate rise; a 2-year gilt might fall only 2%. This is measured by "duration."

### Buying Below Par (Discount Gilt)
Price < £100. You receive:
- Lower running yield than the coupon rate (less income per £ invested)
- A capital gain at maturity (tax-free for conventional gilts)
- Higher YTM than the coupon rate

Common for low-coupon gilts issued when rates were lower (e.g. 0.125% gilt now yields ~4.5%).

### Buying Above Par (Premium Gilt)
Price > £100. You receive:
- Higher running yield than the coupon rate (more income per £ invested)
- A capital loss at maturity (you paid more than £100, get back £100)
- Lower YTM than the coupon rate

---

## Selling Before Maturity

You are not obligated to hold a gilt to maturity. The UK gilt market is highly liquid.

### What you keep
- All coupons already paid to you — unconditionally yours
- Accrued interest since the last coupon — received automatically in the dirty price

### What determines your capital outcome
The clean price at the date you sell. This could be:
- Higher than your purchase price → capital gain (tax-free)
- Lower than your purchase price → capital loss

### Key risk: interest rate risk
If rates have risen since you bought, the price will have fallen. You may need to sell at a loss. The longer the remaining maturity, the larger the potential price move.

### Modelling an early exit in GiltAnalyzer
Use the **Override Price** column (column D in the Inputs sheet) to enter your expected sale price. All cash flow calculations in the Analysis sheet will then reflect that exit price rather than the £100 par redemption value.

---

## DMO Purchase and Sale Service (Retail)

The UK Debt Management Office operates a retail gilt service allowing individual investors to buy and sell gilts directly, without a broker.

- **Buy price (Ask):** slightly above the market mid price
- **Sell price (Bid):** slightly below the market mid price
- The spread represents the DMO's service cost

The **Approx Retail Ask Yield %** column in this tool reflects the yield at the DMO retail ask price. It will always be slightly lower than the Effective Yield (market mid) because you are paying a slightly higher price.

---

## GiltAnalyzer Calculation Notes

### Gross Cash Gain to Maturity
> Annual Coupon Cash × Years to Maturity + Capital Uplift to Par

This is a simple (non-compounding) approximation. It does not discount future cash flows.

### Net Cash Gain (after income tax)
> Gross Cash Gain − (Annual Coupon Cash × Years × Tax Rate)

CGT is always zero for conventional gilts, so capital uplift is never taxed.

### Annual Net Return
> Net Cash Gain ÷ Years to Maturity

Allows fair comparison between gilts with different maturities.

### Annual Net Return per £10k (Best Value sheet)
All figures are normalised to a £10,000 nominal holding, so gilts are comparable regardless of the default nominal amount setting.

### Limitations
- Yield calculations are approximations; the exact Actual/Actual ICMA day count convention is not implemented
- Coupons are modelled as a continuous annual flow, not discrete semi-annual payments
- No discounting of future cash flows (not a DCF model)
- Retail ask yield bisection uses days/365.25, not the exact gilt day count
