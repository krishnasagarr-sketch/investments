# Fixed Deposit Manager

A web app for tracking bank deposits — built with Flask and SQLite. Add each deposit's amount, rate, and tenure, and it calculates the maturity amount, tracks maturity dates, and shows total interest earned across all your deposits.

All amounts are shown in Indian rupees (`₹`), with lakh/crore digit grouping (e.g. `₹12,34,567.89`).

Supports three deposit types:

| Type | How interest works | Maturity formula |
|---|---|---|
| **Cumulative (reinvested)** | Interest is reinvested and paid with the principal at maturity | `A = P × (1 + r/n)^(n×t)` |
| **Simple interest (payout)** | Interest is paid out periodically and *not* reinvested | `A = P + (P × r × t)` |
| **Recurring deposit (RD)** | One fixed installment per month; the balance compounds quarterly | sum of each installment `M × (1 + r/4)^(months_on_deposit / 3)` |

where `P` = principal (or `M` = monthly installment for an RD), `r` = annual rate as a decimal, `n` = compounding periods per year, `t` = tenure in years.

**Tenure** can be entered in **months** or **days** (cumulative and simple-interest deposits). Day tenures convert to years on a 365-day basis (`t = days / 365`) and the maturity date is the start date plus that many days. Recurring deposits are always monthly.

## Setup

1. Folder layout (keep this exact structure):
   ```
   fd-manager/
   ├── app.py
   ├── requirements.txt
   └── templates/
       ├── base.html
       ├── dashboard.html
       ├── add_deposit.html
       ├── depositors.html
       ├── banks.html
       ├── metals.html
       ├── summary.html
       └── chart.html
   ```

2. Create a virtual environment (recommended):
   ```bash
   python -m venv venv
   source venv/bin/activate      # macOS/Linux
   venv\Scripts\activate         # Windows
   ```

3. Install dependencies:
   ```bash
   pip install -r requirements.txt
   ```

## Run

```bash
python app.py
```

Open **http://127.0.0.1:5000**. A `fixed_deposits.db` SQLite file is created automatically on first run — your data persists across restarts. It holds `depositors` (holder ID + name), `banks` (bank ID + name), `deposits` (references a depositor and a bank by id), `metals` (precious-metal holdings), and `metal_prices` (one live ₹/gram rate per metal). On startup the app runs in-place migrations, so an older `.db` from a previous version is upgraded automatically; legacy free-text holder / bank names are promoted into `depositors` / `banks` rows, and `metal_prices` is seeded from each metal's most recent holding.

> **Add at least one depositor and one bank first** (Depositors / Banks tabs) — the Add Deposit form needs both to attach the deposit to.

## Features

- **Dashboard** — the first tab: whole-portfolio totals (deposits **and** metals) — total invested, current value, unrealised gain/loss with return %, deposit maturity value, and **annualised return**. Then breakdowns (each with its own Ann. Return column): **by asset class** (deposits vs metals), **by holder across all assets** (including total metal grams per holder), **deposits by holder & bank**, **metals by type** (grams + value per metal), **metals by holder & type** (grams per holder+metal), and **deposits by bank**.
- **Annualised return** — a CAGR-style figure (`(current/invested)^(365/days held) − 1`) shown per deposit, per metal holding, and as an invested-weighted blend for every group and the whole portfolio. It's a rough blend across positions with different start dates, not a true money-weighted (XIRR) return — treat it as directional. Holdings younger than 7 days show "—" rather than an exaggerated figure (a 1-day gain projected over a year would be misleading).
- **Chart** — the same figures as a chart, covering deposits and metals. Pick the **chart type** (bar or pie) with the toggle at the top:
  - **Bar** — two grouped horizontal bars per row: invested (money in / cost) vs current value. Sections: by holder (all assets), deposits by bank, metals by type.
  - **Pie** — invested *or* current value, showing each holder's / bank's / metal's share of the total, with an amount + percentage legend.

  All charts are pure inline SVG — no JavaScript or chart library. The selection is kept in the URL (`/chart?type=pie&metric=invested`).
- **Depositors** — a separate master list (`depositors` table) of people who hold deposits, each with a unique **holder ID** (customer number, PAN, etc.) and a **name**. Managed on the **Depositors** tab, which also shows each depositor's **total invested** (principal only, no interest), **current value** and **total maturity value**, with a combined total across everyone. A depositor can't be removed while any deposit references it.
- **Banks** — a separate master list (`banks` table), each with a unique **bank ID** (IFSC / branch code / any identifier) and a **name**. Managed on the **Banks** tab; a bank can't be removed while any deposit references it.
- **Metals** — record precious-metal holdings (`metals` table): metal (**Gold 24K**, **Gold 22K**, silver, platinum, palladium, other — 24K and 22K are tracked separately, each with its own market rate), an optional **depositor** (linked to the depositors list), an optional description, weight in **grams**, and the **purchase price in ₹ per gram**. Current value is driven by a separate **market-price table** (`metal_prices`) — one live ₹/gram rate per metal, so you set the rate once and every holding of that metal revalues. A metal with no rate set falls back to each holding's purchase price (value = cost). The tab shows cost, current value, total return %, annualised return, and unrealised gain/loss per holding and overall.
  - **Fetch Live Prices** — pulls the current spot rate for gold, silver, platinum and palladium (24K gold's rate is used to derive 22K at 22/24 purity) and converts it to ₹/gram, using [gold-api.com](https://gold-api.com) for spot prices and [frankfurter.app](https://frankfurter.app) for the USD→INR rate — no API key needed, no new Python dependency (uses the standard library). Network failures show an error and leave existing prices untouched.
  - **Manual entry** stays available in the same panel — type a price directly to override, or to set "Other" (which has no live source). Each price shows whether it's `🌐 live` or `✎ manual`, and when it was last updated.
- **Add deposits** — pick the depositor and the bank from dropdowns, then deposit type, amount (lump-sum principal, or monthly installment for an RD), annual interest rate, tenure (months or days), compounding frequency (cumulative only), start date. The form relabels fields, switches the tenure unit, and shows/hides compounding frequency based on the type you pick.
- **Edit deposits** — the **Edit** link on each dashboard row opens the same form pre-filled; saving updates the row in place.
- **Holder / bank tracking** — each deposit is linked to its depositor and its bank; the dashboard shows the depositor's name + holder ID and the bank's name + bank ID (older, unlinked rows show "—")
- **Automatic calculations** — maturity date, maturity amount, and interest earned, using the formula for the chosen deposit type (see table above)
- **Current value** — for every deposit, the value *today* (principal + interest accrued so far): simple interest accrues linearly, cumulative compounds, and an RD sums each installment paid to date compounded quarterly. Once a deposit matures, current value equals the maturity amount.
- **RD progress** — recurring-deposit rows show the installment amount and how many have been paid so far (`₹1,000.00/mo · 11/24 paid`); the "Invested" figure is that cash paid in to date.
- **Status tracking** — shows "Matured" or days remaining until maturity for each deposit
- **Portfolio summary** — total invested, current value, total maturity value, and total interest at maturity across all deposits (for RDs, "invested" is the installment amount × the number of installments **paid so far**, not the full-term commitment)
- **Remove deposits / depositors** — the **Remove** button is a two-step confirm (click once to arm, again within 4 s to delete). It doesn't use a native `confirm()` dialog, so it still works in embedded browsers that block those.

## Notes

- Interest rate should be entered as a percentage (e.g., `6.5` for 6.5%), not a decimal.
- **Simple-interest deposits:** the "Maturity Amount" column shows principal + total interest over the term. In practice that interest is paid out to you along the way rather than in one lump at maturity.
- **Recurring deposits:** the RD formula assumes quarterly compounding and one installment at the start of each month, which is what most Indian banks use. Your bank's figure may differ by a small amount depending on its exact day-count and rounding.
- This tool doesn't account for tax on interest (e.g., TDS) or premature withdrawal penalties — it assumes the deposit runs to full maturity as entered.
- All figures are for personal tracking only; confirm exact maturity values with your bank.

## Possible next steps

- Track TDS/tax withheld on interest
- Reminders/notifications as FDs approach maturity
- Auto-renewal tracking (roll maturity amount into a new FD)
- Multi-currency support
- Export to CSV for tax filing