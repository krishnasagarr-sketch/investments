# Fixed Deposit Manager

A web app for tracking bank deposits — built with Flask and SQLite. Add each deposit's amount, rate, and tenure, and it calculates the maturity amount, tracks maturity dates, and shows total interest earned across all your deposits.

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
       └── banks.html
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

Open **http://127.0.0.1:5000**. A `fixed_deposits.db` SQLite file is created automatically on first run — your data persists across restarts. It holds three tables: `depositors` (holder ID + name), `banks` (bank ID + name), and `deposits` (which references a depositor and a bank by id). On startup the app runs in-place migrations, so an older `.db` from a previous version is upgraded automatically; any legacy free-text holder or bank name on an existing deposit is promoted into a `depositors` / `banks` row and linked.

> **Add at least one depositor and one bank first** (Depositors / Banks tabs) — the Add Deposit form needs both to attach the deposit to.

## Features

- **Dashboard** — the first tab: portfolio totals plus a **holdings breakdown by holder & bank** (grouped rows with deposit count, invested, current value, maturity value, and a grand total), and separate **by-holder** and **by-bank** rollups.
- **Chart** — the same by-holder and by-bank figures as a chart. Pick the **chart type** (bar or pie) with the toggle at the top:
  - **Bar** — grouped horizontal bars showing invested / current value / maturity value together. Bars share one linear scale, so a single very large deposit will dwarf the rest.
  - **Pie** — one metric at a time (choose invested / current value / maturity value), showing each holder's or bank's share of the total, with a legend of amounts and percentages.

  All charts are pure inline SVG — no JavaScript or chart library. The selection is kept in the URL (`/chart?type=pie&metric=maturity`).
- **Depositors** — a separate master list (`depositors` table) of people who hold deposits, each with a unique **holder ID** (customer number, PAN, etc.) and a **name**. Managed on the **Depositors** tab, which also shows each depositor's **total invested** (principal only, no interest), **current value** and **total maturity value**, with a combined total across everyone. A depositor can't be removed while any deposit references it.
- **Banks** — a separate master list (`banks` table), each with a unique **bank ID** (IFSC / branch code / any identifier) and a **name**. Managed on the **Banks** tab; a bank can't be removed while any deposit references it.
- **Add deposits** — pick the depositor and the bank from dropdowns, then deposit type, amount (lump-sum principal, or monthly installment for an RD), annual interest rate, tenure (months or days), compounding frequency (cumulative only), start date. The form relabels fields, switches the tenure unit, and shows/hides compounding frequency based on the type you pick.
- **Edit deposits** — the **Edit** link on each dashboard row opens the same form pre-filled; saving updates the row in place.
- **Holder / bank tracking** — each deposit is linked to its depositor and its bank; the dashboard shows the depositor's name + holder ID and the bank's name + bank ID (older, unlinked rows show "—")
- **Automatic calculations** — maturity date, maturity amount, and interest earned, using the formula for the chosen deposit type (see table above)
- **Current value** — for every deposit, the value *today* (principal + interest accrued so far): simple interest accrues linearly, cumulative compounds, and an RD sums each installment paid to date compounded quarterly. Once a deposit matures, current value equals the maturity amount.
- **RD progress** — recurring deposits also show how many installments have been paid so far (`11/24 paid`) and the cash actually paid in.
- **Status tracking** — shows "Matured" or days remaining until maturity for each deposit
- **Portfolio summary** — total invested, current value, total maturity value, and total interest at maturity across all deposits (for RDs, "invested" is installment × number of installments)
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