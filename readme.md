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
       ├── metal_prices.html
       ├── investments.html
       ├── summary.html
       ├── chart.html
       ├── calculator.html
       ├── notifications.html
       ├── setup.html
       ├── login.html
       ├── forgot_password.html
       └── reset_password.html
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

Open **http://127.0.0.1:5000**. The first visit takes you to a one-time **account setup** page (see **Login** below) before anything else loads. A `fixed_deposits.db` SQLite file is created automatically on first run — your data persists across restarts. It holds `depositors` (holder ID + name), `banks` (bank ID + name), `deposits` (references a depositor and a bank by id), `metals` (precious-metal holdings), `metal_prices` (one live ₹/gram rate per metal), `investments` (stock/mutual fund holdings, linked to a depositor), `notification_settings` (the maturity-email settings, a single row), and `auth_user` (the single login account, a single row). On startup the app runs in-place migrations, so an older `.db` from a previous version is upgraded automatically; legacy free-text holder / bank names are promoted into `depositors` / `banks` rows, and `metal_prices` is seeded from each metal's most recent holding.

> **Add at least one depositor and one bank first** (Depositors / Banks tabs) — the Add Deposit form needs both to attach the deposit to.

## Features

- **Login** — the whole app sits behind a single login (`auth_user` table, one account). The very first visit shows a **Set up your account** page (email + password) instead of the dashboard; every page after that redirects to **Log in** until you sign in, and stays signed in for 30 days via a signed session cookie. Passwords are hashed with Werkzeug's `generate_password_hash`/`check_password_hash` (never stored in plain text). The session-signing key (`.flask_secret_key`) is generated once on first run and gitignored — don't delete it, or every existing session is invalidated.
  - **Forgot your password?** on the login page emails a one-time reset link (valid for **1 hour**) to the account's email, reusing the same **Gmail sender + App Password already configured on the Notifications tab** — set that up first if you haven't. The link opens a **Reset your password** page; an expired or already-used link shows an explicit "invalid or has expired" message rather than silently failing.
- **Notifications** — email yourself when a deposit is close to maturing. On the **Notifications** tab: turn it on, set a recipient email, a **sender Gmail address + Gmail App Password** (not your normal password — generate one at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) after enabling 2-Step Verification), and how many days before maturity to alert. Sends via Gmail SMTP using only the Python standard library (`smtplib`) — no new dependency.
  - A background thread checks every **12 hours** (and once at startup) for as long as `python app.py` is running — there's no separate scheduler to configure, but it also means nothing fires while the app is stopped.
  - Each maturing deposit is bundled into **one digest email**; a deposit already alerted is skipped for **7 days** even if it's checked again, so you don't get the same reminder daily. The Notifications page shows every deposit within the window and whether it's due for a fresh alert or was recently sent.
  - **Send Test Email** verifies your SMTP settings immediately without touching deposit state; **Check & Send Now** runs the real check on demand. Both report success/failure right on the page — a bad password shows *"Gmail rejected the sender email / app password"* rather than failing silently.
  - The app password is stored in `fixed_deposits.db` (gitignored, never committed) — leaving the password field blank on save keeps whatever is already stored, so it's never echoed back into the page.

- **Calculator** — a standalone interest calculator: pick a deposit type (cumulative / simple interest / recurring), enter the principal or monthly instalment, annual rate, and a **duration in months and days**, and it shows the maturity amount and interest earned — using the exact same math as the rest of the app. For **simple interest**, it also shows the interest payout per **month** and per **quarter** (constant every period, since simple interest doesn't compound). Recurring deposits ignore the days field (RD instalments are always whole months). **Nothing on this page is saved** — it's pure calculation via the URL's query string (`?deposit_type=simple&principal=...`), so a result is shareable/bookmarkable without touching the database.
- **Dashboard** — the first tab: whole-portfolio totals (deposits, metals, **and investments**) — total invested, current value, unrealised gain/loss with return %, deposit maturity value, and **annualised return**. Then breakdowns (each with its own Ann. Return column): **by asset class** (deposits vs metals vs investments), **by holder across all assets** (including total metal grams and investment count per holder), **deposits by holder & bank**, **metals by type**, **metals by holder & type**, **investments by ticker**, **investments by holder & ticker**, and **deposits by bank**.
- **Annualised return** — a CAGR-style figure (`(current/invested)^(365/days held) − 1`) shown per deposit, per metal holding, and as an invested-weighted blend for every group and the whole portfolio. It's a rough blend across positions with different start dates, not a true money-weighted (XIRR) return — treat it as directional. Holdings younger than 7 days show "—" rather than an exaggerated figure (a 1-day gain projected over a year would be misleading).
- **Chart** — the same figures as a chart, covering deposits, metals, and investments. Pick the **chart type** (bar or pie) with the toggle at the top:
  - **Bar** — two grouped horizontal bars per row: invested (money in / cost) vs current value. Sections: by holder (all assets), deposits by bank, metals by type, investments by ticker.
  - **Pie** — invested *or* current value, showing each holder's / bank's / metal's / ticker's share of the total, with an amount + percentage legend.

  All charts are pure inline SVG — no JavaScript or chart library. The selection is kept in the URL (`/chart?type=pie&metric=invested`).
- **Depositors** — a separate master list (`depositors` table) of people who hold deposits, each with a unique **holder ID** (customer number, PAN, etc.) and a **name**. Managed on the **Depositors** tab, which also shows each depositor's **total invested** (principal only, no interest), **current value** and **total maturity value**, with a combined total across everyone. A depositor can't be removed while any deposit references it.
- **Banks** — a separate master list (`banks` table), each with a unique **bank ID** (IFSC / branch code / any identifier) and a **name**. Managed on the **Banks** tab; a bank can't be removed while any deposit references it.
- **Metals** — record precious-metal holdings (`metals` table): metal (**Gold 24K**, **Gold 22K**, silver, platinum, palladium, other — 24K and 22K are tracked separately, each with its own market rate), an optional **depositor** (linked to the depositors list), an optional description, weight in **grams**, and the **purchase price in ₹ per gram**. Current value is driven by a separate **market-price table** (`metal_prices`), managed on its own **Market Prices** tab. A metal with no rate set falls back to each holding's purchase price (value = cost) and the Metals tab shows a banner linking to Market Prices. The Metals tab itself shows cost, current value, total return %, annualised return, and unrealised gain/loss per holding and overall.
- **Market Prices** — a separate tab (`/metal-prices`) holding the "Current market prices" panel: one live ₹/gram rate per metal, so you set the rate once and every holding of that metal revalues everywhere.
  - **Fetch Live Prices** — pulls the current spot rate for gold, silver, platinum and palladium (24K gold's rate is used to derive 22K at 22/24 purity) and converts it to ₹/gram, using [gold-api.com](https://gold-api.com) for spot prices and [frankfurter.app](https://frankfurter.app) for the USD→INR rate — no API key needed, no new Python dependency (uses the standard library). Network failures show an error and leave existing prices untouched.
  - **Manual entry** stays available in the same panel — type a price directly to override, or to set "Other" (which has no live source). Each price shows whether it's `🌐 live` or `✎ manual`, and when it was last updated.
- **Investments** — track stock and mutual fund holdings (`investments` table): ticker, a **required depositor**, shares/units, purchase price in ₹/unit, and purchase date. The **Ticker** field is a searchable dropdown covering every NSE-listed stock and every AMFI-registered mutual fund scheme (~2,500 stocks + ~38,000 schemes) — type a few letters of the name or symbol to search both at once; you can also type a custom ticker (e.g. a US stock like `AAPL`) if it isn't in the list. The stock/fund directory is fetched once per app run from the [NSE equity list](https://archives.nseindia.com) and [mfapi.in](https://www.mfapi.in) and cached to disk (`cache/`, gitignored, refreshed automatically after 30 days). Picking a suggestion also fetches its **current price live and prefills the Purchase price field** with it (already converted to ₹) — edit the value if you actually paid a different price.
  - **Stocks** are priced **live via `yfinance`** on every page view and converted to rupees if the ticker trades in USD — for Indian stocks (NSE `.NS` / BSE `.BO` suffixes) it's already in ₹, so no conversion is applied.
  - **Mutual funds** are priced by their latest **AMFI NAV via mfapi.in**, always in ₹.
  - A ticker that can't be priced (invalid symbol, no data, offline) shows "N/A" with the reason instead of crashing the page, and is left out of the portfolio totals until it prices successfully — it still shows its own row on the Investments tab. Shows cost, current value, total return %, annualised return, and unrealised gain/loss per holding and overall, same as Metals.
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