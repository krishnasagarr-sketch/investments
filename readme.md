# Fixed Deposit Manager

A personal finance web app for an Indian household — built with Flask and SQLite. It started
as a fixed/recurring deposit tracker and has grown into a broader tool: deposits (including
NRI accounts), metals, stock/mutual fund investments, retirement savings, income and expense
logging, and India-specific tax/compliance estimates (TDS, DICGC insurance, an income tax
estimate, capital gains on sold investments, an Income & Expenditure statement). It can also
read your bank mailboxes (Mail Scan), keep the details each bank has on file for each of your
accounts, store and automatically unlock password-protected statements, and — with an Anthropic API
key — read a receipt into a draft deposit. Runs as a local web app, a desktop executable
(Windows/Mac), or embedded in native Android/iOS shells — same code, same database format,
everywhere (the mailbox, unlocking and Claude features are desktop-only; see [Mobile
builds](#mobile-builds)).

All rupee amounts use Indian digit grouping (`₹12,34,567.89`). Foreign-currency (FCNR) amounts
are shown in their own currency and are **never converted to rupees** — see [FCNR / NRE / NRO
deposits](#fcnr--nre--nro-deposits) below.

## Setup

1. Folder layout (keep this exact structure):
   ```
   fd-manager/
   ├── app.py
   ├── requirements.txt
   └── templates/
       └── *.html
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

Open **http://127.0.0.1:5000**. The first visit takes you to a one-time **account setup** page
(see [Login](#login--security) below) before anything else loads.

A `fixed_deposits.db` SQLite file is created automatically on first run — your data persists
across restarts. On startup the app runs in-place schema migrations, so an older `.db` from a
previous version upgrades automatically; legacy free-text holder/bank names are promoted into
proper `depositors`/`banks` records.

> **Add at least one depositor and one bank first** (Depositors / Banks tabs) — the Add Deposit
> form needs both.

### Optional dependencies

`requirements.txt` installs everything the desktop build uses. A few features degrade gracefully
if their library is missing (checked once at startup, never crashes a page):

| Feature | Library | If missing |
|---|---|---|
| PDF export (TDS, DICGC, Income & Expenditure statement) | `fpdf2` | Export PDF button is hidden |
| Excel export (Deposits, Income & Expenditure statement) | `openpyxl` | Export Excel button is hidden |
| Live stock/mutual fund prices (Investments) | `yfinance` | Investments tab shows a banner; holdings still list, just unpriced |
| Unlocking password-protected PDF attachments | `pypdf` (+ `cryptography` for AES) | A protected PDF's unlock page says so and offers the file as saved |
| Claude features: reading a bank's password note (Haiku), drafting a deposit from a receipt and reading a bank statement's transactions (Sonnet) | `anthropic` + `ANTHROPIC_API_KEY` | Unlocking falls back to the built-in regex reader; "Create draft FD" says no key is set |

This matters most on Android/iOS builds, which install a smaller dependency set — see
[Mobile builds](#mobile-builds).

### Configuration

Nothing here is needed to run the app.

- **`ANTHROPIC_API_KEY`** enables the Claude features. Set it as an environment variable, or in a
  `.env` file next to `app.py` (in the packaged apps: in the app's data folder —
  `~/Library/Application Support/FDManager` on Mac, next to the `.exe` on Windows):
  ```
  ANTHROPIC_API_KEY=sk-ant-...
  ```
  `.env` is gitignored — keep the key out of the repository. What gets sent to Anthropic, and when, is
  listed under [Login & Security](#login--security).
- **Switches in `app.py`:** `KEEP_ORIGINAL_LOCKED_FILES` (keep the locked original when a PDF is
  unlocked — on), `SHOW_GENERATED_PASSWORDS` (testing aid — on until unlocking is signed off; see
  [Temporary testing mode](#temporary-testing-mode)), `HAIKU_MODEL` (reads password notes) and
  `EXTRACT_MODEL` (reads receipts into drafts).

## Deposits

The core object: a fixed deposit (FD) or recurring deposit (RD) held at a bank.

### Deposit types

| Type | How interest works | Maturity amount |
|---|---|---|
| **Cumulative (reinvested)** | Interest is reinvested and paid with the principal at maturity | `A = P × (1 + r/n)^(n×t)` |
| **Simple interest (payout)** | Interest is paid out periodically (monthly/quarterly) and *not* reinvested | Equal to the principal — the interest already left the deposit as it accrued |
| **Recurring deposit (RD)** | One fixed installment per month; the balance compounds quarterly | Sum of each installment `M × (1 + r/4)^(months on deposit / 3)` |

`P` = principal (or `M` = monthly installment for an RD), `r` = annual rate as a decimal,
`n` = compounding periods per year, `t` = tenure in years.

**Tenure** can be entered in **months** or **days** (cumulative and simple-interest deposits).
Day tenures convert to years on a 365-day basis. Recurring deposits are always monthly.

**Current value** (the deposit's worth *today*): simple interest accrues linearly, cumulative
compounds, an RD sums each installment paid to date compounded quarterly. A matured deposit's
current value equals its maturity amount.

**Annualised return** — a CAGR-style figure (`(current/invested)^(365/days held) − 1`), shown
per deposit and as an invested-weighted blend for every group and the whole portfolio. It's a
rough blend across positions with different start dates, not a true money-weighted (XIRR)
return. Holdings younger than 7 days show "—" rather than an exaggerated figure.

### FCNR / NRE / NRO deposits

Each deposit has an **Account Category**: Resident (default), NRE, NRO, or FCNR — with real tax
treatment differences, not just a label:

- **NRE** and **FCNR** interest is exempt from Indian income tax and TDS entirely (Section
  10(4)), as long as NRI status is maintained — excluded from the TDS tab and the Tax
  estimate's taxable income.
- **NRO** interest is taxable, but at the NRI rate under Section 195 — **31.2%** (30% + 4%
  cess; a surcharge may also apply above certain income levels, not modelled here) from the
  **first rupee**, not the resident 10%/₹40,000-threshold rules. Shown in its own section on
  the TDS tab, since even the same depositor+bank pair can hold both a Resident and an NRO
  deposit.
- **FCNR** is held in a foreign **currency** (USD, GBP, EUR, AUD, CAD, SGD, CHF, JPY, or HKD)
  from start to finish. This app does no live currency conversion anywhere, so FCNR deposits
  are **excluded from every rupee-denominated total** (Dashboard, DICGC, Tags, Income &
  Expenditure) and shown separately, grouped by their own currency, instead of being silently
  mixed into a rupee figure.

The Tax estimate flags when a depositor holds any NRI account, since the Section 87A rebate it
shows doesn't apply to non-residents.

### Other deposit fields

- **Owned by** — optional; only set this if the money actually belongs to someone other than
  the depositor it's held under (e.g. deposited in a parent's name but really the child's
  money). The Dashboard's **By owner** table classifies every deposit by this — falling back to
  the holder when it's blank — separately from **By holder**, which always groups by who the
  bank has on record. Tax/TDS/DICGC always use the **holder**, never the owner.
- **Deposit / Account Number** — optional, free text — the FD/RD number on the bank's own
  receipt.
- **Tag** — optional, see [Tags](#tags) below.
- **Remarks** — optional, free text for any note worth keeping alongside the deposit (e.g. "kept
  in bank locker", "for daughter's wedding"). Shown on the Deposits tab, History, and the Excel
  export. Every other holding type (Metals, Investments, Retirement accounts) has the same field.

### Bulk CSV import

The **Import** tab moves in a batch of existing deposits from a spreadsheet instead of
one-at-a-time entry. Download the CSV template, fill in a row per deposit, upload it back.
Depositors and banks are matched by name (case-insensitive) or created automatically. Invalid
rows are skipped individually with a specific reason and row number, rather than failing the
whole import.

### Reinvesting, closing, and History

Every active deposit's row on the Deposits tab shows a closure link alongside Edit and
Remove — **Reinvest** (green) once it's matured, or **Close early** (amber) before maturity, for
a premature withdrawal. Both open the same page, summarising what the deposit is worth right now
(principal, interest, and either the maturity amount or, before maturity, today's current value)
and asking what to do next:

- **Reinvest — full amount**, **principal only**, or **interest only**
- **Reinvest — modified amount** — a custom figure you've already worked out, for topping up
  with fresh money or reinvesting a specific amount
- **Withdraw everything** — close it out with no reinvestment
- **Partial withdrawal** — a different kind of option, described below: it doesn't close
  anything

Closing before maturity shows a reminder that banks often pay a lower "penal" rate on premature
withdrawal — the app has no bank-specific penalty terms, so the figure shown assumes the full
contracted rate up to today.

Choosing any "Reinvest"/"Withdraw everything" option opens a form for the new deposit, prefilled
from the old one's depositor, owner, bank, tag, account category, currency, rate, tenure, and
compounding frequency — all editable before saving. A matured RD reinvests as a fresh FD, not a
new RD, since it pays out as a lump sum. The old deposit is then **closed**, not deleted: it
disappears from the Deposits tab, the Dashboard/Chart, DICGC, Tags, and the Excel export, but is
kept on the **History** tab — with its original principal, rate, start date, a "Closed early"
badge if it didn't run to term, its interest/value *as of the actual closure date* (not today, so
an early closure doesn't look like it kept accruing after it stopped existing), its closure type,
and (if reinvested) a link to the new deposit it became. A **Reopen** button on History undoes a
closure if it was done by mistake.

**Partial withdrawal** is different in kind from every option above: it doesn't close the
deposit or create a new one. Say how much to take out, and that exact deposit — same id, same
start date, same rate/tenure/bank — just keeps running with a smaller balance from today onward;
its maturity amount and current value both drop to match, and nothing shows up on History. Under
the hood this only logs the withdrawal's date and amount, rather than editing the deposit's
principal directly: interest already attributed to periods *before* the withdrawal (already
reported on TDS, the Tax estimate, or the Income & Expenditure statement) isn't touched, only
interest from that date forward is computed on the reduced balance. Not available for a
recurring deposit, whose principal is a monthly installment rather than a lump sum. The Dashboard
shows a "withdrawn" note under the deposit's principal whenever this has happened.

Closing a deposit never erases its tax history: **TDS, the Tax estimate, and the Income &
Expenditure statement** compute a deposit's interest from its own start/maturity dates for
whichever financial year they fall in, regardless of whether it's currently open or closed — so
a deposit that was closed (at maturity or early) partway through a year still counts fully for
that year.

## Dashboard & Chart

- **Deposits tab** (`/`) — the deposits list and its own totals, plus a **By owner** breakdown
  and a **Foreign currency deposits (FCNR)** table for anything not counted in the rupee totals.
- **Dashboard tab** (`/summary`) — whole-portfolio totals across deposits, metals, *and*
  investments: total invested, current value, unrealised gain/loss, deposit maturity value, and
  annualised return. Breakdowns: by asset class, by holder (all assets combined), deposits by
  holder & bank, metals by type, metals by holder & type, investments by ticker, investments by
  holder & ticker, deposits by bank. FCNR deposits get their own by-currency table here too,
  excluded from every total on the page for the same reason as above.
- **Chart tab** — the same figures as inline SVG bar or pie charts (no JS charting library):
  invested vs current value, by holder / bank / metal / ticker. FCNR deposits are excluded from
  charts for the same currency-mixing reason.

## Depositors & Banks

Separate master lists. Each **depositor** has a holder ID (customer number, PAN, etc.), a
name and — kept once per person, here — an optional **PAN** and **date of birth** (Edit changes them; every
bank account of theirs uses them, see below); each **bank** has a bank ID (IFSC/branch code/anything) and a name. Deposits, metals, and
investments all link to a depositor by reference, so renaming one updates everywhere it's used.
A depositor or bank can't be removed while anything still references it.

## Bank Accounts

The **Bank Accounts** tab keeps what each bank has on file for one of your accounts: first name,
last name, customer ID, and the email address (with its app password) the bank
writes to — plus the bank, the depositor and an account number/label. It's per account
because the name (and the rest) can differ from bank to bank. The holder's **PAN and date of birth** aren't
asked per account: they're the depositor's (set once on the Depositors tab, and a depositor must be chosen
for each account), and are used with the account's details to:

- **Rebuild a statement's password without typing** — see [Opening a protected
  PDF](#opening-a-protected-pdf). On the unlock page a **Saved bank account** picker fills in the
  details: a deposit's attachment pre-selects the account at its bank (narrowed by its depositor, and
  by an account label that appears in the deposit number); an investment's attachment, the holder's
  own accounts; a retirement account's, the account at the bank its institution names (else the
  holder's accounts); a Mail Scan attachment uses the account its email was sorted into, or matches by
  the sender's bank; a bank account's own Statements use that account. Where several accounts match, **Try every matching account** builds the password from each
  in turn, using whatever you've typed for any detail an account hasn't saved (accounts still missing
  one are skipped and listed).
- **Scan each account's mailbox** — [Mail Scan](#mail-scan--draft-deposits) reads every account that
  has an email and app password saved, and sorts each email into the account it belongs to.
- **Keep the account's bank statements** — a **Statements** link per account opens the same
  attachment page as every other holding (password note, document date, automatic unlock, kept
  original, delete). An account that still has statements can't be deleted.

The customer ID is masked in the list (Edit shows it in full); a saved app password is never
shown again on any page. Everything is stored as plain text in the local database — like the
Notifications app password — and so in its backups; see [Login & Security](#login--security).
Rediffmail has no app passwords and a free account can't be read at all — see [Mail
Scan](#mail-scan--draft-deposits).

## Tags

Cross-cutting purpose/allocation labels (Emergency Fund, Tax-saving, Retirement, Kids'
Education, …) attachable to any deposit, metal holding, investment, or PPF/EPF/NPS account —
one tag per holding. The **Tags** tab shows an allocation view (current value and % of
portfolio per tag, plus an "Untagged" bucket) spanning all four holding types, which none of
the per-tab totals do on their own. This is separate from **Owned by** — a tag is about
*purpose*, ownership is about *whose money it is*.

## Attachments

Another cross-cutting feature alongside Tags: every deposit, metal holding, investment,
PPF/EPF/NPS account, income tax record ([Tax Filings](#tax-filings)) and — as **Statements**, on the
[Bank Accounts](#bank-accounts) tab — every bank account has an **Attachments** link (with a count once anything's attached) for keeping the
contract note, allotment advice, demat statement, deposit receipt, bank statement or a certificate
scan against it for future reference. Accepts PDF, JPEG/PNG, and Word documents up to 20 MB each,
several per upload (they share the password note and document date). A PDF opens inline in a new
tab, other types download; any file can be deleted (its kept original goes with it). Stored on disk,
one folder per holding, under `deposit_attachments/`, `metal_attachments/`, `investment_attachments/`,
`retirement_attachments/`, `bank_account_attachments/` or `tax_attachments/` next to the database (and
`mail_attachments/` for [Mail Scan](#mail-scan--draft-deposits)) — all gitignored, since these are real personal
financial documents. Files are stored, not analysed — except when you ask: a protected PDF is
unlocked when you open it ([below](#opening-a-protected-pdf)), and a Mail Scan PDF can be turned into
a draft deposit ([see Mail Scan](#create-draft-fd-from-an-attachment)).

### Document date

Each attachment can also carry a **document date** — the date the document itself bears (a
statement's month-end, a receipt's date): set it when uploading (it applies to every file chosen
together), or change or clear it per file next to the password note, on every attachments page
(deposits, metals, investments, retirement, a bank account's Statements) and under Mail Scan's
Saved attachments. It's shown beside the file name, files with a date are listed newest first
(undated ones after, by name), it travels with the file when it's moved (to a deposit, or to a
bank account's Statements), and it must be a real date between 1990 and a year from now.

**Auto-filled from the mailbox.** When Mail Scan saves an attachment it gives it a document date
itself — preferring the date the document is *about* over the day it was emailed: the end of a
statement period named in the subject or text ("from September 01, 2026 to September 30, 2026"), an
"as on"/"ended" date, the month named ("Statement for September-2026" → its last day), a date in the
file name (`…30092026…`, `…2026MTH09…`, `…20260930…`), and only then the email's own date. Dates in the
future, older than ~3 years or in reversed ranges are ignored. Each carries where it came from (shown
as "(auto)", with the reason as a tooltip), a date you type is marked "entered by you" and is never
overwritten, and saving a note without touching the date doesn't change that. **Fill in the dates** on
Mail Scan handles what was saved earlier — Mail Scan files from their email, and files on any other
page (e.g. statements already moved to a bank account, which no longer have their email) from their
file names alone.

### Password note

Many bank-issued PDFs are password-protected (a PAN number, a date of birth, an account number),
so each attachment has its own optional **password note** — set it at upload time or edit it later
independently of the file itself. Put in the bank's own wording ("first four letters of your name
in capitals followed by your date of birth as DDMM"), or the real password: either works. Mail Scan
fills it in from the email automatically when the email says how to open the file.

### Opening a protected PDF

Clicking **View** on a password-protected PDF doesn't just hand it to your browser's viewer to
ask for a password it can't work out — the app reads the password note and builds the password
for you. A bank's note is usually a *recipe*, so the unlock page shows how it read the note (here:
1. first 4 letters of name in CAPITALS, 2. date of birth as DDMM) and asks for just the details
that recipe needs — name, date of birth, PAN, registered mobile number, account number,
customer ID/CIF, folio or Aadhaar — in the order the note gives them. Type them in, and it
assembles the password, decrypts the PDF, and opens it.

- Understands "first/last N letters/digits of …", date formats written as `DDMMYYYY`, `DDMM`,
  `YYMM…` and so on (or "year of birth"), "in capitals"/"in lowercase" per part, and a password
  stated outright in the note. Where the note leaves case or date format unstated it tries the
  likely variants (e.g. a name in upper, lower and as typed; `DDMMYYYY` then `DDMMYY`).
- If the note can't be read, or the guess is wrong, it says so and takes the password typed
  directly; **Open the file as saved** hands over the untouched file for your own PDF viewer.
- The details you type are used for that one request and **never stored**, and the app never
  stores a password it finds. A PDF that's only "encrypted" to restrict printing opens straight away
  and is left as it is.
- PDFs only (a protected Word/Excel file isn't unlocked, just shown with its note). Needs `pypdf`,
  and `cryptography` for AES-encrypted PDFs, which is most bank statements — see [Optional
  dependencies](#optional-dependencies). Works identically for [Mail Scan](#mail-scan--draft-deposits)'s
  saved attachments and a bank account's Statements.

**The password comes off the file; the original is kept.** **View** tries, in order: no password,
then a **regenerated** one — the note is read as described below, the details come from the saved
bank account(s) matching the file (a deposit's bank/depositor; an investment's or retirement
account's holder — a retirement account's institution also picks the account at that bank; the
account the email was sorted into; a bank account's own statements use that account), and every candidate is tried, all without typing. If that fails you land on the
unlock page, which says why and takes the details (or the password) by hand. Whatever opens the
file is then **removed from the saved file for good**: an unlocked copy replaces it, so from then
on it just opens, in any viewer, with nothing stored. The swap only happens after a check — the
copy must open with no password, have the same page count, and carry the same text on every page
(the first 200) as the original; if not, the saved file is left untouched and the reason is shown.
The **untouched original, still locked, exactly as the bank sent it, is kept** in an `_originals/`
folder beside it (switch off with `KEEP_ORIGINAL_LOCKED_FILES = False` in `app.py`), and each
file in the attachment lists gets a **View original** link plus a 🔓 line saying what was
verified. Text and page count are what's compared; images, layout and digital signatures aren't
(a signed original is flagged — the copy no longer carries the signature), which is what the
kept original is for. Deleting an attachment deletes its original too, and re-saving the same
email attachment later doesn't create a duplicate.

### Reading the instructions with Claude Haiku

With an API key set (see [Configuration](#configuration)), the unlock page sends the bank's email text (Mail Scan keeps up to 8,000 characters of the
body of any email that carried an attachment; other attachments use their saved password note) to
Claude Haiku with a JSON-schema structured output and gets back, in order, which details make
up the password — `{"fields": [{"type": "date_of_birth", "date_format": "DDMM"},
{"type": "first_name", "start_index": 0, "end_index": 3}]}`. A `date_of_birth` entry carries the
layout to write the date in (`DDMMYYYY`, `DDMMYY`, `DDMM`, `YYYY`, `MMYYYY`, …); `first_name`,
`last_name`, `pan` and `customer_id` entries are 0-based slices, end exclusive (a negative start counts from the
end, so "last 3 letters" is `-3..99`). Haiku only ever sees that email text — **never the
password and never the name/date of birth/PAN you type in**; plain Python formats and slices
those details per the JSON, joins them, and tries the result on the
PDF. Capitalisation isn't in the JSON: "capital letters"/"lowercase" in the text decides it,
otherwise upper, lower and as-typed are all tried. Without a key, the `anthropic` package, or a
working connection, the built-in regex reader takes over and the page says why.

### Temporary testing mode

While unlocking is being tested, `SHOW_GENERATED_PASSWORDS = True`
in `app.py` makes the unlock page *display* the password(s) it built from your details (in the
order tried, with the winner marked), offers a "show the generated password(s) only — don't
open" preview, and waits for an **Open the PDF now** click after a successful unlock (by then the saved file is already unlocked). This
puts a real secret on screen, so set it to `False` (or delete the flag) once signed off — the
page then goes back to unlocking and opening in one step with nothing shown.

## Metals & Market Prices

Record precious-metal holdings: metal (Gold 24K, Gold 22K, silver, platinum, palladium, or
other — 24K and 22K tracked separately, each with its own market rate), an optional depositor
and tag, weight in grams, purchase price per gram, and optional free-text **description** (e.g.
"22K coin, sovereign, bar") and **remarks** for any other note. Current value comes from a
shared **Market Prices** table (one live ₹/gram rate per metal), so setting the rate once
revalues every holding of that metal everywhere.

- **Fetch Live Prices** — gold and silver use the **IBJA (India Bullion & Jewellers
  Association)** daily reference rate, the same benchmark Indian jewellers and banks price
  against — it already includes import duty, GST, and the local market premium, so it reads
  higher than a plain international spot conversion (that's expected — it's what makes it the
  *Indian* market price). Platinum and palladium aren't published by IBJA, so those use the
  global spot rate converted at the live USD→INR rate. IBJA only publishes on business days;
  on a weekend or holiday, gold and silver fall back to IBJA's **last published rate**, and the
  page says so.
- **Manual entry** stays available — type a price directly to override, or to set "Other"
  (which has no live source).

## Investments

Track stock and mutual fund holdings: ticker (searchable across NSE-listed stocks and
AMFI-registered mutual fund schemes, or type a custom ticker like a US stock), a required
depositor, optional tag, shares/units, purchase price, purchase date, and optional **remarks**
(e.g. "long-term hold", "tax-loss harvest candidate").

- **Stocks** are priced live via `yfinance` on every page view, converted to rupees if quoted
  in USD.
- **Mutual funds** are priced by their latest AMFI NAV.
- A ticker that can't be priced shows "N/A" with the reason and is left out of portfolio totals
  until it prices successfully.

### Selling a holding, and Capital Gains

Each holding has a **Sell** link alongside Edit and Remove — enter how many shares (up to what's
still held) and at what price (prefilled from the live quote), and it's logged as a sale rather
than deleting or editing the holding. The original purchase row never changes: it stays the cost
basis for whatever's still held, and "Shares" on the Investments tab shows the *remaining*
amount after any sales, with a small "X sold" note. Sell everything and the holding drops off the
Investments tab entirely — it's fully realised, and its purchase record lives on as history for
the **Capital Gains** tab (which reads it directly, so a holding with recorded sales can't be
removed — the app refuses and says why, since deleting it would erase that history).

The **Capital Gains** tab shows, financial-year by financial year and grouped by depositor, the
realised long-term and short-term gains from every sale, an estimated tax, and a line-by-line
detail table (each with a Delete action, which un-sells those shares back onto the Investments
tab). It assumes every holding here is a listed equity share or equity-oriented mutual fund
taxed under Sections 111A (short-term, more than 12 months holding period is long-term) / 112A
(long-term) — 20% flat for short-term, 12.5% above a ₹1,25,000-per-person-per-year exemption for
long-term (10%/₹1,00,000 for sales before 23 Jul 2024, when Budget 2024 changed both), plus a 4%
cess; surcharge isn't modelled. **Not modelled at all:** debt mutual funds (no LTCG treatment if
bought on/after 1 Apr 2023 — taxed at slab rate instead), foreign shares (24-month LTCG
threshold, different rates), and pre-31-Jan-2018 grandfathering — this app can't reliably tell
those apart from a ticker string, so treat the estimate as equity-only. Losses aren't carried
forward to a later year, and (a narrow edge case specific to the transition year) a loss on one
side of 23 Jul 2024 isn't netted against a gain on the other side within FY2024-25.

An investment drops off the Investments tab once every share is sold, so each sale row on the Capital
Gains tab links to that holding's **Attachments** (contract notes, statements) — they stay reachable.
A holding that still has attachments can't be deleted.

## Calculator

A standalone interest calculator — pick a deposit type, enter principal/installment, rate, and
duration, and see the maturity amount and interest earned, using the same math as the rest of
the app. Nothing here is saved; the result lives in the URL's query string, so it's
shareable/bookmarkable.

## Notifications

Email reminders, sent via Gmail SMTP (an app password, not your normal password — generate one
at [myaccount.google.com/apppasswords](https://myaccount.google.com/apppasswords) after
enabling 2-Step Verification). A background thread checks every **12 hours** (and once at
startup) while the app is running.

- **Maturity alerts** — email when a deposit is within a configurable number of days of
  maturing (default 30).
- **Contribution reminders** — email when a PPF/EPF/NPS account has had **no contribution
  logged** for a configurable number of days (default 45), counted from the most recent
  contribution or the account's opened date.

Both share one resend cooldown (**7 days** — an already-alerted item isn't re-emailed sooner
even if checked again) and are bundled into digest emails. **Send Test Email** verifies SMTP
settings without touching any state; **Check & Send Now** runs the real check on demand.

## Other Income

Log dated income entries per depositor under **Salary, Rent, Business, Capital Gains, or
Other** — anything besides FD/RD interest, which is computed automatically from the deposits
themselves. Feeds into the Tax estimate and the Income & Expenditure statement.

## Expenses

Log dated spending per depositor under **Household, Medical, Education, Travel, Utilities,
Insurance, or Other**, with a running this-financial-year total and category breakdown shown on
the page. Feeds into the Income & Expenditure statement.

## Gifts

Log money given between two depositors already tracked here — i.e. *within the family*, not
to/from an outside party (from, to, date, amount, note). India exempts gifts between specified
relatives (spouse, siblings, parents, in-laws, lineal ascendants/descendants) from gift tax
regardless of amount; this doesn't determine who counts as a relative, so use your own
judgement. Shown as "for information" on the Income & Expenditure statement, not counted as
income or expenditure.

## Income & Expenditure Statement

Pick a financial year (labelled with its matching assessment year, e.g. **FY 2025-26 (AY
2026-27)**) and optionally a single depositor, and get:

- **Income** — deposit interest on an **accrual basis** (what the deposits earned within the
  year, whether or not it was paid out — same convention as TDS and Tax), plus the Other Income
  log by category. NRE interest is listed but marked tax-exempt and excluded from the taxable
  figure; NRO is included as taxable.
- **Expenditure** — the Expenses log by category.
- **Surplus or deficit**, and how much of the income is taxable.
- **For information** — retirement contributions and family gifts (they move money without
  being income or expenditure), and FCNR interest in its own currency.

Exports to PDF and Excel. The current (in-progress) year shows figures to date.

Other Income and Expenses entries can also be generated from your bank statements — see [Income &
Expenditure from bank statements](#income--expenditure-from-bank-statements).

## Tax estimate

Pick a depositor and financial year for a quick estimate — FD/RD interest (excluding
tax-exempt NRE/FCNR) plus Other Income — against **both** tax regimes side by side. Section 80C
(from PPF contributions, capped at ₹1,50,000) and Section 80CCD(1B) (from NPS contributions,
capped at ₹50,000) are filled in automatically from what's logged on the Retirement tab
(Old Regime only, per current rules), and it calls out which regime is cheaper. Flags when the
depositor holds any NRI account, since the Section 87A rebate shown doesn't apply to
non-residents. This is an estimate from what's tracked here — not a substitute for filing
software or a CA; it doesn't know about salary TDS already deducted or other deductions.

## Tax Filings

The **Tax Filings** tab keeps your income tax paperwork by assessment year: the **returns you filed**
(ITR form, original/revised/belated/updated, date filed, acknowledgement number, status, refund or
tax payable) and every **communication from the Income Tax Department** (notice, intimation u/s
143(1), demand, refund advice, order, rectification, your own reply — with section, reference/DIN,
date received, respond-by date, status and amount). Records can be filtered by taxpayer (a
depositor) and assessment year, and are grouped under each year.

- **Open notices are hard to miss.** An open communication with a respond-by date is flagged *Due
  within 15 days* or *Overdue*, open ones sort first (soonest deadline first), and the nav shows
  `Tax Filings (n)` for the number still open — mark one *Response filed* or *Closed* (Edit) and it
  drops out of the count.
- **Documents on every record.** Each return and communication has a **Documents** page — the same
  attachment page as everywhere else (several files at once, document date, password note, View
  original, delete): the ITR-V and computation, Form 26AS / AIS, a notice and your reply. Income Tax
  PDFs are locked with your **PAN in lowercase followed by your date of birth (DDMMYYYY)**, so that
  note is pre-filled, and a taxpayer whose [depositor](#depositors--banks) has the PAN and date of
  birth saved gets them opened automatically. Documents are stored in `tax_attachments/` (gitignored).
  A record that still has documents can't be deleted.
- **From Mail Scan.** Email from the department (`…incometax…` addresses, including one you forwarded
  from another mailbox) has its attachments saved like a bank's. On Mail Scan, such an email's **Belongs to**
  picker comes pre-set to **➕ New tax communication from this email** (it's never tied to a bank account); press **Move files** to create a communication record from it in one
  click — type, section, DIN, assessment year, respond-by date (an "on or before" date, or "within N
  days") and any amount are read from the text, the taxpayer comes from the bank account the email
  was sorted into — and move the file into it; or pick an existing record to add more documents to.
  Check the guessed fields on the Tax Filings tab.

## TDS

Estimates TDS (tax deducted at source) on FD interest, split by the rules that actually apply:

- **Resident deposits** — 10% (20% without PAN on file) once a depositor's total interest from
  one bank in a financial year crosses **₹40,000** (₹50,000 for senior citizens) — a
  configurable threshold on the page.
- **NRO deposits** — their own section, **31.2%** flat from the first rupee, no threshold (see
  [FCNR / NRE / NRO deposits](#fcnr--nre--nro-deposits)).
- **NRE and FCNR** interest doesn't appear here at all — it's tax-exempt.

Exports to PDF.

## DICGC

DICGC insures deposits up to **₹5,00,000 per depositor, per bank** — covering principal and
accrued interest together across every rupee deposit that depositor holds there. This tab
groups deposits the same way and flags anything over the limit. FCNR deposits are excluded from
this rupee total (DICGC does insure them too, converted at claim time, but this app doesn't do
that conversion) — the page notes how many are excluded. Exports to PDF.

## Interest Check

Verifies that a payout ("simple interest") FD's interest actually landed in the bank account as
expected. Pick a depositor + bank, upload that account's bank statement (CSV — a single signed
amount column or separate debit/credit columns are both auto-detected, along with day-first
DD/MM/YYYY dates), and assign each imported credit to the deposit it belongs to. Compares what's
been received against the linear interest the deposit should have accrued to date, flagging any
deposit running short. Statement lines are scoped to a depositor+bank pair, since several FDs
at the same bank pay into the same account.

## Mail Scan & Draft Deposits

Reads the Gmail inbox already set up on the Notifications tab (same App Password, read over
IMAP instead of used to send) — and, **one by one, the mailbox of every [bank account](#bank-accounts)
that has an email and app password saved** — and keeps the **attachments** (PDF, CSV, Excel) of emails from
a **recognised sender**. The text of an email is *not* mined for transactions: an alert's wording varies
too much between banks and says too little (an amount and a date, never the rate or tenure), so interest
credits and FD bookings are no longer guessed from it — the authoritative sources are the documents
themselves (a deposit receipt becomes a draft deposit, below) and, for the account's day-to-day
transactions, its statements.

A sender is "recognised" if it's a bank; the Income Tax Department (see [Tax Filings](#tax-filings));
an NPS / EPF record-keeper (Protean, formerly NSDL e-Gov, and EPFO — see
[Retirement](#retirement-ppf--epf--nps)); or anyone you add under **Recognised senders** (below).
Banks and the rest are matched by a keyword (e.g. `sbi`, `hdfcbank`, `equitas`) against each
dot-separated label of the sender's domain, not the domain as a whole — real bank transactional mail
routinely comes from a dedicated ESP/sub-brand domain (`bounce-zem.equitas.bank.in`,
`alerts.sbi.bank.in`) that looks nothing like the bank's own website. An unrecognised sender's attachment
is never saved. Saved attachments go one subfolder per email, named after its Message-ID, under
`mail_attachments/` next to the database. The Mail Scan tab shows how many have been saved in total and
exactly where, plus a **Saved attachments** list grouped by the email each one came from (subject,
sender, date) with a **View** link per file — a PDF opens inline in a new tab, a CSV/Excel file
downloads, the same as either would from any other site.

How a scan stays cheap and duplicate-free: each distinct address is scanned **once** (compared
case-insensitively; an account whose email is the Notifications one just shares that scan), and
within a scan only each email's Message-ID *header* is fetched first, in batches — mail already
handled is skipped without downloading it, so a repeat scan downloads nothing new. An email's
Message-ID is its identity everywhere, so the same email reaching two mailboxes is one email (read
from the account's own mailbox, which is scanned before the Notifications one). One mailbox failing — a stale app password, say — is reported
without stopping the others. The host comes from the address (Gmail by default; Outlook/Hotmail,
Yahoo, iCloud and Rediffmail Pro addresses use their own IMAP servers). **Rediffmail
(`@rediffmail.com`) has no IMAP** (only Rediffmail Pro does), so it's read over **POP3**
(`pop.rediffmail.com:995`) instead — **but Rediff sells POP3 access as a paid feature, and a free
account is refused with "Login Not Allowed"** (the app says so rather than blaming the password).
For a free account, the realistic options are: ask the bank to use a Gmail address, attach statements
by hand (several files at once, on the account's **Statements** page), or pay for POP3 access / Rediffmail Pro
(free accounts reportedly can't auto-forward either; Gmail's own "check other accounts" would need POP3
too). Clearing the account's password stops the scan trying that mailbox.
POP3 can't search by date, so the newest emails are walked back from the end reading only their
headers (`TOP`) until they're older than the window, and only new ones are downloaded — everything
after that (duplicates, attachments, account matching) is the same. It signs in with the full
address and, if that's rejected, the short form (the part before the `@`) — two attempts at most.
Nothing is ever deleted (`DELE` is never sent), but Rediffmail has **no app passwords**, so the
account's own mailbox password goes in the bank account's password field (plain text, like the
others), and "keep a copy on the server" should be ticked so downloading doesn't remove mail.
Only Gmail has been exercised for real; the POP3 and other-provider paths are tested against
stand-in servers.

**Recognised senders — add your own.** The **Recognised senders** card on the Mail Scan tab lists who
Mail Scan saves attachments for, beyond the built-in banks, Income Tax Department and NPS/EPF
record-keepers (open *Built-in senders* to see them). Add a mutual-fund registrar, an insurer or an
advisor as a **full address** (`cas@kfintech.com` — exact match), a **domain** (`kfintech.com` — it
and its subdomains) or a **word** from the sender's domain (`kfintech` — any domain with that whole
label, like the built-in bank keywords), with an optional note; **Remove** stops it (attachments
already saved stay). Public mail providers (`gmail.com`, `yahoo.com`, `rediffmail.com`…) and generic
words are refused as domains or words, since that would save every friend's attachment — a specific
address at Gmail is fine. Adding a sender makes the next scan look again at emails from it that were
skipped earlier, and a forwarded email counts by its *original* sender.

**A password note for a sender.** Some mail gives no password instruction as text — Protean's NPS statements
arrive as 16 pictures and the word "proteantech", nothing else. Give such a sender a **password note** (a field on the
add form, and editable per row): *"Password is your date of birth in DDMMYYYY format"*. Every saved file from that
sender that has no note yet gets it straight away, and so does each new one, but only when the email itself says
nothing — an instruction found in an email always wins, and a note you typed on a file is never replaced. To set one for
a **built-in** sender, add its name (`proteantech`) with the note. Unlocking reads the note when the email's text
yields no recipe. An NPS statement's file name starts with the PRAN, so if the account number you saved on the
Retirement tab matches it, the statement is attributed to that holder (and the unlock page offers their details).

**Forwarded bank emails.** A bank email you forward to the scanned mailbox (say from a Rediffmail you
can't scan directly) arrives *from you*, so it used to look like personal mail and its PDF was skipped.
For a subject starting `Fw:`/`Fwd:`, the original sender is now read from the forwarded header block at
the top of the text (`From: …`, as Rediffmail, Gmail and Outlook write it) and used for bank
recognition, attachment saving and account matching — that sender is what's shown as the email's
sender. Newer bank addresses of the form `name.bank.in` (`icici.bank.in`, `hdfcbank.bank.in`) are
recognised too. A forward with no such block, or an ordinary email from you, is left alone. Forwards
scanned before this existed get one fresh look on the next scan.

**Which bank account an email belongs to.** For each email with a saved attachment the sender's
bank narrows the candidates; if that leaves one account, that's it. Otherwise the email's text is
checked against each candidate's account number/label (also masked, like `XXXX1234` or "ending
1234"), customer ID, PAN (also masked, like `abcXXXX4f` in a file name), first and last name
(tolerating a split such as "Krishna Sagar" for `krishnasagar`, and an initial as in "R Krishna
Sagar"), and depositor name. The email's subject, text **and the names of its attached files** all
count as evidence. An account wins only on real evidence with a clear lead over the runner-up — an
unrecognised sender is only assumed to belong to the sole account of an account's own mailbox. For
mail that isn't a bank's (the Income Tax Department, NPS/EPF), one person's several accounts share
the PAN and name, so which one is beside the point: if the evidence points at a single holder, the
mail goes to that holder's first account (and says so); evidence split between different holders
is still left for you. Every scan also re-tries mails that are still unsorted. Whatever the matcher
can't settle is left **unsorted**: under *Saved attachments* each email has one **Belongs to** picker (a bank
account, an income-tax filing, a retirement account or an investment; *None of these*, or *Work it out again* to hand
it back to the matcher). **Save** ties the email to the chosen *bank account* (a manual choice is remembered and
never overridden); **Move files** moves the email's files onto whatever is chosen (see below). There's no separate "match" button: the matcher runs as each new email is scanned, and every
scan re-tries the emails still unsorted (so after adding a bank account, the next scan picks up what it
now fits). The unlock page pre-selects the account an email was sorted into.

**Tax and NPS/EPF mail isn't a bank's.** An Income Tax Department email is sorted to **➕ New tax
communication**, and an NPS or EPF statement to your NPS/EPF account (the holder's, judged from the PAN/name
in the email and its file names, when you have more than one) — shown in the picker with the reason, never as a
bank account. Press **Move files** and the record is created (the taxpayer judged the same way) or the file lands on
the account. Unlocking such a file still uses the holder's saved details.

If the email itself says how to open a password-protected attachment — banks routinely spell
this out ("the password is your PAN in capital letters") — that instruction is picked up
automatically and shown under the file. The email is cut into short pieces (lines, bullets, sentences) and
the piece that says what the password is made of is chosen, together with the neighbouring pieces that
complete it (the tax department's instruction plus its worked example, a note that it's in capitals); warnings
such as "we will never ask for your password", log-in help and the *Non-individual* variant are passed over,
and the *Individual* one is preferred. It's shown in the same password field every holding's own
Attachments page has (see [Attachments](#attachments) below); edit it if the automatic guess was
wrong or incomplete. It's only ever set automatically when nothing's noted yet, so a manual edit
is never overwritten by a later scan.

**Reading the email's text.** The password instruction isn't always in the plain-text part: a forwarder can
leave a cut-down one, and ICICI's mail carries it only as the *alt text of its banner image* in the HTML. So the
email's HTML part is converted too (line breaks kept, image alt text included) and is used when it says clearly
more than the plain text. Emails already saved without a password note are looked at once more by the next scan to
fill it in; a file that's already saved isn't saved twice, and a note you typed is never replaced.

An email is normally only ever looked at once, tracked by its Message-ID — but if it was scanned
before attachment-saving existed at all, the very next scan gives it exactly one further check
for an attachment it never got the chance to be considered for, with no need to do anything by
hand. A **Reset scan history** button on the tab is also there for a clean slate anytime — it
clears what's been "looked at", so the next scan re-examines everything in the window again;
attachments already saved, moved or deleted aren't saved a second time (they're remembered by the hash
of the file as received).

Nothing here becomes a real record without you: a deposit receipt is only ever turned into a
**draft** deposit that you check and approve (see [Create draft FD from an
attachment](#create-draft-fd-from-an-attachment)). A "Draft Deposits (N)" badge in the nav shows how many
are waiting. (Earlier versions also queued "interest credited" emails for review and drafted deposits from
the text of "FD booked" alerts; both were dropped — their data was too thin to trust. Rows they left in the
database are simply not shown, and a draft made that way is still there to approve or reject.)

### Working with saved attachments

Each file under **Saved attachments** has these actions:

- **Move files** The email's **Belongs to** picker (grouped: bank accounts, income tax, retirement accounts,
  investments; pre-set from where the email was sorted) and its **Move files** button move the email's files — with
  its kept locked original, password note, document date and unlock status — onto that record's
  attachments: a bank account's **Statements** ([Bank Accounts](#bank-accounts) tab), an
  investment's, retirement account's or tax record's **Attachments**, where it unlocks automatically from
  the holder's saved details. Moving to a bank account also sorts an unsorted email into it. It's a real
  move: the files leave Mail Scan's list, a same-named file already there isn't overwritten, and what
  was moved is remembered (by the hash of the file as received), so scanning the same email again —
  e.g. after a history reset — doesn't bring it back.
- **Delete.** A **Delete** button (two-step confirm) removes the file, its kept locked original and
  its notes. It's refused while a pending draft deposit was made from that file (approve or reject the
  draft first); an email whose last file goes drops out of the lists and its empty folder is tidied
  away; and what was deleted is remembered by the hash of the file as received, so scanning the same
  email again (e.g. after a history reset) doesn't save it back.
- **Document date and password note** — editable per file; see [Attachments](#attachments).
- **Create draft FD** — see the next section.

### Create draft FD from an attachment

Next to every saved PDF in Mail Scan's **Saved attachments** is a **Create draft FD** button (not on a
holding's own Attachments page — that document already belongs to a record, and drafting from it
would duplicate the FD). It opens the PDF the usual way (a locked one is
unlocked first, as above), sends its text — up to 12,000 characters of the first 12 pages — to
Claude **Sonnet** (`EXTRACT_MODEL` in `app.py` — the password note above stays on the cheaper Haiku) with a JSON-schema structured output, and files what it finds as a **pending draft**
on the Draft Deposits tab: amount (or RD instalment), rate, tenure (derived from the start and
maturity dates when only those are given — in months when it lands exactly, else days), start
date, deposit type (cumulative / payout-simple / recurring), compounding, account category and
currency, and the FD number. The bank and depositor are matched by name against your Banks and
Depositors only when exactly one fits (a bank the document names that you haven't added is left
blank; an email's sorted bank account is only a fallback when the document names none). Implausible
values (a rate over 30%, a bad date) are dropped rather than guessed. Many receipts print no
rate at all — then it's **worked out from the maturity amount** (compound rate for a cumulative
deposit, interest over principal and term for a payout one; kept only if it lands between 1% and
15%) and the card says so, as an approximation to check. If neither is available the rate stays
blank. Each card has a **What was read from the document** box with the model's raw JSON, to see
exactly what it found. The text is taken with the PDF's **layout preserved** (a table's
columns stay apart — plain extraction once ran `38063` and `30000` together), and if the read still
comes back without the amount, any rate figure or any date, the PDF **itself** (up to 8 MB) is
handed to the model so it can see the page, headings and all; the box says which was used. A document that isn't an
FD/RD receipt is refused, as is one whose FD number already belongs to a deposit or a pending draft
(compared ignoring spaces and dashes — it's probably the same FD), there's one draft per file (re-clicking points to it), and the draft card
shows the PDF itself in a **viewer pane on the right** (sticky beside the form, stacked below it on
narrow screens; a locked file is unlocked first) so every field can be checked against the document
as you edit it, with an *Open in new tab* link — nothing becomes a real deposit until you check it and **Approve**. On
approval the PDF is **moved** (not copied) onto the new deposit's own attachments, with its kept
locked original, password note and unlock status, so the document stays with the FD for future
reference (it then no longer appears in Mail Scan's list); rejecting the draft leaves it where it is. Needs
`ANTHROPIC_API_KEY` and, like the password note, **sends the document's text to Anthropic** — only
when you click the button. It can misread; check every field against the document.

### Income & Expenditure from bank statements

A saved bank-statement **PDF** (statements are PDF-only) can fill the [Income & Expenditure
statement](#income--expenditure-statement) for you. Once its email is sorted into a bank account, press
**Read transactions** beside the file — on Mail Scan, or later on the account's own **Statements** page after the file has been moved there (a locked PDF is unlocked first). Claude **Sonnet** reads the statement
in page-sized pieces (up to 60 pages; it can take a minute) and every transaction lands on the **Statement
Entries** tab — with a suggested category — as **pending**. Nothing counts until you press **Add these as
chosen**:

- money **in** becomes an *Other Income* entry (Salary, Rent, Business, Savings interest or Other);
- money **out** becomes an *Expenses* entry (Household, Medical, Education, Travel, Utilities, Insurance, Other);
- **FD interest credits** go to [Interest Check](#interest-check) as statement lines (deposit interest is already
  counted in the statement from the deposits themselves, so adding it as income would count it twice);
- anything that isn't income or spending — transfers between your own accounts, FD bookings and maturities,
  investments, refunds — is **ignored**; a line left on *Decide later* stays on the tab.

The entries belong to the bank account's depositor, carry the narration as their note, and can be **undone**
from the *Recently added* list (the entry it created is removed and the line goes back to pending). The model
is told to use the running balance to tell credits from debits, and wherever the balances contradict its
choice the app follows the balance (and says how many it corrected). A statement that overlaps an earlier one
doesn't add lines twice (lines already read from the same account are matched on date, narration, amount and
balance), and **Discard these lines** throws away what's still pending from a file so it can be read again.
A categorisation is only a suggestion — look through before adding. Needs `ANTHROPIC_API_KEY` and **sends
the statement's text (account number, transactions, balances) to Anthropic**, only when you press the button.

## Retirement (PPF / EPF / NPS)

PPF, EPF, and NPS rates are government-notified and change over time, with rules (minimum
balance dates, market-linked NPS returns) this app doesn't try to reproduce — so **current
balance is entered by hand** from the account's own passbook or portal, the same pattern used
for metals' market price. Contributions are logged individually (date, amount, note) to compute
gain, and a PPF account flags when its contributions in the current financial year exceed the
**₹1,50,000** annual limit. Each account also has a **remarks** field (e.g. "employer-matched",
"nominee is spouse"), editable after creation via its own "Update remarks" box. Each account has an
**Attachments** link (statements, passbook scans) with the same automatic unlocking as everywhere
else; an account that still has attachments can't be deleted. Monthly **NPS statements** from Protean
(`…proteantech…`, also the older `cra-nsdl`) and EPFO mail are recognised by [Mail Scan](#mail-scan--draft-deposits),
which saves their PDFs; the **Belongs to** picker comes pre-set to your retirement account of that type
when there's exactly one (an NPS statement → your NPS account), and the statement's period in the
subject becomes its document date.

## Backup & Restore

**Download full backup (.zip)** is one file holding a consistent snapshot of the database *and every
saved document* — each holding's attachments, bank statements, tax documents, the kept locked
originals (`_originals/`) and Mail Scan's saved attachments. **Download the database (.db)** is the
records only, if you don't want the documents. Either contains personal financial information (and
the plain-text secrets described under [Login & Security](#login--security)) — keep it private.

**Restore** accepts either: a full backup brings back the records and the documents; a database-only
`.db` brings back just the records and leaves the documents alone. Before anything is replaced, what's
in the app now is kept — a dated folder `fd-manager-before-restore-…` next to the database (database copy
plus the old document folders) for a zip, or a `….db` copy for a database-only restore. A zip is checked
before it's unpacked: it must contain a valid FD Manager database, and any file name that is absolute, uses
`..` or backslashes, or points outside the known document folders gets the whole restore refused with
nothing changed.

### Reset — clear all data and start again

For when the app has been built and tested and the test data should be replaced with real data. The
card at the bottom of the Backup tab shows what's there now (record counts, number and size of saved
documents) and, once you give your **login password** and type `RESET`, empties every table and removes
every saved document, with ids starting again at 1 (so new records never pick up an old record's
files). **Keep these** — ticked by default, since they're settings rather than data — are your login
account, the notification settings (Gmail sender and app password) and your recognised senders;
untick any to clear it too (without the login you land on the first-run setup page). **Keep a safety
copy** (ticked by default) moves the old records and documents into a dated folder
`fd-manager-before-reset-…` beside the database — a snapshot `fixed_deposits.db` plus the old document
folders and a README on putting them back — instead of deleting them; untick it for a permanent wipe
(a second confirmation reminds you). Download a full backup first if the data matters. After a reset,
Mail Scan has forgotten what it has seen, so its next scan re-reads the mail in its window. The backup,
safety and upload files are gitignored.

## Login & Security

The whole app sits behind a single login (one account, set up once on first run). Passwords are
hashed with Werkzeug (`generate_password_hash`/`check_password_hash`, never stored in plain
text); sessions are signed and last 30 days. The signing key (`.flask_secret_key`) is generated
once on first run and gitignored — don't delete it, or every existing session is invalidated.

**Forgot your password?** emails a one-time reset link (valid 1 hour), reusing the Gmail sender
already configured on the Notifications tab.

**Secrets stored in plain text.** The login password is hashed, but this is a single-user local app,
not a vault: the Notifications Gmail app password, each depositor's PAN and date of birth, each bank account's customer ID,
email and app password (or, for Rediffmail, mailbox password), and any password notes are stored as
plain text in `fixed_deposits.db` — and so in its backups — and the API key sits in `.env`. Keep the
data folder private. A saved app password is never shown back on any page.

**What leaves your machine.** Nothing, unless you've set an API key and use a Claude feature, or you
run Mail Scan (which talks to your mail provider directly):

- *Opening a protected PDF* sends Anthropic (Haiku) the bank's **email text** — or the saved password
  note — up to 6,000 characters. Never the password, and never the name/date of birth/PAN used to build it.
- *Read transactions* sends Anthropic (Sonnet) a bank statement's **text** (up to 60 pages, in pieces) — account
  number, narrations and balances — only when you click the button.
- *Create draft FD* sends Anthropic (Sonnet) the PDF's **text** (up to 12,000 characters from the first
  12 pages) and, only if that read comes back missing the essentials, the **PDF itself** (up to 8 MB) — and
  only when you click the button. A receipt carries your name, address, PAN and account numbers.
- Both calls go straight from this app to the Anthropic API with your key; nothing is sent without a
  click or a scan you started.

## Mobile builds

The same `app.py`/`templates/` are embedded, via symlinks, into native shells so the Flask app
runs on-device with no server to reach over the network:

- **Android** — Chaquopy embeds CPython in a Kotlin app (`android/`). Its dependency set is
  smaller than desktop's for size/build-time reasons: Flask + openpyxl only. No `yfinance`
  (pulls in pandas/numpy) and no `fpdf2` (pulls in Pillow) — so Investments' live pricing and
  PDF export are desktop-only; everything else, including Excel export, works.
- **iOS** — Briefcase/Toga embeds CPython behind a WKWebView-backed WebView widget (`ios/`).
  Same reduced dependency set as Android, for the same reasons (plus MarkupSafe needing a
  locally-built pure-Python wheel, since Briefcase can't build from source for iOS — see
  `ios/pyproject.toml`).

Neither bundles `pypdf`, `cryptography` or `anthropic`, so on mobile the **PDF unlocking** and the
**Claude features** (password-note reading, Create draft FD) are unavailable — a protected PDF just
offers the file as saved. Mail Scan's mailbox reading is part of the same code but is desktop-oriented.

Both are built by GitHub Actions (`.github/workflows/build-android.yml`,
`build-ios.yml`) alongside the desktop build (`build-apps.yml`, which produces the Windows
`.exe` and Mac `.app`) whenever `app.py` or `templates/` changes on `main`.

## Data model

Everything lives in one SQLite file (`fixed_deposits.db`, gitignored). Key tables: `depositors`,
`banks`, `deposits` + `deposit_withdrawals` (partial withdrawals logged against a still-open
deposit — see [Reinvesting, closing, and History](#reinvesting-closing-and-history)), `metals` +
`metal_prices`, `investments` + `investment_sales` (realised sales logged against a holding —
see [Selling a holding, and Capital Gains](#selling-a-holding-and-capital-gains)),
`retirement_accounts` + `retirement_contributions`, `other_income`, `expenses`, `family_gifts`,
`portfolio_tags`, `interest_statement_lines` (Interest Check), `statement_entries` (lines read from bank statements, pending or decided — see above), `bank_accounts` ([Bank
Accounts](#bank-accounts)), `attachment_notes` (per-file password note, document date, unlock status),
`tax_records` ([Tax Filings](#tax-filings)), `mail_sender_rules` (your recognised senders), `moved_attachments` (hashes of files moved or deleted out of Mail Scan, so a re-scan doesn't bring
them back), `processed_emails` + `deposit_drafts` (Mail Scan & Draft Deposits — see above;
`scanned_transactions` is a leftover of the dropped email-text scanner, kept only so old drafts can still be
approved or rejected), `notification_settings`, and `auth_user`. Schema migrations run automatically on startup, so
upgrading from an older version is a normal `git pull` + restart, no manual steps.

## Notes & limitations

- Interest rate should be entered as a percentage (e.g., `6.5`), not a decimal.
- The RD formula assumes quarterly compounding and one installment at the start of each month
  (standard for Indian banks) — your bank's exact figure may differ slightly by day-count and
  rounding.
- This app does no live currency conversion — FCNR amounts are never turned into rupees, by
  design (see [FCNR / NRE / NRO deposits](#fcnr--nre--nro-deposits)).
- TDS, DICGC, Tax, and Capital Gains figures are estimates from what's tracked here, not a
  substitute for your bank's TDS certificate (Form 16A), DICGC's own records, or a CA/filing
  software — in particular, NRO's 31.2% TDS estimate doesn't model surcharge, the Tax estimate
  assumes resident rules (the Section 87A rebate doesn't apply to NRIs), and Capital Gains
  assumes every Investments-tab holding is equity (see [Selling a holding, and Capital
  Gains](#selling-a-holding-and-capital-gains) for what that leaves out).
- Premature-withdrawal penalties and auto-renewal aren't modelled — every deposit is assumed to
  run to its full tenure as entered.
- Mail Scan saves and files documents; it doesn't read transactions out of email text or parse
  statements, so interest received, income and spending aren't picked up from mail (Income &
  Expenditure uses deposit interest from the deposits themselves, plus what you enter under Other
  Income and Expenses). A deposit receipt is read into a draft only on request, and every draft is
  checked and approved by you (see [Mail Scan & Draft Deposits](#mail-scan--draft-deposits)).
- The Claude features can misread a document or a password note: a draft deposit is always a draft
  to check against the receipt shown beside it, and a password recipe it gets wrong falls back to
  asking you. They need an API key and send text to Anthropic (see [Login &
  Security](#login--security)).
- Free Rediffmail can't be read by Mail Scan (Rediff reserves POP3 for paid accounts); statements from
  it are best attached by hand to the account's Statements.
- All figures are for personal tracking only; confirm exact values with your bank/CA.
