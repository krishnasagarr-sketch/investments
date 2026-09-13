import math
import os
import secrets
import smtplib
import socket
import sqlite3
import threading
import time
from datetime import date, datetime, timedelta
from email.mime.text import MIMEText
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, g, session
from werkzeug.security import generate_password_hash, check_password_hash

try:
    import yfinance as yf
    YFINANCE_AVAILABLE = True
except ImportError:
    YFINANCE_AVAILABLE = False

app = Flask(__name__)

DB_PATH = Path(__file__).parent / "fixed_deposits.db"

# Secret key for signed session cookies — generated once and persisted next
# to the DB, so logins survive app restarts. Never committed (gitignored).
SECRET_KEY_PATH = Path(__file__).parent / ".flask_secret_key"
if SECRET_KEY_PATH.exists():
    app.secret_key = SECRET_KEY_PATH.read_text().strip()
else:
    app.secret_key = secrets.token_hex(32)
    SECRET_KEY_PATH.write_text(app.secret_key)
app.config["PERMANENT_SESSION_LIFETIME"] = timedelta(days=30)

PASSWORD_RESET_TOKEN_LIFETIME = timedelta(hours=1)
MIN_PASSWORD_LENGTH = 8

# Paths reachable without being logged in. Everything else redirects to
# /login (or /setup, if no account has been created yet).
AUTH_EXEMPT_PATHS = {"/setup", "/login", "/forgot-password"}
AUTH_EXEMPT_PREFIXES = ("/reset-password/", "/static/")

# All amounts in the app are Indian rupees.
CURRENCY_SYMBOL = "₹"  # ₹


def _indian_group(digits: str) -> str:
    """Group an integer digit-string the Indian way: 12,34,56,789."""
    if len(digits) <= 3:
        return digits
    head, tail = digits[:-3], digits[-3:]
    parts = []
    while len(head) > 2:
        parts.insert(0, head[-2:])
        head = head[:-2]
    if head:
        parts.insert(0, head)
    return ",".join(parts) + "," + tail


def format_rupees(value, decimals=2) -> str:
    """Format a number as ₹12,34,567.89 (Indian grouping)."""
    try:
        value = float(value)
    except (TypeError, ValueError):
        return str(value)
    negative = value < 0
    text = f"{abs(value):.{decimals}f}"
    int_part, _, frac_part = text.partition(".")
    out = CURRENCY_SYMBOL + _indian_group(int_part)
    if frac_part:
        out += "." + frac_part
    return ("−" + out) if negative else out


def format_pct(v, decimals=1) -> str:
    """Render a nullable signed percentage: None -> '—', else e.g. '+8.2%'."""
    if v is None:
        return "—"
    return f"{v:+.{decimals}f}%"


app.jinja_env.filters["money"] = lambda v: format_rupees(v, 2)
app.jinja_env.filters["money0"] = lambda v: format_rupees(v, 0)
app.jinja_env.filters["pct"] = format_pct
app.jinja_env.globals["CURRENCY_SYMBOL"] = CURRENCY_SYMBOL

# Supported deposit types: internal key -> human label
DEPOSIT_TYPES = {
    "cumulative": "Cumulative (reinvested)",
    "simple": "Simple interest (payout)",
    "recurring": "Recurring deposit",
}

# How a tenure figure is expressed. Recurring deposits are always monthly.
TENURE_UNITS = {"months": "Months", "days": "Days"}

DAYS_PER_YEAR = 365  # simple day-count convention used for interest on day tenures

# Precious-metal holdings: internal key -> human label
METAL_TYPES = {
    "gold_24k": "Gold 24K",
    "gold_22k": "Gold 22K",
    "silver": "Silver",
    "platinum": "Platinum",
    "palladium": "Palladium",
    "other": "Other",
}


# ---------- Database helpers ----------
def get_db():
    if "db" not in g:
        g.db = sqlite3.connect(DB_PATH)
        g.db.row_factory = sqlite3.Row
    return g.db


@app.teardown_appcontext
def close_db(exception=None):
    db = g.pop("db", None)
    if db is not None:
        db.close()


def get_auth_user(db):
    return db.execute("SELECT * FROM auth_user WHERE id = 1").fetchone()


@app.before_request
def _require_login():
    path = request.path
    if path.startswith(AUTH_EXEMPT_PREFIXES):
        return None

    db = get_db()
    user = get_auth_user(db)

    if user is None:
        if path != "/setup":
            return redirect(url_for("setup_page"))
        return None

    if path == "/setup":
        return redirect(url_for("login_page"))
    if path in AUTH_EXEMPT_PATHS:
        return None
    if not session.get("logged_in"):
        return redirect(url_for("login_page", next=path))
    return None


def init_db():
    conn = sqlite3.connect(DB_PATH)
    conn.row_factory = sqlite3.Row

    # Master list of depositors (deposit holders).
    conn.execute("""
        CREATE TABLE IF NOT EXISTS depositors (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            holder_id TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL
        )
    """)

    # Master list of banks.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS banks (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            bank_id TEXT NOT NULL UNIQUE,
            name TEXT NOT NULL
        )
    """)

    # Stock/ETF holdings. Prices are fetched live via yfinance and converted
    # to rupees, so nothing is cached here beyond what you paid.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS investments (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            depositor_id INTEGER REFERENCES depositors(id),
            ticker TEXT NOT NULL,
            shares REAL NOT NULL,
            purchase_price REAL NOT NULL,
            purchase_date TEXT NOT NULL
        )
    """)

    # Precious-metal holdings. Prices are per gram.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS metals (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            depositor_id INTEGER REFERENCES depositors(id),
            metal TEXT NOT NULL,
            description TEXT NOT NULL DEFAULT '',
            grams REAL NOT NULL,
            purchase_price REAL NOT NULL,
            current_price REAL NOT NULL,
            purchase_date TEXT NOT NULL
        )
    """)
    if "depositor_id" not in {r[1] for r in conn.execute("PRAGMA table_info(metals)")}:
        conn.execute("ALTER TABLE metals ADD COLUMN depositor_id INTEGER REFERENCES depositors(id)")

    # Live market price per gram for each metal (one row per metal type).
    conn.execute("""
        CREATE TABLE IF NOT EXISTS metal_prices (
            metal TEXT PRIMARY KEY,
            price_per_gram REAL NOT NULL,
            updated_on TEXT NOT NULL,
            source TEXT NOT NULL DEFAULT 'manual'
        )
    """)
    if "source" not in {r[1] for r in conn.execute("PRAGMA table_info(metal_prices)")}:
        conn.execute("ALTER TABLE metal_prices ADD COLUMN source TEXT NOT NULL DEFAULT 'manual'")

    # Gold split into 24K / 22K — migrate the old single "gold" key.
    conn.execute("UPDATE metals SET metal = 'gold_24k' WHERE metal = 'gold'")
    if not conn.execute("SELECT 1 FROM metal_prices WHERE metal = 'gold_24k'").fetchone():
        conn.execute("UPDATE metal_prices SET metal = 'gold_24k' WHERE metal = 'gold'")
    conn.execute("DELETE FROM metal_prices WHERE metal = 'gold'")

    # Seed from existing holdings: use each metal's most recent holding price.
    priced = {r["metal"] for r in conn.execute("SELECT metal FROM metal_prices")}
    for row in conn.execute(
        "SELECT metal, current_price FROM metals ORDER BY purchase_date DESC, id DESC"
    ):
        if row["metal"] not in priced:
            conn.execute(
                "INSERT INTO metal_prices (metal, price_per_gram, updated_on) VALUES (?, ?, ?)",
                (row["metal"], row["current_price"], date.today().isoformat()),
            )
            priced.add(row["metal"])

    conn.execute("""
        CREATE TABLE IF NOT EXISTS deposits (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            depositor_id INTEGER REFERENCES depositors(id),
            bank_ref_id INTEGER REFERENCES banks(id),
            holder_id TEXT NOT NULL DEFAULT '',
            holder_name TEXT NOT NULL DEFAULT '',
            bank_name TEXT NOT NULL,
            principal REAL NOT NULL,
            interest_rate REAL NOT NULL,
            tenure_months INTEGER NOT NULL,
            tenure_days INTEGER NOT NULL DEFAULT 0,
            tenure_unit TEXT NOT NULL DEFAULT 'months',
            compounding_frequency INTEGER NOT NULL,
            start_date TEXT NOT NULL
        )
    """)
    # Migrations: add columns that older databases don't have yet.
    existing_cols = {row[1] for row in conn.execute("PRAGMA table_info(deposits)")}
    migrations = {
        "deposit_type": "ALTER TABLE deposits ADD COLUMN deposit_type TEXT NOT NULL DEFAULT 'cumulative'",
        "holder_id": "ALTER TABLE deposits ADD COLUMN holder_id TEXT NOT NULL DEFAULT ''",
        "holder_name": "ALTER TABLE deposits ADD COLUMN holder_name TEXT NOT NULL DEFAULT ''",
        "tenure_unit": "ALTER TABLE deposits ADD COLUMN tenure_unit TEXT NOT NULL DEFAULT 'months'",
        "tenure_days": "ALTER TABLE deposits ADD COLUMN tenure_days INTEGER NOT NULL DEFAULT 0",
        "depositor_id": "ALTER TABLE deposits ADD COLUMN depositor_id INTEGER REFERENCES depositors(id)",
        "bank_ref_id": "ALTER TABLE deposits ADD COLUMN bank_ref_id INTEGER REFERENCES banks(id)",
        "last_notified_on": "ALTER TABLE deposits ADD COLUMN last_notified_on TEXT",
    }
    for col, ddl in migrations.items():
        if col not in existing_cols:
            conn.execute(ddl)

    # Backfill: promote any legacy free-text holder on a deposit into a
    # depositor record and link it, so old rows show up in the new UI.
    legacy = conn.execute(
        """SELECT DISTINCT holder_id, holder_name FROM deposits
           WHERE depositor_id IS NULL AND TRIM(holder_name) <> ''"""
    ).fetchall()
    for row in legacy:
        depositor_id = find_or_create_depositor(
            conn, row["holder_id"].strip() or row["holder_name"].strip(), row["holder_name"].strip()
        )
        conn.execute(
            "UPDATE deposits SET depositor_id = ? WHERE depositor_id IS NULL AND holder_name = ?",
            (depositor_id, row["holder_name"]),
        )

    # Backfill: promote each distinct free-text bank name into a bank record.
    legacy_banks = conn.execute(
        """SELECT DISTINCT bank_name FROM deposits
           WHERE bank_ref_id IS NULL AND TRIM(bank_name) <> ''"""
    ).fetchall()
    for row in legacy_banks:
        name = row["bank_name"].strip()
        bank_id = find_or_create_bank(conn, name, name)
        conn.execute(
            "UPDATE deposits SET bank_ref_id = ? WHERE bank_ref_id IS NULL AND bank_name = ?",
            (bank_id, row["bank_name"]),
        )

    # Maturity-reminder email settings (a single row, id=1).
    conn.execute("""
        CREATE TABLE IF NOT EXISTS notification_settings (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            enabled INTEGER NOT NULL DEFAULT 0,
            recipient_email TEXT NOT NULL DEFAULT '',
            sender_email TEXT NOT NULL DEFAULT '',
            sender_app_password TEXT NOT NULL DEFAULT '',
            days_before INTEGER NOT NULL DEFAULT 30,
            last_check_at TEXT,
            last_check_result TEXT
        )
    """)
    conn.execute("INSERT OR IGNORE INTO notification_settings (id) VALUES (1)")

    # Single admin login for the app (a single row, id=1). No row yet means
    # the app hasn't been through first-run /setup.
    conn.execute("""
        CREATE TABLE IF NOT EXISTS auth_user (
            id INTEGER PRIMARY KEY CHECK (id = 1),
            email TEXT NOT NULL,
            password_hash TEXT NOT NULL,
            reset_token TEXT,
            reset_token_expires TEXT
        )
    """)

    conn.commit()
    conn.close()


def find_or_create_depositor(conn, holder_id: str, name: str) -> int:
    """Return the id of the depositor with this holder_id, creating it if needed."""
    existing = conn.execute(
        "SELECT id FROM depositors WHERE holder_id = ?", (holder_id,)
    ).fetchone()
    if existing:
        return existing["id"]
    cur = conn.execute(
        "INSERT INTO depositors (holder_id, name) VALUES (?, ?)", (holder_id, name)
    )
    return cur.lastrowid


def find_or_create_bank(conn, bank_id: str, name: str) -> int:
    """Return the id of the bank with this bank_id, creating it if needed."""
    existing = conn.execute(
        "SELECT id FROM banks WHERE bank_id = ?", (bank_id,)
    ).fetchone()
    if existing:
        return existing["id"]
    cur = conn.execute(
        "INSERT INTO banks (bank_id, name) VALUES (?, ?)", (bank_id, name)
    )
    return cur.lastrowid


# ---------- Date math ----------
def add_months(d: date, months: int) -> date:
    month_index = d.month - 1 + months
    year = d.year + month_index // 12
    month = month_index % 12 + 1
    day = min(d.day, [31,29 if year % 4 == 0 and (year % 100 != 0 or year % 400 == 0) else 28,
                       31,30,31,30,31,31,30,31,30,31][month - 1])
    return date(year, month, day)


# ---------- FD / RD math ----------
def calculate_cumulative(principal: float, annual_rate: float, t_years: float,
                         compounding_frequency: int):
    """Interest is reinvested. A = P * (1 + r/n)^(n*t), t in years."""
    r = annual_rate / 100
    n = compounding_frequency
    maturity_amount = principal * (1 + r / n) ** (n * t_years)
    return maturity_amount, maturity_amount - principal


def calculate_simple(principal: float, annual_rate: float, t_years: float):
    """Interest is paid out periodically, not reinvested. I = P * r * t."""
    r = annual_rate / 100
    interest_earned = principal * r * t_years
    return principal + interest_earned, interest_earned


def calculate_recurring(monthly_installment: float, annual_rate: float, tenure_months: int):
    """One fixed installment per month; balance compounds quarterly (standard bank RD).

    Each installment earns interest for the whole number of months it stays on
    deposit, converted to quarters: installment k is invested for
    (tenure_months - k + 1) months.
    """
    r = annual_rate / 100
    quarterly_rate = r / 4
    maturity_amount = 0.0
    for k in range(1, tenure_months + 1):
        months_on_deposit = tenure_months - k + 1
        quarters = months_on_deposit / 3
        maturity_amount += monthly_installment * (1 + quarterly_rate) ** quarters
    invested = monthly_installment * tenure_months
    return maturity_amount, maturity_amount - invested


def months_between(start: date, today: date) -> int:
    """Whole months elapsed from `start` up to `today` (0 if today is on/before start)."""
    if today <= start:
        return 0
    m = (today.year - start.year) * 12 + (today.month - start.month)
    if today.day < start.day:
        m -= 1
    return max(m, 0)


def recurring_value_to_date(monthly_installment: float, annual_rate: float,
                            tenure_months: int, installments_paid: int,
                            months_elapsed: int) -> float:
    """Value of an RD today: each installment paid so far, compounded quarterly
    from the month it was paid up to today."""
    quarterly_rate = annual_rate / 100 / 4
    value = 0.0
    for k in range(1, installments_paid + 1):
        months_on_deposit = min(months_elapsed - (k - 1), tenure_months - (k - 1))
        months_on_deposit = max(months_on_deposit, 0)
        value += monthly_installment * (1 + quarterly_rate) ** (months_on_deposit / 3)
    return value


MIN_DAYS_TO_ANNUALISE = 7  # shorter holds swing wildly when annualised; show "—" instead


def annualised_return_pct(invested: float, current: float, days_held: float):
    """CAGR-style annualised return: ((current/invested)^(365/days) - 1) * 100.
    Returns None when it can't be sensibly computed (nothing invested, not
    started yet, or too new to annualise without a misleading swing)."""
    if invested is None or invested <= 0 or days_held is None or days_held < MIN_DAYS_TO_ANNUALISE:
        return None
    years = days_held / DAYS_PER_YEAR
    try:
        return ((current / invested) ** (1.0 / years) - 1.0) * 100.0
    except (OverflowError, ValueError, ZeroDivisionError):
        return None


def weighted_annualised_return(rows, invested_key: str, current_key: str, days_key: str):
    """Blend several holdings' annualised returns into one figure: the holding
    period used is the invested-weighted average of the individual periods
    (so a big, long-held position moves the number more than a small, new
    one). `rows` may be dicts or sqlite3.Row-like objects."""
    total_invested = sum(r[invested_key] for r in rows)
    total_current = sum(r[current_key] for r in rows)
    if total_invested <= 0:
        return None
    weighted_days = sum(r[invested_key] * r[days_key] for r in rows) / total_invested
    return annualised_return_pct(total_invested, total_current, weighted_days)


def _row_get(d, key, default):
    return d[key] if key in d.keys() else default


def summarise_deposit(d) -> dict:
    """Turn a raw deposits row into a display dict with computed figures."""
    dtype = _row_get(d, "deposit_type", "cumulative")
    if dtype not in DEPOSIT_TYPES:
        dtype = "cumulative"
    start = date.fromisoformat(d["start_date"])

    # Resolve the tenure into a length in years and a maturity date.
    unit = _row_get(d, "tenure_unit", "months")
    tenure_days = _row_get(d, "tenure_days", 0)
    if unit == "days" and dtype != "recurring":
        t_years = tenure_days / DAYS_PER_YEAR
        maturity_date = start + timedelta(days=tenure_days)
        tenure_label = f"{tenure_days} days"
    else:
        unit = "months"
        t_years = d["tenure_months"] / 12
        maturity_date = add_months(start, d["tenure_months"])
        tenure_label = f"{d['tenure_months']} mo"

    if dtype == "simple":
        maturity_amount, interest_earned = calculate_simple(
            d["principal"], d["interest_rate"], t_years
        )
        invested = d["principal"]
    elif dtype == "recurring":
        maturity_amount, interest_earned = calculate_recurring(
            d["principal"], d["interest_rate"], d["tenure_months"]
        )
        invested = 0.0  # set below to installments actually paid so far
    else:
        maturity_amount, interest_earned = calculate_cumulative(
            d["principal"], d["interest_rate"], t_years, d["compounding_frequency"]
        )
        invested = d["principal"]

    # Prefer the linked master records; fall back to any legacy free text.
    holder_name = _row_get(d, "depositor_name", None) or _row_get(d, "holder_name", "")
    holder_id = _row_get(d, "depositor_holder_id", None) or _row_get(d, "holder_id", "")
    bank_name = _row_get(d, "bank_ref_name", None) or _row_get(d, "bank_name", "")
    bank_code = _row_get(d, "bank_code", "") or ""

    today = date.today()
    is_matured = today >= maturity_date

    # ----- Progress + current value (principal + interest accrued to date) -----
    months_elapsed = months_between(start, today)
    elapsed_years = min(max((today - start).days, 0) / DAYS_PER_YEAR, t_years)

    if dtype == "recurring":
        total_installments = d["tenure_months"]
        if today < start:
            installments_paid = 0
        else:
            installments_paid = min(months_elapsed + 1, total_installments)
        paid_in = installments_paid * d["principal"]
        invested = paid_in  # RD "invested" = instalment amount x instalments paid to date
        if is_matured:
            current_value = maturity_amount
        else:
            current_value = recurring_value_to_date(
                d["principal"], d["interest_rate"], d["tenure_months"],
                installments_paid, months_elapsed,
            )
    else:
        total_installments = None
        installments_paid = None
        paid_in = d["principal"]
        if is_matured:
            current_value = maturity_amount
        elif dtype == "simple":
            current_value = d["principal"] + d["principal"] * (d["interest_rate"] / 100) * elapsed_years
        else:  # cumulative
            r = d["interest_rate"] / 100
            n = d["compounding_frequency"]
            current_value = d["principal"] * (1 + r / n) ** (n * elapsed_years)

    accrued_interest = current_value - paid_in
    days_held = max((today - start).days, 0)
    annualised_return = annualised_return_pct(invested, current_value, days_held)

    return {
        "id": d["id"],
        "holder_id": holder_id,
        "holder_name": holder_name,
        "bank_name": bank_name,
        "bank_code": bank_code,
        "deposit_type": dtype,
        "deposit_type_label": DEPOSIT_TYPES[dtype],
        "principal": d["principal"],
        "monthly_installment": d["principal"] if dtype == "recurring" else None,
        "invested": invested,
        "interest_rate": d["interest_rate"],
        "tenure_unit": unit,
        "tenure_months": d["tenure_months"],
        "tenure_days": tenure_days,
        "tenure_label": tenure_label,
        "compounding_frequency": d["compounding_frequency"],
        "start_date": d["start_date"],
        "maturity_date": maturity_date.isoformat(),
        "maturity_amount": maturity_amount,
        "interest_earned": interest_earned,
        "is_matured": is_matured,
        "days_remaining": (maturity_date - today).days,
        "installments_paid": installments_paid,
        "total_installments": total_installments,
        "paid_in": paid_in,
        "current_value": current_value,
        "accrued_interest": accrued_interest,
        "days_held": days_held,
        "annualised_return": annualised_return,
    }


# ---------- Master-list helpers ----------
def list_depositors(db):
    return db.execute(
        """SELECT d.id, d.holder_id, d.name,
                  (SELECT COUNT(*) FROM deposits    WHERE depositor_id = d.id) AS deposit_count,
                  (SELECT COUNT(*) FROM metals      WHERE depositor_id = d.id) AS metal_count,
                  (SELECT COUNT(*) FROM investments WHERE depositor_id = d.id) AS investment_count
           FROM depositors d
           ORDER BY d.name COLLATE NOCASE"""
    ).fetchall()


def list_banks(db):
    return db.execute(
        """SELECT b.id, b.bank_id, b.name, COUNT(dep.id) AS deposit_count
           FROM banks b
           LEFT JOIN deposits dep ON dep.bank_ref_id = b.id
           GROUP BY b.id ORDER BY b.name COLLATE NOCASE"""
    ).fetchall()


DEPOSITS_WITH_REFS = """
    SELECT deposits.*,
           depositors.name AS depositor_name,
           depositors.holder_id AS depositor_holder_id,
           banks.name AS bank_ref_name,
           banks.bank_id AS bank_code
    FROM deposits
    LEFT JOIN depositors ON deposits.depositor_id = depositors.id
    LEFT JOIN banks ON deposits.bank_ref_id = banks.id
"""


def depositors_with_totals(db):
    """Depositor list plus each one's total invested (principal only),
    current value (principal + interest accrued to date) and total maturity
    value, summed over their linked deposits."""
    totals = {}  # depositor_id -> {"invested": x, "current": c, "maturity": y}
    for d in db.execute(DEPOSITS_WITH_REFS).fetchall():
        if d["depositor_id"] is None:
            continue
        s = summarise_deposit(d)
        acc = totals.setdefault(
            d["depositor_id"], {"invested": 0.0, "current": 0.0, "maturity": 0.0}
        )
        acc["invested"] += s["invested"]
        acc["current"] += s["current_value"]
        acc["maturity"] += s["maturity_amount"]

    rows = []
    for dep in list_depositors(db):
        t = totals.get(dep["id"], {"invested": 0.0, "current": 0.0, "maturity": 0.0})
        rows.append({
            "id": dep["id"],
            "holder_id": dep["holder_id"],
            "name": dep["name"],
            "deposit_count": dep["deposit_count"],
            "metal_count": dep["metal_count"],
            "investment_count": dep["investment_count"],
            "total_invested": t["invested"],
            "total_current": t["current"],
            "total_maturity": t["maturity"],
        })
    return rows


def _agg_blank():
    return {"count": 0, "invested": 0.0, "current": 0.0, "maturity": 0.0, "days_x_invested": 0.0}


def _agg_add(acc, s):
    acc["count"] += 1
    acc["invested"] += s["invested"]
    acc["current"] += s["current_value"]
    acc["maturity"] += s["maturity_amount"]
    acc["days_x_invested"] += s["invested"] * s["days_held"]


def _metal_blank():
    return {"count": 0, "grams": 0.0, "invested": 0.0, "current": 0.0, "days_x_invested": 0.0}


def _inv_blank():
    return {"count": 0, "shares": 0.0, "invested": 0.0, "current": 0.0, "days_x_invested": 0.0}


def _with_annualised(acc: dict) -> dict:
    """Add an "annualised_return" key: the invested-weighted average holding
    period, fed into annualised_return_pct. Leaves the input untouched."""
    weighted_days = (acc["days_x_invested"] / acc["invested"]) if acc["invested"] else 0
    return {**acc, "annualised_return": annualised_return_pct(acc["invested"], acc["current"], weighted_days)}


def portfolio_summary(db):
    """Deposits + metal holdings + investments, aggregated for the Dashboard
    and Chart.

    For metals and investments, "invested" is cost (grams/shares x purchase
    price) and "current" is today's market value; neither has a maturity
    value. Investments whose live price couldn't be fetched right now are
    left out of these totals (they still show on the Investments tab).
    """
    _key = lambda kv: kv[0].lower()

    # ----- deposits -----
    dep_pair, dep_holder, dep_bank = {}, {}, {}
    dep_overall = _agg_blank()
    for d in db.execute(DEPOSITS_WITH_REFS).fetchall():
        s = summarise_deposit(d)
        holder = s["holder_name"] or "—"
        bank = s["bank_name"] or "—"
        _agg_add(dep_pair.setdefault((holder, bank), _agg_blank()), s)
        _agg_add(dep_holder.setdefault(holder, _agg_blank()), s)
        _agg_add(dep_bank.setdefault(bank, _agg_blank()), s)
        _agg_add(dep_overall, s)

    # ----- metals -----
    met_pair, met_holder, met_metal = {}, {}, {}
    met_overall = _metal_blank()
    for m in list_metals(db):
        holder = m["depositor_name"] or "—"
        for bucket, k in (
            (met_pair, (holder, m["metal_label"])),
            (met_holder, holder),
            (met_metal, m["metal_label"]),
        ):
            acc = bucket.setdefault(k, _metal_blank())
            acc["count"] += 1
            acc["grams"] += m["grams"]
            acc["invested"] += m["cost"]
            acc["current"] += m["value"]
            acc["days_x_invested"] += m["cost"] * m["days_held"]
        met_overall["count"] += 1
        met_overall["grams"] += m["grams"]
        met_overall["invested"] += m["cost"]
        met_overall["current"] += m["value"]
        met_overall["days_x_invested"] += m["cost"] * m["days_held"]

    # ----- investments -----
    inv_pair, inv_holder, inv_ticker = {}, {}, {}
    inv_overall = _inv_blank()
    for iv in list_investments(db):
        if iv["value"] is None:
            continue
        holder = iv["depositor_name"] or "—"
        for bucket, k in (
            (inv_pair, (holder, iv["ticker"])),
            (inv_holder, holder),
            (inv_ticker, iv["ticker"]),
        ):
            acc = bucket.setdefault(k, _inv_blank())
            acc["count"] += 1
            acc["shares"] += iv["shares"]
            acc["invested"] += iv["cost"]
            acc["current"] += iv["value"]
            acc["days_x_invested"] += iv["cost"] * iv["days_held"]
        inv_overall["count"] += 1
        inv_overall["shares"] += iv["shares"]
        inv_overall["invested"] += iv["cost"]
        inv_overall["current"] += iv["value"]
        inv_overall["days_x_invested"] += iv["cost"] * iv["days_held"]

    # ----- combined by holder (deposits + metals + investments) -----
    combined = {}
    def _c(h):
        return combined.setdefault(h, {"deposit_count": 0, "metal_count": 0, "metal_grams": 0.0,
                                       "investment_count": 0,
                                       "invested": 0.0, "current": 0.0, "maturity": 0.0,
                                       "days_x_invested": 0.0})
    for h, v in dep_holder.items():
        c = _c(h)
        c["deposit_count"] += v["count"]
        c["invested"] += v["invested"]; c["current"] += v["current"]; c["maturity"] += v["maturity"]
        c["days_x_invested"] += v["days_x_invested"]
    for h, v in met_holder.items():
        c = _c(h)
        c["metal_count"] += v["count"]
        c["metal_grams"] += v["grams"]
        c["invested"] += v["invested"]; c["current"] += v["current"]
        c["days_x_invested"] += v["days_x_invested"]
    for h, v in inv_holder.items():
        c = _c(h)
        c["investment_count"] += v["count"]
        c["invested"] += v["invested"]; c["current"] += v["current"]
        c["days_x_invested"] += v["days_x_invested"]

    total_invested = dep_overall["invested"] + met_overall["invested"] + inv_overall["invested"]
    total_current = dep_overall["current"] + met_overall["current"] + inv_overall["current"]
    total_days_x_invested = (
        dep_overall["days_x_invested"] + met_overall["days_x_invested"] + inv_overall["days_x_invested"]
    )
    total_weighted_days = (total_days_x_invested / total_invested) if total_invested else 0
    total_annualised_return = annualised_return_pct(total_invested, total_current, total_weighted_days)

    return {
        "dep_pairs": [_with_annualised({"holder": h, "bank": b, **v})
                      for (h, b), v in sorted(dep_pair.items(),
                                              key=lambda kv: (kv[0][0].lower(), kv[0][1].lower()))],
        "dep_banks": [_with_annualised({"name": b, **v}) for b, v in sorted(dep_bank.items(), key=_key)],
        "dep_overall": _with_annualised(dep_overall),
        "met_pairs": [_with_annualised({"holder": h, "metal": mt, **v})
                      for (h, mt), v in sorted(met_pair.items(),
                                               key=lambda kv: (kv[0][0].lower(), kv[0][1].lower()))],
        "met_by_metal": [_with_annualised({"name": mt, **v}) for mt, v in sorted(met_metal.items(), key=_key)],
        "met_overall": _with_annualised(met_overall),
        "inv_pairs": [_with_annualised({"holder": h, "ticker": tk, **v})
                      for (h, tk), v in sorted(inv_pair.items(),
                                               key=lambda kv: (kv[0][0].lower(), kv[0][1].lower()))],
        "inv_by_ticker": [_with_annualised({"name": tk, **v}) for tk, v in sorted(inv_ticker.items(), key=_key)],
        "inv_overall": _with_annualised(inv_overall),
        "combined_holders": [_with_annualised({"name": h, **v}) for h, v in sorted(combined.items(), key=_key)],
        "total_invested": total_invested,
        "total_current": total_current,
        "total_gain": total_current - total_invested,
        "total_annualised_return": total_annualised_return,
    }


# ---------- Authentication ----------
def build_password_reset_email(reset_url: str) -> tuple:
    subject = "Password reset — Fixed Deposit Manager"
    body = (
        "A password reset was requested for your Fixed Deposit Manager login.\n\n"
        f"Reset your password: {reset_url}\n\n"
        "This link expires in 1 hour and can only be used once. "
        "If you didn't request this, you can safely ignore this email.\n\n"
        "— sent automatically by Fixed Deposit Manager"
    )
    return subject, body


def send_password_reset_email(db, to_email: str, reset_url: str) -> None:
    """Reuses the Gmail sender configured on the Notifications tab. Raises
    RuntimeError (with a user-facing message) if that isn't set up, or if
    sending fails."""
    ns = get_notification_settings(db)
    settings = {
        "sender_email": ns.get("sender_email", ""),
        "sender_app_password": ns.get("sender_app_password", ""),
        "recipient_email": to_email,
    }
    subject, body = build_password_reset_email(reset_url)
    send_email(settings, subject, body)


@app.route("/setup", methods=["GET", "POST"])
def setup_page():
    db = get_db()
    if get_auth_user(db) is not None:
        return redirect(url_for("login_page"))

    error = None
    email = ""
    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        if not email or "@" not in email:
            error = "Enter a valid email address."
        elif len(password) < MIN_PASSWORD_LENGTH:
            error = f"Password must be at least {MIN_PASSWORD_LENGTH} characters."
        elif password != confirm_password:
            error = "Passwords don't match."
        else:
            db.execute(
                "INSERT INTO auth_user (id, email, password_hash) VALUES (1, ?, ?)",
                (email, generate_password_hash(password)),
            )
            db.commit()
            session.clear()
            session["logged_in"] = True
            session.permanent = True
            return redirect(url_for("summary_page"))

    return render_template("setup.html", error=error, email=email)


@app.route("/login", methods=["GET", "POST"])
def login_page():
    db = get_db()
    user = get_auth_user(db)
    error = None

    if request.method == "POST":
        email = request.form.get("email", "").strip()
        password = request.form.get("password", "")
        if user is not None and email.lower() == user["email"].lower() \
                and check_password_hash(user["password_hash"], password):
            session.clear()
            session["logged_in"] = True
            session.permanent = True
            next_url = request.form.get("next") or url_for("summary_page")
            return redirect(next_url)
        error = "Incorrect email or password."

    return render_template("login.html", error=error, next=request.args.get("next", ""))


@app.route("/logout", methods=["POST"])
def logout():
    session.clear()
    return redirect(url_for("login_page"))


@app.route("/forgot-password", methods=["GET", "POST"])
def forgot_password_page():
    db = get_db()
    user = get_auth_user(db)
    error = None
    submitted = False

    if request.method == "POST":
        email = request.form.get("email", "").strip()
        if user is not None and email.lower() == user["email"].lower():
            token = secrets.token_urlsafe(32)
            expires = (datetime.utcnow() + PASSWORD_RESET_TOKEN_LIFETIME).isoformat()
            db.execute(
                "UPDATE auth_user SET reset_token = ?, reset_token_expires = ? WHERE id = 1",
                (token, expires),
            )
            db.commit()
            reset_url = url_for("reset_password_page", token=token, _external=True)
            try:
                send_password_reset_email(db, user["email"], reset_url)
            except RuntimeError as e:
                error = f"Couldn't send the reset email: {e}"
        submitted = True

    return render_template("forgot_password.html", error=error, submitted=submitted)


@app.route("/reset-password/<token>", methods=["GET", "POST"])
def reset_password_page(token):
    db = get_db()
    user = get_auth_user(db)
    valid = (
        user is not None
        and user["reset_token"]
        and secrets.compare_digest(user["reset_token"], token)
        and user["reset_token_expires"]
        and datetime.fromisoformat(user["reset_token_expires"]) > datetime.utcnow()
    )
    if not valid:
        return render_template("reset_password.html", invalid=True, error=None)

    error = None
    if request.method == "POST":
        password = request.form.get("password", "")
        confirm_password = request.form.get("confirm_password", "")
        if len(password) < MIN_PASSWORD_LENGTH:
            error = f"Password must be at least {MIN_PASSWORD_LENGTH} characters."
        elif password != confirm_password:
            error = "Passwords don't match."
        else:
            db.execute(
                """UPDATE auth_user SET password_hash = ?, reset_token = NULL, reset_token_expires = NULL
                   WHERE id = 1""",
                (generate_password_hash(password),),
            )
            db.commit()
            session.clear()
            return redirect(url_for("login_page"))

    return render_template("reset_password.html", invalid=False, error=error)


# ---------- Routes ----------
@app.route("/summary")
def summary_page():
    db = get_db()
    return render_template(
        "summary.html", active_tab="summary", **portfolio_summary(db)
    )


CHART_TYPES = {"bar": "Bar", "pie": "Pie"}
CHART_METRICS = {"invested": "Invested", "current": "Current value"}
PIE_COLORS = [
    "#1e4d6b", "#1c7c45", "#b8720b", "#7d5ba6", "#c0392b",
    "#2c8c99", "#8a6d3b", "#5b6b8c", "#4a7c59", "#9b3b6a",
]


def build_pie(rows, metric, cx=90.0, cy=90.0, r=80.0):
    """Turn aggregate rows into SVG pie-slice paths for the chosen metric."""
    total = sum(row[metric] for row in rows)
    slices = []
    angle = -90.0
    for i, row in enumerate(rows):
        value = row[metric]
        frac = (value / total) if total > 0 else 0.0
        sweep = frac * 360.0
        colour = PIE_COLORS[i % len(PIE_COLORS)]
        if len(rows) == 1 or frac >= 0.999999:
            # a lone / full slice: draw a complete circle
            path = (f"M {cx - r:.2f} {cy:.2f} A {r} {r} 0 1 1 {cx + r:.2f} {cy:.2f} "
                    f"A {r} {r} 0 1 1 {cx - r:.2f} {cy:.2f} Z")
        else:
            a0, a1 = math.radians(angle), math.radians(angle + sweep)
            x0, y0 = cx + r * math.cos(a0), cy + r * math.sin(a0)
            x1, y1 = cx + r * math.cos(a1), cy + r * math.sin(a1)
            large = 1 if sweep > 180 else 0
            path = (f"M {cx:.2f} {cy:.2f} L {x0:.2f} {y0:.2f} "
                    f"A {r} {r} 0 {large} 1 {x1:.2f} {y1:.2f} Z")
        slices.append({
            "label": row["name"], "value": value,
            "pct": frac * 100.0, "colour": colour, "path": path,
        })
        angle += sweep
    return {"slices": slices, "total": total, "size": cx * 2}


@app.route("/chart")
def chart_page():
    db = get_db()
    ps = portfolio_summary(db)

    chart_type = request.args.get("type", "bar")
    if chart_type not in CHART_TYPES:
        chart_type = "bar"
    metric = request.args.get("metric", "current")
    if metric not in CHART_METRICS:
        metric = "current"

    def series(items):
        return [{"name": it["name"], "invested": it["invested"], "current": it["current"]}
                for it in items]

    holder_data = series(ps["combined_holders"])   # deposits + metals + investments
    bank_data = series(ps["dep_banks"])            # deposits only
    metal_data = series(ps["met_by_metal"])        # metals only
    investment_data = series(ps["inv_by_ticker"])  # investments only
    axis_max = max(
        [0.0]
        + [row[k] for row in holder_data + bank_data + metal_data + investment_data
           for k in ("invested", "current")]
    )
    return render_template(
        "chart.html",
        active_tab="chart",
        chart_type=chart_type,
        metric=metric,
        chart_types=CHART_TYPES,
        chart_metrics=CHART_METRICS,
        holder_data=holder_data,
        bank_data=bank_data,
        metal_data=metal_data,
        investment_data=investment_data,
        axis_max=axis_max,
        holder_pie=build_pie(holder_data, metric),
        bank_pie=build_pie(bank_data, metric),
        metal_pie=build_pie(metal_data, metric),
        investment_pie=build_pie(investment_data, metric),
    )


@app.route("/calculator")
def calculator_page():
    """Standalone interest calculator — nothing here is saved to the database."""
    args = request.args
    deposit_type = args.get("deposit_type", "cumulative")
    if deposit_type not in DEPOSIT_TYPES:
        deposit_type = "cumulative"

    form_data = {
        "deposit_type": deposit_type,
        "principal": args.get("principal", ""),
        "interest_rate": args.get("interest_rate", ""),
        "duration_months": args.get("duration_months", ""),
        "duration_days": args.get("duration_days", ""),
        "compounding_frequency": args.get("compounding_frequency", "4"),
    }

    result = None
    error = None
    if form_data["principal"]:
        try:
            principal = float(form_data["principal"])
            rate = float(form_data["interest_rate"])
            months = int(form_data["duration_months"] or 0)
            days = int(form_data["duration_days"] or 0)
            compounding_frequency = int(form_data["compounding_frequency"])

            amount_label = "Monthly installment" if deposit_type == "recurring" else "Principal"
            if principal <= 0:
                raise ValueError(f"{amount_label} must be greater than 0.")
            if rate <= 0:
                raise ValueError("Interest rate must be greater than 0.")
            if months < 0 or days < 0:
                raise ValueError("Duration can't be negative.")
            if months == 0 and days == 0:
                raise ValueError("Enter a duration in months and/or days.")
            if deposit_type == "recurring" and months == 0:
                raise ValueError("A recurring deposit needs a whole number of months.")

            if deposit_type == "recurring":
                maturity_amount, interest_earned = calculate_recurring(principal, rate, months)
                result = {
                    "deposit_type": deposit_type,
                    "invested": principal * months,
                    "maturity_amount": maturity_amount,
                    "interest_earned": interest_earned,
                    "installments": months,
                }
            else:
                total_years = months / 12 + days / DAYS_PER_YEAR
                if deposit_type == "simple":
                    maturity_amount, interest_earned = calculate_simple(principal, rate, total_years)
                    result = {
                        "deposit_type": deposit_type,
                        "invested": principal,
                        "maturity_amount": maturity_amount,
                        "interest_earned": interest_earned,
                        "monthly_interest": principal * (rate / 100) / 12,
                        "quarterly_interest": principal * (rate / 100) / 4,
                    }
                else:  # cumulative
                    maturity_amount, interest_earned = calculate_cumulative(
                        principal, rate, total_years, compounding_frequency
                    )
                    result = {
                        "deposit_type": deposit_type,
                        "invested": principal,
                        "maturity_amount": maturity_amount,
                        "interest_earned": interest_earned,
                    }
        except ValueError as e:
            msg = str(e)
            error = msg if ("could not convert" not in msg and "invalid literal" not in msg) \
                else "Please enter valid numbers for amount, rate and duration."

    return render_template(
        "calculator.html",
        active_tab="calculator",
        deposit_types=DEPOSIT_TYPES,
        form_data=form_data,
        result=result,
        error=error,
    )


@app.route("/")
def dashboard():
    db = get_db()
    deposits = db.execute(
        DEPOSITS_WITH_REFS + " ORDER BY deposits.start_date DESC"
    ).fetchall()

    rows = [summarise_deposit(d) for d in deposits]
    total_invested = sum(r["invested"] for r in rows)
    total_current = sum(r["current_value"] for r in rows)
    total_maturity = sum(r["maturity_amount"] for r in rows)
    total_interest = sum(r["interest_earned"] for r in rows)
    total_annualised_return = weighted_annualised_return(rows, "invested", "current_value", "days_held")

    return render_template(
        "dashboard.html",
        rows=rows,
        total_invested=total_invested,
        total_current=total_current,
        total_maturity=total_maturity,
        total_interest=total_interest,
        total_annualised_return=total_annualised_return,
        active_tab="dashboard",
        wide_page=True,
    )


def parse_deposit_form(form_data, db) -> dict:
    """Validate the shared deposit form. Returns a dict of DB column values,
    or raises ValueError with a user-facing message."""
    deposit_type = form_data["deposit_type"]
    if deposit_type not in DEPOSIT_TYPES:
        raise ValueError("Please choose a valid deposit type.")

    depositor = db.execute(
        "SELECT id, holder_id, name FROM depositors WHERE id = ?",
        (form_data["depositor_id"],),
    ).fetchone()
    if depositor is None:
        raise ValueError("Please choose a depositor.")

    bank = db.execute(
        "SELECT id, bank_id, name FROM banks WHERE id = ?", (form_data["bank_ref_id"],)
    ).fetchone()
    if bank is None:
        raise ValueError("Please choose a bank.")

    try:
        principal = float(form_data["principal"])
        interest_rate = float(form_data["interest_rate"])
        tenure_value = int(form_data["tenure_value"])
        compounding_frequency = int(form_data["compounding_frequency"])
    except (TypeError, ValueError):
        raise ValueError("Please enter valid numbers for principal, rate, and tenure.")

    if principal <= 0:
        amount_label = "Monthly installment" if deposit_type == "recurring" else "Principal"
        raise ValueError(f"{amount_label} must be greater than 0.")
    if interest_rate <= 0:
        raise ValueError("Interest rate must be greater than 0.")

    tenure_unit = form_data["tenure_unit"]
    if tenure_unit not in TENURE_UNITS:
        raise ValueError("Please choose a valid tenure unit.")
    if deposit_type == "recurring":
        tenure_unit = "months"  # RDs are inherently monthly
    if tenure_value <= 0:
        raise ValueError(f"Tenure must be greater than 0 {tenure_unit}.")

    if tenure_unit == "days":
        tenure_months, tenure_days = 0, tenure_value
    else:
        tenure_months, tenure_days = tenure_value, 0

    # Compounding frequency only matters for cumulative deposits.
    if deposit_type == "simple":
        compounding_frequency = 1
    elif deposit_type == "recurring":
        compounding_frequency = 4

    return {
        "depositor_id": depositor["id"],
        "holder_id": depositor["holder_id"],
        "holder_name": depositor["name"],
        "bank_ref_id": bank["id"],
        "bank_name": bank["name"],
        "deposit_type": deposit_type,
        "principal": principal,
        "interest_rate": interest_rate,
        "tenure_months": tenure_months,
        "tenure_days": tenure_days,
        "tenure_unit": tenure_unit,
        "compounding_frequency": compounding_frequency,
        "start_date": form_data["start_date"],
    }


def _row_to_form_data(row) -> dict:
    """Prefill the deposit form from an existing row."""
    unit = row["tenure_unit"] or "months"
    return {
        "depositor_id": str(row["depositor_id"] or ""),
        "bank_ref_id": str(row["bank_ref_id"] or ""),
        "deposit_type": row["deposit_type"],
        "principal": _trim_number(row["principal"]),
        "interest_rate": _trim_number(row["interest_rate"]),
        "tenure_value": str(row["tenure_days"] if unit == "days" else row["tenure_months"]),
        "tenure_unit": unit,
        "compounding_frequency": str(row["compounding_frequency"]),
        "start_date": row["start_date"],
    }


def _trim_number(x):
    """Render a float without a trailing '.0' so it round-trips cleanly in the form."""
    return str(int(x)) if float(x).is_integer() else str(x)


BLANK_DEPOSIT_FORM = {
    "depositor_id": "", "bank_ref_id": "", "deposit_type": "cumulative",
    "principal": "", "interest_rate": "", "tenure_value": "",
    "tenure_unit": "months", "compounding_frequency": "4",
    "start_date": None,  # filled with today's date at request time
}


def _render_deposit_form(db, **kwargs):
    return render_template(
        "add_deposit.html",
        deposit_types=DEPOSIT_TYPES,
        tenure_units=TENURE_UNITS,
        depositors=db.execute(
            "SELECT id, holder_id, name FROM depositors ORDER BY name COLLATE NOCASE"
        ).fetchall(),
        banks=db.execute(
            "SELECT id, bank_id, name FROM banks ORDER BY name COLLATE NOCASE"
        ).fetchall(),
        active_tab="dashboard",
        **kwargs,
    )


@app.route("/add", methods=["GET", "POST"])
def add_deposit():
    db = get_db()
    error = None
    form_data = dict(BLANK_DEPOSIT_FORM)
    form_data["start_date"] = str(date.today())

    if request.method == "POST":
        for key in form_data:
            form_data[key] = request.form.get(key, form_data[key])
        try:
            cols = parse_deposit_form(form_data, db)
            db.execute(
                """INSERT INTO deposits
                   (depositor_id, bank_ref_id, holder_id, holder_name, bank_name, deposit_type,
                    principal, interest_rate, tenure_months, tenure_days, tenure_unit,
                    compounding_frequency, start_date)
                   VALUES (:depositor_id, :bank_ref_id, :holder_id, :holder_name, :bank_name, :deposit_type,
                    :principal, :interest_rate, :tenure_months, :tenure_days, :tenure_unit,
                    :compounding_frequency, :start_date)""",
                cols,
            )
            db.commit()
            return redirect(url_for("dashboard"))
        except ValueError as e:
            error = str(e)

    return _render_deposit_form(db, error=error, form_data=form_data, editing=False)


@app.route("/edit/<int:deposit_id>", methods=["GET", "POST"])
def edit_deposit(deposit_id):
    db = get_db()
    row = db.execute("SELECT * FROM deposits WHERE id = ?", (deposit_id,)).fetchone()
    if row is None:
        return redirect(url_for("dashboard"))

    error = None
    form_data = _row_to_form_data(row)

    if request.method == "POST":
        for key in form_data:
            form_data[key] = request.form.get(key, form_data[key])
        try:
            cols = parse_deposit_form(form_data, db)
            cols["id"] = deposit_id
            db.execute(
                """UPDATE deposits SET
                     depositor_id = :depositor_id, bank_ref_id = :bank_ref_id,
                     holder_id = :holder_id, holder_name = :holder_name,
                     bank_name = :bank_name, deposit_type = :deposit_type,
                     principal = :principal, interest_rate = :interest_rate,
                     tenure_months = :tenure_months, tenure_days = :tenure_days,
                     tenure_unit = :tenure_unit,
                     compounding_frequency = :compounding_frequency, start_date = :start_date
                   WHERE id = :id""",
                cols,
            )
            db.commit()
            return redirect(url_for("dashboard"))
        except ValueError as e:
            error = str(e)

    return _render_deposit_form(
        db, error=error, form_data=form_data, editing=True, deposit_id=deposit_id
    )


@app.route("/depositors", methods=["GET", "POST"])
def depositors_page():
    db = get_db()
    error = None
    form_data = {"holder_id": "", "name": ""}

    if request.method == "POST":
        form_data["holder_id"] = request.form.get("holder_id", "").strip()
        form_data["name"] = request.form.get("name", "").strip()
        try:
            if not form_data["holder_id"]:
                raise ValueError("Deposit holder ID is required.")
            if not form_data["name"]:
                raise ValueError("Depositor name is required.")
            existing = db.execute(
                "SELECT id FROM depositors WHERE holder_id = ?", (form_data["holder_id"],)
            ).fetchone()
            if existing:
                raise ValueError(f"A depositor with ID '{form_data['holder_id']}' already exists.")
            db.execute(
                "INSERT INTO depositors (holder_id, name) VALUES (?, ?)",
                (form_data["holder_id"], form_data["name"]),
            )
            db.commit()
            return redirect(url_for("depositors_page"))
        except ValueError as e:
            error = str(e)

    depositors = depositors_with_totals(db)
    return render_template(
        "depositors.html",
        depositors=depositors,
        grand_invested=sum(d["total_invested"] for d in depositors),
        grand_current=sum(d["total_current"] for d in depositors),
        grand_maturity=sum(d["total_maturity"] for d in depositors),
        error=error, form_data=form_data, active_tab="depositors",
    )


@app.route("/depositors/<int:depositor_id>/delete", methods=["POST"])
def delete_depositor(depositor_id):
    db = get_db()
    in_use = (
        db.execute("SELECT COUNT(*) AS n FROM deposits WHERE depositor_id = ?", (depositor_id,)).fetchone()["n"]
        + db.execute("SELECT COUNT(*) AS n FROM metals WHERE depositor_id = ?", (depositor_id,)).fetchone()["n"]
        + db.execute("SELECT COUNT(*) AS n FROM investments WHERE depositor_id = ?", (depositor_id,)).fetchone()["n"]
    )
    if in_use == 0:
        db.execute("DELETE FROM depositors WHERE id = ?", (depositor_id,))
        db.commit()
    return redirect(url_for("depositors_page"))


@app.route("/banks", methods=["GET", "POST"])
def banks_page():
    db = get_db()
    error = None
    form_data = {"bank_id": "", "name": ""}

    if request.method == "POST":
        form_data["bank_id"] = request.form.get("bank_id", "").strip()
        form_data["name"] = request.form.get("name", "").strip()
        try:
            if not form_data["bank_id"]:
                raise ValueError("Bank ID is required.")
            if not form_data["name"]:
                raise ValueError("Bank name is required.")
            existing = db.execute(
                "SELECT id FROM banks WHERE bank_id = ?", (form_data["bank_id"],)
            ).fetchone()
            if existing:
                raise ValueError(f"A bank with ID '{form_data['bank_id']}' already exists.")
            db.execute(
                "INSERT INTO banks (bank_id, name) VALUES (?, ?)",
                (form_data["bank_id"], form_data["name"]),
            )
            db.commit()
            return redirect(url_for("banks_page"))
        except ValueError as e:
            error = str(e)

    return render_template(
        "banks.html", banks=list_banks(db), error=error,
        form_data=form_data, active_tab="banks",
    )


@app.route("/banks/<int:bank_id>/delete", methods=["POST"])
def delete_bank(bank_id):
    db = get_db()
    in_use = db.execute(
        "SELECT COUNT(*) AS n FROM deposits WHERE bank_ref_id = ?", (bank_id,)
    ).fetchone()["n"]
    if in_use == 0:
        db.execute("DELETE FROM banks WHERE id = ?", (bank_id,))
        db.commit()
    return redirect(url_for("banks_page"))


@app.route("/delete/<int:deposit_id>", methods=["POST"])
def delete_deposit(deposit_id):
    db = get_db()
    db.execute("DELETE FROM deposits WHERE id = ?", (deposit_id,))
    db.commit()
    return redirect(url_for("dashboard"))


# ---------- Maturity email notifications ----------
# Don't re-alert on the same deposit more than once a week, so an "enabled"
# check that runs every few hours doesn't spam the same reminder daily.
NOTIFY_RESEND_COOLDOWN_DAYS = 7


def get_notification_settings(db) -> dict:
    row = db.execute("SELECT * FROM notification_settings WHERE id = 1").fetchone()
    return dict(row) if row else {}


def deposits_within_window(db, days_before: int):
    """Un-matured deposits maturing within `days_before` days, each tagged
    with whether it's actually due for an alert (i.e. not in the resend
    cooldown) so the page can preview upcoming maturities honestly."""
    today = date.today()
    rows = []
    for d in db.execute(DEPOSITS_WITH_REFS).fetchall():
        s = summarise_deposit(d)
        if s["is_matured"] or s["days_remaining"] > days_before:
            continue
        last = d["last_notified_on"]
        in_cooldown = bool(last) and (today - date.fromisoformat(last)).days < NOTIFY_RESEND_COOLDOWN_DAYS
        s["last_notified_on"] = last
        s["alert_due"] = not in_cooldown
        rows.append(s)
    return sorted(rows, key=lambda r: r["days_remaining"])


def deposits_due_for_alert(db, days_before: int):
    return [r for r in deposits_within_window(db, days_before) if r["alert_due"]]


def build_maturity_email(rows) -> tuple:
    """Returns (subject, plain-text body) for a digest of deposits due."""
    n = len(rows)
    subject = f"\U0001F3E6 {n} fixed deposit{'s' if n != 1 else ''} maturing soon"
    lines = [f"{n} deposit(s) are approaching maturity:", ""]
    for r in rows:
        lines.append(
            f"- {r['holder_name'] or 'Unknown holder'} / {r['bank_name'] or 'Unknown bank'}: "
            f"{format_rupees(r['maturity_amount'])} matures {r['maturity_date']} "
            f"({r['days_remaining']} day{'s' if r['days_remaining'] != 1 else ''} left)"
        )
    lines += ["", "— sent automatically by Fixed Deposit Manager"]
    return subject, "\n".join(lines)


def send_email(settings: dict, subject: str, body: str) -> None:
    """Send via Gmail SMTP with an App Password. Raises RuntimeError with a
    user-facing message on any failure."""
    if not settings.get("sender_email") or not settings.get("sender_app_password"):
        raise RuntimeError("Sender email / app password isn't set.")
    if not settings.get("recipient_email"):
        raise RuntimeError("Recipient email isn't set.")

    msg = MIMEText(body)
    msg["Subject"] = subject
    msg["From"] = settings["sender_email"]
    msg["To"] = settings["recipient_email"]

    try:
        with smtplib.SMTP("smtp.gmail.com", 587, timeout=15) as server:
            server.starttls()
            server.login(settings["sender_email"], settings["sender_app_password"])
            server.sendmail(settings["sender_email"], [settings["recipient_email"]], msg.as_string())
    except smtplib.SMTPAuthenticationError:
        raise RuntimeError("Gmail rejected the sender email / app password.")
    except socket.gaierror:
        raise RuntimeError(
            "Could not look up smtp.gmail.com — the machine running this app "
            "doesn't seem to have a working internet/DNS connection right now. "
            "Check your Wi-Fi/network and try again."
        )
    except (smtplib.SMTPException, OSError) as e:
        raise RuntimeError(f"Could not send email ({e}).")


def _record_check_result(db, result: str) -> None:
    db.execute(
        "UPDATE notification_settings SET last_check_at = ?, last_check_result = ? WHERE id = 1",
        (datetime.now().isoformat(timespec="seconds"), result),
    )
    db.commit()


def run_maturity_check(db) -> str:
    """Email a digest of any deposits newly due for a maturity reminder.
    Always updates last_check_at / last_check_result. Returns the result."""
    settings = get_notification_settings(db)
    if not settings.get("enabled"):
        result = "Notifications are turned off."
        _record_check_result(db, result)
        return result

    due = deposits_due_for_alert(db, settings["days_before"])
    if not due:
        result = "No deposits due for a reminder."
        _record_check_result(db, result)
        return result

    subject, body = build_maturity_email(due)
    try:
        send_email(settings, subject, body)
    except RuntimeError as e:
        result = f"Failed to send: {e}"
        _record_check_result(db, result)
        return result

    today = str(date.today())
    for r in due:
        db.execute("UPDATE deposits SET last_notified_on = ? WHERE id = ?", (today, r["id"]))
    result = f"Emailed a reminder for {len(due)} deposit(s)."
    _record_check_result(db, result)
    return result


@app.route("/notifications")
def notifications_page():
    db = get_db()
    settings = get_notification_settings(db)
    upcoming = deposits_within_window(db, settings["days_before"])
    return render_template(
        "notifications.html",
        settings=settings,
        upcoming=upcoming,
        cooldown_days=NOTIFY_RESEND_COOLDOWN_DAYS,
        active_tab="notifications",
        error=None,
    )


@app.route("/notifications/settings", methods=["POST"])
def save_notification_settings():
    db = get_db()
    current = get_notification_settings(db)

    enabled = 1 if request.form.get("enabled") == "on" else 0
    recipient_email = request.form.get("recipient_email", "").strip()
    sender_email = request.form.get("sender_email", "").strip()
    sender_app_password = request.form.get("sender_app_password", "").strip()
    # Leaving the password field blank keeps whatever is already saved,
    # so the page never has to (and never does) echo it back into the HTML.
    if not sender_app_password:
        sender_app_password = current.get("sender_app_password", "")

    error = None
    try:
        days_before = int(request.form.get("days_before", "").strip())
        if days_before <= 0:
            raise ValueError
    except ValueError:
        days_before = current.get("days_before", 30)
        error = "Alert window must be a whole number of days greater than 0."

    if error is None and enabled:
        if not recipient_email:
            error = "Recipient email is required to turn notifications on."
        elif not sender_email or not sender_app_password:
            error = "Sender Gmail address and app password are required to turn notifications on."

    if error:
        upcoming = deposits_within_window(db, days_before)
        return render_template(
            "notifications.html",
            settings={
                "enabled": enabled, "recipient_email": recipient_email,
                "sender_email": sender_email, "sender_app_password": sender_app_password,
                "days_before": days_before,
                "last_check_at": current.get("last_check_at"),
                "last_check_result": current.get("last_check_result"),
            },
            upcoming=upcoming,
            cooldown_days=NOTIFY_RESEND_COOLDOWN_DAYS,
            active_tab="notifications",
            error=error,
        )

    db.execute(
        """UPDATE notification_settings SET
             enabled = ?, recipient_email = ?, sender_email = ?,
             sender_app_password = ?, days_before = ?
           WHERE id = 1""",
        (enabled, recipient_email, sender_email, sender_app_password, days_before),
    )
    db.commit()
    return redirect(url_for("notifications_page"))


@app.route("/notifications/check", methods=["POST"])
def check_notifications_now():
    run_maturity_check(get_db())
    return redirect(url_for("notifications_page"))


@app.route("/notifications/test", methods=["POST"])
def send_test_email():
    db = get_db()
    settings = get_notification_settings(db)
    try:
        send_email(
            settings,
            "\U0001F3E6 FD Manager test email",
            "This is a test email from your Fixed Deposit Manager app.\n\n"
            "If you're reading this, your email settings are working.",
        )
        result = "Test email sent successfully."
    except RuntimeError as e:
        result = f"Test email failed: {e}"
    _record_check_result(db, result)
    return redirect(url_for("notifications_page"))


def _background_maturity_loop(interval_seconds: int = 12 * 60 * 60):
    """Runs for the lifetime of the process, checking periodically so
    reminders go out even if nobody opens the Notifications tab."""
    while True:
        try:
            conn = sqlite3.connect(DB_PATH)
            conn.row_factory = sqlite3.Row
            run_maturity_check(conn)
            conn.close()
        except Exception:
            pass  # never let a background hiccup take the app down
        time.sleep(interval_seconds)


def start_background_maturity_checker():
    """Start the loop once per real process — Flask's debug reloader forks a
    child with WERKZEUG_RUN_MAIN=true, so gate on that (or no debug at all)
    to avoid two loops running when the reloader is active."""
    if os.environ.get("WERKZEUG_RUN_MAIN") == "true" or not app.debug:
        threading.Thread(target=_background_maturity_loop, daemon=True).start()


# ---------- Metals ----------
def get_metal_prices(db) -> dict:
    """metal key -> {"price": per-gram rate, "updated_on": iso date, "source": manual/live}."""
    return {
        r["metal"]: {"price": r["price_per_gram"], "updated_on": r["updated_on"], "source": r["source"]}
        for r in db.execute("SELECT * FROM metal_prices").fetchall()
    }


# Spot-price API (gold-api.com, no key required) symbol for each metal we can
# fetch automatically; troy-ounce quotes are converted to rupees per gram.
METAL_API_SYMBOLS = {
    "gold_24k": "XAU",
    "silver": "XAG",
    "platinum": "XPT",
    "palladium": "XPD",
}
TROY_OUNCE_GRAMS = 31.1034768


def _fetch_json_url(url: str) -> dict:
    """GET a URL and parse it as JSON. Raises RuntimeError with a user-facing
    message on any network/parsing failure — shared by every "fetch live
    price" feature in the app."""
    import json
    import urllib.request
    import urllib.error

    req = urllib.request.Request(url, headers={"User-Agent": "fd-manager/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=8) as resp:
            return json.loads(resp.read().decode())
    except urllib.error.URLError as e:
        raise RuntimeError(f"Could not reach {url.split('/')[2]}: {e.reason}")
    except (TimeoutError, OSError) as e:
        raise RuntimeError(f"Could not reach {url.split('/')[2]}: {e}")
    except (ValueError, json.JSONDecodeError):
        raise RuntimeError(f"{url.split('/')[2]} returned an unexpected response.")


def get_usd_to_inr_rate() -> float:
    """Live USD->INR rate via frankfurter.app. Raises RuntimeError on failure."""
    fx = _fetch_json_url("https://api.frankfurter.app/latest?from=USD&to=INR")
    try:
        return float(fx["rates"]["INR"])
    except (KeyError, TypeError, ValueError):
        raise RuntimeError("Could not read the USD/INR exchange rate.")


def fetch_live_metal_prices() -> dict:
    """Fetch spot prices (USD/troy oz) and the USD->INR rate, return
    {metal: price_per_gram_in_rupees}. Raises RuntimeError with a
    user-facing message on any network/parsing failure."""
    usd_to_inr = get_usd_to_inr_rate()

    prices = {}
    for metal, symbol in METAL_API_SYMBOLS.items():
        data = _fetch_json_url(f"https://api.gold-api.com/price/{symbol}")
        try:
            usd_per_oz = float(data["price"])
        except (KeyError, TypeError, ValueError):
            raise RuntimeError(f"Could not read the spot price for {metal}.")
        prices[metal] = usd_per_oz / TROY_OUNCE_GRAMS * usd_to_inr

    # 22K gold is 22/24ths pure; derive it from the 24K spot rate.
    if "gold_24k" in prices:
        prices["gold_22k"] = prices["gold_24k"] * 22 / 24

    return prices


def list_metals(db):
    """Metal holdings with cost, current value and gain/loss. The current price
    per gram comes from the metal_prices table (falling back to the price
    recorded on the holding itself if that metal has no market price set)."""
    prices = get_metal_prices(db)
    rows = []
    for m in db.execute("""
        SELECT metals.*, depositors.name AS depositor_name
        FROM metals LEFT JOIN depositors ON metals.depositor_id = depositors.id
        ORDER BY metals.purchase_date DESC, metals.id DESC
    """).fetchall():
        info = prices.get(m["metal"])
        current_price = info["price"] if info else m["purchase_price"]
        cost = m["grams"] * m["purchase_price"]
        value = m["grams"] * current_price
        gain = value - cost
        days_held = max((date.today() - date.fromisoformat(m["purchase_date"])).days, 0)
        rows.append({
            "id": m["id"],
            "metal": m["metal"],
            "metal_label": METAL_TYPES.get(m["metal"], m["metal"].title()),
            "depositor_id": m["depositor_id"],
            "depositor_name": m["depositor_name"],
            "description": m["description"],
            "grams": m["grams"],
            "purchase_price": m["purchase_price"],
            "current_price": current_price,
            "price_is_market": info is not None,
            "price_updated_on": info["updated_on"] if info else None,
            "purchase_date": m["purchase_date"],
            "cost": cost,
            "value": value,
            "gain": gain,
            "gain_pct": (gain / cost * 100.0) if cost else 0.0,
            "days_held": days_held,
            "annualised_return": annualised_return_pct(cost, value, days_held),
        })
    return rows


def parse_metal_form(form_data, db) -> dict:
    """Validate the holding form. Returns DB column values or raises ValueError."""
    metal = form_data["metal"]
    if metal not in METAL_TYPES:
        raise ValueError("Please choose a metal.")

    depositor_id = None
    if form_data.get("depositor_id"):
        dep = db.execute(
            "SELECT id FROM depositors WHERE id = ?", (form_data["depositor_id"],)
        ).fetchone()
        if dep is None:
            raise ValueError("Please choose a valid depositor.")
        depositor_id = dep["id"]

    try:
        grams = float(form_data["grams"])
        purchase_price = float(form_data["purchase_price"])
    except (TypeError, ValueError):
        raise ValueError("Please enter valid numbers for grams and price.")
    if grams <= 0:
        raise ValueError("Weight in grams must be greater than 0.")
    if purchase_price <= 0:
        raise ValueError("Purchase price must be greater than 0.")
    return {
        "metal": metal,
        "depositor_id": depositor_id,
        "description": form_data["description"].strip(),
        "grams": grams,
        "purchase_price": purchase_price,
        "purchase_date": form_data["purchase_date"] or str(date.today()),
    }


BLANK_METAL_FORM = {
    "metal": "gold_24k", "depositor_id": "", "description": "", "grams": "",
    "purchase_price": "", "purchase_date": None,
}


def _metal_to_form_data(row) -> dict:
    return {
        "metal": row["metal"],
        "depositor_id": str(row["depositor_id"] or ""),
        "description": row["description"],
        "grams": _trim_number(row["grams"]),
        "purchase_price": _trim_number(row["purchase_price"]),
        "purchase_date": row["purchase_date"],
    }


def _render_metals(db, **kwargs):
    metals = list_metals(db)
    prices = get_metal_prices(db)
    return render_template(
        "metals.html",
        metals=metals,
        metal_types=METAL_TYPES,
        metal_prices=prices,
        depositors=db.execute(
            "SELECT id, name FROM depositors ORDER BY name COLLATE NOCASE"
        ).fetchall(),
        prices_missing=sorted({m["metal"] for m in metals if not m["price_is_market"]}),
        total_cost=sum(m["cost"] for m in metals),
        total_value=sum(m["value"] for m in metals),
        total_gain=sum(m["gain"] for m in metals),
        total_annualised_return=weighted_annualised_return(metals, "cost", "value", "days_held"),
        active_tab="metals",
        wide_page=True,
        **kwargs,
    )


def _render_metal_prices(db, **kwargs):
    metals = list_metals(db)
    return render_template(
        "metal_prices.html",
        metal_types=METAL_TYPES,
        metal_prices=get_metal_prices(db),
        live_metal_keys=set(METAL_API_SYMBOLS) | {"gold_22k"},
        prices_missing=sorted({m["metal"] for m in metals if not m["price_is_market"]}),
        active_tab="metal_prices",
        **kwargs,
    )


@app.route("/metal-prices")
def metal_prices_page():
    return _render_metal_prices(get_db(), fetch_error=None)


@app.route("/metals", methods=["GET", "POST"])
def metals_page():
    db = get_db()
    error = None
    form_data = dict(BLANK_METAL_FORM)
    form_data["purchase_date"] = str(date.today())

    if request.method == "POST":
        for key in form_data:
            form_data[key] = request.form.get(key, form_data[key])
        try:
            cols = parse_metal_form(form_data, db)
            # Freeze a current_price on the row (legacy column); the live value
            # comes from metal_prices, but seed it with the market rate if known.
            market = get_metal_prices(db).get(cols["metal"])
            cols["current_price"] = market["price"] if market else cols["purchase_price"]
            db.execute(
                """INSERT INTO metals
                   (metal, depositor_id, description, grams, purchase_price, current_price, purchase_date)
                   VALUES (:metal, :depositor_id, :description, :grams, :purchase_price, :current_price, :purchase_date)""",
                cols,
            )
            db.commit()
            return redirect(url_for("metals_page"))
        except ValueError as e:
            error = str(e)

    return _render_metals(db, error=error, form_data=form_data, editing=None)


@app.route("/metals/prices", methods=["POST"])
def update_metal_prices():
    db = get_db()
    today = str(date.today())
    for metal in METAL_TYPES:
        raw = request.form.get(f"price_{metal}", "").strip()
        if raw == "":
            continue
        try:
            price = float(raw)
        except ValueError:
            continue
        if price <= 0:
            continue
        db.execute(
            """INSERT INTO metal_prices (metal, price_per_gram, updated_on, source)
               VALUES (?, ?, ?, 'manual')
               ON CONFLICT(metal) DO UPDATE SET price_per_gram = excluded.price_per_gram,
                                                updated_on = excluded.updated_on,
                                                source = excluded.source""",
            (metal, price, today),
        )
    db.commit()
    return redirect(url_for("metal_prices_page"))


@app.route("/metals/prices/fetch", methods=["POST"])
def fetch_metal_prices():
    db = get_db()
    try:
        live_prices = fetch_live_metal_prices()
    except RuntimeError as e:
        return _render_metal_prices(db, fetch_error=str(e))

    today = str(date.today())
    for metal, price in live_prices.items():
        db.execute(
            """INSERT INTO metal_prices (metal, price_per_gram, updated_on, source)
               VALUES (?, ?, ?, 'live')
               ON CONFLICT(metal) DO UPDATE SET price_per_gram = excluded.price_per_gram,
                                                updated_on = excluded.updated_on,
                                                source = excluded.source""",
            (metal, price, today),
        )
    db.commit()
    return redirect(url_for("metal_prices_page"))


@app.route("/metals/<int:metal_id>/edit", methods=["GET", "POST"])
def edit_metal(metal_id):
    db = get_db()
    row = db.execute("SELECT * FROM metals WHERE id = ?", (metal_id,)).fetchone()
    if row is None:
        return redirect(url_for("metals_page"))

    error = None
    form_data = _metal_to_form_data(row)

    if request.method == "POST":
        for key in form_data:
            form_data[key] = request.form.get(key, form_data[key])
        try:
            cols = parse_metal_form(form_data, db)
            cols["id"] = metal_id
            db.execute(
                """UPDATE metals SET
                     metal = :metal, depositor_id = :depositor_id, description = :description,
                     grams = :grams, purchase_price = :purchase_price, purchase_date = :purchase_date
                   WHERE id = :id""",
                cols,
            )
            db.commit()
            return redirect(url_for("metals_page"))
        except ValueError as e:
            error = str(e)

    return _render_metals(db, error=error, form_data=form_data, editing=metal_id)


@app.route("/metals/<int:metal_id>/delete", methods=["POST"])
def delete_metal(metal_id):
    db = get_db()
    db.execute("DELETE FROM metals WHERE id = ?", (metal_id,))
    db.commit()
    return redirect(url_for("metals_page"))


# ---------- Ticker directory (NSE stocks + AMFI mutual funds, for the Investments search box) ----------
TICKER_CACHE_DIR = Path(__file__).parent / "cache"
NSE_EQUITY_LIST_URL = "https://archives.nseindia.com/content/equities/EQUITY_L.csv"
MF_SCHEME_LIST_URL = "https://api.mfapi.in/mf"
TICKER_CACHE_MAX_AGE_DAYS = 30
MF_TICKER_PREFIX = "MF:"

_ticker_index = None          # list of {"ticker", "name", "type"} — lazy-built, process-lifetime cache
_ticker_name_lookup = None    # ticker -> name, built alongside the index


def _fetch_text_url(url: str) -> str:
    """GET a URL and return decoded text. Raises RuntimeError on failure —
    same contract as _fetch_json_url, just without the JSON parse."""
    import urllib.request
    import urllib.error

    req = urllib.request.Request(url, headers={"User-Agent": "fd-manager/1.0"})
    try:
        with urllib.request.urlopen(req, timeout=15) as resp:
            return resp.read().decode()
    except urllib.error.URLError as e:
        raise RuntimeError(f"Could not reach {url.split('/')[2]}: {e.reason}")
    except (TimeoutError, OSError) as e:
        raise RuntimeError(f"Could not reach {url.split('/')[2]}: {e}")


def _read_ticker_cache(filename: str):
    """Returns the cached list if the file exists and isn't stale, else None."""
    import json

    path = TICKER_CACHE_DIR / filename
    if not path.exists():
        return None
    try:
        payload = json.loads(path.read_text())
        fetched_at = date.fromisoformat(payload["fetched_at"])
        if (date.today() - fetched_at).days > TICKER_CACHE_MAX_AGE_DAYS:
            return None
        return payload["items"]
    except Exception:
        return None


def _write_ticker_cache(filename: str, items: list):
    import json

    TICKER_CACHE_DIR.mkdir(exist_ok=True)
    path = TICKER_CACHE_DIR / filename
    path.write_text(json.dumps({"fetched_at": str(date.today()), "items": items}))


def _load_nse_stocks() -> list:
    """NSE-listed equities as {"ticker": "SYMBOL.NS", "name", "type": "stock"}.
    Cached to disk for TICKER_CACHE_MAX_AGE_DAYS; falls back to a stale cache
    (rather than an empty list) if a refresh attempt fails offline."""
    import csv
    import io

    cached = _read_ticker_cache("nse_stocks.json")
    if cached is not None:
        return cached
    try:
        text = _fetch_text_url(NSE_EQUITY_LIST_URL)
        reader = csv.DictReader(io.StringIO(text))
        items = [
            {"ticker": f"{row['SYMBOL'].strip()}.NS", "name": row["NAME OF COMPANY"].strip(), "type": "stock"}
            for row in reader if row.get("SYMBOL")
        ]
        _write_ticker_cache("nse_stocks.json", items)
        return items
    except Exception:
        path = TICKER_CACHE_DIR / "nse_stocks.json"
        if path.exists():
            import json
            return json.loads(path.read_text())["items"]
        return []


def _load_mf_schemes() -> list:
    """AMFI-registered mutual fund schemes (via mfapi.in) as
    {"ticker": "MF:<schemeCode>", "name", "type": "mf"}. Cached like stocks."""
    import json

    cached = _read_ticker_cache("mf_schemes.json")
    if cached is not None:
        return cached
    try:
        schemes = json.loads(_fetch_text_url(MF_SCHEME_LIST_URL))
        items = [
            {"ticker": f"{MF_TICKER_PREFIX}{s['schemeCode']}", "name": s["schemeName"].strip(), "type": "mf"}
            for s in schemes if s.get("schemeName")
        ]
        _write_ticker_cache("mf_schemes.json", items)
        return items
    except Exception:
        path = TICKER_CACHE_DIR / "mf_schemes.json"
        if path.exists():
            return json.loads(path.read_text())["items"]
        return []


def get_ticker_index() -> list:
    """Combined NSE stock + AMFI mutual fund directory, built once per process."""
    global _ticker_index, _ticker_name_lookup
    if _ticker_index is None:
        _ticker_index = _load_nse_stocks() + _load_mf_schemes()
        _ticker_name_lookup = {item["ticker"]: item["name"] for item in _ticker_index}
    return _ticker_index


def get_ticker_display_name(ticker: str):
    """Company / scheme name for a ticker already in the directory, else None."""
    get_ticker_index()  # ensures _ticker_name_lookup is built
    return _ticker_name_lookup.get(ticker)


def search_tickers(query: str, limit: int = 25) -> list:
    """Prefix matches first, then substring matches, across ticker + name."""
    q = query.strip().lower()
    if not q:
        return []
    prefix_hits, other_hits = [], []
    for item in get_ticker_index():
        ticker_l, name_l = item["ticker"].lower(), item["name"].lower()
        if ticker_l.startswith(q) or name_l.startswith(q):
            prefix_hits.append(item)
        elif q in ticker_l or q in name_l:
            other_hits.append(item)
        if len(prefix_hits) >= limit:
            break
    return (prefix_hits + other_hits)[:limit]


@app.route("/api/tickers/search")
def api_ticker_search():
    from flask import jsonify
    return jsonify(search_tickers(request.args.get("q", ""), limit=25))


# ---------- Investments (stocks / ETFs / mutual funds) ----------
def get_mf_quote(scheme_code: str) -> dict:
    """Latest NAV for an AMFI scheme code via mfapi.in. Never raises."""
    import json

    try:
        payload = json.loads(_fetch_text_url(f"https://api.mfapi.in/mf/{scheme_code}/latest"))
        nav_rows = payload.get("data") or []
        if not nav_rows:
            return {"price": None, "currency": None, "error": "No NAV data found for this scheme."}
        return {"price": float(nav_rows[0]["nav"]), "currency": "INR", "error": None}
    except RuntimeError as e:
        return {"price": None, "currency": None, "error": str(e)}
    except Exception as e:
        return {"price": None, "currency": None, "error": str(e)[:200]}


def get_stock_quote(ticker: str) -> dict:
    """Latest price for a ticker via yfinance (or mfapi.in for MF: scheme
    codes). Never raises — always returns
    {"price": float|None, "currency": str|None, "error": str|None}."""
    if ticker.startswith(MF_TICKER_PREFIX):
        return get_mf_quote(ticker[len(MF_TICKER_PREFIX):])
    if not YFINANCE_AVAILABLE:
        return {"price": None, "currency": None, "error": "yfinance isn't installed."}
    try:
        t = yf.Ticker(ticker)
        hist = t.history(period="1d")
        if hist.empty:
            return {"price": None, "currency": None, "error": "No price data found for this ticker."}
        price = float(hist["Close"].iloc[-1])
        try:
            currency = t.fast_info.get("currency")
        except Exception:
            currency = None
        return {"price": price, "currency": currency, "error": None}
    except Exception as e:
        return {"price": None, "currency": None, "error": str(e)[:200]}


def quote_to_inr(quote: dict, usd_rate: float = None) -> dict:
    """Resolves a get_stock_quote()/get_mf_quote() result to a rupee price.
    Pass a pre-fetched usd_rate to avoid refetching it for every row in a
    batch; omit it to fetch on demand for a single lookup. Returns
    {"price": float|None, "price_note": str|None}."""
    if quote["error"]:
        return {"price": None, "price_note": quote["error"]}
    currency = quote["currency"]
    if currency in (None, "INR"):
        return {"price": quote["price"], "price_note": None}
    if currency == "USD":
        if usd_rate is None:
            try:
                usd_rate = get_usd_to_inr_rate()
            except RuntimeError as e:
                return {"price": None, "price_note": f"USD rate unavailable ({e})"}
        return {"price": quote["price"] * usd_rate, "price_note": None}
    return {"price": None, "price_note": f"priced in {currency}, not converted to ₹"}


@app.route("/api/tickers/quote")
def api_ticker_quote():
    from flask import jsonify
    ticker = request.args.get("ticker", "").strip().upper()
    if not ticker:
        return jsonify({"price": None, "price_note": "No ticker given."})
    return jsonify(quote_to_inr(get_stock_quote(ticker)))


def list_investments(db):
    """Investment holdings with cost, current value (converted to rupees) and
    gain/loss. Prices are fetched live via yfinance on every call — a USD
    quote is converted with one shared USD->INR lookup per render; other
    currencies are shown un-converted with a note rather than guessed at."""
    usd_rate = None
    usd_rate_error = None
    rows = []

    for h in db.execute("""
        SELECT investments.*, depositors.name AS depositor_name
        FROM investments LEFT JOIN depositors ON investments.depositor_id = depositors.id
        ORDER BY investments.ticker
    """).fetchall():
        quote = get_stock_quote(h["ticker"])
        currency = quote["currency"]

        if currency == "USD" and usd_rate is None and usd_rate_error is None:
            try:
                usd_rate = get_usd_to_inr_rate()
            except RuntimeError as e:
                usd_rate_error = str(e)

        if currency == "USD" and usd_rate is None:
            current_price, price_note = None, f"USD rate unavailable ({usd_rate_error})"
        else:
            resolved = quote_to_inr(quote, usd_rate=usd_rate)
            current_price, price_note = resolved["price"], resolved["price_note"]

        cost = h["shares"] * h["purchase_price"]
        value = h["shares"] * current_price if current_price is not None else None
        gain = (value - cost) if value is not None else None
        gain_pct = (gain / cost * 100.0) if (gain is not None and cost) else None
        days_held = max((date.today() - date.fromisoformat(h["purchase_date"])).days, 0)
        annualised_return = annualised_return_pct(cost, value, days_held) if value is not None else None

        rows.append({
            "id": h["id"],
            "ticker": h["ticker"],
            "display_name": get_ticker_display_name(h["ticker"]),
            "depositor_id": h["depositor_id"],
            "depositor_name": h["depositor_name"],
            "shares": h["shares"],
            "purchase_price": h["purchase_price"],
            "purchase_date": h["purchase_date"],
            "current_price": current_price,
            "currency": currency,
            "price_note": price_note,
            "cost": cost,
            "value": value,
            "gain": gain,
            "gain_pct": gain_pct,
            "days_held": days_held,
            "annualised_return": annualised_return,
        })
    return rows


def parse_investment_form(form_data, db) -> dict:
    """Validate the holding form. Returns DB column values or raises ValueError."""
    ticker = form_data["ticker"].strip().upper()
    if not ticker:
        raise ValueError("Ticker symbol is required.")

    if not form_data.get("depositor_id"):
        raise ValueError("Please choose a depositor.")
    dep = db.execute(
        "SELECT id FROM depositors WHERE id = ?", (form_data["depositor_id"],)
    ).fetchone()
    if dep is None:
        raise ValueError("Please choose a valid depositor.")
    depositor_id = dep["id"]

    try:
        shares = float(form_data["shares"])
        purchase_price = float(form_data["purchase_price"])
    except (TypeError, ValueError):
        raise ValueError("Please enter valid numbers for shares and price.")
    if shares <= 0:
        raise ValueError("Shares must be greater than 0.")
    if purchase_price <= 0:
        raise ValueError("Purchase price must be greater than 0.")

    return {
        "ticker": ticker,
        "depositor_id": depositor_id,
        "shares": shares,
        "purchase_price": purchase_price,
        "purchase_date": form_data["purchase_date"] or str(date.today()),
    }


BLANK_INVESTMENT_FORM = {
    "ticker": "", "depositor_id": "", "shares": "", "purchase_price": "", "purchase_date": None,
}


def _investment_to_form_data(row) -> dict:
    return {
        "ticker": row["ticker"],
        "depositor_id": str(row["depositor_id"] or ""),
        "shares": _trim_number(row["shares"]),
        "purchase_price": _trim_number(row["purchase_price"]),
        "purchase_date": row["purchase_date"],
    }


def _render_investments(db, **kwargs):
    investments = list_investments(db)
    # Cost/value/gain tiles are all computed over the same "priced" subset,
    # so they stay internally consistent (value - cost == gain). A holding
    # with no live price still shows its own cost/N-A in the table below,
    # it just isn't folded into these totals until it prices successfully.
    priced = [r for r in investments if r["value"] is not None]
    return render_template(
        "investments.html",
        investments=investments,
        depositors=db.execute(
            "SELECT id, name FROM depositors ORDER BY name COLLATE NOCASE"
        ).fetchall(),
        total_cost=sum(r["cost"] for r in priced),
        total_value=sum(r["value"] for r in priced),
        total_gain=sum(r["gain"] for r in priced),
        total_annualised_return=weighted_annualised_return(priced, "cost", "value", "days_held"),
        priced_count=len(priced),
        total_count=len(investments),
        yfinance_available=YFINANCE_AVAILABLE,
        active_tab="investments",
        wide_page=True,
        **kwargs,
    )


@app.route("/investments", methods=["GET", "POST"])
def investments_page():
    db = get_db()
    error = None
    form_data = dict(BLANK_INVESTMENT_FORM)
    form_data["purchase_date"] = str(date.today())

    if request.method == "POST":
        for key in form_data:
            form_data[key] = request.form.get(key, form_data[key])
        try:
            cols = parse_investment_form(form_data, db)
            db.execute(
                """INSERT INTO investments (ticker, depositor_id, shares, purchase_price, purchase_date)
                   VALUES (:ticker, :depositor_id, :shares, :purchase_price, :purchase_date)""",
                cols,
            )
            db.commit()
            return redirect(url_for("investments_page"))
        except ValueError as e:
            error = str(e)

    return _render_investments(db, error=error, form_data=form_data, editing=None)


@app.route("/investments/<int:investment_id>/edit", methods=["GET", "POST"])
def edit_investment(investment_id):
    db = get_db()
    row = db.execute("SELECT * FROM investments WHERE id = ?", (investment_id,)).fetchone()
    if row is None:
        return redirect(url_for("investments_page"))

    error = None
    form_data = _investment_to_form_data(row)

    if request.method == "POST":
        for key in form_data:
            form_data[key] = request.form.get(key, form_data[key])
        try:
            cols = parse_investment_form(form_data, db)
            cols["id"] = investment_id
            db.execute(
                """UPDATE investments SET
                     ticker = :ticker, depositor_id = :depositor_id, shares = :shares,
                     purchase_price = :purchase_price, purchase_date = :purchase_date
                   WHERE id = :id""",
                cols,
            )
            db.commit()
            return redirect(url_for("investments_page"))
        except ValueError as e:
            error = str(e)

    return _render_investments(db, error=error, form_data=form_data, editing=investment_id)


@app.route("/investments/<int:investment_id>/delete", methods=["POST"])
def delete_investment(investment_id):
    db = get_db()
    db.execute("DELETE FROM investments WHERE id = ?", (investment_id,))
    db.commit()
    return redirect(url_for("investments_page"))


if __name__ == "__main__":
    init_db()
    start_background_maturity_checker()
    app.run(debug=True)
