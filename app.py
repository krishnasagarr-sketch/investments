import sqlite3
from datetime import date, timedelta
from pathlib import Path

from flask import Flask, render_template, request, redirect, url_for, g

app = Flask(__name__)

DB_PATH = Path(__file__).parent / "fixed_deposits.db"

# Supported deposit types: internal key -> human label
DEPOSIT_TYPES = {
    "cumulative": "Cumulative (reinvested)",
    "simple": "Simple interest (payout)",
    "recurring": "Recurring deposit",
}

# How a tenure figure is expressed. Recurring deposits are always monthly.
TENURE_UNITS = {"months": "Months", "days": "Days"}

DAYS_PER_YEAR = 365  # simple day-count convention used for interest on day tenures


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
        invested = d["principal"] * d["tenure_months"]
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
    }


# ---------- Master-list helpers ----------
def list_depositors(db):
    return db.execute(
        """SELECT d.id, d.holder_id, d.name, COUNT(dep.id) AS deposit_count
           FROM depositors d
           LEFT JOIN deposits dep ON dep.depositor_id = d.id
           GROUP BY d.id ORDER BY d.name COLLATE NOCASE"""
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
            "total_invested": t["invested"],
            "total_current": t["current"],
            "total_maturity": t["maturity"],
        })
    return rows


def _agg_blank():
    return {"count": 0, "invested": 0.0, "current": 0.0, "maturity": 0.0}


def _agg_add(acc, s):
    acc["count"] += 1
    acc["invested"] += s["invested"]
    acc["current"] += s["current_value"]
    acc["maturity"] += s["maturity_amount"]


def holdings_summary(db):
    """Aggregate every deposit by (holder, bank), by holder, and by bank."""
    by_pair, by_holder, by_bank = {}, {}, {}
    overall = _agg_blank()

    for d in db.execute(DEPOSITS_WITH_REFS).fetchall():
        s = summarise_deposit(d)
        holder = s["holder_name"] or "—"
        bank = s["bank_name"] or "—"
        _agg_add(by_pair.setdefault((holder, bank), _agg_blank()), s)
        _agg_add(by_holder.setdefault(holder, _agg_blank()), s)
        _agg_add(by_bank.setdefault(bank, _agg_blank()), s)
        _agg_add(overall, s)

    pairs = [
        {"holder": h, "bank": b, **v}
        for (h, b), v in sorted(by_pair.items(), key=lambda kv: (kv[0][0].lower(), kv[0][1].lower()))
    ]
    holders = [{"name": h, **v} for h, v in sorted(by_holder.items(), key=lambda kv: kv[0].lower())]
    banks = [{"name": b, **v} for b, v in sorted(by_bank.items(), key=lambda kv: kv[0].lower())]
    return {"pairs": pairs, "holders": holders, "banks": banks, "overall": overall}


# ---------- Routes ----------
@app.route("/summary")
def summary_page():
    db = get_db()
    return render_template(
        "summary.html", active_tab="summary", **holdings_summary(db)
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

    return render_template(
        "dashboard.html",
        rows=rows,
        total_invested=total_invested,
        total_current=total_current,
        total_maturity=total_maturity,
        total_interest=total_interest,
        active_tab="dashboard",
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
    in_use = db.execute(
        "SELECT COUNT(*) AS n FROM deposits WHERE depositor_id = ?", (depositor_id,)
    ).fetchone()["n"]
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


if __name__ == "__main__":
    init_db()
    app.run(debug=True)
