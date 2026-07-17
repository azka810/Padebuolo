
import base64
import io
import json
import os
import sqlite3
import time
from datetime import date, datetime
from pathlib import Path

import pandas as pd
import requests
import streamlit as st

APP_TITLE = "V.6 Padebuolo Fresh"
APP_VERSION = "V.6.3 Beyond Super!!!"
DEFAULT_PASSWORD = "rumdin123"

BASE_DIR = Path(__file__).resolve().parent
DATA_DIR = BASE_DIR / "data"
INSTANCE_DIR = BASE_DIR / "instance"
INSTANCE_DIR.mkdir(exist_ok=True)
DB_PATH = INSTANCE_DIR / "padebuolo.db"

SEED_TX_PATH = DATA_DIR / "seed_transactions.csv"
SEED_BUDGET_PATH = DATA_DIR / "seed_budgets.csv"
SEED_XLSX_PATH = DATA_DIR / "Rekap Kas Rumdin.xlsx"

CATEGORIES = [
    "Saldo Awal",
    "Inject Dana / Top Up",
    "Internet",
    "Laundry",
    "Listrik",
    "Air / PDAM",
    "Sewa",
    "Perlengkapan Rumah",
    "Pemeliharaan",
    "Renovasi",
    "Aset Rumah",
    "Konsumsi",
    "Transport / Parkir",
    "Iuran / Patungan",
    "Reimbursement",
    "Lainnya",
]

FUNDS_DEFAULT = ["Kas Azka", "Kas Rayhan"]
METHODS = ["Kas", "Transfer", "QRIS", "Tunai", "Lainnya"]

BUDGET_STANDARD = [
    {"fund": "Kas Rayhan", "component": "Sewa", "amount": 200000, "note": ""},
    {"fund": "Kas Rayhan", "component": "Indihome", "amount": 165000, "note": ""},
    {"fund": "Kas Rayhan", "component": "Listrik", "amount": 300000, "note": ""},
    {"fund": "Kas Rayhan", "component": "Air", "amount": 50000, "note": ""},
    {"fund": "Kas Azka", "component": "Sewa", "amount": 245000, "note": ""},
    {"fund": "Kas Azka", "component": "Indihome", "amount": 165000, "note": ""},
    {"fund": "Kas Azka", "component": "Listrik", "amount": 300000, "note": ""},
    {"fund": "Kas Azka", "component": "Air", "amount": 50000, "note": ""},
]

KEYWORD_RULES = {
    "Internet": ["wifi", "indihome", "internet"],
    "Laundry": ["laundry", "cuci karpet", "cuci"],
    "Listrik": ["listrik", "pln", "token"],
    "Air / PDAM": ["pdam", "air", "galon"],
    "Sewa": ["sewa", "rumdin"],
    "Perlengkapan Rumah": ["lampu", "kran", "pipa", "lem", "perlengkapan", "sapu", "pel"],
    "Pemeliharaan": ["rumput", "pemeliharaan", "obat rumput", "bensin mesin"],
    "Renovasi": ["cat", "roller", "amplas", "kunci jendela"],
    "Aset Rumah": ["kulkas", "ongkir kulkas", "double tip"],
    "Inject Dana / Top Up": ["inject", "top up", "tambah dana", "setor"],
    "Saldo Awal": ["saldo awal"],
}


# ----------------------
# Helpers
# ----------------------
def get_secret(name, default=None):
    try:
        return st.secrets.get(name, os.environ.get(name, default))
    except Exception:
        return os.environ.get(name, default)


def parse_number(value, default=0):
    if value is None:
        return default
    if isinstance(value, (int, float)):
        return float(value)
    text = str(value).strip()
    if text == "":
        return default
    text = text.replace("Rp", "").replace("rp", "").replace(" ", "")
    if "," in text and "." in text:
        # Indonesian-style 1.234.567,89
        text = text.replace(".", "").replace(",", ".")
    elif "," in text and "." not in text:
        # Could be 123,000 or 123,45. For this app, comma is usually thousands.
        text = text.replace(",", "")
    else:
        text = text.replace(",", "")
    try:
        return float(text)
    except Exception:
        return default


def parse_date_any(value):
    if value is None or str(value).strip() == "":
        return date.today().isoformat()
    if isinstance(value, datetime):
        return value.date().isoformat()
    if isinstance(value, date):
        return value.isoformat()
    if isinstance(value, (int, float)) and not pd.isna(value):
        # Excel serial date
        try:
            dt = pd.Timestamp("1899-12-30") + pd.to_timedelta(int(value), unit="D")
            return dt.date().isoformat()
        except Exception:
            pass

    text = str(value).strip()
    # ISO must remain year-month-day
    try:
        if len(text) >= 10 and text[4] in ["-", "/"] and text[:4].isdigit():
            return pd.to_datetime(text[:10], yearfirst=True, errors="raise").date().isoformat()
    except Exception:
        pass

    # Indonesian format must be day/month/year
    for fmt in ["%d/%m/%Y", "%d-%m-%Y", "%d.%m.%Y", "%d/%m/%y", "%d-%m-%y"]:
        try:
            return datetime.strptime(text, fmt).date().isoformat()
        except Exception:
            pass

    # Fallback dayfirst
    try:
        return pd.to_datetime(text, dayfirst=True, errors="raise").date().isoformat()
    except Exception:
        return date.today().isoformat()


def format_date_id(value):
    try:
        return pd.to_datetime(value, errors="coerce").strftime("%d/%m/%Y")
    except Exception:
        return ""


def format_month_id(period):
    try:
        dt = pd.to_datetime(str(period) + "-01", errors="coerce")
        if pd.isna(dt):
            return str(period)
        month_map = {
            1: "Jan", 2: "Feb", 3: "Mar", 4: "Apr", 5: "Mei", 6: "Jun",
            7: "Jul", 8: "Agu", 9: "Sep", 10: "Okt", 11: "Nov", 12: "Des"
        }
        return f"{month_map[int(dt.month)]} {int(dt.year)}"
    except Exception:
        return str(period)


def rp(value):
    value = parse_number(value, 0)
    return f"Rp{value:,.0f}".replace(",", ".")


def rp_compact(value):
    value = parse_number(value, 0)
    abs_value = abs(value)
    sign = "-" if value < 0 else ""
    if abs_value >= 1_000_000:
        s = f"{abs_value/1_000_000:.1f}".replace(".", ",")
        if s.endswith(",0"):
            s = s[:-2]
        return f"{sign}Rp{s} jt"
    if abs_value >= 1_000:
        s = f"{abs_value/1_000:.0f}"
        return f"{sign}Rp{s} rb"
    return f"{sign}Rp{abs_value:.0f}"


def df_display(df, money_cols=None, date_cols=None):
    out = df.copy()
    money_cols = money_cols or []
    date_cols = date_cols or []
    for col in date_cols:
        if col in out.columns:
            out[col] = out[col].apply(format_date_id)
    for col in money_cols:
        if col in out.columns:
            out[col] = out[col].apply(rp)
    return out


def show_df(df, height=None):
    """Render dataframe safely across Streamlit versions.

    Newer Streamlit versions can error when height=None is passed explicitly,
    so height is only sent when it has a real value.
    """
    kwargs = {"use_container_width": True, "hide_index": True}
    if height is not None:
        kwargs["height"] = height

    try:
        st.dataframe(df, **kwargs)
    except TypeError:
        kwargs.pop("hide_index", None)
        st.dataframe(df, **kwargs)


def safe_toast(msg):
    try:
        st.toast(msg)
    except Exception:
        st.success(msg)


def normalize_text(s):
    return str(s or "").strip()


def auto_category(description):
    text = normalize_text(description).lower()
    for cat, keywords in KEYWORD_RULES.items():
        if any(k in text for k in keywords):
            return cat
    return "Lainnya"


# ----------------------
# SQLite
# ----------------------
def get_conn():
    conn = sqlite3.connect(DB_PATH, check_same_thread=False)
    conn.row_factory = sqlite3.Row
    return conn


def init_db():
    with get_conn() as conn:
        conn.execute("""
            CREATE TABLE IF NOT EXISTS transactions (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                date TEXT NOT NULL,
                fund TEXT NOT NULL,
                type TEXT NOT NULL,
                amount REAL NOT NULL DEFAULT 0,
                category TEXT NOT NULL DEFAULT 'Lainnya',
                description TEXT NOT NULL DEFAULT '',
                method TEXT NOT NULL DEFAULT 'Kas',
                note TEXT NOT NULL DEFAULT '',
                created_at TEXT NOT NULL,
                updated_at TEXT NOT NULL
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS budgets (
                id INTEGER PRIMARY KEY AUTOINCREMENT,
                fund TEXT NOT NULL,
                component TEXT NOT NULL,
                amount REAL NOT NULL DEFAULT 0,
                note TEXT NOT NULL DEFAULT ''
            )
        """)
        conn.execute("""
            CREATE TABLE IF NOT EXISTS meta (
                key TEXT PRIMARY KEY,
                value TEXT
            )
        """)
        conn.commit()


def table_count(table):
    with get_conn() as conn:
        return conn.execute(f"SELECT COUNT(*) AS n FROM {table}").fetchone()["n"]


def get_meta(key, default=None):
    with get_conn() as conn:
        row = conn.execute("SELECT value FROM meta WHERE key=?", (key,)).fetchone()
        return row["value"] if row else default


def set_meta(key, value):
    with get_conn() as conn:
        conn.execute(
            "INSERT INTO meta(key, value) VALUES(?, ?) ON CONFLICT(key) DO UPDATE SET value=excluded.value",
            (key, str(value)),
        )
        conn.commit()


def read_transactions():
    with get_conn() as conn:
        df = pd.read_sql_query("SELECT * FROM transactions ORDER BY date, id", conn)
    if df.empty:
        return pd.DataFrame(columns=["id", "date", "fund", "type", "amount", "category", "description", "method", "note", "created_at", "updated_at"])
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    return df


def read_budgets():
    with get_conn() as conn:
        df = pd.read_sql_query("SELECT * FROM budgets ORDER BY fund, id", conn)
    if df.empty:
        return pd.DataFrame(columns=["id", "fund", "component", "amount", "note"])
    df["amount"] = pd.to_numeric(df["amount"], errors="coerce").fillna(0)
    return df


def replace_transactions(df):
    now = datetime.utcnow().isoformat(timespec="seconds")
    rows = []
    for _, r in df.iterrows():
        rows.append((
            parse_date_any(r.get("date")),
            normalize_text(r.get("fund") or "Kas Azka"),
            normalize_text(r.get("type") or "Keluar"),
            parse_number(r.get("amount")),
            normalize_text(r.get("category") or "Lainnya"),
            normalize_text(r.get("description") or r.get("keterangan") or ""),
            normalize_text(r.get("method") or "Kas"),
            normalize_text(r.get("note") or r.get("catatan") or ""),
            now,
            now,
        ))
    with get_conn() as conn:
        conn.execute("DELETE FROM transactions")
        conn.executemany("""
            INSERT INTO transactions(date, fund, type, amount, category, description, method, note, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)
        conn.commit()


def append_transactions(df):
    now = datetime.utcnow().isoformat(timespec="seconds")
    rows = []
    for _, r in df.iterrows():
        rows.append((
            parse_date_any(r.get("date")),
            normalize_text(r.get("fund") or "Kas Azka"),
            normalize_text(r.get("type") or "Keluar"),
            parse_number(r.get("amount")),
            normalize_text(r.get("category") or "Lainnya"),
            normalize_text(r.get("description") or r.get("keterangan") or ""),
            normalize_text(r.get("method") or "Kas"),
            normalize_text(r.get("note") or r.get("catatan") or ""),
            now,
            now,
        ))
    with get_conn() as conn:
        conn.executemany("""
            INSERT INTO transactions(date, fund, type, amount, category, description, method, note, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, rows)
        conn.commit()


def replace_budgets(df):
    rows = []
    for _, r in df.iterrows():
        rows.append((
            normalize_text(r.get("fund") or "Kas Azka"),
            normalize_text(r.get("component") or r.get("komponen") or ""),
            parse_number(r.get("amount") or r.get("total")),
            normalize_text(r.get("note") or r.get("catatan") or ""),
        ))
    with get_conn() as conn:
        conn.execute("DELETE FROM budgets")
        conn.executemany("INSERT INTO budgets(fund, component, amount, note) VALUES (?, ?, ?, ?)", rows)
        conn.commit()


def reset_standard_budget():
    replace_budgets(pd.DataFrame(BUDGET_STANDARD))


def add_transaction(date_value, fund, trx_type, amount, category, description, method="Kas", note=""):
    now = datetime.utcnow().isoformat(timespec="seconds")
    with get_conn() as conn:
        conn.execute("""
            INSERT INTO transactions(date, fund, type, amount, category, description, method, note, created_at, updated_at)
            VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
        """, (parse_date_any(date_value), fund, trx_type, parse_number(amount), category, description, method, note, now, now))
        conn.commit()
    cloud_auto_backup()


def update_transaction(trx_id, date_value, fund, trx_type, amount, category, description, method, note):
    now = datetime.utcnow().isoformat(timespec="seconds")
    with get_conn() as conn:
        conn.execute("""
            UPDATE transactions
            SET date=?, fund=?, type=?, amount=?, category=?, description=?, method=?, note=?, updated_at=?
            WHERE id=?
        """, (parse_date_any(date_value), fund, trx_type, parse_number(amount), category, description, method, note, now, int(trx_id)))
        conn.commit()
    cloud_auto_backup()


def delete_transaction(trx_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM transactions WHERE id=?", (int(trx_id),))
        conn.commit()
    cloud_auto_backup()


def add_budget(fund, component, amount, note=""):
    with get_conn() as conn:
        conn.execute("INSERT INTO budgets(fund, component, amount, note) VALUES (?, ?, ?, ?)", (fund, component, parse_number(amount), note))
        conn.commit()
    cloud_auto_backup()


def update_budget(budget_id, fund, component, amount, note=""):
    with get_conn() as conn:
        conn.execute("UPDATE budgets SET fund=?, component=?, amount=?, note=? WHERE id=?", (fund, component, parse_number(amount), note, int(budget_id)))
        conn.commit()
    cloud_auto_backup()


def delete_budget(budget_id):
    with get_conn() as conn:
        conn.execute("DELETE FROM budgets WHERE id=?", (int(budget_id),))
        conn.commit()
    cloud_auto_backup()


def seed_from_files(force=False):
    if force or table_count("transactions") == 0:
        if SEED_TX_PATH.exists():
            tx = pd.read_csv(SEED_TX_PATH)
            replace_transactions(tx)
        elif SEED_XLSX_PATH.exists():
            tx, bud = parse_excel_to_dataframes(SEED_XLSX_PATH)
            replace_transactions(tx)
            if table_count("budgets") == 0:
                replace_budgets(bud)
    if force or table_count("budgets") == 0:
        if SEED_BUDGET_PATH.exists():
            replace_budgets(pd.read_csv(SEED_BUDGET_PATH))
        else:
            reset_standard_budget()
    set_meta("last_seed_at", datetime.utcnow().isoformat(timespec="seconds"))


# ----------------------
# Import / Export / Persistence
# ----------------------
def parse_excel_to_dataframes(file):
    master = pd.read_excel(file, sheet_name="Master Kas")
    cols = {str(c).strip().lower(): c for c in master.columns}
    def pick(*names):
        for n in names:
            if n.lower() in cols:
                return cols[n.lower()]
        return None

    date_col = pick("date", "tanggal")
    fund_col = pick("fund", "sumber dana", "sumber_dana")
    desc_col = pick("description", "keterangan")
    cat_col = pick("category", "kategori")
    type_col = pick("type", "jenis")
    amount_col = pick("amount", "jumlah", "nominal")
    masuk_col = pick("masuk")
    keluar_col = pick("keluar")
    method_col = pick("method", "metode")
    note_col = pick("note", "catatan")

    rows = []
    for _, r in master.iterrows():
        if date_col is None or fund_col is None:
            continue
        if pd.isna(r.get(date_col)) and pd.isna(r.get(fund_col)):
            continue
        if type_col and amount_col:
            trx_type = normalize_text(r.get(type_col) or "Keluar")
            amount = parse_number(r.get(amount_col))
        else:
            masuk = parse_number(r.get(masuk_col), 0) if masuk_col else 0
            keluar = parse_number(r.get(keluar_col), 0) if keluar_col else 0
            if masuk > 0 and keluar <= 0:
                trx_type, amount = "Masuk", masuk
            elif keluar > 0 and masuk <= 0:
                trx_type, amount = "Keluar", keluar
            elif masuk == 0 and keluar == 0:
                trx_type, amount = "Keluar", 0
            else:
                if masuk >= keluar:
                    trx_type, amount = "Masuk", masuk - keluar
                else:
                    trx_type, amount = "Keluar", keluar - masuk

        desc = normalize_text(r.get(desc_col) if desc_col else "")
        cat = normalize_text(r.get(cat_col) if cat_col else "") or auto_category(desc)
        rows.append({
            "date": parse_date_any(r.get(date_col)),
            "fund": normalize_text(r.get(fund_col) or "Kas Azka"),
            "type": trx_type if trx_type in ["Masuk", "Keluar"] else ("Masuk" if trx_type.lower().startswith("m") else "Keluar"),
            "amount": amount,
            "category": cat,
            "description": desc,
            "method": normalize_text(r.get(method_col) if method_col else "Kas") or "Kas",
            "note": normalize_text(r.get(note_col) if note_col else ""),
        })

    tx_df = pd.DataFrame(rows)
    if tx_df.empty:
        tx_df = pd.DataFrame(columns=["date", "fund", "type", "amount", "category", "description", "method", "note"])

    bud_df = pd.DataFrame(BUDGET_STANDARD)
    try:
        tmp = pd.read_excel(file, sheet_name="Biaya Bulanan")
        # Keep standard budget unless the sheet contains normalized columns.
        lower_cols = [str(c).lower() for c in tmp.columns]
        if {"fund", "component", "amount"}.issubset(set(lower_cols)):
            tmp.columns = [str(c).lower() for c in tmp.columns]
            bud_df = tmp[["fund", "component", "amount"]].copy()
            bud_df["note"] = tmp["note"] if "note" in tmp.columns else ""
    except Exception:
        pass
    return tx_df, bud_df


def export_payload():
    return {
        "app": APP_TITLE,
        "version": APP_VERSION,
        "exported_at": datetime.utcnow().isoformat(timespec="seconds"),
        "transactions": read_transactions().to_dict(orient="records"),
        "budgets": read_budgets().to_dict(orient="records"),
    }


def import_payload(payload, replace=True):
    tx = pd.DataFrame(payload.get("transactions", []))
    bud = pd.DataFrame(payload.get("budgets", []))
    if replace:
        if not tx.empty:
            replace_transactions(tx)
        if not bud.empty:
            replace_budgets(bud)
    else:
        if not tx.empty:
            append_transactions(tx)
        if not bud.empty:
            current = read_budgets()
            replace_budgets(pd.concat([current[["fund", "component", "amount", "note"]], bud], ignore_index=True))
    set_meta("last_import_at", datetime.utcnow().isoformat(timespec="seconds"))


def make_excel_bytes():
    output = io.BytesIO()
    tx = read_transactions()
    bud = read_budgets()
    with pd.ExcelWriter(output, engine="xlsxwriter") as writer:
        tx_disp = tx.copy()
        if not tx_disp.empty:
            tx_disp["Tanggal"] = tx_disp["date"].apply(format_date_id)
            tx_disp["Masuk"] = tx_disp.apply(lambda r: r["amount"] if r["type"] == "Masuk" else 0, axis=1)
            tx_disp["Keluar"] = tx_disp.apply(lambda r: r["amount"] if r["type"] == "Keluar" else 0, axis=1)
            tx_disp = tx_disp[["Tanggal", "fund", "type", "category", "description", "method", "Masuk", "Keluar", "note"]]
            tx_disp.columns = ["Tanggal", "Sumber Dana", "Jenis", "Kategori", "Keterangan", "Metode", "Masuk", "Keluar", "Catatan"]
        tx_disp.to_excel(writer, index=False, sheet_name="Transaksi")
        bud.to_excel(writer, index=False, sheet_name="Budget")
        movement = monthly_category_data(tx, ["Semua"], ["Semua"])
        if not movement.empty:
            movement.to_excel(writer, index=False, sheet_name="Belanja Bulanan")
    output.seek(0)
    return output.getvalue()


def is_github_enabled():
    provider = str(get_secret("PERSISTENCE_PROVIDER", "") or "").lower()
    return provider == "github" and bool(get_secret("GITHUB_TOKEN")) and bool(get_secret("GITHUB_REPO"))


def github_config():
    return {
        "token": get_secret("GITHUB_TOKEN", ""),
        "repo": get_secret("GITHUB_REPO", ""),
        "branch": get_secret("GITHUB_BRANCH", "main"),
        "path": get_secret("GITHUB_DATA_FILE", "padebuolo_live_backup.json"),
    }


def github_headers(token):
    return {
        "Authorization": f"Bearer {token}",
        "Accept": "application/vnd.github+json",
        "X-GitHub-Api-Version": "2022-11-28",
    }


def github_get_file():
    cfg = github_config()
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}"
    res = requests.get(url, headers=github_headers(cfg["token"]), params={"ref": cfg["branch"]}, timeout=20)
    if res.status_code == 404:
        return None, None
    res.raise_for_status()
    data = res.json()
    content = base64.b64decode(data.get("content", "")).decode("utf-8")
    return json.loads(content), data.get("sha")


def github_put_file(payload):
    cfg = github_config()
    url = f"https://api.github.com/repos/{cfg['repo']}/contents/{cfg['path']}"
    _, sha = github_get_file()
    body = {
        "message": f"backup data Padebuolo {datetime.utcnow().isoformat(timespec='seconds')}",
        "content": base64.b64encode(json.dumps(payload, ensure_ascii=False, indent=2).encode("utf-8")).decode("utf-8"),
        "branch": cfg["branch"],
    }
    if sha:
        body["sha"] = sha
    res = requests.put(url, headers=github_headers(cfg["token"]), json=body, timeout=30)
    res.raise_for_status()
    set_meta("last_cloud_backup_at", datetime.utcnow().isoformat(timespec="seconds"))
    return res.json()


def cloud_backup_now():
    if not is_github_enabled():
        return False, "GitHub backup belum dikonfigurasi di Secrets."
    try:
        github_put_file(export_payload())
        return True, "Backup cloud berhasil disimpan."
    except Exception as e:
        return False, f"Backup cloud gagal: {type(e).__name__}: {e}"


def cloud_restore_now():
    if not is_github_enabled():
        return False, "GitHub backup belum dikonfigurasi di Secrets."
    try:
        payload, _ = github_get_file()
        if not payload:
            return False, "File backup belum ada di repo data."
        import_payload(payload, replace=True)
        set_meta("last_cloud_restore_at", datetime.utcnow().isoformat(timespec="seconds"))
        return True, "Restore dari cloud berhasil."
    except Exception as e:
        return False, f"Restore cloud gagal: {type(e).__name__}: {e}"


def cloud_auto_backup():
    if is_github_enabled():
        try:
            github_put_file(export_payload())
        except Exception as e:
            set_meta("last_cloud_backup_error", f"{type(e).__name__}: {e}")


def first_boot_restore_or_seed():
    init_db()
    if get_meta("bootstrapped") == "1":
        return
    if is_github_enabled():
        ok, msg = cloud_restore_now()
        if ok and table_count("transactions") > 0:
            set_meta("bootstrapped", "1")
            return
    seed_from_files(force=False)
    set_meta("bootstrapped", "1")


# ----------------------
# Data views
# ----------------------
def get_funds(df=None):
    funds = []
    if df is None:
        df = read_transactions()
    if not df.empty:
        funds.extend([x for x in df["fund"].dropna().unique().tolist() if x])
    for f in FUNDS_DEFAULT:
        if f not in funds:
            funds.append(f)
    return sorted(funds)


def get_categories(df=None):
    cats = list(CATEGORIES)
    if df is None:
        df = read_transactions()
    if not df.empty:
        for c in df["category"].dropna().unique().tolist():
            if c and c not in cats:
                cats.append(c)
    return cats


def balances_by_fund(df):
    if df.empty:
        return pd.DataFrame(columns=["fund", "saldo"])
    work = df.copy()
    work["signed"] = work.apply(lambda r: r["amount"] if r["type"] == "Masuk" else -r["amount"], axis=1)
    return work.groupby("fund", as_index=False)["signed"].sum().rename(columns={"signed": "saldo"})


def monthly_category_data(df, selected_categories=None, selected_funds=None):
    """Return monthly expense movement with a real chronological sort key.

    Important: the display label (Apr 2026, Mei 2026, etc.) must not be used
    as the chart index because Streamlit/Altair can sort text labels
    alphabetically. We keep BulanSort as an actual month-start Timestamp so
    charts always run Apr -> Mei -> Jun -> Jul, etc.
    """
    cols = ["BulanSort", "Bulan", "Kategori", "Sumber Dana", "Total"]
    if df.empty:
        return pd.DataFrame(columns=cols)

    work = df.copy()
    work = work[work["type"].eq("Keluar")].copy()
    if selected_categories and "Semua" not in selected_categories:
        work = work[work["category"].isin(selected_categories)]
    if selected_funds and "Semua" not in selected_funds:
        work = work[work["fund"].isin(selected_funds)]
    if work.empty:
        return pd.DataFrame(columns=cols)

    work["date_dt"] = pd.to_datetime(work["date"], errors="coerce")
    work = work.dropna(subset=["date_dt"])
    if work.empty:
        return pd.DataFrame(columns=cols)

    work["BulanSort"] = work["date_dt"].dt.to_period("M").dt.to_timestamp()
    work["Bulan"] = work["BulanSort"].dt.to_period("M").astype(str).map(format_month_id)

    agg = (
        work.groupby(["BulanSort", "Bulan", "category", "fund"], as_index=False)["amount"]
        .sum()
        .sort_values(["BulanSort", "category", "fund"])
    )
    agg = agg.rename(columns={"category": "Kategori", "fund": "Sumber Dana", "amount": "Total"})
    return agg[cols]


def filter_transactions(df, start=None, end=None, funds=None, types=None, cats=None, keyword=""):
    if df.empty:
        return df
    work = df.copy()
    work["date_dt"] = pd.to_datetime(work["date"], errors="coerce")
    if start:
        work = work[work["date_dt"] >= pd.to_datetime(start)]
    if end:
        work = work[work["date_dt"] <= pd.to_datetime(end)]
    if funds and "Semua" not in funds:
        work = work[work["fund"].isin(funds)]
    if types and "Semua" not in types:
        work = work[work["type"].isin(types)]
    if cats and "Semua" not in cats:
        work = work[work["category"].isin(cats)]
    if keyword:
        k = keyword.lower()
        work = work[work["description"].str.lower().str.contains(k, na=False) | work["note"].str.lower().str.contains(k, na=False)]
    return work.drop(columns=["date_dt"], errors="ignore")


# ----------------------
# Pages
# ----------------------
def require_login():
    password = get_secret("APP_PASSWORD", DEFAULT_PASSWORD)
    if "logged_in" not in st.session_state:
        st.session_state.logged_in = False
    if st.session_state.logged_in:
        return True

    st.title("🏠 V.6 Padebuolo Fresh")
    st.caption("Aplikasi Kas Rumah Dinas")
    with st.form("login_form"):
        pwd = st.text_input("Password", type="password")
        submitted = st.form_submit_button("Masuk")
    if submitted:
        if pwd == password:
            st.session_state.logged_in = True
            st.rerun()
        else:
            st.error("Password salah.")
    st.info(f"Default password: `{DEFAULT_PASSWORD}`. Ganti via Streamlit Secrets `APP_PASSWORD`.")
    return False


def page_dashboard(df, budgets):
    st.title("🏠 Dashboard Padebuolo")
    balances = balances_by_fund(df)
    total_saldo = balances["saldo"].sum() if not balances.empty else 0
    total_masuk = df.loc[df["type"].eq("Masuk"), "amount"].sum() if not df.empty else 0
    total_keluar = df.loc[df["type"].eq("Keluar"), "amount"].sum() if not df.empty else 0

    cols = st.columns(4)
    cols[0].metric("Total Sisa Kas", rp_compact(total_saldo), help=rp(total_saldo))
    for i, fund in enumerate(["Kas Azka", "Kas Rayhan"], start=1):
        val = balances.loc[balances["fund"].eq(fund), "saldo"].sum() if not balances.empty else 0
        cols[i].metric(f"Sisa {fund.replace('Kas ', '')}", rp_compact(val), help=rp(val))
    cols[3].metric("Total Keluar", rp_compact(total_keluar), help=rp(total_keluar))

    st.divider()

    if df.empty:
        st.warning("Belum ada transaksi.")
        return

    work = df.copy()
    work["date_dt"] = pd.to_datetime(work["date"], errors="coerce")
    latest = work["date_dt"].max()
    current_month = latest.to_period("M") if pd.notna(latest) else pd.Timestamp.today().to_period("M")
    month_df = work[work["date_dt"].dt.to_period("M").eq(current_month)]
    month_keluar = month_df.loc[month_df["type"].eq("Keluar"), "amount"].sum()
    month_masuk = month_df.loc[month_df["type"].eq("Masuk"), "amount"].sum()

    c1, c2, c3 = st.columns(3)
    c1.metric(f"Keluar {format_month_id(str(current_month))}", rp_compact(month_keluar), help=rp(month_keluar))
    c2.metric(f"Masuk {format_month_id(str(current_month))}", rp_compact(month_masuk), help=rp(month_masuk))
    monthly_budget = budgets["amount"].sum() if not budgets.empty else 0
    c3.metric("Budget Bulanan", rp_compact(monthly_budget), help=rp(monthly_budget))

    st.subheader("Belanja per Kategori")
    expense = work[work["type"].eq("Keluar")].copy()
    top = expense.groupby("category", as_index=False)["amount"].sum().sort_values("amount", ascending=False).head(8)
    if not top.empty:
        st.bar_chart(top.set_index("category")["amount"])
        show_df(df_display(top.rename(columns={"category": "Kategori", "amount": "Total"}), money_cols=["Total"]))
    else:
        st.info("Belum ada pengeluaran.")

    st.subheader("Transaksi Terbaru")
    recent = df.sort_values(["date", "id"], ascending=False).head(10)
    recent["Masuk"] = recent.apply(lambda r: r["amount"] if r["type"] == "Masuk" else 0, axis=1)
    recent["Keluar"] = recent.apply(lambda r: r["amount"] if r["type"] == "Keluar" else 0, axis=1)
    recent_disp = recent[["date", "fund", "type", "category", "description", "Masuk", "Keluar", "note"]].rename(columns={
        "date": "Tanggal", "fund": "Sumber Dana", "type": "Jenis", "category": "Kategori", "description": "Keterangan", "note": "Catatan"
    })
    show_df(df_display(recent_disp, money_cols=["Masuk", "Keluar"], date_cols=["Tanggal"]))


def page_input(df):
    st.title("➕ Input Transaksi")
    funds = get_funds(df)
    categories = get_categories(df)

    tab1, tab2, tab3 = st.tabs(["Transaksi Biasa", "Inject Dana / Top Up", "Split Pengeluaran"])

    with tab1:
        with st.form("form_add_transaction"):
            c1, c2, c3 = st.columns(3)
            d = c1.date_input("Tanggal", value=date.today(), format="DD/MM/YYYY")
            fund = c2.selectbox("Sumber Dana", funds)
            trx_type = c3.selectbox("Jenis", ["Keluar", "Masuk"])

            c4, c5 = st.columns(2)
            amount = c4.number_input("Nominal", min_value=0, step=1000, format="%d")
            category = c5.selectbox("Kategori", categories, index=categories.index("Lainnya") if "Lainnya" in categories else 0)

            desc = st.text_input("Keterangan transaksi")
            method = st.selectbox("Metode", METHODS)
            note = st.text_area("Catatan tambahan", height=80)
            submitted = st.form_submit_button("Simpan Transaksi")
        if submitted:
            if amount <= 0:
                st.error("Nominal harus lebih dari 0.")
            else:
                add_transaction(d, fund, trx_type, amount, category, desc, method, note)
                st.success("Transaksi tersimpan.")
                st.rerun()

    with tab2:
        with st.form("form_inject"):
            c1, c2 = st.columns(2)
            d = c1.date_input("Tanggal Top Up", value=date.today(), format="DD/MM/YYYY", key="inject_date")
            fund = c2.selectbox("Sumber Dana", funds, key="inject_fund")
            amount = st.number_input("Nominal Inject", min_value=0, step=1000, format="%d", key="inject_amount")
            desc = st.text_input("Keterangan", value="Inject Dana")
            note = st.text_area("Catatan", height=80, key="inject_note")
            submitted = st.form_submit_button("Simpan Inject Dana")
        if submitted:
            if amount <= 0:
                st.error("Nominal harus lebih dari 0.")
            else:
                add_transaction(d, fund, "Masuk", amount, "Inject Dana / Top Up", desc, "Kas", note)
                st.success("Inject dana tersimpan.")
                st.rerun()

    with tab3:
        with st.form("form_split"):
            d = st.date_input("Tanggal", value=date.today(), format="DD/MM/YYYY", key="split_date")
            desc = st.text_input("Keterangan", value="Split Pengeluaran")
            total = st.number_input("Total Pengeluaran", min_value=0, step=1000, format="%d", key="split_total")
            category = st.selectbox("Kategori", categories, index=categories.index("Lainnya") if "Lainnya" in categories else 0, key="split_cat")
            participants = st.multiselect("Peserta Split", funds, default=[f for f in funds if f in ["Kas Azka", "Kas Rayhan"]])
            mode = st.radio("Pembagian", ["Rata", "Manual"], horizontal=True)
            manual = {}
            if mode == "Manual":
                st.caption("Isi nominal masing-masing. Total manual sebaiknya sama dengan total pengeluaran.")
                for p in participants:
                    manual[p] = st.number_input(f"Nominal {p}", min_value=0, step=1000, format="%d", key=f"manual_{p}")
            note = st.text_area("Catatan", height=80, key="split_note")
            submitted = st.form_submit_button("Simpan Split")
        if submitted:
            if total <= 0 or not participants:
                st.error("Total harus lebih dari 0 dan peserta harus dipilih.")
            else:
                if mode == "Rata":
                    share = round(total / len(participants))
                    amounts = {p: share for p in participants}
                    # adjust rounding to first participant
                    amounts[participants[0]] += total - sum(amounts.values())
                else:
                    amounts = manual
                    if sum(amounts.values()) != total:
                        st.warning(f"Total manual {rp(sum(amounts.values()))}, total pengeluaran {rp(total)}. Tetap disimpan sesuai nominal manual.")
                for p, amt in amounts.items():
                    if amt > 0:
                        add_transaction(d, p, "Keluar", amt, category, desc, "Kas", note)
                st.success("Split pengeluaran tersimpan.")
                st.rerun()


def page_buku_besar(df):
    st.title("📒 Buku Besar")
    balances = balances_by_fund(df)
    c1, c2, c3 = st.columns(3)
    for col, fund in zip([c1, c2], ["Kas Azka", "Kas Rayhan"]):
        val = balances.loc[balances["fund"].eq(fund), "saldo"].sum() if not balances.empty else 0
        col.metric(f"Sisa {fund.replace('Kas ', '')}", rp_compact(val), help=rp(val))
    c3.metric("Total Sisa Kas", rp_compact(balances["saldo"].sum() if not balances.empty else 0), help=rp(balances["saldo"].sum() if not balances.empty else 0))

    if df.empty:
        st.warning("Belum ada transaksi.")
        return

    work = df.copy()
    min_d = pd.to_datetime(work["date"]).min().date()
    max_d = pd.to_datetime(work["date"]).max().date()
    with st.expander("Filter", expanded=True):
        c1, c2, c3 = st.columns(3)
        start = c1.date_input("Tanggal awal", value=min_d, format="DD/MM/YYYY")
        end = c2.date_input("Tanggal akhir", value=max_d, format="DD/MM/YYYY")
        keyword = c3.text_input("Cari keterangan/catatan")
        c4, c5, c6 = st.columns(3)
        funds = c4.multiselect("Sumber dana", ["Semua"] + get_funds(df), default=["Semua"])
        types = c5.multiselect("Jenis", ["Semua", "Masuk", "Keluar"], default=["Semua"])
        cats = c6.multiselect("Kategori", ["Semua"] + get_categories(df), default=["Semua"])

    filt = filter_transactions(df, start, end, funds, types, cats, keyword)
    show = filt.copy()
    show["Masuk"] = show.apply(lambda r: r["amount"] if r["type"] == "Masuk" else 0, axis=1)
    show["Keluar"] = show.apply(lambda r: r["amount"] if r["type"] == "Keluar" else 0, axis=1)
    show["Netto"] = show["Masuk"] - show["Keluar"]
    display = show[["id", "date", "fund", "type", "category", "description", "Masuk", "Keluar", "Netto", "note"]].rename(columns={
        "id": "ID", "date": "Tanggal", "fund": "Sumber Dana", "type": "Jenis", "category": "Kategori", "description": "Keterangan", "note": "Catatan"
    })
    show_df(df_display(display, money_cols=["Masuk", "Keluar", "Netto"], date_cols=["Tanggal"]), height=480)

    with st.expander("Edit / Hapus Transaksi"):
        if filt.empty:
            st.info("Tidak ada transaksi sesuai filter.")
            return
        ids = filt["id"].astype(int).tolist()
        trx_id = st.selectbox("Pilih ID transaksi", ids)
        row = df[df["id"].eq(trx_id)].iloc[0]
        with st.form("form_edit_trx"):
            c1, c2, c3 = st.columns(3)
            d = c1.date_input("Tanggal", value=pd.to_datetime(row["date"]).date(), format="DD/MM/YYYY", key="edit_date")
            fund = c2.selectbox("Sumber Dana", get_funds(df), index=get_funds(df).index(row["fund"]) if row["fund"] in get_funds(df) else 0)
            trx_type = c3.selectbox("Jenis", ["Keluar", "Masuk"], index=["Keluar", "Masuk"].index(row["type"]) if row["type"] in ["Keluar", "Masuk"] else 0)
            c4, c5 = st.columns(2)
            amount = c4.number_input("Nominal", min_value=0, step=1000, value=int(row["amount"]), format="%d", key="edit_amount")
            cats = get_categories(df)
            category = c5.selectbox("Kategori", cats, index=cats.index(row["category"]) if row["category"] in cats else cats.index("Lainnya"))
            desc = st.text_input("Keterangan", value=row["description"])
            method = st.selectbox("Metode", METHODS, index=METHODS.index(row["method"]) if row["method"] in METHODS else 0)
            note = st.text_area("Catatan", value=row["note"], height=80)
            save = st.form_submit_button("Simpan Perubahan")
        if save:
            update_transaction(trx_id, d, fund, trx_type, amount, category, desc, method, note)
            st.success("Transaksi berhasil diupdate.")
            st.rerun()

        st.warning("Hapus transaksi tidak bisa dibatalkan.")
        confirm_delete = st.checkbox("Saya yakin mau hapus transaksi ini", key=f"del_confirm_{trx_id}")
        if st.button("Hapus Transaksi", disabled=not confirm_delete):
            delete_transaction(trx_id)
            st.success("Transaksi dihapus.")
            st.rerun()


def page_kategorisasi(df):
    st.title("🏷️ Kategorisasi Massal")
    if df.empty:
        st.warning("Belum ada transaksi.")
        return

    st.subheader("Auto-kategorisasi")
    only_lainnya = st.checkbox("Hanya ubah transaksi yang kategorinya masih Lainnya", value=True)
    candidate = df.copy()
    if only_lainnya:
        candidate = candidate[candidate["category"].eq("Lainnya")]
    candidate["Kategori Usulan"] = candidate["description"].apply(auto_category)
    candidate = candidate[candidate["Kategori Usulan"].ne("Lainnya")]
    st.caption(f"Transaksi yang bisa dikategorikan otomatis: {len(candidate)}")
    if not candidate.empty:
        preview = candidate[["id", "date", "fund", "description", "category", "Kategori Usulan"]].rename(columns={
            "id": "ID", "date": "Tanggal", "fund": "Sumber Dana", "description": "Keterangan", "category": "Kategori Lama"
        })
        show_df(df_display(preview, date_cols=["Tanggal"]))
        if st.button("Terapkan Auto-kategorisasi"):
            with get_conn() as conn:
                for _, r in candidate.iterrows():
                    conn.execute("UPDATE transactions SET category=?, updated_at=? WHERE id=?", (r["Kategori Usulan"], datetime.utcnow().isoformat(timespec="seconds"), int(r["id"])))
                conn.commit()
            cloud_auto_backup()
            st.success(f"{len(candidate)} transaksi berhasil dikategorikan.")
            st.rerun()
    else:
        st.info("Tidak ada kandidat auto-kategorisasi.")

    st.divider()
    st.subheader("Ubah berdasarkan keyword sendiri")
    with st.form("keyword_category"):
        keywords = st.text_input("Keyword, pisahkan dengan koma", placeholder="contoh: indihome, wifi, internet")
        new_cat = st.selectbox("Kategori tujuan", get_categories(df))
        only_lainnya2 = st.checkbox("Hanya ubah yang masih Lainnya", value=True, key="only_lainnya2")
        submitted = st.form_submit_button("Terapkan Keyword")
    if submitted:
        words = [w.strip().lower() for w in keywords.split(",") if w.strip()]
        if not words:
            st.error("Keyword belum diisi.")
        else:
            target = df.copy()
            if only_lainnya2:
                target = target[target["category"].eq("Lainnya")]
            mask = target["description"].str.lower().apply(lambda x: any(w in str(x) for w in words))
            ids = target.loc[mask, "id"].tolist()
            with get_conn() as conn:
                for tid in ids:
                    conn.execute("UPDATE transactions SET category=?, updated_at=? WHERE id=?", (new_cat, datetime.utcnow().isoformat(timespec="seconds"), int(tid)))
                conn.commit()
            cloud_auto_backup()
            st.success(f"{len(ids)} transaksi diubah ke kategori {new_cat}.")
            st.rerun()


def page_budget(df, budgets):
    st.title("📅 Budget Bulanan")
    st.caption("Budget standar Padebuolo: Rayhan Rp715.000 dan Azka Rp760.000 per bulan.")

    if budgets.empty:
        st.warning("Budget belum ada.")
    else:
        for fund in ["Kas Rayhan", "Kas Azka"]:
            sub = budgets[budgets["fund"].eq(fund)].copy()
            st.subheader(fund)
            total = sub["amount"].sum()
            st.metric("Total Budget", rp_compact(total), help=rp(total))
            disp = sub[["component", "amount", "note"]].rename(columns={"component": "Komponen", "amount": "Total", "note": "Catatan"})
            show_df(df_display(disp, money_cols=["Total"]))

        total_all = budgets["amount"].sum()
        st.info(f"Total budget bulanan seluruh kas: **{rp(total_all)}**")

    c1, c2 = st.columns(2)
    with c1:
        if st.button("Reset ke Budget Standar Padebuolo"):
            reset_standard_budget()
            cloud_auto_backup()
            st.success("Budget direset ke standar Padebuolo.")
            st.rerun()

    with st.expander("Tambah / Edit / Hapus Budget"):
        mode = st.radio("Mode", ["Tambah", "Edit", "Hapus"], horizontal=True)
        funds = get_funds(df)
        if mode == "Tambah":
            with st.form("add_budget"):
                fund = st.selectbox("Sumber Dana", funds)
                comp = st.text_input("Komponen")
                amount = st.number_input("Nominal", min_value=0, step=1000, format="%d")
                note = st.text_input("Catatan")
                submitted = st.form_submit_button("Tambah Budget")
            if submitted:
                add_budget(fund, comp, amount, note)
                st.success("Budget ditambahkan.")
                st.rerun()
        elif mode == "Edit" and not budgets.empty:
            bid = st.selectbox("Pilih ID Budget", budgets["id"].astype(int).tolist())
            row = budgets[budgets["id"].eq(bid)].iloc[0]
            with st.form("edit_budget"):
                fund = st.selectbox("Sumber Dana", funds, index=funds.index(row["fund"]) if row["fund"] in funds else 0)
                comp = st.text_input("Komponen", value=row["component"])
                amount = st.number_input("Nominal", min_value=0, step=1000, value=int(row["amount"]), format="%d")
                note = st.text_input("Catatan", value=row["note"])
                submitted = st.form_submit_button("Simpan Budget")
            if submitted:
                update_budget(bid, fund, comp, amount, note)
                st.success("Budget diupdate.")
                st.rerun()
        elif mode == "Hapus" and not budgets.empty:
            bid = st.selectbox("Pilih ID Budget", budgets["id"].astype(int).tolist(), key="hapus_budget_id")
            if st.button("Hapus Budget"):
                delete_budget(bid)
                st.success("Budget dihapus.")
                st.rerun()


def page_pergerakan(df):
    st.title("📈 Pergerakan Belanja Bulanan")
    if df.empty:
        st.warning("Belum ada transaksi.")
        return
    expense = df[df["type"].eq("Keluar")].copy()
    if expense.empty:
        st.info("Belum ada pengeluaran.")
        return

    expense["date_dt"] = pd.to_datetime(expense["date"], errors="coerce")
    expense = expense.dropna(subset=["date_dt"]).copy()
    if expense.empty:
        st.info("Tanggal transaksi pengeluaran belum valid.")
        return

    expense["BulanSort"] = expense["date_dt"].dt.to_period("M").dt.to_timestamp()
    month_lookup = (
        expense[["BulanSort"]]
        .drop_duplicates()
        .sort_values("BulanSort")
        .assign(Bulan=lambda x: x["BulanSort"].dt.to_period("M").astype(str).map(format_month_id))
    )
    month_options = month_lookup["Bulan"].tolist()
    month_map = dict(zip(month_lookup["Bulan"], month_lookup["BulanSort"]))

    min_d = expense["date_dt"].min().date()
    max_d = expense["date_dt"].max().date()
    with st.expander("Filter", expanded=True):
        st.caption("Pakai filter bulan kalau mau lihat belanja bulan tertentu, misalnya hanya Mei 2026.")
        selected_months = st.multiselect(
            "Bulan belanja",
            ["Semua"] + month_options,
            default=["Semua"],
            key="mov_months",
        )
        c1, c2 = st.columns(2)
        start = c1.date_input("Tanggal awal", value=min_d, format="DD/MM/YYYY", key="mov_start")
        end = c2.date_input("Tanggal akhir", value=max_d, format="DD/MM/YYYY", key="mov_end")
        cats = st.multiselect("Kategori", ["Semua"] + get_categories(df), default=["Semua"], key="mov_cats")
        funds = st.multiselect("Sumber Dana", ["Semua"] + get_funds(df), default=["Semua"], key="mov_funds")
        chart_type = st.radio("Jenis grafik", ["Line chart", "Bar chart"], horizontal=True)

    filtered = filter_transactions(expense.drop(columns=["date_dt", "BulanSort"], errors="ignore"), start, end, funds, ["Keluar"], cats)
    if selected_months and "Semua" not in selected_months:
        selected_month_starts = [month_map[m] for m in selected_months if m in month_map]
        filtered_dt = pd.to_datetime(filtered["date"], errors="coerce")
        filtered_month = filtered_dt.dt.to_period("M").dt.to_timestamp()
        filtered = filtered[filtered_month.isin(selected_month_starts)].copy()

    movement = monthly_category_data(filtered, cats, funds)
    if movement.empty:
        st.info("Tidak ada data sesuai filter.")
        return

    selected_month_label = "Semua bulan" if (not selected_months or "Semua" in selected_months) else ", ".join(selected_months)
    total_filtered = movement["Total"].sum()
    st.metric("Total belanja sesuai filter", rp_compact(total_filtered), help=rp(total_filtered))
    st.caption(f"Bulan terpilih: **{selected_month_label}**")

    # Build chart from a real datetime index. Do NOT use the displayed month label
    # as the index, because Streamlit/Altair can sort text labels
    # alphabetically.
    pivot = (
        movement.groupby(["BulanSort", "Kategori"], as_index=False)["Total"]
        .sum()
        .pivot(index="BulanSort", columns="Kategori", values="Total")
        .fillna(0)
        .sort_index()
    )

    if chart_type == "Line chart":
        st.line_chart(pivot)
    else:
        st.bar_chart(pivot)

    st.subheader("Tabel Pivot Bulan x Kategori")
    pivot_disp = pivot.reset_index().rename(columns={"BulanSort": "Bulan"})
    pivot_disp["Bulan"] = pivot_disp["Bulan"].dt.to_period("M").astype(str).map(format_month_id)
    show_df(df_display(pivot_disp, money_cols=[c for c in pivot_disp.columns if c != "Bulan"]))

    st.subheader("Ranking Belanja Kategori")
    rank = movement.groupby("Kategori", as_index=False)["Total"].sum().sort_values("Total", ascending=False)
    show_df(df_display(rank, money_cols=["Total"]))

    st.subheader("Detail Transaksi Pengeluaran")
    detail = filtered.sort_values("date").copy()
    detail_disp = detail[["date", "fund", "category", "description", "amount", "note"]].rename(columns={
        "date": "Tanggal",
        "fund": "Sumber Dana",
        "category": "Kategori",
        "description": "Keterangan",
        "amount": "Keluar",
        "note": "Catatan",
    })
    show_df(df_display(detail_disp, money_cols=["Keluar"], date_cols=["Tanggal"]), height=360)

    export_movement = movement.copy()
    export_movement["BulanSort"] = export_movement["BulanSort"].dt.strftime("%Y-%m")
    st.download_button(
        "Download data pergerakan CSV",
        data=export_movement.to_csv(index=False).encode("utf-8"),
        file_name="pergerakan_belanja_bulanan.csv",
        mime="text/csv",
    )


def page_import_export(df, budgets):
    st.title("⬆️⬇️ Import / Export")
    st.subheader("Export")
    c1, c2, c3 = st.columns(3)
    c1.download_button("Download JSON Backup", data=json.dumps(export_payload(), ensure_ascii=False, indent=2).encode("utf-8"), file_name="padebuolo_backup.json", mime="application/json")
    c2.download_button("Download Transaksi CSV", data=read_transactions().to_csv(index=False).encode("utf-8"), file_name="transactions.csv", mime="text/csv")
    c3.download_button("Download Excel", data=make_excel_bytes(), file_name="padebuolo_export.xlsx", mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")

    st.divider()
    st.subheader("Import")
    st.warning("Gunakan Replace kalau mau upload ulang total dari file basis. Gunakan Append kalau cuma nambah transaksi.")
    upload = st.file_uploader("Upload Excel / CSV / JSON", type=["xlsx", "csv", "json"])
    mode = st.radio("Mode import", ["Replace semua data", "Append transaksi"], horizontal=True)
    if upload is not None:
        if st.button("Proses Import"):
            try:
                name = upload.name.lower()
                replace = mode.startswith("Replace")
                if name.endswith(".json"):
                    payload = json.loads(upload.getvalue().decode("utf-8"))
                    import_payload(payload, replace=replace)
                elif name.endswith(".csv"):
                    tmp = pd.read_csv(upload)
                    if replace:
                        replace_transactions(tmp)
                    else:
                        append_transactions(tmp)
                elif name.endswith(".xlsx"):
                    tx, bud = parse_excel_to_dataframes(upload)
                    if replace:
                        replace_transactions(tx)
                        replace_budgets(bud if not bud.empty else pd.DataFrame(BUDGET_STANDARD))
                    else:
                        append_transactions(tx)
                cloud_auto_backup()
                st.success("Import berhasil.")
                st.rerun()
            except Exception as e:
                st.error(f"Import gagal: {type(e).__name__}: {e}")


def page_pengaturan():
    st.title("⚙️ Pengaturan")
    st.subheader("Penyimpanan")
    st.code(str(DB_PATH))
    st.write(f"Jumlah transaksi: **{table_count('transactions')}**")
    st.write(f"Jumlah budget: **{table_count('budgets')}**")
    st.write(f"Last cloud backup: `{get_meta('last_cloud_backup_at', '-')}`")
    st.write(f"Last cloud restore: `{get_meta('last_cloud_restore_at', '-')}`")
    err = get_meta("last_cloud_backup_error", "")
    if err:
        st.warning(f"Last backup error: {err}")

    st.subheader("Cloud Persistence")
    if is_github_enabled():
        st.success("GitHub backup aktif.")
        st.caption(f"Repo data: `{github_config()['repo']}` | File: `{github_config()['path']}`")
    else:
        st.warning("GitHub backup belum aktif. Data di Streamlit Cloud bisa hilang setelah reboot/redeploy kalau tidak dikonfigurasi.")
        with st.expander("Template Secrets"):
            st.code("""
APP_PASSWORD = "ganti_password_app"
PERSISTENCE_PROVIDER = "github"
GITHUB_TOKEN = "github_pat_xxx"
GITHUB_REPO = "username/padebuolo-data"
GITHUB_BRANCH = "main"
GITHUB_DATA_FILE = "padebuolo_live_backup.json"
""".strip(), language="toml")
            st.caption("Saran: pakai repo data terpisah dan private supaya backup tidak memicu redeploy app.")

    c1, c2 = st.columns(2)
    if c1.button("Simpan backup cloud sekarang"):
        ok, msg = cloud_backup_now()
        if ok:
            st.success(msg)
        else:
            st.error(msg)
    if c2.button("Pulihkan dari backup cloud"):
        ok, msg = cloud_restore_now()
        if ok:
            st.success(msg)
            st.rerun()
        else:
            st.error(msg)

    st.divider()
    st.subheader("Reset Database")
    st.warning("Reset akan mengganti database lokal dari file seed di folder data.")
    confirm_seed = st.checkbox("Saya paham, reset database dari file seed GitHub", key="confirm_seed")
    if st.button("Reset Database dari File Seed", disabled=not confirm_seed):
        seed_from_files(force=True)
        cloud_auto_backup()
        st.success("Database direset dari file seed.")
        st.rerun()

    st.subheader("Hapus Semua Data")
    confirm_clear = st.checkbox("Saya paham, hapus semua transaksi dan budget", key="confirm_clear_all")
    if st.button("Hapus Semua Data", disabled=not confirm_clear):
        with get_conn() as conn:
            conn.execute("DELETE FROM transactions")
            conn.execute("DELETE FROM budgets")
            conn.commit()
        cloud_auto_backup()
        st.success("Semua data dihapus.")
        st.rerun()

    st.divider()
    st.caption(APP_VERSION)


def main():
    st.set_page_config(page_title=APP_TITLE, page_icon="🏠", layout="wide")
    first_boot_restore_or_seed()
    if not require_login():
        return

    df = read_transactions()
    budgets = read_budgets()

    st.sidebar.title("🏠 Padebuolo")
    st.sidebar.caption(APP_VERSION)
    menu = st.sidebar.radio(
        "Menu",
        [
            "Dashboard",
            "Input Transaksi",
            "Buku Besar",
            "Kategorisasi Massal",
            "Budget Bulanan",
            "Pergerakan Belanja",
            "Import / Export",
            "Pengaturan",
        ],
    )
    st.sidebar.divider()
    if is_github_enabled():
        st.sidebar.success("Cloud backup aktif")
    else:
        st.sidebar.warning("Cloud backup belum aktif")
    st.sidebar.caption(f"Default password: {DEFAULT_PASSWORD}")
    if st.sidebar.button("Logout"):
        st.session_state.logged_in = False
        st.rerun()

    if menu == "Dashboard":
        page_dashboard(df, budgets)
    elif menu == "Input Transaksi":
        page_input(df)
    elif menu == "Buku Besar":
        page_buku_besar(df)
    elif menu == "Kategorisasi Massal":
        page_kategorisasi(df)
    elif menu == "Budget Bulanan":
        page_budget(df, budgets)
    elif menu == "Pergerakan Belanja":
        page_pergerakan(df)
    elif menu == "Import / Export":
        page_import_export(df, budgets)
    elif menu == "Pengaturan":
        page_pengaturan()


if __name__ == "__main__":
    main()
