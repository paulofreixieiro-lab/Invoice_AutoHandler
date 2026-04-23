
from __future__ import annotations

import csv
import hashlib
import os
import re
import getpass
import shutil
import sqlite3
import sys
from dataclasses import dataclass, field
from datetime import datetime
from decimal import Decimal, ROUND_HALF_UP, InvalidOperation
from pathlib import Path
from typing import Any, Optional
import xml.etree.ElementTree as ET

import pandas as pd
from pypdf import PdfReader
import tkinter as tk
from tkinter import ttk, messagebox, filedialog, simpledialog


# ============================================================
# BASE / CONFIG
# ============================================================
if getattr(sys, "frozen", False):
    APP_DIR = Path(sys.executable).resolve().parent
else:
    APP_DIR = Path(__file__).resolve().parent

BASE_DIR = APP_DIR
BIN_DIR = BASE_DIR / ".bin"
EXCEL_FILE = BASE_DIR / "faturas_resumo.xlsx"
DB_FILE = BIN_DIR / "faturas_history.db"
BACKUP_DIR = BIN_DIR / "backup"

EMAIL_TO_DEFAULT = "faturas@deco.proteste.pt"
EMAIL_CC_DEFAULT = "luis.quaresma@deco.proteste.pt"

APP_VERSION = "Prod 1.1.2"

EDP_SUMMARY_COLUMNS = [
    "InvoiceNumber", "CA", "Estado", "kWh", "AV (€)", "Valor(€)",
    "Periodo", "DocType", "PdfFile", "FinalFile", "ProcessedBy", "ProcessedAt"
]

VISIBLE_FIELDS_COMMON_REDUCED = ("agresso", "description", "produit", "unite", "periode", "nombre", "prixunit", "mnt", "iva", "code_iva", "project", "resno", "ana4", "ana5", "dep", "ct")
VISIBLE_FIELDS_VIAVERDE = ("agresso", "description", "produit", "unite", "periode", "nombre", "prixunit", "mnt", "project", "resno", "ana4", "ana5", "dep", "ct")
EDITABLE_FIELDS_VIAVERDE = ("description", "produit", "unite", "periode", "nombre", "prixunit", "mnt", "project", "resno", "ana4", "ana5", "dep", "ct")
APP_DISPLAY_NAME = "DECO PROteste - Auto Invoice Handler"
SPLASH_FILE_NAME = "splash_saas.png"
SPLASH_DELAY_MS = 2500

ADMIN_USERS_DEFAULT = {"pfr", "srm"}
ADMIN_USERS_FILE = BIN_DIR / "admin_users.txt"


def resource_path(*parts: str) -> Path:
    base = Path(sys._MEIPASS) if getattr(sys, "frozen", False) and hasattr(sys, "_MEIPASS") else APP_DIR
    return base.joinpath(*parts)


def get_splash_candidates() -> list[Path]:
    return [
        BIN_DIR / SPLASH_FILE_NAME,
        APP_DIR / SPLASH_FILE_NAME,
        resource_path(".bin", SPLASH_FILE_NAME),
        resource_path(SPLASH_FILE_NAME),
    ]


def get_splash_file() -> Optional[Path]:
    for candidate in get_splash_candidates():
        try:
            if candidate.exists():
                return candidate
        except Exception:
            pass
    return None


def show_splash_and_start(app: tk.Tk, delay_ms: int = SPLASH_DELAY_MS):
    splash_path = get_splash_file()
    if not splash_path:
        app.deiconify()
        app.lift()
        app.focus_force()
        app.mainloop()
        return

    app.withdraw()
    splash = tk.Toplevel(app)
    splash.overrideredirect(True)
    splash.configure(bg="white")
    try:
        splash.attributes("-topmost", True)
    except Exception:
        pass

    try:
        splash_img = tk.PhotoImage(file=str(splash_path))
    except Exception:
        splash.destroy()
        app.deiconify()
        app.lift()
        app.focus_force()
        app.mainloop()
        return

    width = splash_img.width()
    height = splash_img.height()
    screen_w = splash.winfo_screenwidth()
    screen_h = splash.winfo_screenheight()
    x = max((screen_w // 2) - (width // 2), 0)
    y = max((screen_h // 2) - (height // 2), 0)
    splash.geometry(f"{width}x{height}+{x}+{y}")

    label = tk.Label(splash, image=splash_img, bg="white", bd=0, highlightthickness=0)
    label.image = splash_img
    label.pack()

    def launch_main():
        try:
            splash.destroy()
        except Exception:
            pass
        app.deiconify()
        app.lift()
        app.focus_force()

    app.after(delay_ms, launch_main)
    app.mainloop()


EDP_MAP_FILE = BIN_DIR / "edp_ca_map.csv"
EPAL_MAP_FILE = BIN_DIR / "epal_ca_map.csv"
GALP_ADMIN_FILE = BIN_DIR / "Galp_admin_info.csv"
GALP_MAPPING_FILE = BIN_DIR / "Galp_vehicle_mapping.csv"
DELTA_ADMIN_FILE = BIN_DIR / "Delta_admin_info.csv"
DELTA_MAPPING_FILE = BIN_DIR / "Delta_product_mapping.csv"
SAMSIC_ADMIN_FILE = BIN_DIR / "Samsic_admin_info.csv"
EVIO_ADMIN_FILE = BIN_DIR / "Evio_admin_info.csv"
EVIO_MAPPING_FILE = BIN_DIR / "Evio_vehicle_mapping.csv"
VIAVERDE_ADMIN_FILE = BIN_DIR / "ViaVerde_admin_info.csv"
VIAVERDE_RELATION_FILE = BIN_DIR / "ViaVerde_relation_map.csv"
VIAVERDE_CA_FILE = BIN_DIR / "ViaVerde_ca_mapping.csv"
AYVENS_ADMIN_XLSX = BASE_DIR / "Ayvens_admin.xlsx"
AYVENS_EXAMPLE_XLSX = BASE_DIR / "ayvens_exemplo.xlsx"
AYVENS_ADMIN_FILE = BIN_DIR / "Ayvens_admin_info.csv"
AYVENS_TEMPLATE_FILE = BIN_DIR / "Ayvens_monthly_template.csv"
AYVENS_RELATION_FILE = BIN_DIR / "Ayvens_relation_map.csv"
AYVENS_CA_FILE = BIN_DIR / "Ayvens_ca_mapping.csv"
AYVENS_SPECIAL_FULL_VAT_RENT_PLATES = {"AX06SZ", "BA21FV"}

MONTHS_PT = {
    "JAN": "01", "FEV": "02", "MAR": "03", "ABR": "04", "MAI": "05", "JUN": "06",
    "JUL": "07", "AGO": "08", "SET": "09", "OUT": "10", "NOV": "11", "DEZ": "12",
}

DELTA_DEFAULT_MAPPING = [
    {"material": "adoçante", "produto_agresso": "63-PT"},
    {"material": "açucar", "produto_agresso": "64-PT"},
    {"material": "açúcar", "produto_agresso": "64-PT"},
    {"material": "colher", "produto_agresso": "65-PT"},
    {"material": "chá", "produto_agresso": "consochaud"},
    {"material": "infusão", "produto_agresso": "consochaud"},
    {"material": "café", "produto_agresso": "62-PT"},
    {"material": "descafeinado", "produto_agresso": "62-PT"},
    {"material": "copos", "produto_agresso": "65-pt"},
    {"material": "cacau", "produto_agresso": "consochaud"},
    {"material": "cappuccino", "produto_agresso": "62-pt"},
    {"material": "paletina", "produto_agresso": "65-pt"},
    {"material": "delta business solutions", "produto_agresso": "62-PT"},
    {"material": "delta solúvel s/c", "produto_agresso": "62-PT"},
    {"material": "delta soluvel s/c", "produto_agresso": "62-PT"},
    {"material": "inf", "produto_agresso": "consochaud"},
]

GALP_DEFAULT_ADMIN = [
    {"key": "annual_card_ca", "value": "211007492", "notes": "CA anual das anuidades GALP"},
    {"key": "last_fuel_ca", "value": "", "notes": "Último CA usado para combustível"},
    {"key": "email_to", "value": EMAIL_TO_DEFAULT, "notes": "Destinatário do rascunho"},
]

DELTA_DEFAULT_ADMIN = [
    {"key": "last_ca", "value": "", "notes": "Último CA usado para DELTA"},
    {"key": "email_to", "value": EMAIL_TO_DEFAULT, "notes": "Destinatário do rascunho"},
]

SAMSIC_DEFAULT_ADMIN = [
    {"key": "current_annual_ca", "value": "", "notes": "CA anual activo da SAMSIC"},
    {"key": "email_to", "value": EMAIL_TO_DEFAULT, "notes": "Destinatário do rascunho"},
]


EVIO_DEFAULT_ADMIN = [
    {"key": "last_ca", "value": "", "notes": "Último CA usado para EVIO"},
    {"key": "email_to", "value": EMAIL_TO_DEFAULT, "notes": "Destinatário do rascunho"},
]

VIAVERDE_DEFAULT_ADMIN = [
    {"key": "last_ca", "value": "", "notes": "Último CA usado para Via Verde"},
    {"key": "email_to", "value": EMAIL_TO_DEFAULT, "notes": "Destinatário do rascunho"},
    {"key": "digital_map_alexandra.silva", "value": "BT-42-CJ", "notes": "Mapeamento Via Verde serviços digitais"},
    {"key": "digital_map_ricardo.rosa", "value": "BG-06-PM", "notes": "Mapeamento Via Verde serviços digitais"},
    {"key": "digital_map_luis.quaresma", "value": "BC-35-EJ", "notes": "Mapeamento Via Verde serviços digitais"},
    {"key": "digital_map_mariajoao.abreu", "value": "BA-21-FV", "notes": "Mapeamento Via Verde serviços digitais"},
]

VIAVERDE_RELATION_HEADER = [
    "description", "produit", "prodfourn", "unite", "compte", "ana1",
    "project", "resno", "ana4", "ana5", "dep", "interco", "ct", "st", "t", "active"
]

VIAVERDE_UNKNOWN_OVERRIDE_FILE = BIN_DIR / "ViaVerde_unknown_overrides.csv"
VIAVERDE_UNKNOWN_OVERRIDE_HEADER = [
    "period", "identifier", "reference", "special", "description", "produit", "prodfourn", "unite", "compte", "ana1",
    "project", "resno", "ana4", "ana5", "dep", "interco", "ct", "st", "t", "active"
]

VIAVERDE_DEFAULT_RELATIONS = [
    {"description":"74-ZL-81","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5015","resno":"DEP1000","ana4":"616220","ana5":"74-ZL-81","dep":"5015","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"AA-75-AJ","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5015","resno":"DEP5015","ana4":"616220","ana5":"AA-75-AJ","dep":"5015","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BC-35-EJ","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5015","resno":"DEP5015","ana4":"616220","ana5":"BC-35-EJ","dep":"5015","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BB-78-JG","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5000-4000","resno":"8642","ana4":"620616","ana5":"BB-78-JG","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"77-ZX-36","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5000-4000","resno":"7554","ana4":"620616","ana5":"77-ZX-36","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BM-49-CR","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5000-7000","resno":"6855","ana4":"620616","ana5":"BM-49-CR","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BF-68-TL","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5000-7000","resno":"6587","ana4":"620616","ana5":"BF-68-TL","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BD-61-ZH","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"1642","ana4":"620616","ana5":"BD-61-ZH","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BG-76-MM","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"1763","ana4":"620616","ana5":"BG-76-MM","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BG-84-MM","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"8869","ana4":"620616","ana5":"BG-84-MM","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BG-93-MM","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"8464","ana4":"620616","ana5":"BG-93-MM","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BG-01-MN","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"5933","ana4":"620616","ana5":"BG-01-MN","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BG-02-MN","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"1724","ana4":"620616","ana5":"BG-02-MN","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BG-06-PM","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"9154","ana4":"620616","ana5":"BG-06-PM","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"AX-06-SZ","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5015","resno":"DEP0100","ana4":"616220","ana5":"AX-06-SZ","dep":"5015","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BA-21-FV","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5015","resno":"9213","ana4":"616220","ana5":"BA-21-FV","dep":"5015","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"BT-42-CJ","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5010","resno":"9214","ana4":"620616","ana5":"BT-42-CJ","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"CB-99-TU","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5000-4000","resno":"7554","ana4":"620616","ana5":"CB-99-TU","dep":"5011","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
    {"description":"CE-96-AT","produit":"PEAGE_NAT","prodfourn":"PEAGE_NAT","unite":"US","compte":"62510200","ana1":"NV1","project":"DO5015","resno":"DEP5015","ana4":"616220","ana5":"CE-96-AT","dep":"5015","interco":"9","ct":"NV","st":"","t":"D","active":"1"},
]

AYVENS_DEFAULT_ADMIN = [
    {"key": "email_to", "value": EMAIL_TO_DEFAULT, "notes": "Destinatário do rascunho"},
    {"key": "extra_produit", "value": "PORTAGENS", "notes": "Produit por defeito para extras Ayvens"},
    {"key": "extra_prodfourn", "value": "PORTAGENS", "notes": "ProdFourn por defeito para extras Ayvens"},
    {"key": "extra_unite", "value": "US", "notes": "Unidade por defeito para extras Ayvens"},
    {"key": "extra_compte", "value": "62260300", "notes": "Compte por defeito para extras Ayvens"},
    {"key": "extra_ana1", "value": "9", "notes": "Ana1 por defeito para extras Ayvens"},
    {"key": "extra_project", "value": "DO5015", "notes": "PROJECT por defeito para extras Ayvens"},
    {"key": "extra_resno", "value": "DEP5015", "notes": "RESNO por defeito para extras Ayvens"},
    {"key": "extra_ana4", "value": "616220", "notes": "Ana4 por defeito para extras Ayvens"},
    {"key": "extra_ana5", "value": "PORTAGENS", "notes": "Ana5 por defeito para extras Ayvens"},
    {"key": "extra_dep", "value": "5015", "notes": "DEP por defeito para extras Ayvens"},
    {"key": "extra_interco", "value": "9", "notes": "INTERCO por defeito para extras Ayvens"},
    {"key": "extra_ct", "value": "BR", "notes": "CT por defeito para extras Ayvens"},
    {"key": "extra_st", "value": "", "notes": "ST por defeito para extras Ayvens"},
    {"key": "extra_t", "value": "D", "notes": "T por defeito para extras Ayvens"},
]

GALP_MAPPING_HEADER = [
    "description", "fuel_type", "produit", "prodfourn", "unite", "f", "s",
    "compte", "ana1", "project", "resno", "ana4", "ana5", "dep",
    "interco", "ct", "st", "t", "active"
]


EVIO_MAPPING_HEADER = [
    "description", "produit", "prodfourn", "unite", "compte", "ana1",
    "project", "resno", "ana4", "ana5", "dep", "interco", "ct", "st", "t", "active"
]

EVIO_DEFAULT_MAPPING = [
    {"description": "BB-78-JG", "produit": "CARB_ELEC_PERS", "prodfourn": "AN951", "unite": "US", "compte": "62410001", "ana1": "9", "project": "DO5000-4000", "resno": "8642", "ana4": "620616", "ana5": "BB-78-JG", "dep": "5011", "interco": "9", "ct": "BG", "st": "", "t": "D", "active": "1"},
    {"description": "BD-61-ZH", "produit": "CARB_ELEC_PERS", "prodfourn": "AN951", "unite": "US", "compte": "62410001", "ana1": "9", "project": "DO5010", "resno": "1642", "ana4": "620616", "ana5": "BD-61-ZH", "dep": "5011", "interco": "9", "ct": "BG", "st": "", "t": "D", "active": "1"},
    {"description": "BT-42-CJ", "produit": "CARB_ELEC_PERS", "prodfourn": "AN951", "unite": "US", "compte": "62410001", "ana1": "9", "project": "DO5000-4000", "resno": "9214", "ana4": "620616", "ana5": "BT-42-CJ", "dep": "5011", "interco": "9", "ct": "BG", "st": "", "t": "D", "active": "1"},
]


# ============================================================
# MODELS
# ============================================================
@dataclass
class InvoiceRecord:
    supplier: str
    source_path: Path
    file_name: str
    selected: bool = False
    registrada_agresso: bool = False
    status: str = "Pendente"
    errors: list[str] = field(default_factory=list)

    invoice_number: str = ""
    invoice_date: str = ""
    billing_period: str = ""
    period: str = ""
    file_hash: str = ""
    invoice_key: str = ""
    final_name: str = ""
    final_path: str = ""
    processed_at: str = ""
    doc_type: str = ""

    # EDP / EPAL
    ca: str = ""
    piso: str = ""
    cpe: str = ""
    cpe_suffix: str = ""
    av: float = 0.0
    total_before_iva_23: float = 0.0
    kwh: float = 0.0
    cl: str = ""
    invoice_digits: str = ""
    total: float = 0.0
    a_faturar: float = 0.0
    a_deduzir: float = 0.0
    m3: float = 0.0
    abastecimento: float = 0.0
    saneamento: float = 0.0
    residuos: float = 0.0
    adicional: float = 0.0
    taxas: float = 0.0

    # GALP / DELTA
    aux_path: Optional[Path] = None
    rows: list[dict[str, Any]] = field(default_factory=list)


# ============================================================
# UTIL
# ============================================================
def ensure_dir(path: Path):
    path.mkdir(parents=True, exist_ok=True)


def backup_file(path: Path):
    if path.exists():
        ensure_dir(BACKUP_DIR)
        stamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        shutil.copy2(path, BACKUP_DIR / f"{stamp}_{path.name}")


def normalize_text(s: str) -> str:
    repl = {
        "á": "a", "à": "a", "ã": "a", "â": "a",
        "é": "e", "ê": "e",
        "í": "i",
        "ó": "o", "ô": "o", "õ": "o",
        "ú": "u",
        "ç": "c",
    }
    out = (s or "").strip().lower()
    for a, b in repl.items():
        out = out.replace(a, b)
    out = re.sub(r"\s+", " ", out)
    return out


def normalize_plate(s: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", (s or "").upper())


def safe_float(value: Any) -> float:
    if value is None:
        return 0.0
    s = str(value).strip().replace(" ", "")
    if not s:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    else:
        s = s.replace(",", ".")
    try:
        return float(s)
    except Exception:
        return 0.0


def pt_to_float(value: str) -> float:
    return safe_float(value)


def round_money(value: Any) -> float:
    num = safe_float(value)
    try:
        dec = Decimal(str(num))
        return float(dec.quantize(Decimal("0.01"), rounding=ROUND_HALF_UP))
    except (InvalidOperation, ValueError, TypeError):
        return 0.0


def money_str(value: Any) -> str:
    return f"{round_money(value):.2f}"


def format_amount_pt(value: Any) -> str:
    num = round_money(value)
    s = f"{num:,.2f}"
    return s.replace(",", "_").replace(".", ",").replace("_", ".")


def normalize_money_value(value: Any) -> str:
    return money_str(value)


def normalize_money_fields_in_row(row: dict[str, Any], fields: tuple[str, ...] = ("prixunit", "mnt")) -> dict[str, Any]:
    for field in fields:
        if field in row and str(row.get(field, "")).strip() != "":
            row[field] = normalize_money_value(row[field])
    return row


def normalize_currency_df(sheet_name: str, df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return df
    df = df.copy()
    money_columns_by_sheet = {
        "EDP": ["AV (€)", "Valor(€)"],
        "EPAL": ["Valor", "Valor (€)"],
        "GALP": ["Mnt HT en dev.", "TotalValorHT"],
        "GALP_OUTPUT": ["Mnt HT en dev."],
        "DELTA": ["TotalValorHT"],
        "DELTA_OUTPUT": ["Mnt HT en dev."],
        "SAMSIC": ["TotalValorHT"],
        "SAMSIC_OUTPUT": ["PrixUnit", "Mnt HT en dev."],
        "EVIO": ["TotalValorHT"],
        "EVIO_OUTPUT": ["PrixUnit", "Mnt HT en dev."],
        "AYVENS": ["TotalValorHT"],
        "AYVENS_OUTPUT": ["PrixUnit", "Mnt HT en dev."],
        "VIAVERDE": ["TotalValorHT"],
        "VIAVERDE_OUTPUT": ["PrixUnit", "Mnt HT en dev."],
        "Histórico": [],
        "Erros": [],
    }
    for col in money_columns_by_sheet.get(sheet_name, []):
        if col in df.columns:
            df[col] = df[col].apply(normalize_money_value)
    return df


def load_admin_users() -> set[str]:
    users = {u.lower() for u in ADMIN_USERS_DEFAULT}
    try:
        if ADMIN_USERS_FILE.exists():
            for line in ADMIN_USERS_FILE.read_text(encoding="utf-8").splitlines():
                line = line.strip().lower()
                if line and not line.startswith("#"):
                    users.add(line)
    except Exception:
        pass
    return users


def get_local_username() -> str:
    candidates = [
        str(BASE_DIR),
        str(APP_DIR),
        os.environ.get("USERPROFILE", ""),
        os.environ.get("HOMEPATH", ""),
    ]
    for candidate in candidates:
        if candidate:
            m = re.search(r"[\\/]Users[\\/]([^\\/]+)", candidate, re.IGNORECASE)
            if m:
                return m.group(1).strip().lower()

    for getter in (getpass.getuser, os.getlogin):
        try:
            username = (getter() or "").strip().lower()
            if username:
                return username
        except Exception:
            pass
    try:
        return Path.home().name.strip().lower()
    except Exception:
        return "desconhecido"


def extract_pdf_text(pdf_path: Path) -> str:
    try:
        reader = PdfReader(str(pdf_path))
        return "\n".join(page.extract_text() or "" for page in reader.pages)
    except Exception:
        return ""


def file_sha256(path: Path) -> str:
    h = hashlib.sha256()
    with open(path, "rb") as f:
        for chunk in iter(lambda: f.read(65536), b""):
            h.update(chunk)
    return h.hexdigest()


def read_csv_df(path: Path, columns: Optional[list[str]] = None) -> pd.DataFrame:
    if not path.exists():
        return pd.DataFrame(columns=columns or [])
    try:
        return pd.read_csv(path, dtype=str).fillna("")
    except Exception:
        return pd.DataFrame(columns=columns or [])


def write_csv_df(path: Path, df: pd.DataFrame):
    ensure_dir(path.parent)
    backup_file(path)
    df.to_csv(path, index=False, encoding="utf-8-sig")


def read_admin_info(path: Path, default_rows: list[dict[str, str]]) -> tuple[dict[str, str], pd.DataFrame]:
    if not path.exists():
        df = pd.DataFrame(default_rows)
        write_csv_df(path, df)
    df = read_csv_df(path, ["key", "value", "notes"])
    if df.empty:
        df = pd.DataFrame(default_rows)
        write_csv_df(path, df)
    info: dict[str, str] = {}
    for _, row in df.iterrows():
        info[str(row.get("key", "")).strip()] = str(row.get("value", "")).strip()
    return info, df


def write_admin_info(path: Path, df: pd.DataFrame):
    if "key" not in df.columns:
        df["key"] = ""
    if "value" not in df.columns:
        df["value"] = ""
    if "notes" not in df.columns:
        df["notes"] = ""
    write_csv_df(path, df[["key", "value", "notes"]])


def create_outlook_draft(subject: str, body: str, to_addr: str = "", cc_addr: str = "", attachments: Optional[list[Path]] = None) -> tuple[bool, str]:
    attachments = attachments or []
    try:
        import win32com.client  # type: ignore
        outlook = win32com.client.Dispatch("Outlook.Application")
        mail = outlook.CreateItem(0)
        mail.To = to_addr
        mail.CC = cc_addr
        mail.Subject = subject
        mail.Body = body
        for attachment in attachments:
            if attachment and Path(attachment).exists():
                mail.Attachments.Add(str(attachment))
        mail.Save()
        return True, "Rascunho criado no Outlook."
    except Exception as e:
        fallback = BASE_DIR / f"draft_{datetime.now().strftime('%Y%m%d_%H%M%S')}.txt"
        with open(fallback, "w", encoding="utf-8") as f:
            f.write(f"TO: {to_addr}\nCC: {cc_addr}\nSUBJECT: {subject}\n\n{body}\n\nANEXOS:\n")
            for attachment in attachments:
                f.write(f"- {attachment}\n")
        return False, f"Outlook indisponível. Foi criado o ficheiro: {fallback} ({e})"

def build_standard_email_body() -> str:
    return (
        "Boa tarde,\n\n"
        "Segue(m) em anexo a(s) fatura(s) referente(s) ao período indicado no assunto deste email, para efeitos de contabilização.\n\n"
        "Confirmamos que a(s) fatura(s) foram validadas e são consideradas fidedignas face aos documentos originais.\n\n"
        "Cumprimentos,\n"
    )



# ============================================================
# DB
# ============================================================
def init_db():
    ensure_dir(BIN_DIR)
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        CREATE TABLE IF NOT EXISTS processed_invoices (
            id INTEGER PRIMARY KEY AUTOINCREMENT,
            supplier TEXT NOT NULL,
            invoice_key TEXT NOT NULL UNIQUE,
            file_hash TEXT NOT NULL UNIQUE,
            invoice_number TEXT,
            period TEXT,
            doc_type TEXT,
            ca_used TEXT,
            source_filename TEXT,
            final_filename TEXT,
            processed_by TEXT,
            processed_at TEXT NOT NULL
        )
    """)
    cur.execute("PRAGMA table_info(processed_invoices)")
    processed_cols = [row[1] for row in cur.fetchall()]
    if "processed_by" not in processed_cols:
        cur.execute("ALTER TABLE processed_invoices ADD COLUMN processed_by TEXT")
    cur.execute("""
        CREATE TABLE IF NOT EXISTS pending_selections (
            supplier TEXT NOT NULL,
            invoice_key TEXT NOT NULL,
            is_selected INTEGER NOT NULL DEFAULT 0,
            updated_at TEXT NOT NULL,
            PRIMARY KEY (supplier, invoice_key)
        )
    """)
    conn.commit()
    conn.close()


def invoice_already_processed(invoice_key: str, file_hash: str):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        SELECT supplier, processed_at, final_filename
        FROM processed_invoices
        WHERE invoice_key = ? OR file_hash = ?
        LIMIT 1
    """, (invoice_key, file_hash))
    row = cur.fetchone()
    conn.close()
    return row


def register_processed_invoice(supplier: str, invoice_key: str, file_hash: str, invoice_number: str,
                               period: str, doc_type: str, ca_used: str, source_filename: str,
                               final_filename: str, processed_by: str = ""):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO processed_invoices (
            supplier, invoice_key, file_hash, invoice_number, period, doc_type,
            ca_used, source_filename, final_filename, processed_by, processed_at
        )
        VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?, ?, ?)
    """, (
        supplier, invoice_key, file_hash, invoice_number, period, doc_type, ca_used,
        source_filename, final_filename, processed_by, datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))
    conn.commit()
    conn.close()


def get_saved_selection(supplier: str, invoice_key: str, default: bool = False) -> bool:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        SELECT is_selected
        FROM pending_selections
        WHERE supplier = ? AND invoice_key = ?
        LIMIT 1
    """, (supplier, invoice_key))
    row = cur.fetchone()
    conn.close()
    if row is None:
        return default
    return bool(row[0])


def save_pending_selection(supplier: str, invoice_key: str, is_selected: bool):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        INSERT INTO pending_selections (supplier, invoice_key, is_selected, updated_at)
        VALUES (?, ?, ?, ?)
        ON CONFLICT(supplier, invoice_key) DO UPDATE SET
            is_selected = excluded.is_selected,
            updated_at = excluded.updated_at
    """, (
        supplier, invoice_key, 1 if is_selected else 0,
        datetime.now().strftime("%Y-%m-%d %H:%M:%S")
    ))
    conn.commit()
    conn.close()


def clear_pending_selection(supplier: str, invoice_key: str):
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    cur.execute("""
        DELETE FROM pending_selections
        WHERE supplier = ? AND invoice_key = ?
    """, (supplier, invoice_key))
    conn.commit()
    conn.close()


def get_last_processed_ca(supplier: str, doc_type: str = "") -> str:
    conn = sqlite3.connect(DB_FILE)
    cur = conn.cursor()
    if doc_type:
        cur.execute("""
            SELECT ca_used FROM processed_invoices
            WHERE supplier = ? AND doc_type = ? AND ca_used <> ''
            ORDER BY id DESC LIMIT 1
        """, (supplier, doc_type))
    else:
        cur.execute("""
            SELECT ca_used FROM processed_invoices
            WHERE supplier = ? AND ca_used <> ''
            ORDER BY id DESC LIMIT 1
        """, (supplier,))
    row = cur.fetchone()
    conn.close()
    return row[0] if row and row[0] else ""


# ============================================================
# EXCEL
# ============================================================
def ensure_excel():
    if EXCEL_FILE.exists():
        return
    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl") as writer:
        pd.DataFrame(columns=EDP_SUMMARY_COLUMNS).to_excel(writer, sheet_name="EDP", index=False)

        pd.DataFrame(columns=[
            "Periodo", "CA", "A_Faturar", "A_Deduzir", "M3", "Abastecimento",
            "Saneamento", "Residuos", "Adicional", "Taxas", "Valor", "CL",
            "Piso", "InvoiceNumber", "DocType", "Estado", "PdfFile",
            "FinalFile", "ProcessedBy", "ProcessedAt"
        ]).to_excel(writer, sheet_name="EPAL", index=False)

        pd.DataFrame(columns=[
            "Supplier", "CA", "Periodo", "InvoiceNumber", "DocType", "RowsCount",
            "TotalLitros", "TotalValorHT", "Estado", "PdfFile", "AuxFile",
            "FinalPdf", "FinalAux", "ProcessedBy", "ProcessedAt"
        ]).to_excel(writer, sheet_name="GALP", index=False)

        pd.DataFrame(columns=[
            "Description", "Tipo", "Produit", "ProdFourn", "Unité", "Période",
            "Nombre", "Compte", "Ana1", "PROJECT", "RESNO", "Ana4", "Ana5",
            "DEP", "CT", "Mnt HT en dev."
        ]).to_excel(writer, sheet_name="GALP_OUTPUT", index=False)

        pd.DataFrame(columns=[
            "Supplier", "CA", "Periodo", "InvoiceNumber", "DocType", "RowsCount",
            "TotalValorHT", "Estado", "PdfFile", "FinalPdf", "ProcessedBy", "ProcessedAt"
        ]).to_excel(writer, sheet_name="DELTA", index=False)

        pd.DataFrame(columns=[
            "Description", "Produit", "Période", "Nombre", "Mnt HT en dev.",
            "IVA %", "Code IVA"
        ]).to_excel(writer, sheet_name="DELTA_OUTPUT", index=False)

        pd.DataFrame(columns=[
            "Supplier", "CA", "Periodo", "InvoiceNumber", "DocType", "RowsCount",
            "TotalValorHT", "Estado", "PdfFile", "FinalPdf", "ProcessedBy", "ProcessedAt"
        ]).to_excel(writer, sheet_name="SAMSIC", index=False)

        pd.DataFrame(columns=[
            "Description", "Produit", "ProdFourn", "Unité", "Période", "Nombre",
            "PrixUnit", "Mnt HT en dev.", "IVA %", "Code IVA"
        ]).to_excel(writer, sheet_name="SAMSIC_OUTPUT", index=False)

        pd.DataFrame(columns=[
            "Supplier", "CA", "Periodo", "InvoiceNumber", "DocType", "RowsCount",
            "TotalValorHT", "Estado", "PdfFile", "AuxFile", "FinalPdf", "FinalAux",
            "ProcessedAt"
        ]).to_excel(writer, sheet_name="EVIO", index=False)

        pd.DataFrame(columns=[
            "Description", "Produit", "ProdFourn", "Unité", "Période", "Nombre",
            "PrixUnit", "Mnt HT en dev.", "IVA %", "Code IVA"
        ]).to_excel(writer, sheet_name="EVIO_OUTPUT", index=False)

        pd.DataFrame(columns=[
            "Supplier", "CA", "Periodo", "InvoiceNumber", "DocType", "RowsCount",
            "TotalValorHT", "Estado", "PdfFile", "FinalPdf", "ProcessedBy", "ProcessedAt"
        ]).to_excel(writer, sheet_name="AYVENS", index=False)

        pd.DataFrame(columns=[
            "Description", "Type", "Produit", "ProdFourn", "Unité", "Période", "Nombre",
            "PrixUnit", "Mnt HT en dev.", "IVA %", "Code IVA", "Compte", "Ana1",
            "PROJECT", "RESNO", "Ana4", "Ana5", "DEP", "INTERCO", "CT", "ST", "T"
        ]).to_excel(writer, sheet_name="AYVENS_OUTPUT", index=False)

        pd.DataFrame(columns=[
            "Supplier", "CA", "Periodo", "InvoiceNumber", "DocType", "RowsCount",
            "TotalValorHT", "Estado", "PdfFile", "FinalPdf", "FinalXml", "ProcessedBy", "ProcessedAt"
        ]).to_excel(writer, sheet_name="VIAVERDE", index=False)

        pd.DataFrame(columns=[
            "Description", "Produit", "ProdFourn", "Unité", "Période", "Nombre",
            "PrixUnit", "Mnt HT en dev.", "Compte", "Ana1", "PROJECT", "RESNO", "Ana4",
            "Ana5", "DEP", "INTERCO", "CT", "ST", "T"
        ]).to_excel(writer, sheet_name="VIAVERDE_OUTPUT", index=False)

        pd.DataFrame(columns=[
            "Supplier", "CA", "Periodo", "InvoiceNumber", "DocType", "Estado",
            "PdfFile", "FinalFile", "ProcessedBy", "ProcessedAt"
        ]).to_excel(writer, sheet_name="Histórico", index=False)

        pd.DataFrame(columns=["Supplier", "File", "Error"]).to_excel(writer, sheet_name="Erros", index=False)


def append_df_to_sheet(sheet_name: str, df: pd.DataFrame):
    ensure_excel()
    df = normalize_currency_df(sheet_name, df)

    if sheet_name == "EDP":
        df = df.copy()
        df = df.rename(columns={
            "Valor (€)": "Valor(€)",
            "TotalBeforeIVA23": "Valor(€)",
            "AV": "AV (€)",
            "Periodo_Agresso": "Periodo",
            "FileName": "PdfFile",
            "PathFinal": "FinalFile",
            "DataProcessamento": "ProcessedAt",
        })
        for col in EDP_SUMMARY_COLUMNS:
            if col not in df.columns:
                df[col] = ""
        df = df[EDP_SUMMARY_COLUMNS]

    if EXCEL_FILE.exists():
        try:
            existing = pd.read_excel(EXCEL_FILE, sheet_name=sheet_name, dtype=str).fillna("")
        except Exception:
            existing = pd.DataFrame()
    else:
        existing = pd.DataFrame()

    if sheet_name == "EPAL":
        df = df.copy()
        df = df.rename(columns={
            "A Faturar": "A_Faturar",
            "A Deduzir": "A_Deduzir",
            "Abastec.": "Abastecimento",
            "Saneam.": "Saneamento",
            "Resíduos": "Residuos",
            "Valor (€)": "Valor",
        })

    if sheet_name == "EPAL":
        existing = existing.copy()
        existing = existing.rename(columns={
            "A Faturar": "A_Faturar",
            "A Deduzir": "A_Deduzir",
            "Abastec.": "Abastecimento",
            "Saneam.": "Saneamento",
            "Resíduos": "Residuos",
            "Valor (€)": "Valor",
        })

    if sheet_name == "EDP":
        existing = existing.copy()
        existing = existing.rename(columns={
            "Valor (€)": "Valor(€)",
            "TotalBeforeIVA23": "Valor(€)",
            "AV": "AV (€)",
            "Periodo_Agresso": "Periodo",
            "FileName": "PdfFile",
            "PathFinal": "FinalFile",
            "DataProcessamento": "ProcessedAt",
        })
        for col in EDP_SUMMARY_COLUMNS:
            if col not in existing.columns:
                existing[col] = ""
        existing = existing[EDP_SUMMARY_COLUMNS]

    combined = pd.concat([existing, df.astype(str)], ignore_index=True)

    if sheet_name == "EDP":
        combined = combined.fillna("")
        for col in EDP_SUMMARY_COLUMNS:
            if col not in combined.columns:
                combined[col] = ""
        combined = combined[EDP_SUMMARY_COLUMNS]
    elif sheet_name == "EPAL":
        combined = combined.fillna("")
        combined = combined.rename(columns={
            "A Faturar": "A_Faturar",
            "A Deduzir": "A_Deduzir",
            "Abastec.": "Abastecimento",
            "Saneam.": "Saneamento",
            "Resíduos": "Residuos",
            "Valor (€)": "Valor",
        })
        epal_columns = [
            "Periodo", "CA", "A_Faturar", "A_Deduzir", "M3", "Abastecimento",
            "Saneamento", "Residuos", "Adicional", "Taxas", "Valor", "CL",
            "Piso", "InvoiceNumber", "DocType", "Estado", "PdfFile",
            "FinalFile", "ProcessedBy", "ProcessedAt"
        ]
        for col in epal_columns:
            if col not in combined.columns:
                combined[col] = ""
        combined = combined[epal_columns]

    with pd.ExcelWriter(EXCEL_FILE, engine="openpyxl", mode="a", if_sheet_exists="replace") as writer:
        combined.to_excel(writer, sheet_name=sheet_name, index=False)


def append_history_row(row: dict[str, Any]):
    append_df_to_sheet("Histórico", pd.DataFrame([row]))


# ============================================================
# AYVENS / LEASEPLAN BOOTSTRAP & PARSERS
# ============================================================
def _ayvens_default_template_columns() -> list[str]:
    return [
        "Pos", "Produit", "Description", "ProdFourn", "Unité", "Période", "F", "Nombre",
        "PrixUnit", "Escompte", "Mnt HT en dev.", "Devise", "S", "Compte", "Ana1", "PROJECT",
        "RESNO", "Ana4", "Ana5", "DEP", "INTERCO", "CT", "ST", "T", "Lock"
    ]


def _ayvens_relation_columns() -> list[str]:
    return [
        "matricula", "line_type", "description", "produit", "prodfourn", "unite",
        "compte", "ana1", "project", "resno", "ana4", "ana5", "dep", "interco",
        "ct", "st", "t", "active"
    ]


def _ayvens_candidate_file(*names: str) -> Optional[Path]:
    for name in names:
        for candidate in [BASE_DIR / name, BIN_DIR / name]:
            if candidate.exists():
                return candidate
    return None


def _normalize_ayvens_relation_df(df: pd.DataFrame) -> pd.DataFrame:
    cols = _ayvens_relation_columns()
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    rename_map = {
        "plate": "matricula",
        "description": "description",
        "line_type": "line_type",
        "produit": "produit",
        "prodfourn": "prodfourn",
        "unite": "unite",
        "compte": "compte",
        "project": "project",
        "resno": "resno",
        "ana4": "ana4",
        "ana5": "ana5",
        "dep": "dep",
        "interco": "interco",
        "ct": "ct",
        "st": "st",
        "t": "t",
        "active": "active",
    }
    df = df.rename(columns=rename_map).fillna("")
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    df = df[cols].copy()
    for col in df.columns:
        df[col] = df[col].astype(str).fillna("")
    df["matricula"] = df["matricula"].apply(normalize_plate)
    df["line_type"] = df["line_type"].astype(str).str.strip().str.upper()
    df["active"] = df["active"].replace("", "1")
    return df


def _build_ayvens_relations_from_template_df(df: pd.DataFrame) -> pd.DataFrame:
    rel_cols = _ayvens_relation_columns()
    if df is None or df.empty:
        return pd.DataFrame(columns=rel_cols)
    df = _ayvens_normalize_template_df(df)
    out = []
    for _, row in df.iterrows():
        rec = {str(k): str(v) for k, v in row.to_dict().items()}
        plate = normalize_plate(rec.get("Description", ""))
        if not plate:
            continue
        row_type = classify_ayvens_row_type(rec)
        out.append({
            "matricula": plate,
            "line_type": row_type,
            "description": rec.get("Description", ""),
            "produit": rec.get("Produit", ""),
            "prodfourn": rec.get("ProdFourn", ""),
            "unite": rec.get("Unité", "US"),
            "compte": rec.get("Compte", ""),
            "ana1": rec.get("Ana1", ""),
            "project": rec.get("PROJECT", ""),
            "resno": rec.get("RESNO", ""),
            "ana4": rec.get("Ana4", ""),
            "ana5": rec.get("Ana5", rec.get("Description", "")),
            "dep": rec.get("DEP", ""),
            "interco": rec.get("INTERCO", ""),
            "ct": rec.get("CT", ""),
            "st": rec.get("ST", ""),
            "t": rec.get("T", "D"),
            "active": "1",
        })
    return _normalize_ayvens_relation_df(pd.DataFrame(out, columns=rel_cols))

def _ayvens_normalize_template_df(df: pd.DataFrame) -> pd.DataFrame:
    if df is None or df.empty:
        return pd.DataFrame(columns=_ayvens_default_template_columns())
    rename_map = {
        "Mnt HT en dev,": "Mnt HT en dev.",
        "Mnt HT en dev": "Mnt HT en dev.",
        "Unite": "Unité",
        "Periode": "Période",
        "Produit ": "Produit",
    }
    df = df.rename(columns=rename_map).fillna("")
    for col in _ayvens_default_template_columns():
        if col not in df.columns:
            df[col] = ""
    df = df[_ayvens_default_template_columns()].copy()
    for col in df.columns:
        df[col] = df[col].astype(str).fillna("")
    return df


def _load_ayvens_template_sources() -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    if AYVENS_ADMIN_XLSX.exists():
        try:
            frames.append(_ayvens_normalize_template_df(pd.read_excel(AYVENS_ADMIN_XLSX, sheet_name="CA_mensal", dtype=str)))
        except Exception:
            pass

    if AYVENS_EXAMPLE_XLSX.exists():
        try:
            frames.append(_ayvens_normalize_template_df(pd.read_excel(AYVENS_EXAMPLE_XLSX, sheet_name="agresso_ca", dtype=str)))
        except Exception:
            pass

    if AYVENS_TEMPLATE_FILE.exists():
        try:
            frames.append(_ayvens_normalize_template_df(read_csv_df(AYVENS_TEMPLATE_FILE)))
        except Exception:
            pass

    if not frames:
        return pd.DataFrame(columns=_ayvens_default_template_columns())

    merged = pd.concat(frames, ignore_index=True).fillna("")
    merged["__plate"] = merged["Description"].astype(str).apply(normalize_plate)
    merged["__row_type"] = merged.apply(lambda r: classify_ayvens_row_type(r.to_dict()), axis=1)
    merged["__priority"] = range(len(merged))
    merged = merged.sort_values(["__plate", "__row_type", "__priority"]).drop_duplicates(["__plate", "__row_type"], keep="last")
    merged = merged.drop(columns=["__plate", "__row_type", "__priority"])
    return _ayvens_normalize_template_df(merged)



def _load_ayvens_relation_sources() -> pd.DataFrame:
    frames: list[pd.DataFrame] = []

    admin_xlsx = _ayvens_candidate_file("Ayvens_admin.xlsx")
    example_xlsx = _ayvens_candidate_file("ayvens_exemplo.xlsx")

    if admin_xlsx:
        try:
            frames.append(_build_ayvens_relations_from_template_df(pd.read_excel(admin_xlsx, sheet_name="CA_mensal", dtype=str)))
        except Exception:
            pass

    if example_xlsx:
        try:
            frames.append(_build_ayvens_relations_from_template_df(pd.read_excel(example_xlsx, sheet_name="agresso_ca", dtype=str)))
        except Exception:
            pass

    if AYVENS_RELATION_FILE.exists():
        try:
            frames.append(_normalize_ayvens_relation_df(read_csv_df(AYVENS_RELATION_FILE)))
        except Exception:
            pass

    if not frames:
        return pd.DataFrame(columns=_ayvens_relation_columns())

    merged = pd.concat(frames, ignore_index=True).fillna("")
    merged["__matricula"] = merged["matricula"].apply(normalize_plate)
    merged["__line_type"] = merged["line_type"].astype(str).str.strip().str.upper()
    merged["__priority"] = range(len(merged))
    merged = merged.sort_values(["__matricula", "__line_type", "__priority"]).drop_duplicates(["__matricula", "__line_type"], keep="last")
    merged = merged.drop(columns=["__matricula", "__line_type", "__priority"])
    return _normalize_ayvens_relation_df(merged)


def read_ayvens_relations() -> pd.DataFrame:
    bootstrap_ayvens_from_excel()
    df = _load_ayvens_relation_sources()
    if df.empty:
        return pd.DataFrame(columns=_ayvens_relation_columns())
    return df.fillna("")


def bootstrap_ayvens_from_excel():
    """Sincroniza a gestão AYVENS a partir dos ficheiros de apoio."""
    default_template_columns = _ayvens_default_template_columns()
    default_relation_columns = _ayvens_relation_columns()

    admin_xlsx = _ayvens_candidate_file("Ayvens_admin.xlsx")
    example_xlsx = _ayvens_candidate_file("ayvens_exemplo.xlsx")

    if not admin_xlsx and not example_xlsx:
        if not AYVENS_ADMIN_FILE.exists():
            pd.DataFrame(AYVENS_DEFAULT_ADMIN).to_csv(AYVENS_ADMIN_FILE, index=False, encoding="utf-8-sig")
        if not AYVENS_TEMPLATE_FILE.exists():
            pd.DataFrame(columns=default_template_columns).to_csv(AYVENS_TEMPLATE_FILE, index=False, encoding="utf-8-sig")
        if not AYVENS_RELATION_FILE.exists():
            pd.DataFrame(columns=default_relation_columns).to_csv(AYVENS_RELATION_FILE, index=False, encoding="utf-8-sig")
        return

    merged_template = _load_ayvens_template_sources()
    if merged_template.empty:
        if not AYVENS_TEMPLATE_FILE.exists():
            pd.DataFrame(columns=default_template_columns).to_csv(AYVENS_TEMPLATE_FILE, index=False, encoding="utf-8-sig")
    else:
        write_csv_df(AYVENS_TEMPLATE_FILE, merged_template)

    merged_relations = _load_ayvens_relation_sources()
    if merged_relations.empty:
        if not AYVENS_RELATION_FILE.exists():
            pd.DataFrame(columns=default_relation_columns).to_csv(AYVENS_RELATION_FILE, index=False, encoding="utf-8-sig")
    else:
        write_csv_df(AYVENS_RELATION_FILE, merged_relations)

    existing_info, existing_df = read_admin_info(AYVENS_ADMIN_FILE, AYVENS_DEFAULT_ADMIN) if AYVENS_ADMIN_FILE.exists() else ({}, pd.DataFrame(AYVENS_DEFAULT_ADMIN))
    rows: list[dict[str, str]] = []
    keys_seen: set[str] = set()

    def put(key: str, value: str, notes: str):
        rows.append({"key": key, "value": value, "notes": notes})
        keys_seen.add(key)

    for row in AYVENS_DEFAULT_ADMIN:
        key = str(row.get("key", "")).strip()
        if key:
            put(key, existing_info.get(key, str(row.get("value", "")).strip()), str(row.get("notes", "")).strip())

    if not existing_df.empty:
        for _, row in existing_df.iterrows():
            key = str(row.get("key", "")).strip()
            if key and key not in keys_seen:
                put(key, str(row.get("value", "")).strip(), str(row.get("notes", "")).strip())

    ca_lista = pd.DataFrame(columns=["No.commande", "Ref. Ext"])
    if admin_xlsx:
        try:
            ca_lista = pd.read_excel(admin_xlsx, sheet_name="CA_lista", dtype=str).fillna("")
        except Exception:
            ca_lista = pd.DataFrame(columns=["No.commande", "Ref. Ext"])

    month_to_num = MONTHS_PT.copy()
    for _, row in ca_lista.iterrows():
        ca = str(row.get("No.commande", "")).strip()
        ref_ext = normalize_text(str(row.get("Ref. Ext", "")))
        if not ca or not ref_ext:
            continue
        year_match = re.search(r"(20\d{2})", ref_ext)
        year = year_match.group(1) if year_match else ""
        if "extras" in ref_ext and year:
            key = f"extra_ca_{year}"
            if key in keys_seen:
                for rec in rows:
                    if rec["key"] == key:
                        rec["value"] = ca
                        break
            else:
                put(key, ca, f"CA de extras AYVENS para {year}")
            continue
        month_abbr = ""
        for abbr in ["JAN", "FEV", "MAR", "ABR", "MAI", "JUN", "JUL", "AGO", "SET", "OUT", "NOV", "DEZ"]:
            if normalize_text(abbr.lower()) in ref_ext:
                month_abbr = abbr
                break
        if year and month_abbr:
            key = f"rent_ca_{year}{month_to_num[month_abbr]}"
            if key in keys_seen:
                for rec in rows:
                    if rec["key"] == key:
                        rec["value"] = ca
                        break
            else:
                put(key, ca, f"CA de renda AYVENS para {year}{month_to_num[month_abbr]}")

    write_admin_info(AYVENS_ADMIN_FILE, pd.DataFrame(rows))

def read_ayvens_template() -> pd.DataFrame:
    bootstrap_ayvens_from_excel()
    df = _load_ayvens_template_sources()
    if df.empty:
        return pd.DataFrame(columns=_ayvens_default_template_columns())
    return df.fillna("")


def classify_ayvens_row_type(row: dict[str, Any]) -> str:
    produit = normalize_text(str(row.get("Produit", row.get("produit", ""))))
    prodfourn = normalize_text(str(row.get("ProdFourn", row.get("prodfourn", ""))))
    ct = str(row.get("CT", row.get("ct", ""))).strip().upper()
    if ct == "IS" or "flottevehicules" in produit or "flottevehicules" in prodfourn or "assurances_pers" in produit:
        return "SERVICE_EXEMPT"
    if "location" in produit or "location" in prodfourn or "loc_perso" in produit:
        return "RENT"
    return "SERVICE_VAT"



def build_ayvens_template_index() -> tuple[dict[tuple[str, str], dict[str, str]], dict[str, dict[str, str]]]:
    df = read_ayvens_relations()
    exact: dict[tuple[str, str], dict[str, str]] = {}
    generic: dict[str, dict[str, str]] = {}
    for _, row in df.iterrows():
        rec = {str(k): str(v) for k, v in row.to_dict().items()}
        plate = normalize_plate(rec.get("matricula", ""))
        row_type = str(rec.get("line_type", "")).strip().upper()
        if str(rec.get("active", "1")).strip() == "0" or not row_type:
            continue
        if row_type not in generic:
            generic[row_type] = rec.copy()
        if plate:
            exact[(plate, row_type)] = rec.copy()
    return exact, generic

def read_galp_plate_defaults() -> dict[str, dict[str, str]]:
    try:
        df = read_galp_mapping()
    except Exception:
        return {}
    out: dict[str, dict[str, str]] = {}
    if df.empty:
        return out
    for _, row in df.iterrows():
        rec = {str(k): str(v) for k, v in row.to_dict().items()}
        plate = normalize_plate(rec.get("description", ""))
        if plate:
            out[plate] = rec
    return out



def build_ayvens_fallback_template(plate: str, row_type: str, generic_index: dict[str, dict[str, str]], galp_defaults: dict[str, dict[str, str]]) -> Optional[dict[str, str]]:
    base = generic_index.get(row_type)
    rec = dict(base) if base else {}
    rec["matricula"] = normalize_plate(plate)
    rec["description"] = plate
    rec["ana5"] = plate
    rec["line_type"] = row_type
    if not rec:
        rec = {
            "matricula": normalize_plate(plate),
            "line_type": row_type,
            "description": plate,
            "unite": "US",
            "ana1": "9",
            "interco": "9",
            "st": "",
            "t": "D",
            "active": "1",
        }

    galp = galp_defaults.get(normalize_plate(plate), {})
    if galp:
        overlay_map = {
            "compte": "compte",
            "ana1": "ana1",
            "project": "project",
            "resno": "resno",
            "ana4": "ana4",
            "dep": "dep",
            "ct": "ct",
        }
        for tgt, src in overlay_map.items():
            if galp.get(src, ""):
                rec[tgt] = str(galp.get(src, ""))

    if row_type == "RENT":
        rec["produit"] = rec.get("produit") or "LOCATION"
        rec["prodfourn"] = rec.get("prodfourn") or rec["produit"]
        rec["compte"] = rec.get("compte") or "62612100"
        rec["ct"] = rec.get("ct") or "BG"
    elif row_type == "SERVICE_VAT":
        rec["produit"] = rec.get("produit") or "ENTRETIEN"
        rec["prodfourn"] = rec.get("prodfourn") or rec["produit"]
        rec["compte"] = rec.get("compte") or "62260300"
        rec["ct"] = rec.get("ct") or "NV"
    elif row_type == "SERVICE_EXEMPT":
        rec["produit"] = rec.get("produit") or "FLOTTEVEHICULES"
        rec["prodfourn"] = rec.get("prodfourn") or rec["produit"]
        rec["compte"] = rec.get("compte") or "62630110"
        rec["ct"] = "IS"
    elif row_type == "EXTRA":
        rec["produit"] = rec.get("produit") or "PORTAGENS"
        rec["prodfourn"] = rec.get("prodfourn") or rec["produit"]
        rec["compte"] = rec.get("compte") or "62260300"
        rec["ct"] = rec.get("ct") or "BR"
    return rec

def resolve_ayvens_ca(period: str, doc_type: str, admin: dict[str, str]) -> str:
    if doc_type == "extra":
        return admin.get(f"extra_ca_{period[:4]}", "") or admin.get("extra_ca", "")
    return admin.get(f"rent_ca_{period}", "") or admin.get("rent_ca_default", "")


def parse_ayvens_period(invoice_date: str) -> str:
    if invoice_date and re.match(r"^20\d{2}/\d{2}/\d{2}$", invoice_date):
        return invoice_date[:4] + invoice_date[5:7]
    return datetime.now().strftime("%Y%m")



def resolve_ayvens_ct(template: dict[str, str], row_type: str, description: str = "") -> str:
    plate_norm = normalize_plate(description or template.get("description", template.get("Description", "")))
    row_type = str(row_type or "").strip().upper()
    template_ct = str(template.get("ct", template.get("CT", ""))).strip().upper()

    if row_type == "SERVICE_EXEMPT":
        return "IS"
    if row_type == "SERVICE_VAT":
        return "BG" if plate_norm == "AA75AJ" else "NV"
    if row_type == "RENT":
        return "NV" if plate_norm in AYVENS_SPECIAL_FULL_VAT_RENT_PLATES else "BG"
    if row_type == "EXTRA":
        return template_ct or "BR"

    return template_ct or "NV"


def build_ayvens_row_from_template(template: dict[str, str], period: str, amount: float, row_type: str) -> dict[str, Any]:
    produit = template.get("produit", template.get("Produit", "")).strip()
    prodfourn = template.get("prodfourn", template.get("ProdFourn", produit)).strip() or produit
    unite = template.get("unite", template.get("Unité", "US")).strip() or "US"
    description = template.get("description", template.get("Description", "")).strip()
    ct = resolve_ayvens_ct(template, row_type, description)
    iva = "0" if ct == "IS" else "23"
    return {
        "confirmed": False,
        "agresso": "☐",
        "description": description,
        "type": row_type,
        "produit": produit,
        "prodfourn": prodfourn,
        "unite": unite,
        "periode": period,
        "nombre": 1,
        "prixunit": round_money(amount),
        "mnt": round_money(amount),
        "iva": iva,
        "code_iva": ct,
        "compte": template.get("compte", template.get("Compte", "")).strip(),
        "ana1": template.get("ana1", template.get("Ana1", "")).strip(),
        "project": template.get("project", template.get("PROJECT", "")).strip(),
        "resno": template.get("resno", template.get("RESNO", "")).strip(),
        "ana4": template.get("ana4", template.get("Ana4", "")).strip(),
        "ana5": template.get("ana5", template.get("Ana5", template.get("description", ""))).strip(),
        "dep": template.get("dep", template.get("DEP", "")).strip(),
        "interco": template.get("interco", template.get("INTERCO", "")).strip(),
        "ct": ct,
        "st": template.get("st", template.get("ST", "")).strip(),
        "t": template.get("t", template.get("T", "D")).strip() or "D",
    }

def parse_ayvens_rent_lines(text: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[tuple[str, str]] = set()
    for raw_line in (text or "").splitlines():
        line = " ".join(raw_line.split())
        if not line:
            continue
        m = re.match(r"^([A-Z0-9]{2}-[A-Z0-9]{2}-[A-Z0-9]{1,2})(\d{2}-\d{2}/\d{2}/\d{4})\s+(.+)$", line)
        if not m:
            continue
        plate = m.group(1)
        billing_range = m.group(2)
        tail = m.group(3).split()
        if len(tail) < 8:
            continue
        # O texto extraído do PDF LeasePlan vem frequentemente reordenado.
        # Os 4 últimos tokens mantêm-se estáveis e são os que precisamos:
        # iva_pct, custo locação, exploração sujeita a IVA, isento de IVA.
        try:
            iva_pct = int(tail[-4])
        except Exception:
            continue
        key = (plate, billing_range)
        if key in seen:
            continue
        seen.add(key)
        rows.append({
            "billing_range": billing_range,
            "plate": plate,
            "locacao": safe_float(tail[-3]),
            "exploracao_sujeita": safe_float(tail[-2]),
            "isento": safe_float(tail[-1]),
            "comb_prev": safe_float(tail[0]) if len(tail) >= 1 else 0.0,
            "comb_real": safe_float(tail[4]) if len(tail) >= 5 else 0.0,
            "iva_pct": iva_pct,
            "iva_renda": safe_float(tail[2]) if len(tail) >= 3 else 0.0,
            "iva_comb": safe_float(tail[3]) if len(tail) >= 4 else 0.0,
            "total_c_iva": safe_float(tail[1]) if len(tail) >= 2 else 0.0,
        })
    return rows

def parse_ayvens_extra_lines(text: str) -> list[dict[str, Any]]:
    rows: list[dict[str, Any]] = []
    seen: set[tuple[str, str, float]] = set()

    for raw_line in (text or "").splitlines():
        line = " ".join(raw_line.split())
        if not line:
            continue

        # Formato clássico:
        # "CA-18-PG 6 Serviço não contratado 4,85 4,58 0,27"
        m = re.match(
            r"^([A-Z0-9]{2}-[A-Z0-9]{2}-[A-Z0-9]{1,2})\s+(\d{1,2})?\s*(Servi[cç]o.+?)\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)$",
            line,
            re.IGNORECASE,
        )
        if m:
            plate = m.group(1)
            iva_pct = int(m.group(2) or 0)
            description = " ".join(m.group(3).split())

            if "servico" not in normalize_text(description):
                continue

            total = safe_float(m.group(4))
            net = safe_float(m.group(5))
            iva = safe_float(m.group(6))

            key = (plate, description, total)
            if key in seen:
                continue
            seen.add(key)

            rows.append(
                {
                    "plate": plate,
                    "description": description,
                    "billing_range": "",
                    "net": net,
                    "iva_pct": iva_pct,
                    "iva": iva,
                    "isento": 0.0,
                    "total": total,
                }
            )
            continue

        # Formato AYVENS FLEX / Flexiplan:
        # "01-30/04/2026 CA-18-PG 558,54 23 128,46 0,00 687,00"
        m2 = re.match(
            r"^(\d{2}-\d{2}/\d{2}/\d{4})\s+([A-Z0-9]{2}-[A-Z0-9]{2}-[A-Z0-9]{1,2})\s+([\d.,]+)\s+(\d{1,2})\s+([\d.,]+)\s+([\d.,]+)\s+([\d.,]+)$",
            line,
        )
        if m2:
            billing_range = m2.group(1)
            plate = m2.group(2)
            net = safe_float(m2.group(3))
            iva_pct = int(m2.group(4))
            iva = safe_float(m2.group(5))
            isento = safe_float(m2.group(6))
            total = safe_float(m2.group(7))

            description = f"AYVENS FLEX {plate} {billing_range}"
            key = (plate, description, total)
            if key in seen:
                continue
            seen.add(key)

            rows.append(
                {
                    "plate": plate,
                    "description": description,
                    "billing_range": billing_range,
                    "net": net,
                    "iva_pct": iva_pct,
                    "iva": iva,
                    "isento": isento,
                    "total": total,
                }
            )

    return rows


def parse_ayvens_document(
    pdf_path: Path, ayvens_admin: dict[str, str]
) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    invoice_number = re.sub(
        r"\s+", " ", find_first(r"(FT\s*\d+/\d+)", text, re.IGNORECASE)
    ).strip()
    invoice_date = find_first(r"\b(20\d{2}/\d{2}/\d{2})\b", text)
    norm_text = normalize_text(text)

    # Regra de negócio:
    # - AYVENS FLEX / Flexiplan = extra
    # - "servicos nao contratados" = extra
    is_flex = "ayvens flex" in norm_text or "flexiplan" in norm_text
    has_non_contracted_services = "servicos nao contratados" in norm_text
    doc_type = "extra" if (is_flex or has_non_contracted_services) else "rent"

    parsed_lines_rent = parse_ayvens_rent_lines(text) if doc_type == "rent" else []
    parsed_lines_extra = parse_ayvens_extra_lines(text) if doc_type == "extra" else []

    if doc_type == "rent" and parsed_lines_rent:
        period = _period_from_billing_range(
            parsed_lines_rent[0]["billing_range"], invoice_date
        )
    elif (
        doc_type == "extra"
        and parsed_lines_extra
        and parsed_lines_extra[0].get("billing_range")
    ):
        period = _period_from_billing_range(
            parsed_lines_extra[0]["billing_range"], invoice_date
        )
    else:
        period = parse_ayvens_period(invoice_date)

    file_hash = file_sha256(pdf_path)
    invoice_key = f"AYVENS|{invoice_number}|{period}|{doc_type}"

    rec = InvoiceRecord(
        supplier="AYVENS",
        source_path=pdf_path,
        file_name=pdf_path.name,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        period=period,
        doc_type=doc_type,
        file_hash=file_hash,
        invoice_key=invoice_key,
    )

    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
        return rec

    rec.ca = resolve_ayvens_ca(period, doc_type, ayvens_admin)
    template_index, generic_index = build_ayvens_template_index()
    galp_defaults = read_galp_plate_defaults()

    rows: list[dict[str, Any]] = []
    missing: list[str] = []

    if doc_type == "rent":
        if not parsed_lines_rent:
            rec.errors.append(
                "Nenhuma linha de renda AYVENS detectada. Verifica o layout extraído do PDF."
            )
            rec.status = "Erro"
            return rec

        for item in parsed_lines_rent:
            plate = item["plate"]
            plate_norm = normalize_plate(plate)

            amt_map = {
                "RENT": (
                    item["locacao"] * 1.23
                    if plate_norm in AYVENS_SPECIAL_FULL_VAT_RENT_PLATES
                    else item["locacao"]
                ),
                "SERVICE_VAT": (
                    item["exploracao_sujeita"]
                    if plate_norm == "AA75AJ"
                    else (
                        item["exploracao_sujeita"] * 1.23
                        if item["exploracao_sujeita"]
                        else 0.0
                    )
                ),
                "SERVICE_EXEMPT": item["isento"],
            }

            for row_type, amount in amt_map.items():
                if amount <= 0:
                    continue

                    # unreachable? no, keep structure below
                tpl = template_index.get(
                    (plate_norm, row_type)
                ) or build_ayvens_fallback_template(
                    plate, row_type, generic_index, galp_defaults
                )

                if tpl:
                    rows.append(
                        build_ayvens_row_from_template(tpl, period, amount, row_type)
                    )
                else:
                    missing.append(f"Linha {row_type} sem template para {plate}")

    else:
        if not parsed_lines_extra:
            rec.errors.append(
                "Nenhuma linha de extra AYVENS detectada. Verifica o layout extraído do PDF."
            )
            rec.status = "Erro"
            return rec

        for item in parsed_lines_extra:
            tpl = (
                template_index.get(("EXTRA_DEFAULT", "EXTRA"))
                or generic_index.get("EXTRA")
                or build_ayvens_fallback_template(
                    item["plate"], "EXTRA", generic_index, galp_defaults
                )
            )

            if tpl:
                amount = item["total"] if item["total"] else item["net"]
                row = build_ayvens_row_from_template(tpl, period, amount, "EXTRA")
                row["description"] = item["description"]
                row["iva"] = str(item.get("iva_pct", 0))
                row["ct"] = resolve_ayvens_ct(tpl, "EXTRA", item["description"])
                row["code_iva"] = row["ct"]
                rows.append(row)
            else:
                missing.append(f"Linha EXTRA sem template para {item['plate']}")

    rec.rows = rows

    if missing:
        rec.errors.extend(missing)
        rec.status = "Erro"

    if not rec.ca:
        rec.status = rec.status if rec.status == "Erro" else "Pendente"
    elif not rec.errors:
        rec.status = "Pendente"

    return rec


# ============================================================
# ENV / DEFAULT FILES
# ============================================================
def _bootstrap_admin_defaults(path: Path, defaults: list[dict[str, str]]):
    info, df = read_admin_info(path, defaults)
    changed = False
    existing_keys = set(df["key"].astype(str).str.strip().tolist()) if not df.empty else set()
    for row in defaults:
        key = str(row.get("key", "")).strip()
        if key and key not in existing_keys:
            df.loc[len(df)] = [row.get("key", ""), row.get("value", ""), row.get("notes", "")]
            changed = True
    if changed:
        write_admin_info(path, df)
    return read_admin_info(path, defaults)

def _bootstrap_evio_mapping_file():
    df = read_csv_df(EVIO_MAPPING_FILE)
    if df is None or df.empty:
        df = pd.DataFrame(EVIO_DEFAULT_MAPPING)
    rename_map = {}
    for col in list(df.columns):
        norm = normalize_text(str(col))
        if norm in {"matricula", "descricao", "description"}:
            rename_map[col] = "description"
        elif norm in {"produit", "produto"}:
            rename_map[col] = "produit"
        elif norm in {"prodfourn", "fornecedor", "prodfourn"}:
            rename_map[col] = "prodfourn"
        elif norm in {"unite", "unidade"}:
            rename_map[col] = "unite"
        elif norm == "compte":
            rename_map[col] = "compte"
        elif norm == "ana1":
            rename_map[col] = "ana1"
        elif norm == "project":
            rename_map[col] = "project"
        elif norm == "resno":
            rename_map[col] = "resno"
        elif norm == "ana4":
            rename_map[col] = "ana4"
        elif norm == "ana5":
            rename_map[col] = "ana5"
        elif norm == "dep":
            rename_map[col] = "dep"
        elif norm == "interco":
            rename_map[col] = "interco"
        elif norm == "ct":
            rename_map[col] = "ct"
        elif norm == "st":
            rename_map[col] = "st"
        elif norm == "t":
            rename_map[col] = "t"
        elif norm == "active":
            rename_map[col] = "active"
    if rename_map:
        df = df.rename(columns=rename_map)
    for col in EVIO_MAPPING_HEADER:
        if col not in df.columns:
            df[col] = ""
    df = df[EVIO_MAPPING_HEADER].fillna("")
    if "active" in df.columns:
        df["active"] = df["active"].replace("", "1")
    existing = set(df["description"].astype(str).apply(normalize_plate))
    for rec in EVIO_DEFAULT_MAPPING:
        plate = normalize_plate(rec["description"])
        if plate not in existing:
            df.loc[len(df)] = rec
    write_csv_df(EVIO_MAPPING_FILE, df[EVIO_MAPPING_HEADER])


def _add_viaverde_agg(agg: dict[tuple[str, bool], dict[str, Any]], rel_map: dict[str, dict[str, Any]], period: str, raw_plate: str, amount: float,
                      special: bool = False, conta_digital: str = "", admin: Optional[dict[str, str]] = None, errors: Optional[list[str]] = None,
                      identifier: str = "", reference: str = "", unknown_override_map: Optional[dict[tuple[str, str, str, str], dict[str, Any]]] = None):
    if amount <= 0:
        return
    admin = admin or {}
    errors = errors if errors is not None else []
    digital_map = _viaverde_digital_mapping(admin)
    email_key = _viaverde_email_key(conta_digital)
    mapped_plate = digital_map.get(email_key, "") if email_key else ""

    norm_plate = normalize_plate(raw_plate)
    unknown_plate = norm_plate in ("", "DESCONHECIDA")

    final_plate = ""
    if mapped_plate:
        final_plate = mapped_plate.strip()
        special = True
    elif not unknown_plate:
        final_plate = raw_plate.strip()
    else:
        final_plate = ""

    # unknown with manual persisted override
    override = None
    if unknown_plate and unknown_override_map is not None:
        okey = (str(period).strip(), str(identifier).strip(), str(reference).strip(), "1" if special else "0")
        override = unknown_override_map.get(okey)

    if override:
        display_desc = str(override.get("description", "")).strip() or ("DESCONHECIDA_est" if special else "DESCONHECIDA")
        key = (normalize_plate(display_desc), special)
        if key not in agg:
            row = _build_viaverde_row(override, display_desc, period, 0.0)
            row["vv_identifier"] = str(identifier).strip()
            row["vv_reference"] = str(reference).strip()
            row["vv_special"] = "1" if special else "0"
            row["manual_required"] = False
            agg[key] = row
        agg[key]["mnt"] = f"{safe_float(agg[key].get('mnt',0)) + amount:.2f}"
        agg[key]["prixunit"] = agg[key]["mnt"]
        return

    if not final_plate:
        display = "DESCONHECIDA_est" if special else "DESCONHECIDA"
        key = (display, special)
        if key not in agg:
            note = "Matrícula desconhecida: identifica manualmente na app antes de processar."
            agg[key] = _viaverde_unresolved_row(display, period, 0.0, note, identifier=identifier, reference=reference, special=special)
            errors.append(note)
        agg[key]["mnt"] = f"{safe_float(agg[key].get('mnt',0)) + amount:.2f}"
        agg[key]["prixunit"] = agg[key]["mnt"]
        # keep metadata
        if identifier and not agg[key].get("vv_identifier"):
            agg[key]["vv_identifier"] = str(identifier).strip()
        if reference and not agg[key].get("vv_reference"):
            agg[key]["vv_reference"] = str(reference).strip()
        return

    display_desc = f"{final_plate}_est" if special else final_plate
    relation_key = normalize_plate(final_plate)
    rel = rel_map.get(relation_key)
    key = (normalize_plate(display_desc), special)

    if rel is None:
        if key not in agg:
            note = f"Sem relation map para {final_plate}. Edita o mapa de gestão."
            agg[key] = _viaverde_unresolved_row(display_desc, period, 0.0, note, identifier=identifier, reference=reference, special=special)
            errors.append(note)
        agg[key]["mnt"] = f"{safe_float(agg[key].get('mnt',0)) + amount:.2f}"
        agg[key]["prixunit"] = agg[key]["mnt"]
        return

    if key not in agg:
        row = _build_viaverde_row(rel, display_desc, period, 0.0)
        row["vv_identifier"] = str(identifier).strip()
        row["vv_reference"] = str(reference).strip()
        row["vv_special"] = "1" if special else "0"
        agg[key] = row
    agg[key]["mnt"] = f"{safe_float(agg[key].get('mnt',0)) + amount:.2f}"
    agg[key]["prixunit"] = agg[key]["mnt"]
def _build_viaverde_row(rel: dict[str, Any], display_desc: str, period: str, amount: float) -> dict[str, Any]:
    return {
        "confirmed": False, "agresso": "☐",
        "description": display_desc,
        "produit": rel.get("produit", ""),
        "prodfourn": rel.get("prodfourn", ""),
        "unite": rel.get("unite", "US"),
        "periode": period,
        "nombre": "1.00",
        "prixunit": money_str(amount),
        "mnt": money_str(amount),
        "compte": rel.get("compte", ""),
        "ana1": rel.get("ana1", ""),
        "project": rel.get("project", ""),
        "resno": rel.get("resno", ""),
        "ana4": rel.get("ana4", ""),
        "ana5": display_desc,
        "dep": rel.get("dep", ""),
        "interco": rel.get("interco", ""),
        "ct": rel.get("ct", ""),
        "st": rel.get("st", ""),
        "t": rel.get("t", "D"),
        "manual_required": False,
    }

def _find_viaverde_associated_files(xml_path: Path) -> tuple[list[Path], list[Path]]:
    stem = xml_path.stem
    m = re.search(r"ViaVerde_EXTRACTO_(\d+)_(\d{4})_(\d{2})", stem, re.IGNORECASE)
    pdfs=[]; others=[xml_path]
    if not m:
        return pdfs, others
    contract, year, month = m.group(1), m.group(2), m.group(3)
    patterns = [
        f"ViaVerde_{contract}_{year}_{month}.pdf",
        f"ViaVerde_Detalhe_{contract}_{year}_{month}.pdf",
    ]
    for pat in patterns:
        p = BASE_DIR / pat
        if p.exists():
            pdfs.append(p)
            others.append(p)
    return pdfs, others

def _find_viaverde_associated_files_from_detail(detail_pdf: Path) -> tuple[list[Path], list[Path]]:
    stem = detail_pdf.stem
    m = re.search(r"ViaVerde_Detalhe_(\d+)_(\d{4})_(\d{2})", stem, re.IGNORECASE)
    pdfs = [detail_pdf]
    others = [detail_pdf]
    if not m:
        return pdfs, others
    contract, year, month = m.group(1), m.group(2), m.group(3)
    patterns = [
        f"ViaVerde_{contract}_{year}_{month}.pdf",
        f"ViaVerde_EXTRACTO_{contract}_{year}_{month}.xml",
        f"ViaVerde_EXTRACTO_{contract}_{year}_{month}.csv",
    ]
    for pat in patterns:
        p = BASE_DIR / pat
        if p.exists():
            others.append(p)
            if p.suffix.lower() == '.pdf':
                pdfs.append(p)
    return pdfs, others

def _iter_viaverde_page1_lines(text: str) -> list[str]:
    # keep only page 1 until VALORES DETALHADOS
    head = text.split('VALORES DETALHADOS', 1)[0]
    lines = []
    for raw in head.splitlines():
        line = re.sub(r"\s+", " ", raw).strip()
        if line:
            lines.append(line)
    return lines

def _viaverde_digital_mapping(admin: dict[str, str]) -> dict[str, str]:
    out = {}
    for k, v in (admin or {}).items():
        ks = str(k).strip()
        if ks.lower().startswith("digital_map_") and str(v).strip():
            out[_viaverde_email_key(ks[len("digital_map_"):])] = str(v).strip()
    return out

def _viaverde_email_key(email: str) -> str:
    email = (email or "").strip().lower()
    if "@" in email:
        email = email.split("@", 1)[0]
    return email

def _viaverde_is_special(service_type: str = "", operador: str = "", entrada: str = "", saida: str = "", conta_digital: str = "") -> bool:
    texto = " ".join([service_type or "", operador or "", entrada or "", saida or "", conta_digital or ""]).lower()
    if "portagens" in (service_type or "").lower():
        return False
    special_keys = [
        "estacionamento", "parque", "parques", "digital", "transacção via verde", "transacção", "transacao via verde",
        "envio", "anuidades", "mensalidades", "mobilidade", "serviço", "servico", "fixação", "fixacao"
    ]
    return any(k in texto for k in special_keys)


def _viaverde_unresolved_row(display_desc: str, period: str, amount: float, note: str = "", identifier: str = "", reference: str = "", special: bool = False) -> dict[str, Any]:
    return {
        "confirmed": False, "agresso": "☐", "description": display_desc, "produit": "", "prodfourn": "", "unite": "US",
        "periode": period, "nombre": "1.00", "prixunit": money_str(amount), "mnt": money_str(amount),
        "compte": "", "ana1": "", "project": "", "resno": "", "ana4": "", "ana5": display_desc,
        "dep": "", "interco": "", "ct": "", "st": "", "t": "D", "manual_required": True, "note": note,
        "vv_identifier": str(identifier).strip(), "vv_reference": str(reference).strip(), "vv_special": "1" if special else "0",
    }

def combine_viaverde_records(records: list[InvoiceRecord], admin: dict[str, str]) -> list[InvoiceRecord]:
    """Agrupa todos os documentos Via Verde do mesmo período num único documento mensal."""
    if not records:
        return []

    grouped: dict[str, list[InvoiceRecord]] = {}
    for rec in records:
        key = getattr(rec, "period", "") or ""
        if not key:
            key = "SEM_PERIODO"
        grouped.setdefault(key, []).append(rec)

    combined_records: list[InvoiceRecord] = []
    for period, group in sorted(grouped.items(), key=lambda kv: kv[0]):
        if len(group) == 1:
            combined_records.append(group[0])
            continue

        all_rows: list[dict[str, Any]] = []
        all_errors: list[str] = []
        pdf_files: list[Path] = []
        all_files: list[Path] = []
        source_names: list[str] = []
        source_paths: list[Path] = []
        hashes: list[str] = []

        for rec in group:
            all_rows.extend(getattr(rec, "rows", []) or [])
            all_errors.extend(getattr(rec, "errors", []) or [])
            pdf_files.extend(getattr(rec, "pdf_files", []) or [])
            all_files.extend(getattr(rec, "all_files", []) or [rec.source_path])
            source_names.append(rec.file_name)
            source_paths.append(rec.source_path)
            if getattr(rec, "file_hash", ""):
                hashes.append(rec.file_hash)

        # deduplicar ficheiros preservando ordem
        seen = set()
        dedup_pdfs = []
        for p in pdf_files:
            sp = str(p)
            if sp not in seen:
                seen.add(sp)
                dedup_pdfs.append(p)
        seen = set()
        dedup_all = []
        for p in all_files:
            sp = str(p)
            if sp not in seen:
                seen.add(sp)
                dedup_all.append(p)

        # combinar linhas por matrícula, somando o montante se aparecer em mais do que um XML do mesmo período
        merged_rows: list[dict[str, Any]] = []
        row_map: dict[tuple[str, str], dict[str, Any]] = {}
        for row in all_rows:
            key = (
                normalize_plate(str(row.get("description", ""))),
                str(row.get("produit", "")).strip().upper(),
            )
            if key in row_map:
                existing = row_map[key]
                mnt = safe_float(existing.get("mnt", 0)) + safe_float(row.get("mnt", 0))
                existing["mnt"] = money_str(mnt)
                existing["prixunit"] = money_str(mnt)
            else:
                clone = dict(row)
                row_map[key] = clone
                merged_rows.append(clone)

        period_label = period or (getattr(group[0], "period", "") or "")
        merged_rows = apply_viaverde_unknown_overrides(merged_rows, period_label)
        invoice_number = f"VIA VERDE {period_label}"
        combined_hash = hashlib.sha256("".join(sorted(hashes)).encode("utf-8")).hexdigest() if hashes else file_sha256(source_paths[0])
        invoice_key = f"VIAVERDE|{period_label}|MENSAL"

        combined = InvoiceRecord(
            supplier="VIAVERDE",
            source_path=source_paths[0],
            file_name=", ".join(source_names),
            invoice_number=invoice_number,
            period=period_label,
            doc_type="standard",
            file_hash=combined_hash,
            invoice_key=invoice_key,
        )
        combined.rows = merged_rows
        combined.errors = all_errors
        combined.pdf_files = dedup_pdfs
        combined.all_files = dedup_all
        combined.source_paths = source_paths
        combined.source_names = source_names
        combined.ca, year = resolve_viaverde_ca(period_label, admin)
        combined.year = year
        combined.status = "Pendente" if merged_rows or all_errors else "Erro"
        combined_records.append(combined)

    return combined_records



class CsvEditorWindow(tk.Toplevel):
    def __init__(self, master, title: str, file_path: Path, columns: list[str]):
        super().__init__(master)
        self.title(title)
        self.geometry("1200x700")
        self.file_path = file_path
        self.columns = columns
        self.df = read_csv_df(file_path, columns)
        if self.df.empty:
            self.df = pd.DataFrame(columns=columns)
        for col in columns:
            if col not in self.df.columns:
                self.df[col] = ""

        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=10)
        ttk.Button(btns, text="Adicionar", command=self.add_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Editar", command=self.edit_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Remover", command=self.remove_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Guardar", command=self.save).pack(side="right", padx=4)

        self.protocol("WM_DELETE_WINDOW", self.on_close)
        self.refresh()

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, (_, row) in enumerate(self.df.iterrows()):
            self.tree.insert("", "end", iid=str(idx), values=[row.get(c, "") for c in self.columns])

    def ask_values(self, initial: Optional[dict[str, str]] = None) -> Optional[dict[str, str]]:
        initial = initial or {}
        values: dict[str, str] = {}
        for col in self.columns:
            v = simpledialog.askstring(self.title(), col, initialvalue=initial.get(col, ""), parent=self)
            if v is None:
                return None
            values[col] = v.strip()
        return values

    def add_row(self):
        values = self.ask_values()
        if values is None:
            return
        self.df.loc[len(self.df)] = values
        self.refresh()
        self.save(silent=True)

    def edit_row(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return
        idx = int(sel[0])
        values = self.ask_values({c: str(self.df.iloc[idx].get(c, "")) for c in self.columns})
        if values is None:
            return
        for col, val in values.items():
            self.df.at[idx, col] = val
        self.refresh()
        self.save(silent=True)

    def remove_row(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return
        idx = int(sel[0])
        self.df = self.df.drop(self.df.index[idx]).reset_index(drop=True)
        self.refresh()
        self.save(silent=True)

    def save(self, silent: bool = False):
        write_csv_df(self.file_path, self.df[self.columns])
        if not silent:
            messagebox.showinfo("Sucesso", f"Guardado: {self.file_path.name}", parent=self)

    def on_close(self):
        try:
            self.save(silent=True)
        finally:
            self.destroy()


# ============================================================
# APP

# ============================================================
class FaturasFacilitiesV12(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_DISPLAY_NAME} v{APP_VERSION}")
        self.geometry("1860x980")

        ensure_environment()
        self.current_user = get_local_username()
        self.admin_users = load_admin_users()
        self.is_admin = self.current_user in self.admin_users
        self.user_role = "Admin" if self.is_admin else "Operação"
        self.email_to = tk.StringVar(value=EMAIL_TO_DEFAULT)
        self.email_cc = tk.StringVar(value=EMAIL_CC_DEFAULT)

        self.edp_map: dict[str, dict[str, str]] = {}
        self.epal_map: dict[str, dict[str, str]] = {}
        self.galp_admin, self.galp_admin_df = read_admin_info(GALP_ADMIN_FILE, GALP_DEFAULT_ADMIN)
        self.delta_admin, self.delta_admin_df = read_admin_info(DELTA_ADMIN_FILE, DELTA_DEFAULT_ADMIN)
        self.samsic_admin, self.samsic_admin_df = read_admin_info(SAMSIC_ADMIN_FILE, SAMSIC_DEFAULT_ADMIN)
        self.evio_admin, self.evio_admin_df = read_admin_info(EVIO_ADMIN_FILE, EVIO_DEFAULT_ADMIN)
        self.viaverde_admin, self.viaverde_admin_df = read_admin_info(VIAVERDE_ADMIN_FILE, VIAVERDE_DEFAULT_ADMIN)
        self.ayvens_admin, self.ayvens_admin_df = read_admin_info(AYVENS_ADMIN_FILE, AYVENS_DEFAULT_ADMIN)

        self.pending_edp: list[InvoiceRecord] = []
        self.pending_epal: list[InvoiceRecord] = []
        self.pending_galp: list[InvoiceRecord] = []
        self.pending_delta: list[InvoiceRecord] = []
        self.pending_samsic: list[InvoiceRecord] = []
        self.pending_evio: list[InvoiceRecord] = []
        self.pending_viaverde: list[InvoiceRecord] = []
        self.pending_ayvens: list[InvoiceRecord] = []
        self.error_records: list[InvoiceRecord] = []

        self.galp_index = 0
        self.delta_index = 0
        self.samsic_index = 0
        self.evio_index = 0
        self.viaverde_index = 0
        self.ayvens_index = 0

        self.galp_previous_ca_var = tk.StringVar()
        self.galp_ca_var = tk.StringVar()
        self.galp_period_var = tk.StringVar()
        self.delta_ca_var = tk.StringVar()
        self.delta_period_var = tk.StringVar()
        self.samsic_ca_var = tk.StringVar()
        self.samsic_period_var = tk.StringVar()
        self.evio_ca_var = tk.StringVar()
        self.evio_period_var = tk.StringVar()
        self.viaverde_previous_ca_var = tk.StringVar()
        self.viaverde_ca_var = tk.StringVar()
        self.viaverde_period_var = tk.StringVar()
        self.ayvens_ca_var = tk.StringVar()
        self.ayvens_period_var = tk.StringVar()
        self.galp_total_var = tk.StringVar(value="Resumo total: -")
        self.delta_total_var = tk.StringVar(value="Resumo total: -")
        self.samsic_total_var = tk.StringVar(value="Resumo total: -")
        self.evio_total_var = tk.StringVar(value="Resumo total: -")
        self.viaverde_total_var = tk.StringVar(value="Resumo total: -")
        self.ayvens_total_var = tk.StringVar(value="Resumo total: -")

        self._build_ui()
        self._configure_treeview_style()
        self.load_all()


    def _configure_treeview_style(self):
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=24)
        style.map("Treeview", background=[("selected", "#d9d9d9")], foreground=[("selected", "black")])

    def _pack_tree(self, parent, tree):
        # IMPORTANT: do not mix grid and pack inside notebook tabs.
        # The tree is usually created with `parent` as its master, so keep all
        # widgets in the same container and manage them only with `pack`.
        yscroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        yscroll.pack(side="right", fill="y", padx=(0, 8), pady=(8, 0))
        xscroll.pack(side="bottom", fill="x", padx=8, pady=(0, 8))
        tree.pack(fill="both", expand=True, padx=8, pady=8)
        return parent

    def _apply_tree_tags(self, tree):
        for idx, item in enumerate(tree.get_children()):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            tree.item(item, tags=(tag,))
        tree.tag_configure("evenrow", background="#f7f7f7")
        tree.tag_configure("oddrow", background="#ffffff")

    # ---------------- UI
    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        self.base_var = tk.StringVar(value=str(BASE_DIR))
        self.bin_var = tk.StringVar(value=str(BIN_DIR))
        self.excel_var = tk.StringVar(value=str(EXCEL_FILE))

        ttk.Label(top, text="Pasta base:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.base_var, width=100).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(top, text="Escolher", command=self.pick_base).grid(row=0, column=2, padx=4)

        ttk.Label(top, text="Pasta .bin:").grid(row=1, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.bin_var, width=100).grid(row=1, column=1, sticky="ew", padx=4)

        ttk.Label(top, text="Excel histórico:").grid(row=2, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.excel_var, width=100).grid(row=2, column=1, sticky="ew", padx=4)

        ttk.Label(top, text="Email To:").grid(row=0, column=3, sticky="e", padx=(25, 4))
        ttk.Entry(top, textvariable=self.email_to, width=35).grid(row=0, column=4, sticky="ew")
        ttk.Label(top, text="CC:").grid(row=1, column=3, sticky="e", padx=(25, 4))
        ttk.Entry(top, textvariable=self.email_cc, width=35).grid(row=1, column=4, sticky="ew")
        self.user_info_var = tk.StringVar(value=f"Utilizador app: {self.current_user} | Perfil: {self.user_role}")
        user_lbl = ttk.Label(top, textvariable=self.user_info_var)
        user_lbl.grid(row=0, column=5, rowspan=2, sticky="e", padx=(20, 0))

        btns = ttk.Frame(top)
        btns.grid(row=2, column=3, columnspan=2, sticky="e", pady=4)
        ttk.Button(btns, text="Actualizar / Ler PDFs", command=self.load_all).pack(side="left", padx=4)
        ttk.Button(btns, text="Marcar todos", command=self.mark_all_current).pack(side="left", padx=4)
        ttk.Button(btns, text="Desmarcar todos", command=self.unmark_all_current).pack(side="left", padx=4)
        ttk.Button(btns, text="Editar linha", command=self.edit_current_line).pack(side="left", padx=4)
        self.btn_management = ttk.Button(btns, text="Administração", command=self.open_management)
        self.btn_management.pack(side="left", padx=4)
        if not self.is_admin:
            self.btn_management.state(["disabled"])
        ttk.Button(btns, text="Processar tab actual", command=self.process_current_tab).pack(side="left", padx=4)

        top.columnconfigure(1, weight=1)
        top.columnconfigure(4, weight=1)
        top.columnconfigure(5, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_history = ttk.Frame(self.notebook)
        self.tab_maps = ttk.Frame(self.notebook)
        self.tab_errors = ttk.Frame(self.notebook)

        self._build_history_tab()
        self._build_maps_tab()
        self._build_errors_tab()

        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill="x", padx=8, pady=(0, 8))
        self.log_text = tk.Text(log_frame, height=4)
        self.log_text.pack(fill="both", expand=True)

    def _build_history_tab(self):
        cols = ("supplier", "invoice", "period", "ca", "processed_by", "final", "processed")
        body = ttk.Frame(self.tab_history)
        body.pack(fill="both", expand=True)
        self.tree_history = ttk.Treeview(body, columns=cols, show="headings")
        headers = {
            "supplier": "Supplier", "invoice": "Invoice", "period": "Período",
            "ca": "CA", "processed_by": "Utilizador", "final": "Ficheiro Final", "processed": "Processado em"
        }
        for c in cols:
            self.tree_history.heading(c, text=headers[c])
            self.tree_history.column(c, width=220 if c == "final" else 140, anchor="center")
        self._pack_tree(body, self.tree_history)

    def _build_maps_tab(self):
        frame = ttk.Frame(self.tab_maps, padding=10)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Ficheiros de gestão activos na pasta .bin").pack(anchor="center")
        self.map_list = tk.Listbox(frame, height=10)
        self.map_list.pack(fill="both", expand=True, pady=10)
        for p in [EDP_MAP_FILE, EPAL_MAP_FILE, GALP_ADMIN_FILE, GALP_MAPPING_FILE, DELTA_ADMIN_FILE, DELTA_MAPPING_FILE, SAMSIC_ADMIN_FILE, EVIO_ADMIN_FILE, EVIO_MAPPING_FILE, VIAVERDE_ADMIN_FILE, VIAVERDE_CA_FILE, VIAVERDE_RELATION_FILE, AYVENS_ADMIN_FILE, AYVENS_TEMPLATE_FILE]:
            self.map_list.insert("end", p.name)

    def _build_errors_tab(self):
        cols = ("supplier", "file", "error")
        body = ttk.Frame(self.tab_errors)
        body.pack(fill="both", expand=True)
        self.tree_errors = ttk.Treeview(body, columns=cols, show="headings")
        for c, h, w in [("supplier", "Supplier", 120), ("file", "Ficheiro", 260), ("error", "Erro", 1000)]:
            self.tree_errors.heading(c, text=h)
            self.tree_errors.column(c, width=w, anchor="center")
        self._pack_tree(body, self.tree_errors)

    def _build_edp_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="EDP pendentes").pack(side="left")
        cols = ("agresso", "invoice", "period", "suffix", "ca", "kwh", "valor", "av", "file", "piso")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "invoice": "Fatura", "period": "Período", "suffix": "Suffix",
            "ca": "CA", "kwh": "kWh", "valor": "Valor", "av": "AV (€)", "file": "Ficheiro", "piso": "Piso"
        }
        widths = {
            "agresso": 70, "invoice": 170, "period": 90, "suffix": 80, "ca": 110,
            "kwh": 110, "valor": 90, "av": 85, "file": 240, "piso": 90
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=widths.get(c, 120), anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_pending_tree_click(e, "EDP"))
        self.tree_edp = tree
        self.tab_edp = tab

    def _build_epal_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="EPAL pendentes").pack(side="left")
        cols = (
            "agresso", "period", "ca", "a_faturar", "a_deduzir", "m3", "abastecimento",
            "saneamento", "residuos", "adicional", "taxas", "valor",
            "file", "invoice", "cl", "piso"
        )
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "period": "Período", "ca": "CA", "a_faturar": "A Faturar",
            "a_deduzir": "A Deduzir", "m3": "M3", "abastecimento": "Abastec.",
            "saneamento": "Saneam.", "residuos": "Resíduos", "adicional": "Adicional",
            "taxas": "Taxas", "valor": "Valor", "file": "Ficheiro",
            "invoice": "Fatura", "cl": "CL", "piso": "Piso"
        }
        widths = {
            "agresso": 70, "period": 90, "ca": 100, "a_faturar": 95, "a_deduzir": 95, "m3": 85,
            "abastecimento": 95, "saneamento": 95, "residuos": 95, "adicional": 95,
            "taxas": 85, "valor": 85, "file": 220, "invoice": 160, "cl": 90, "piso": 90
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=widths.get(c, 120), anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_pending_tree_click(e, "EPAL"))
        self.tree_epal = tree
        self.tab_epal = tab

    def _build_galp_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_galp).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_galp).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("GALP", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("GALP", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA anterior:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.galp_previous_ca_var, width=18, state="readonly").pack(side="left")
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.galp_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.galp_period_var, width=10).pack(side="left")
        self.galp_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.galp_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.galp_total_var).pack(side="left", padx=16)

        self.galp_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.galp_status).pack(anchor="center", padx=10)

        cols = ("agresso", "description", "tipo", "produit", "prodfourn", "unite", "periode",
                "nombre", "mnt", "compte", "ana1", "project", "resno", "ana4", "ana5", "dep", "ct")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "tipo": "Tipo", "produit": "Produit",
            "prodfourn": "ProdFourn", "unite": "Unité", "periode": "Période", "nombre": "Nombre",
            "compte": "Compte", "ana1": "Ana1", "project": "PROJECT", "resno": "RESNO",
            "ana4": "Ana4", "ana5": "Ana5", "dep": "DEP", "ct": "CT", "mnt": "Mnt HT en dev."
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=120, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "GALP"))
        self.tree_galp = tree
        self.tab_galp = tab

    def _build_delta_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_delta).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_delta).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("DELTA", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("DELTA", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.delta_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.delta_period_var, width=10).pack(side="left")
        self.delta_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.delta_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.delta_total_var).pack(side="left", padx=16)

        self.delta_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.delta_status).pack(anchor="center", padx=10)

        cols = ("agresso", "description", "produit", "periode", "nombre", "mnt", "iva", "code_iva")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "produit": "Produit", "periode": "Période",
            "nombre": "Nombre", "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=150 if c == "description" else 120, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "DELTA"))
        self.tree_delta = tree
        self.tab_delta = tab


    def _build_samsic_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_samsic).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_samsic).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("SAMSIC", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("SAMSIC", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.samsic_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.samsic_period_var, width=10).pack(side="left")
        self.samsic_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.samsic_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.samsic_total_var).pack(side="left", padx=16)

        self.samsic_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.samsic_status).pack(anchor="center", padx=10)

        cols = VISIBLE_FIELDS_COMMON_REDUCED
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "produit": "Produit", "prodfourn": "ProdFourn",
            "unite": "Unité", "periode": "Période", "nombre": "Nombre", "prixunit": "PrixUnit",
            "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=130 if c == "description" else 100, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "SAMSIC"))
        self.tree_samsic = tree
        self.tab_samsic = tab


    def _build_evio_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_evio).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_evio).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("EVIO", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("EVIO", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.evio_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.evio_period_var, width=10).pack(side="left")
        self.evio_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.evio_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.evio_total_var).pack(side="left", padx=16)

        self.evio_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.evio_status).pack(anchor="center", padx=10)

        cols = VISIBLE_FIELDS_COMMON_REDUCED
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "produit": "Produit", "prodfourn": "ProdFourn",
            "unite": "Unité", "periode": "Période", "nombre": "Nombre", "prixunit": "PrixUnit",
            "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA", "compte": "Compte",
            "ana1": "Ana1", "project": "PROJECT", "resno": "RESNO", "ana4": "Ana4", "ana5": "Ana5", "dep": "DEP"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=160 if c == "description" else 110, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "EVIO"))
        self.tree_evio = tree
        self.tab_evio = tab


    def _build_viaverde_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_viaverde).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_viaverde).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("VIAVERDE", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("VIAVERDE", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA anterior:").pack(side="left", padx=(20,4))
        ttk.Entry(top, textvariable=self.viaverde_previous_ca_var, width=18, state="readonly").pack(side="left")
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(12,4))
        ttk.Entry(top, textvariable=self.viaverde_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12,4))
        ttk.Entry(top, textvariable=self.viaverde_period_var, width=10).pack(side="left")
        self.viaverde_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.viaverde_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.viaverde_total_var).pack(side="left", padx=16)

        self.viaverde_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.viaverde_status).pack(anchor="center", padx=10)

        cols = VISIBLE_FIELDS_VIAVERDE
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {"agresso":"Agresso","description":"Description","produit":"Produit","prodfourn":"ProdFourn","unite":"Unité","periode":"Période","nombre":"Nombre","prixunit":"PrixUnit","mnt":"Mnt HT en dev.","compte":"Compte","ana1":"Ana1","project":"PROJECT","resno":"RESNO","ana4":"Ana4","ana5":"Ana5","dep":"DEP","interco":"INTERCO","ct":"CT","st":"ST","t":"T"}
        for c in cols:
            tree.heading(c, text=hdr.get(c,c))
            tree.column(c, width=120, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "VIAVERDE"))
        self.tree_viaverde = tree
        self.tab_viaverde = tab

    def _build_ayvens_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_ayvens).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_ayvens).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("AYVENS", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("AYVENS", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.ayvens_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.ayvens_period_var, width=10).pack(side="left")
        self.ayvens_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.ayvens_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.ayvens_total_var).pack(side="left", padx=16)

        self.ayvens_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.ayvens_status).pack(anchor="center", padx=10)

        cols = ("agresso", "description", "type", "produit", "prodfourn", "unite", "periode", "nombre", "prixunit", "mnt", "iva", "code_iva")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "type": "Type", "produit": "Produit",
            "prodfourn": "ProdFourn", "unite": "Unité", "periode": "Période", "nombre": "Nombre",
            "prixunit": "PrixUnit", "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=150 if c in ("description","produit","prodfourn") else 110, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "AYVENS"))
        self.tree_ayvens = tree
        self.tab_ayvens = tab

    # ---------------- management
    def open_management(self):
        if not self.is_admin:
            messagebox.showwarning("Acesso condicionado", "A área de administração está reservada aos perfis admin desta instalação.", parent=self)
            return
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label="EDP Mapping", command=lambda: CsvEditorWindow(self, "EDP Mapping", EDP_MAP_FILE, ["Suffix", "Piso", "CA"]))
        menu.add_command(label="EPAL Mapping", command=lambda: CsvEditorWindow(self, "EPAL Mapping", EPAL_MAP_FILE, ["CL", "Piso", "CA"]))
        menu.add_separator()
        menu.add_command(label="GALP Admin", command=lambda: CsvEditorWindow(self, "GALP Admin", GALP_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="GALP Vehicle Mapping", command=lambda: CsvEditorWindow(self, "GALP Vehicle Mapping", GALP_MAPPING_FILE, GALP_MAPPING_HEADER))
        menu.add_separator()
        menu.add_command(label="DELTA Admin", command=lambda: CsvEditorWindow(self, "DELTA Admin", DELTA_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="DELTA Product Mapping", command=lambda: CsvEditorWindow(self, "DELTA Product Mapping", DELTA_MAPPING_FILE, ["material", "produto_agresso"]))
        menu.add_separator()
        menu.add_command(label="SAMSIC Admin", command=lambda: CsvEditorWindow(self, "SAMSIC Admin", SAMSIC_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_separator()
        menu.add_command(label="EVIO Admin", command=lambda: CsvEditorWindow(self, "EVIO Admin", EVIO_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="EVIO Vehicle Mapping", command=lambda: CsvEditorWindow(self, "EVIO Vehicle Mapping", EVIO_MAPPING_FILE, EVIO_MAPPING_HEADER))
        menu.add_separator()
        menu.add_command(label="VIA VERDE Admin", command=lambda: CsvEditorWindow(self, "VIA VERDE Admin", VIAVERDE_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="VIA VERDE CA Mapping", command=lambda: CsvEditorWindow(self, "VIA VERDE CA Mapping", VIAVERDE_CA_FILE, ["TIPO","ANO","MES","PERIODO","CA","DESCRICAO","ACTIVE"]))
        menu.add_command(label="VIA VERDE Relation Map", command=lambda: CsvEditorWindow(self, "VIA VERDE Relation Map", VIAVERDE_RELATION_FILE, VIAVERDE_RELATION_HEADER))
        menu.add_separator()
        menu.add_command(label="AYVENS Admin", command=lambda: CsvEditorWindow(self, "AYVENS Admin", AYVENS_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="AYVENS Monthly Template", command=lambda: CsvEditorWindow(self, "AYVENS Monthly Template", AYVENS_TEMPLATE_FILE, list(read_ayvens_template().columns) if not read_ayvens_template().empty else ["Pos", "Produit", "Description", "ProdFourn", "Unité", "Période", "F", "Nombre", "PrixUnit", "Escompte", "Mnt HT en dev.", "Devise", "S", "Compte", "Ana1", "PROJECT", "RESNO", "Ana4", "Ana5", "DEP", "INTERCO", "CT", "ST", "T", "Lock"]))
        menu.add_command(label="AYVENS Relations", command=lambda: CsvEditorWindow(self, "AYVENS Relations", AYVENS_RELATION_FILE, _ayvens_relation_columns()))
        try:
            menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        finally:
            menu.grab_release()

    # ---------------- pickers
    def pick_base(self):
        path = filedialog.askdirectory(initialdir=str(BASE_DIR), title="Seleccionar pasta base")
        if path:
            messagebox.showinfo("Informação", "A Prod 1.0 usa a pasta do executável/script como base. Move o EXE para a pasta pretendida.", parent=self)

    # ---------------- load
    def load_all(self):
        self.edp_map = self.load_map_file(EDP_MAP_FILE, "Suffix")
        self.epal_map = self.load_map_file(EPAL_MAP_FILE, "CL")
        self.galp_admin, self.galp_admin_df = read_admin_info(GALP_ADMIN_FILE, GALP_DEFAULT_ADMIN)
        self.delta_admin, self.delta_admin_df = read_admin_info(DELTA_ADMIN_FILE, DELTA_DEFAULT_ADMIN)
        self.samsic_admin, self.samsic_admin_df = read_admin_info(SAMSIC_ADMIN_FILE, SAMSIC_DEFAULT_ADMIN)
        self.evio_admin, self.evio_admin_df = read_admin_info(EVIO_ADMIN_FILE, EVIO_DEFAULT_ADMIN)
        self.viaverde_admin, self.viaverde_admin_df = read_admin_info(VIAVERDE_ADMIN_FILE, VIAVERDE_DEFAULT_ADMIN)
        self.ayvens_admin, self.ayvens_admin_df = read_admin_info(AYVENS_ADMIN_FILE, AYVENS_DEFAULT_ADMIN)

        self.pending_edp = []
        self.pending_epal = []
        self.pending_galp = []
        self.pending_delta = []
        self.pending_samsic = []
        self.pending_evio = []
        self.pending_viaverde = []
        self.pending_ayvens = []
        self.error_records = []

        for pdf in sorted(BASE_DIR.glob("*.pdf")):
            name = pdf.name.upper()
            try:
                if name.startswith("EDP_"):
                    rec = parse_edp_pdf(pdf, self.edp_map)
                    self.route_record(rec, self.pending_edp)
                elif name.startswith("EPAL_"):
                    rec = parse_epal_pdf(pdf, self.epal_map)
                    self.route_record(rec, self.pending_epal)
                elif name.startswith("GALP_"):
                    rec = parse_galp_document(pdf)
                    self.route_record(rec, self.pending_galp)
                elif name.startswith("DELTA_"):
                    rec = parse_delta_pdf(pdf)
                    self.route_record(rec, self.pending_delta)
                elif name.startswith("SAMSIC"):
                    rec = parse_samsic_pdf(pdf)
                    self.route_record(rec, self.pending_samsic)
                elif name.startswith("EVIO"):
                    rec = parse_evio_document(pdf)
                    self.route_record(rec, self.pending_evio)
                elif name.startswith("VIAVERDE"):
                    pass
                else:
                    is_ayvens = any(k in name for k in ["AYVENS", "LEASEPLAN", "LEASE_PLAN", "LEASE PLAN"])
                    if not is_ayvens:
                        sample = normalize_text(extract_pdf_text(pdf)[:4000])
                        is_ayvens = any(k in sample for k in ["lease plan portugal", "ayvens", "lpptft003", "lpptft010"])
                    if is_ayvens:
                        rec = parse_ayvens_document(pdf, self.ayvens_admin)
                        self.route_record(rec, self.pending_ayvens)
            except Exception as e:
                rec = InvoiceRecord(supplier="UNKNOWN", source_path=pdf, file_name=pdf.name, status="Erro", errors=[str(e)])
                self.error_records.append(rec)

        for detail_pdf in sorted(BASE_DIR.glob("ViaVerde_Detalhe_*_*.pdf")):
            try:
                rec = parse_viaverde_detail_pdf_document(detail_pdf, self.viaverde_admin)
                self.route_record(rec, self.pending_viaverde)
            except Exception as e:
                rec = InvoiceRecord(supplier="VIAVERDE", source_path=detail_pdf, file_name=detail_pdf.name, status="Erro", errors=[str(e)])
                self.error_records.append(rec)

        self.pending_viaverde = combine_viaverde_records(self.pending_viaverde, self.viaverde_admin)

        self.rebuild_notebook()
        self.populate_pending_trees()
        self.populate_history_tree()
        self.populate_error_tree()
        self.log(
            f"Pendentes EDP: {len(self.pending_edp)} | Pendentes EPAL: {len(self.pending_epal)} | "
            f"GALP: {len(self.pending_galp)} | DELTA: {len(self.pending_delta)} | SAMSIC: {len(self.pending_samsic)} | EVIO: {len(self.pending_evio)} | VIA VERDE: {len(self.pending_viaverde)} | AYVENS: {len(self.pending_ayvens)} | Erros: {len(self.error_records)}"
        )
        self.log("Leitura concluída.")

    def route_record(self, rec: InvoiceRecord, bucket: list[InvoiceRecord]):
        if rec.status == "Erro":
            self.error_records.append(rec)
        elif rec.status == "Duplicado":
            self.log(f"Ignorado duplicado: {rec.file_name}")
        else:
            if rec.supplier in ("EDP", "EPAL"):
                rec.selected = get_saved_selection(rec.supplier, rec.invoice_key, default=False)
            bucket.append(rec)

    def rebuild_notebook(self):
        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)

        if self.pending_edp:
            self._build_edp_tab()
            self.notebook.add(self.tab_edp, text="EDP")
        if self.pending_epal:
            self._build_epal_tab()
            self.notebook.add(self.tab_epal, text="EPAL")
        if self.pending_galp:
            self._build_galp_tab()
            self.notebook.add(self.tab_galp, text="GALP")
        if self.pending_delta:
            self._build_delta_tab()
            self.notebook.add(self.tab_delta, text="DELTA")
        if self.pending_samsic:
            self._build_samsic_tab()
            self.notebook.add(self.tab_samsic, text="SAMSIC")
        if self.pending_evio:
            self._build_evio_tab()
            self.notebook.add(self.tab_evio, text="EVIO")
        if self.pending_viaverde:
            self._build_viaverde_tab()
            self.notebook.add(self.tab_viaverde, text="VIA VERDE")
        if self.pending_ayvens:
            self._build_ayvens_tab()
            self.notebook.add(self.tab_ayvens, text="AYVENS")

        self.notebook.add(self.tab_history, text="Histórico")
        self.notebook.add(self.tab_maps, text="Mapeamentos")
        self.notebook.add(self.tab_errors, text="Erros")

    def load_map_file(self, path: Path, key_field: str) -> dict[str, dict[str, str]]:
        out: dict[str, dict[str, str]] = {}
        df = read_csv_df(path)
        if df.empty:
            return out
        for _, row in df.iterrows():
            key = str(row.get(key_field, "")).strip()
            if key:
                out[key] = {"Piso": str(row.get("Piso", "")).strip(), "CA": str(row.get("CA", "")).strip()}
        return out

    # ---------------- populate
    def populate_pending_trees(self):
        if hasattr(self, "tree_edp"):
            for i in self.tree_edp.get_children():
                self.tree_edp.delete(i)
            for idx, rec in enumerate(self.pending_edp):
                self.tree_edp.insert("", "end", iid=f"EDP::{idx}", values=(
                    "☑" if rec.selected else "☐",
                    rec.invoice_number, rec.period, rec.cpe_suffix, rec.ca,
                    f"{rec.kwh:.4f}", money_str(rec.total_before_iva_23), money_str(rec.av),
                    rec.file_name, rec.piso
                ))
            self._apply_tree_tags(self.tree_edp)

        if hasattr(self, "tree_epal"):
            for i in self.tree_epal.get_children():
                self.tree_epal.delete(i)
            for idx, rec in enumerate(self.pending_epal):
                self.tree_epal.insert("", "end", iid=f"EPAL::{idx}", values=(
                    "☑" if rec.selected else "☐",
                    rec.period,
                    rec.ca,
                    f"{rec.a_faturar:.3f}",
                    f"{rec.a_deduzir:.3f}",
                    f"{rec.m3:.3f}",
                    f"{rec.abastecimento:.4f}",
                    f"{rec.saneamento:.4f}",
                    f"{rec.residuos:.4f}",
                    f"{rec.adicional:.4f}",
                    f"{rec.taxas:.4f}",
                    money_str(rec.total),
                    rec.file_name,
                    rec.invoice_number,
                    rec.cl,
                    rec.piso,
                ))
            self._apply_tree_tags(self.tree_epal)

        if hasattr(self, "tree_galp"):
            self.load_galp_current(refresh_only=True)

        if hasattr(self, "tree_delta"):
            self.load_delta_current(refresh_only=True)
        if hasattr(self, "tree_samsic"):
            self.load_samsic_current(refresh_only=True)

        if hasattr(self, "tree_evio"):
            self.load_evio_current(refresh_only=True)
        if hasattr(self, "tree_viaverde"):
            self.load_viaverde_current(refresh_only=True)
        if hasattr(self, "tree_ayvens"):
            self.load_ayvens_current(refresh_only=True)

    def populate_history_tree(self):
        for i in self.tree_history.get_children():
            self.tree_history.delete(i)
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("""
            SELECT supplier, invoice_number, period, ca_used, processed_by, final_filename, processed_at
            FROM processed_invoices ORDER BY id DESC
        """)
        for idx, row in enumerate(cur.fetchall()):
            self.tree_history.insert("", "end", iid=str(idx), values=row)
        conn.close()
        self._apply_tree_tags(self.tree_history)

    def populate_error_tree(self):
        for i in self.tree_errors.get_children():
            self.tree_errors.delete(i)
        all_rows = []
        for rec in self.error_records:
            for err in rec.errors:
                all_rows.append((rec.supplier, rec.file_name, err))
        for idx, row in enumerate(all_rows):
            self.tree_errors.insert("", "end", iid=str(idx), values=row)
        self._apply_tree_tags(self.tree_errors)

    # ---------------- current tab helpers
    def current_tab_name(self) -> str:
        try:
            tab_id = self.notebook.select()
            return self.notebook.tab(tab_id, "text")
        except Exception:
            return ""

    def mark_all_current(self):
        tab = self.current_tab_name()
        if tab == "EDP":
            for rec in self.pending_edp:
                rec.selected = True
                save_pending_selection(rec.supplier, rec.invoice_key, True)
        elif tab == "EPAL":
            for rec in self.pending_epal:
                rec.selected = True
                save_pending_selection(rec.supplier, rec.invoice_key, True)
        elif tab == "GALP":
            self.mark_rows("GALP", True)
            return
        elif tab == "DELTA":
            self.mark_rows("DELTA", True)
            return
        elif tab == "SAMSIC":
            self.mark_rows("SAMSIC", True)
            return
        elif tab == "EVIO":
            self.mark_rows("EVIO", True)
            return
        elif tab == "VIA VERDE":
            self.mark_rows("VIAVERDE", True)
            return
        elif tab == "AYVENS":
            self.mark_rows("AYVENS", True)
            return
        self.populate_pending_trees()

    def unmark_all_current(self):
        tab = self.current_tab_name()
        if tab == "EDP":
            for rec in self.pending_edp:
                rec.selected = False
                save_pending_selection(rec.supplier, rec.invoice_key, False)
        elif tab == "EPAL":
            for rec in self.pending_epal:
                rec.selected = False
                save_pending_selection(rec.supplier, rec.invoice_key, False)
        elif tab == "GALP":
            self.mark_rows("GALP", False)
            return
        elif tab == "DELTA":
            self.mark_rows("DELTA", False)
            return
        elif tab == "SAMSIC":
            self.mark_rows("SAMSIC", False)
            return
        elif tab == "EVIO":
            self.mark_rows("EVIO", False)
            return
        elif tab == "VIA VERDE":
            self.mark_rows("VIAVERDE", False)
            return
        elif tab == "AYVENS":
            self.mark_rows("AYVENS", False)
            return
        self.populate_pending_trees()

    def edit_current_line(self):
        tab = self.current_tab_name()
        if tab in ("EDP", "EPAL"):
            self.edit_simple_record(tab)
        elif tab == "GALP":
            self.edit_doc_row("GALP")
        elif tab == "DELTA":
            self.edit_doc_row("DELTA")
        elif tab == "SAMSIC":
            self.edit_doc_row("SAMSIC")
        elif tab == "EVIO":
            self.edit_doc_row("EVIO")
        elif tab == "VIA VERDE":
            self.edit_doc_row("VIAVERDE")
        elif tab == "AYVENS":
            self.edit_doc_row("AYVENS")
        else:
            messagebox.showinfo("Informação", "Escolhe uma tab com linhas editáveis.", parent=self)

    def edit_simple_record(self, supplier: str):
        tree = self.tree_edp if supplier == "EDP" else self.tree_epal
        records = self.pending_edp if supplier == "EDP" else self.pending_epal
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return
        idx = int(sel[0].split("::", 1)[1])
        rec = records[idx]
        new_piso = simpledialog.askstring("Piso", "Piso:", initialvalue=rec.piso, parent=self)
        if new_piso is not None:
            rec.piso = new_piso.strip()
        new_ca = simpledialog.askstring("CA", "CA:", initialvalue=rec.ca, parent=self)
        if new_ca is not None:
            rec.ca = new_ca.strip()
        new_period = simpledialog.askstring("Período", "Período (YYYYMM):", initialvalue=rec.period, parent=self)
        if new_period is not None and new_period.strip():
            rec.period = new_period.strip()
        if supplier == "EDP" and rec.cpe_suffix and rec.ca:
            rec.final_name = f"EDP_{rec.cpe_suffix}_{rec.source_path.stem}_CA{rec.ca}.pdf"
        if supplier == "EPAL" and rec.cl and rec.invoice_digits and rec.ca:
            rec.final_name = f"EPAL_{rec.cl}_{rec.invoice_digits}_CA{rec.ca}.pdf"
        self.populate_pending_trees()

    def _persist_viaverde_unknown_override(self, row: dict[str, Any]):
        identifier = str(row.get("vv_identifier", "")).strip()
        reference = str(row.get("vv_reference", "")).strip()
        period = str(row.get("periode", "")).strip()
        if not period or (not identifier and not reference):
            return
        desc = str(row.get("description", "")).strip()
        df = read_csv_df(VIAVERDE_UNKNOWN_OVERRIDE_FILE, VIAVERDE_UNKNOWN_OVERRIDE_HEADER)
        if df.empty:
            df = pd.DataFrame(columns=VIAVERDE_UNKNOWN_OVERRIDE_HEADER)
        for col in VIAVERDE_UNKNOWN_OVERRIDE_HEADER:
            if col not in df.columns:
                df[col] = ""
        target_idx = None
        special = "1" if str(row.get("vv_special", "0")).strip() == "1" else "0"
        for idx, rec in df.iterrows():
            if str(rec.get("period","")).strip() == period and str(rec.get("identifier","")).strip() == identifier and str(rec.get("reference","")).strip() == reference and str(rec.get("special","0")).strip() == special:
                target_idx = idx
                break
        payload = {col: str(row.get(col, "")).strip() for col in VIAVERDE_UNKNOWN_OVERRIDE_HEADER if col not in ("period","identifier","reference","special","active")}
        payload["period"] = period
        payload["identifier"] = identifier
        payload["reference"] = reference
        payload["special"] = special
        payload["active"] = "1"
        if target_idx is None:
            df.loc[len(df)] = payload
        else:
            for col, val in payload.items():
                df.at[target_idx, col] = val
        write_csv_df(VIAVERDE_UNKNOWN_OVERRIDE_FILE, df[VIAVERDE_UNKNOWN_OVERRIDE_HEADER])


    def _persist_viaverde_relation_row(self, row: dict[str, Any]):
        desc = str(row.get("description", "")).strip()
        if not desc:
            return
        base_plate = desc[:-4] if desc.lower().endswith("_est") else desc
        norm = normalize_plate(base_plate)
        df = read_csv_df(VIAVERDE_RELATION_FILE, VIAVERDE_RELATION_HEADER)
        if df.empty:
            df = pd.DataFrame(columns=VIAVERDE_RELATION_HEADER)
        for col in VIAVERDE_RELATION_HEADER:
            if col not in df.columns:
                df[col] = ""
        target_idx = None
        for idx, rec in df.iterrows():
            if normalize_plate(str(rec.get("description", ""))) == norm:
                target_idx = idx
                break
        payload = {col: str(row.get(col, "")).strip() for col in VIAVERDE_RELATION_HEADER if col != "active"}
        payload["description"] = base_plate
        payload["ana5"] = base_plate
        payload["active"] = "1"
        if target_idx is None:
            df.loc[len(df)] = payload
        else:
            for col, val in payload.items():
                df.at[target_idx, col] = val
        write_csv_df(VIAVERDE_RELATION_FILE, df[VIAVERDE_RELATION_HEADER])

    def _persist_evio_mapping_row(self, row: dict[str, Any]):
        desc = str(row.get("description", "")).strip()
        if not desc:
            return
        norm = normalize_plate(desc)
        df = read_csv_df(EVIO_MAPPING_FILE, EVIO_MAPPING_HEADER)
        if df.empty:
            df = pd.DataFrame(columns=EVIO_MAPPING_HEADER)
        for col in EVIO_MAPPING_HEADER:
            if col not in df.columns:
                df[col] = ""
        target_idx = None
        for idx, rec in df.iterrows():
            if normalize_plate(str(rec.get("description", ""))) == norm:
                target_idx = idx
                break
        payload = {col: str(row.get(col, "")).strip() for col in EVIO_MAPPING_HEADER if col != "active"}
        payload["description"] = desc
        payload["ana5"] = desc
        payload["active"] = "1"
        if target_idx is None:
            df.loc[len(df)] = payload
        else:
            for col, val in payload.items():
                df.at[target_idx, col] = val
        write_csv_df(EVIO_MAPPING_FILE, df[EVIO_MAPPING_HEADER])

    def _normalize_viaverde_hidden_defaults(self, row: dict[str, Any]) -> dict[str, Any]:
        row["prodfourn"] = str(row.get("prodfourn", "")).strip() or str(row.get("produit", "")).strip()
        row["compte"] = str(row.get("compte", "")).strip() or "62510200"
        row["ana1"] = str(row.get("ana1", "")).strip() or "NV1"
        row["interco"] = str(row.get("interco", "")).strip()
        row["st"] = str(row.get("st", "")).strip()
        row["t"] = str(row.get("t", "")).strip() or "D"
        return row

    def _ask_viaverde_visible_fields(self, row: dict[str, Any]) -> Optional[dict[str, str]]:
        dlg = tk.Toplevel(self)
        dlg.title("Editar linha VIA VERDE")
        dlg.transient(self)
        dlg.grab_set()
        dlg.resizable(False, False)

        entries: dict[str, tk.Entry] = {}
        for i, field in enumerate(EDITABLE_FIELDS_VIAVERDE):
            ttk.Label(dlg, text=field).grid(row=i, column=0, sticky="w", padx=8, pady=4)
            ent = ttk.Entry(dlg, width=40)
            ent.grid(row=i, column=1, sticky="ew", padx=8, pady=4)
            ent.insert(0, str(row.get(field, "")))
            entries[field] = ent

        note = ttk.Label(
            dlg,
            text="Nota: apenas os campos visíveis são editáveis. Os campos técnicos ocultos são preenchidos automaticamente.",
            foreground="#555"
        )
        note.grid(row=len(EDITABLE_FIELDS_VIAVERDE), column=0, columnspan=2, sticky="w", padx=8, pady=(8,4))

        result = {"ok": False}
        def on_ok():
            result["ok"] = True
            dlg.destroy()

        def on_cancel():
            dlg.destroy()

        btns = ttk.Frame(dlg)
        btns.grid(row=len(EDITABLE_FIELDS_VIAVERDE)+1, column=0, columnspan=2, sticky="e", padx=8, pady=8)
        ttk.Button(btns, text="OK", command=on_ok).pack(side="left", padx=4)
        ttk.Button(btns, text="Cancelar", command=on_cancel).pack(side="left", padx=4)

        dlg.columnconfigure(1, weight=1)
        entries[EDITABLE_FIELDS_VIAVERDE[0]].focus_set()
        dlg.wait_window()

        if not result["ok"]:
            return None
        return {field: entries[field].get().strip() for field in EDITABLE_FIELDS_VIAVERDE}

    def edit_doc_row(self, supplier: str):
        if supplier == "GALP":
            doc = self.current_galp_doc()
            tree = self.tree_galp
        elif supplier == "DELTA":
            doc = self.current_delta_doc()
            tree = self.tree_delta
        elif supplier == "SAMSIC":
            doc = self.current_samsic_doc()
            tree = self.tree_samsic
        elif supplier == "EVIO":
            doc = self.current_evio_doc()
            tree = self.tree_evio
        elif supplier in ("VIAVERDE", "VIA VERDE"):
            doc = self.current_viaverde_doc()
            tree = self.tree_viaverde
        else:
            doc = self.current_ayvens_doc()
            tree = self.tree_ayvens
        if not doc:
            return

        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return

        idx = int(sel[0])
        row = dict(doc.rows[idx])

        if supplier in ("VIAVERDE", "VIA VERDE"):
            values = self._ask_viaverde_visible_fields(row)
            if values is None:
                return
            changed = any(str(row.get(field, "")) != values[field] for field in EDITABLE_FIELDS_VIAVERDE)
            if not changed:
                messagebox.showinfo("Informação", "Sem alterações.", parent=self)
                return
            for field, value in values.items():
                row[field] = value

            desc = str(row.get("description", "")).strip() or "DESCONHECIDA"
            row["description"] = desc
            if not str(row.get("ana5", "")).strip():
                row["ana5"] = desc
            row["manual_required"] = False
            self._normalize_viaverde_hidden_defaults(row)

            base_plate = desc[:-4] if desc.lower().endswith("_est") else desc
            is_unknown = normalize_plate(base_plate) == "DESCONHECIDA" or not str(doc.rows[idx].get("vv_identifier", "")).strip()
            if is_unknown:
                self._persist_viaverde_unknown_override(row)
            else:
                self._persist_viaverde_relation_row(row)

            doc.rows[idx] = row
            self.load_viaverde_current(refresh_only=True)
            messagebox.showinfo("Sucesso", "Alteração gravada.", parent=self)
            return

        fields = [k for k in row.keys() if k not in ("confirmed", "agresso")]
        changed = False
        for field in fields:
            current = str(row.get(field, ""))
            new_val = simpledialog.askstring("Editar linha", field, initialvalue=current, parent=self)
            if new_val is None:
                return
            new_val = new_val.strip()
            if new_val != current:
                row[field] = new_val
                changed = True

        if not changed:
            messagebox.showinfo("Informação", "Sem alterações.", parent=self)
            return

        doc.rows[idx] = row
        if supplier == "EVIO":
            self._persist_evio_mapping_row(row)
            self.load_evio_current(refresh_only=True)
        elif supplier == "SAMSIC":
            self.load_samsic_current(refresh_only=True)
        elif supplier == "DELTA":
            self.load_delta_current(refresh_only=True)
        elif supplier == "GALP":
            self.load_galp_current(refresh_only=True)
        elif supplier == "AYVENS":
            self.load_ayvens_current(refresh_only=True)
        messagebox.showinfo("Sucesso", "Alteração gravada.", parent=self)
    def on_pending_tree_click(self, event, supplier: str):
        tree = self.tree_edp if supplier == "EDP" else self.tree_epal
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        allowed_columns = ("#1",)
        if not item_id or column_id not in allowed_columns:
            return
        idx = int(item_id.split("::", 1)[1])
        records = self.pending_edp if supplier == "EDP" else self.pending_epal
        rec = records[idx]
        if column_id == "#1":
            rec.selected = not rec.selected
            save_pending_selection(rec.supplier, rec.invoice_key, rec.selected)

        self.populate_pending_trees()
        return "break"

    def on_row_toggle(self, event, supplier: str):
        tree = self.tree_galp if supplier == "GALP" else self.tree_delta if supplier == "DELTA" else self.tree_samsic if supplier == "SAMSIC" else self.tree_evio if supplier == "EVIO" else self.tree_viaverde if supplier == "VIAVERDE" else self.tree_ayvens
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or col != "#1":
            return
        idx = int(item)
        doc = self.current_galp_doc() if supplier == "GALP" else self.current_delta_doc() if supplier == "DELTA" else self.current_samsic_doc() if supplier == "SAMSIC" else self.current_evio_doc() if supplier == "EVIO" else self.current_viaverde_doc() if supplier == "VIAVERDE" else self.current_ayvens_doc()
        if not doc or idx >= len(doc.rows):
            return "break"
        doc.rows[idx]["confirmed"] = not doc.rows[idx]["confirmed"]
        doc.rows[idx]["agresso"] = "☑" if doc.rows[idx]["confirmed"] else "☐"
        self.populate_pending_trees()
        return "break"

    # ---------------- doc navigation
    def save_current_delta_doc_state(self):
        doc = self.current_delta_doc()
        if doc:
            doc.ca = self.delta_ca_var.get().strip()
            doc.period = self.delta_period_var.get().strip() or doc.period

    def save_current_samsic_doc_state(self):
        doc = self.current_samsic_doc()
        if doc:
            doc.ca = self.samsic_ca_var.get().strip()
            doc.period = self.samsic_period_var.get().strip() or doc.period

    def save_current_evio_doc_state(self):
        doc = self.current_evio_doc()
        if doc:
            doc.ca = self.evio_ca_var.get().strip()
            doc.period = self.evio_period_var.get().strip() or doc.period

    def save_current_galp_ca_to_doc(self):
        doc = self.current_galp_doc()
        if doc:
            doc.ca = self.galp_ca_var.get().strip()
            doc.period = self.galp_period_var.get().strip() or doc.period

    def current_galp_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_galp:
            return None
        if self.galp_index >= len(self.pending_galp):
            self.galp_index = 0
        return self.pending_galp[self.galp_index]

    def current_delta_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_delta:
            return None
        if self.delta_index >= len(self.pending_delta):
            self.delta_index = 0
        return self.pending_delta[self.delta_index]

    def prev_galp(self):
        if self.pending_galp:
            self.save_current_galp_ca_to_doc()
            self.galp_index = (self.galp_index - 1) % len(self.pending_galp)
            self.load_galp_current(refresh_only=True)

    def next_galp(self):
        if self.pending_galp:
            self.save_current_galp_ca_to_doc()
            self.galp_index = (self.galp_index + 1) % len(self.pending_galp)
            self.load_galp_current(refresh_only=True)

    def prev_delta(self):
        if self.pending_delta:
            self.save_current_delta_doc_state()
            self.delta_index = (self.delta_index - 1) % len(self.pending_delta)
            self.load_delta_current(refresh_only=True)

    def next_delta(self):
        if self.pending_delta:
            self.save_current_delta_doc_state()
            self.delta_index = (self.delta_index + 1) % len(self.pending_delta)
            self.load_delta_current(refresh_only=True)


    def current_samsic_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_samsic:
            return None
        if self.samsic_index >= len(self.pending_samsic):
            self.samsic_index = 0
        return self.pending_samsic[self.samsic_index]

    def current_evio_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_evio:
            return None
        if self.evio_index >= len(self.pending_evio):
            self.evio_index = 0
        return self.pending_evio[self.evio_index]

    def prev_samsic(self):
        if self.pending_samsic:
            self.save_current_samsic_doc_state()
            self.samsic_index = (self.samsic_index - 1) % len(self.pending_samsic)
            self.load_samsic_current(refresh_only=True)

    def next_samsic(self):
        if self.pending_samsic:
            self.save_current_samsic_doc_state()
            self.samsic_index = (self.samsic_index + 1) % len(self.pending_samsic)
            self.load_samsic_current(refresh_only=True)

    def prev_evio(self):
        if self.pending_evio:
            self.save_current_evio_doc_state()
            self.evio_index = (self.evio_index - 1) % len(self.pending_evio)
            self.load_evio_current(refresh_only=True)

    def next_evio(self):
        if self.pending_evio:
            self.save_current_evio_doc_state()
            self.evio_index = (self.evio_index + 1) % len(self.pending_evio)
            self.load_evio_current(refresh_only=True)


    def save_current_viaverde_doc_state(self):
        doc = self.current_viaverde_doc()
        if doc:
            doc.ca = self.viaverde_ca_var.get().strip()
            doc.period = self.viaverde_period_var.get().strip() or doc.period

    def current_viaverde_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_viaverde:
            return None
        if self.viaverde_index >= len(self.pending_viaverde):
            self.viaverde_index = 0
        return self.pending_viaverde[self.viaverde_index]

    def prev_viaverde(self):
        if self.pending_viaverde:
            self.save_current_viaverde_doc_state()
            self.viaverde_index = (self.viaverde_index - 1) % len(self.pending_viaverde)
            self.load_viaverde_current(refresh_only=True)

    def next_viaverde(self):
        if self.pending_viaverde:
            self.save_current_viaverde_doc_state()
            self.viaverde_index = (self.viaverde_index + 1) % len(self.pending_viaverde)
            self.load_viaverde_current(refresh_only=True)

    def save_current_ayvens_doc_state(self):
        doc = self.current_ayvens_doc()
        if doc:
            doc.ca = self.ayvens_ca_var.get().strip()
            doc.period = self.ayvens_period_var.get().strip() or doc.period

    def current_ayvens_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_ayvens:
            return None
        if self.ayvens_index >= len(self.pending_ayvens):
            self.ayvens_index = 0
        return self.pending_ayvens[self.ayvens_index]

    def prev_ayvens(self):
        if self.pending_ayvens:
            self.save_current_ayvens_doc_state()
            self.ayvens_index = (self.ayvens_index - 1) % len(self.pending_ayvens)
            self.load_ayvens_current(refresh_only=True)

    def next_ayvens(self):
        if self.pending_ayvens:
            self.save_current_ayvens_doc_state()
            self.ayvens_index = (self.ayvens_index + 1) % len(self.pending_ayvens)
            self.load_ayvens_current(refresh_only=True)

    def mark_rows(self, supplier: str, value: bool):
        doc = self.current_galp_doc() if supplier == "GALP" else self.current_delta_doc() if supplier == "DELTA" else self.current_samsic_doc() if supplier == "SAMSIC" else self.current_evio_doc() if supplier == "EVIO" else self.current_viaverde_doc() if supplier == "VIAVERDE" else self.current_ayvens_doc()
        if not doc:
            return
        for row in doc.rows:
            row["confirmed"] = value
            row["agresso"] = "☑" if value else "☐"
        self.populate_pending_trees()

    def load_galp_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_galp"):
            return
        for i in self.tree_galp.get_children():
            self.tree_galp.delete(i)
        doc = self.current_galp_doc()
        if not doc:
            self.galp_doc_info.set("-")
            self.galp_status.set("")
            self.galp_total_var.set("Resumo total: -")
            return

        if doc.doc_type == "annual":
            annual_ca = self.galp_admin.get("annual_card_ca", "")
            self.galp_previous_ca_var.set("")
            if not doc.ca:
                doc.ca = annual_ca
        elif doc.doc_type == "fuel":
            previous_ca = self.galp_admin.get("last_fuel_ca", "") or get_last_processed_ca("GALP", "fuel")
            self.galp_previous_ca_var.set(previous_ca)
            if not getattr(doc, "ca", ""):
                doc.ca = ""
        else:
            self.galp_previous_ca_var.set("")
        self.galp_ca_var.set(getattr(doc, "ca", "") or "")
        self.galp_period_var.set(getattr(doc, "period", "") or "")

        self.galp_doc_info.set(
            f"Documento {self.galp_index + 1}/{len(self.pending_galp)} | Nº: {doc.invoice_number} | "
            f"Período: {doc.period} | Tipo: {doc.doc_type}"
        )
        extra = " CA anual fixo." if doc.doc_type == "annual" else " Introduz o novo CA actual para renomear e enviar para faturas." if doc.doc_type == "fuel" else ""
        self.galp_status.set(f"Foram carregadas {len(doc.rows)} linhas sumarizadas por matrícula.{extra}")
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.galp_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        for idx, row in enumerate(doc.rows):
            row = normalize_money_fields_in_row(dict(row))
            self.tree_galp.insert("", "end", iid=str(idx), values=(
                row.get("agresso", "☐"), row.get("description", ""), row.get("tipo", ""),
                row.get("produit", ""), row.get("prodfourn", ""), row.get("unite", ""),
                row.get("periode", ""), row.get("nombre", ""), row.get("mnt", ""),
                row.get("compte", ""), row.get("ana1", ""), row.get("project", ""), row.get("resno", ""),
                row.get("ana4", ""), row.get("ana5", ""), row.get("dep", ""),
                row.get("ct", "")
            ))
        self._apply_tree_tags(self.tree_galp)

    def load_delta_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_delta"):
            return
        for i in self.tree_delta.get_children():
            self.tree_delta.delete(i)
        doc = self.current_delta_doc()
        if not doc:
            self.delta_doc_info.set("-")
            self.delta_status.set("")
            self.delta_total_var.set("Resumo total: -")
            return

        if not self.delta_ca_var.get().strip():
            self.delta_ca_var.set(self.delta_admin.get("last_ca", "") or get_last_processed_ca("DELTA"))

        self.delta_period_var.set(getattr(doc, "period", "") or "")
        self.delta_doc_info.set(
            f"Documento {self.delta_index + 1}/{len(self.pending_delta)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.delta_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        sem_mapping = sum(1 for r in doc.rows if not str(r.get("produit", "")).strip())
        self.delta_status.set(f"DELTA carregada com {len(doc.rows)} linhas úteis. Sem mapping: {sem_mapping}.")
        for idx, row in enumerate(doc.rows):
            row = normalize_money_fields_in_row(dict(row))
            self.tree_delta.insert("", "end", iid=str(idx), values=(
                row.get("agresso", "☐"), row.get("description", ""), row.get("produit", ""),
                row.get("periode", ""), row.get("nombre", ""), row.get("mnt", ""),
                row.get("iva", ""), row.get("code_iva", "")
            ))
        self._apply_tree_tags(self.tree_delta)


    def load_samsic_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_samsic"):
            return
        for i in self.tree_samsic.get_children():
            self.tree_samsic.delete(i)
        doc = self.current_samsic_doc()
        if not doc:
            self.samsic_doc_info.set("-")
            self.samsic_status.set("")
            self.samsic_total_var.set("Resumo total: -")
            return

        if not self.samsic_ca_var.get().strip():
            self.samsic_ca_var.set(self.samsic_admin.get("current_annual_ca", ""))

        self.samsic_period_var.set(getattr(doc, "period", "") or "")
        self.samsic_doc_info.set(
            f"Documento {self.samsic_index + 1}/{len(self.pending_samsic)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.samsic_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        self.samsic_status.set("Confirma os valores e, se necessário, actualiza o CA anual antes de processar.")
        for idx, row in enumerate(doc.rows):
            row = normalize_money_fields_in_row(dict(row))
            self.tree_samsic.insert("", "end", iid=str(idx), values=tuple(row.get(c, "☐" if c=="agresso" else "") for c in VISIBLE_FIELDS_COMMON_REDUCED))
        self._apply_tree_tags(self.tree_samsic)


    def load_evio_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_evio"):
            return
        for i in self.tree_evio.get_children():
            self.tree_evio.delete(i)
        doc = self.current_evio_doc()
        if not doc:
            self.evio_doc_info.set("-")
            self.evio_status.set("")
            self.evio_total_var.set("Resumo total: -")
            return

        if not getattr(doc, "ca", ""):
            doc.ca = self.evio_admin.get("last_ca", "") or get_last_processed_ca("EVIO")
        self.evio_ca_var.set(getattr(doc, "ca", "") or "")
        self.evio_period_var.set(getattr(doc, "period", "") or "")
        self.evio_doc_info.set(
            f"Documento {self.evio_index + 1}/{len(self.pending_evio)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.evio_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        sem_mapping = sum(1 for r in doc.rows if not str(r.get("produit", "")).strip())
        self.evio_status.set(f"EVIO carregada com {len(doc.rows)} linhas sumarizadas por matrícula. Sem mapping: {sem_mapping}.")
        for idx, row in enumerate(doc.rows):
            row = normalize_money_fields_in_row(dict(row))
            self.tree_evio.insert("", "end", iid=str(idx), values=tuple(row.get(c, "☐" if c=="agresso" else "") for c in VISIBLE_FIELDS_COMMON_REDUCED))
        self._apply_tree_tags(self.tree_evio)


    def load_viaverde_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_viaverde"):
            return
        for i in self.tree_viaverde.get_children():
            self.tree_viaverde.delete(i)
        doc = self.current_viaverde_doc()
        if not doc:
            self.viaverde_doc_info.set("-")
            self.viaverde_status.set("")
            self.viaverde_total_var.set("Resumo total: -")
            return

        prev_ca = self.viaverde_admin.get("last_ca", "") or get_last_processed_ca("VIAVERDE")
        self.viaverde_previous_ca_var.set(prev_ca)
        if not getattr(doc, "ca", ""):
            doc.ca = prev_ca
        self.viaverde_ca_var.set(getattr(doc, "ca", "") or "")
        self.viaverde_period_var.set(getattr(doc, "period", "") or "")
        self.viaverde_doc_info.set(
            f"Documento {self.viaverde_index + 1}/{len(self.pending_viaverde)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.viaverde_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        sem_mapping = sum(1 for r in doc.rows if not str(r.get("produit", "")).strip())
        manual_req = sum(1 for r in doc.rows if bool(r.get("manual_required", False)))
        ficheiros_origem = len(getattr(doc, "source_names", []) or [])
        self.viaverde_status.set(
            f"VIA VERDE mensal consolidada com {len(doc.rows)} linhas com consumo > 0, a partir de {ficheiros_origem or 1} ficheiro(s) de suporte. "
            f"Sem mapping: {sem_mapping}. Identificação manual obrigatória: {manual_req}. Serviços especiais surgem como linha separada com sufixo _est."
        )
        for idx, row in enumerate(doc.rows):
            row = normalize_money_fields_in_row(dict(row))
            self.tree_viaverde.insert("", "end", iid=str(idx), values=tuple(row.get(c, "☐" if c=="agresso" else "") for c in VISIBLE_FIELDS_VIAVERDE))
        self._apply_tree_tags(self.tree_viaverde)

    def load_ayvens_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_ayvens"):
            return
        for i in self.tree_ayvens.get_children():
            self.tree_ayvens.delete(i)
        doc = self.current_ayvens_doc()
        if not doc:
            self.ayvens_doc_info.set("-")
            self.ayvens_status.set("")
            self.ayvens_total_var.set("Resumo total: -")
            return

        if not getattr(doc, "ca", ""):
            doc.ca = resolve_ayvens_ca(doc.period, doc.doc_type, self.ayvens_admin)
        self.ayvens_ca_var.set(getattr(doc, "ca", "") or "")
        self.ayvens_period_var.set(getattr(doc, "period", "") or "")
        self.ayvens_doc_info.set(
            f"Documento {self.ayvens_index + 1}/{len(self.pending_ayvens)} | Nº: {doc.invoice_number} | Período: {doc.period} | Tipo: {doc.doc_type}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.ayvens_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        self.ayvens_status.set("AYVENS carregada. Rendas: locação sem IVA, excepto AX-06-SZ e BA-21-FV; serviços sujeitos x1,23; isentos directos da fatura.")
        for idx, row in enumerate(doc.rows):
            row = normalize_money_fields_in_row(dict(row))
            self.tree_ayvens.insert("", "end", iid=str(idx), values=(
                row.get("agresso", "☐"), row.get("description", ""), row.get("type", ""),
                row.get("produit", ""), row.get("prodfourn", ""), row.get("unite", ""),
                row.get("periode", ""), row.get("nombre", ""), row.get("prixunit", ""),
                row.get("mnt", ""), row.get("iva", ""), row.get("code_iva", "")
            ))
        self._apply_tree_tags(self.tree_ayvens)

    # ---------------- processing

    def validate_period_value(self, period: str) -> bool:
        return bool(re.match(r"^\d{6}$", period or ""))

    def process_current_tab(self):
        tab = self.current_tab_name()
        if tab == "EDP":
            self.process_simple_supplier("EDP", self.pending_edp)
        elif tab == "EPAL":
            self.process_simple_supplier("EPAL", self.pending_epal)
        elif tab == "GALP":
            self.process_galp()
        elif tab == "DELTA":
            self.process_delta()
        elif tab == "SAMSIC":
            self.process_samsic()
        elif tab == "EVIO":
            self.process_evio()
        elif tab == "VIA VERDE":
            self.process_viaverde()
        elif tab == "AYVENS":
            self.process_ayvens()
        else:
            messagebox.showinfo("Informação", "Não há nada para processar nesta tab.", parent=self)

    def process_simple_supplier(self, supplier: str, records: list[InvoiceRecord]):
        to_process = [r for r in records if r.selected]
        if not to_process:
            messagebox.showwarning("Aviso", "Não existem linhas seleccionadas para processar.", parent=self)
            return

        processed = 0
        for rec in to_process:
            if not rec.final_name:
                messagebox.showerror("Erro", f"Nome final não gerado para {rec.file_name}", parent=self)
                return
            if not rec.ca:
                messagebox.showerror("Erro", f"CA vazio para {rec.file_name}", parent=self)
                return

        for rec in to_process:
            destination_dir = BASE_DIR / supplier / rec.period
            ensure_dir(destination_dir)
            final_path = destination_dir / rec.final_name
            if final_path.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro final: {final_path.name}", parent=self)
                return

        for rec in to_process:
            destination_dir = BASE_DIR / supplier / rec.period
            final_path = destination_dir / rec.final_name
            shutil.move(str(rec.source_path), str(final_path))
            register_processed_invoice(supplier, rec.invoice_key, rec.file_hash, rec.invoice_number, rec.period,
                                       rec.doc_type, rec.ca, rec.file_name, rec.final_name, self.current_user)

            supplier_row = {
                "InvoiceNumber": rec.invoice_number,
                "CA": rec.ca,
                "Estado": "Processado",
                "Periodo": rec.period,
                "DocType": rec.doc_type or "standard",
                "PdfFile": rec.file_name,
                "FinalFile": rec.final_name,
                "ProcessedBy": self.current_user,
                "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            if supplier == "EDP":
                supplier_row = {
                    "InvoiceNumber": rec.invoice_number,
                    "CA": rec.ca,
                    "Estado": "Processado",
                    "kWh": f"{rec.kwh:.4f}",
                    "AV (€)": money_str(rec.av),
                    "Valor(€)": money_str(rec.total_before_iva_23),
                    "Periodo": rec.period,
                    "DocType": rec.doc_type or "standard",
                    "PdfFile": rec.file_name,
                    "FinalFile": rec.final_name,
                    "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            elif supplier == "EPAL":
                supplier_row = {
                    "Periodo": rec.period,
                    "CA": rec.ca,
                    "A_Faturar": f"{rec.a_faturar:.3f}",
                    "A_Deduzir": f"{rec.a_deduzir:.3f}",
                    "M3": f"{rec.m3:.3f}",
                    "Abastecimento": f"{rec.abastecimento:.4f}",
                    "Saneamento": f"{rec.saneamento:.4f}",
                    "Residuos": f"{rec.residuos:.4f}",
                    "Adicional": f"{rec.adicional:.4f}",
                    "Taxas": f"{rec.taxas:.4f}",
                    "Valor": money_str(rec.total),
                    "CL": rec.cl,
                    "Piso": rec.piso,
                    "InvoiceNumber": rec.invoice_number,
                    "DocType": rec.doc_type or "standard",
                    "Estado": "Processado",
                    "PdfFile": rec.file_name,
                    "FinalFile": rec.final_name,
                    "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            append_df_to_sheet(supplier, pd.DataFrame([supplier_row]))
            clear_pending_selection(rec.supplier, rec.invoice_key)

            append_history_row({
                "Supplier": supplier,
                "CA": rec.ca,
                "Periodo": rec.period,
                "InvoiceNumber": rec.invoice_number,
                "DocType": rec.doc_type or "standard",
                "Estado": "Processado",
                "PdfFile": rec.file_name,
                "FinalFile": rec.final_name,
                "ProcessedBy": self.current_user,
                "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            processed += 1
        messagebox.showinfo("Sucesso", f"{supplier}: {processed} fatura(s) processada(s).", parent=self)
        self.load_all()

    def process_galp(self):
        doc = self.current_galp_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas GALP para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas GALP por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_galp_ca_to_doc()
        period = self.galp_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.galp_ca_var.get().strip()
        if not ca:
            if doc.doc_type == "fuel":
                messagebox.showerror("Erro", "Indica o novo CA actual para o combustível. O CA anterior é apenas referência.", parent=self)
            else:
                messagebox.showerror("Erro", "Indica o CA actual.", parent=self)
            return
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "GALP" / period
        ensure_dir(dest_dir)

        pdf_new = dest_dir / f"GALP_{doc.invoice_number}_{period}_{ca_suffix}{doc.source_path.suffix}"
        aux_new = None
        if pdf_new.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {pdf_new.name}.", parent=self)
            return

        if doc.aux_path:
            aux_new = dest_dir / f"GALP_{doc.invoice_number}_{period}_{ca_suffix}{doc.aux_path.suffix}"
            if aux_new.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro {aux_new.name}.", parent=self)
                return

        shutil.move(str(doc.source_path), str(pdf_new))
        if doc.aux_path and doc.aux_path.exists():
            shutil.move(str(doc.aux_path), str(aux_new))

        if doc.doc_type == "fuel":
            self.set_admin_value(GALP_ADMIN_FILE, GALP_DEFAULT_ADMIN, "last_fuel_ca", ca)

        register_processed_invoice("GALP", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, pdf_new.name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "GALP", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalLitros": round(sum(float(r.get("nombre", 0)) for r in doc.rows), 2),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "AuxFile": doc.aux_path.name if doc.aux_path else "",
            "FinalPdf": pdf_new.name, "FinalAux": aux_new.name if aux_new else "",
            "ProcessedBy": self.current_user,
            "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Tipo": r.get("tipo", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "Compte": r.get("compte", ""),
            "Ana1": r.get("ana1", ""),
            "PROJECT": r.get("project", ""),
            "RESNO": r.get("resno", ""),
            "Ana4": r.get("ana4", ""),
            "Ana5": r.get("ana5", ""),
            "DEP": r.get("dep", ""),
            "CT": r.get("ct", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
        } for r in doc.rows])
        append_df_to_sheet("GALP", summary_df)
        output_df = normalize_currency_df("GALP_OUTPUT", output_df)
        append_df_to_sheet("GALP_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"GALP - Fatura {doc.invoice_number} - {doc.period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.galp_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[pdf_new]
        )
        messagebox.showinfo("Sucesso", f"Fatura GALP processada.\n{msg}", parent=self)
        self.load_all()


    def process_delta(self):
        doc = self.current_delta_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas DELTA para processar.", parent=self)
            return

        # Sincroniza primeiro o estado visível do ecrã para doc.rows
        self.save_current_delta_doc_state()

        rows_to_process = [r for r in doc.rows if bool(r.get("confirmed", False))]
        if not rows_to_process:
            messagebox.showerror("Erro", "Marca pelo menos uma linha DELTA antes de processar.", parent=self)
            return

        def _is_missing(v: object) -> bool:
            s = str(v).strip()
            return s in ("", "None", "nan", "NaN")

        unresolved = [
            r.get("description", "") or "(sem descrição)"
            for r in rows_to_process
            if bool(r.get("manual_required", False))
            or _is_missing(r.get("description", ""))
            or _is_missing(r.get("produit", ""))
            or _is_missing(r.get("periode", ""))
            or _is_missing(r.get("nombre", ""))
            or _is_missing(r.get("mnt", ""))
            or _is_missing(r.get("iva", ""))
            or _is_missing(r.get("code_iva", ""))
        ]
        if unresolved:
            messagebox.showerror(
                "Erro",
                "Existem linhas DELTA por completar ou rever antes de processar:\n- "
                + "\n- ".join(map(str, unresolved[:10])),
                parent=self,
            )
            return

        period = self.delta_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in rows_to_process:
            row["periode"] = period

        ca = self.delta_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual para a DELTA.", parent=self)
            return

        self.set_admin_value(DELTA_ADMIN_FILE, DELTA_DEFAULT_ADMIN, "last_ca", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "DELTA" / period
        ensure_dir(dest_dir)

        final_name = f"DELTA_{doc.source_path.stem}_{ca_suffix}{doc.source_path.suffix}"
        final_path = dest_dir / final_name
        if final_path.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
            return

        shutil.move(str(doc.source_path), str(final_path))
        register_processed_invoice(
            "DELTA",
            doc.invoice_key,
            doc.file_hash,
            doc.invoice_number,
            period,
            doc.doc_type,
            ca,
            doc.file_name,
            final_name,
            self.current_user,
        )

        summary_df = pd.DataFrame([{
            "Supplier": "DELTA", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(rows_to_process),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in rows_to_process)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "FinalPdf": final_name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Produit": r.get("produit", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
        } for r in rows_to_process])
        append_df_to_sheet("DELTA", summary_df)
        output_df = normalize_currency_df("DELTA_OUTPUT", output_df)
        append_df_to_sheet("DELTA_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"DELTA - Fatura {doc.invoice_number} - {doc.period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.delta_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_path]
        )
        messagebox.showinfo("Sucesso", f"Fatura DELTA processada.\n{msg}", parent=self)
        self.load_all()


    def process_evio(self):
        doc = self.current_evio_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas EVIO para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas EVIO por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_evio_doc_state()
        period = self.evio_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.evio_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da EVIO.", parent=self)
            return

        self.set_admin_value(EVIO_ADMIN_FILE, EVIO_DEFAULT_ADMIN, "last_ca", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "EVIO" / period
        ensure_dir(dest_dir)

        short_inv = short_evio_invoice_number(doc.invoice_number)
        final_pdf_name = f"EVIO_{short_inv}_{ca_suffix}{doc.source_path.suffix}"
        final_pdf = dest_dir / final_pdf_name
        if final_pdf.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_pdf_name}.", parent=self)
            return

        final_aux = None
        if doc.aux_path and doc.aux_path.exists():
            final_aux = dest_dir / f"EVIO_{short_inv}_{ca_suffix}{doc.aux_path.suffix}"
            if final_aux.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro {final_aux.name}.", parent=self)
                return

        shutil.move(str(doc.source_path), str(final_pdf))
        if doc.aux_path and doc.aux_path.exists():
            shutil.move(str(doc.aux_path), str(final_aux))

        register_processed_invoice("EVIO", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, final_pdf_name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "EVIO", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "AuxFile": doc.aux_path.name if doc.aux_path else "",
            "FinalPdf": final_pdf_name, "FinalAux": final_aux.name if final_aux else "",
            "ProcessedBy": self.current_user,
            "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "PrixUnit": r.get("prixunit", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
            "Compte": r.get("compte", ""),
            "Ana1": r.get("ana1", ""),
            "PROJECT": r.get("project", ""),
            "RESNO": r.get("resno", ""),
            "Ana4": r.get("ana4", ""),
            "Ana5": r.get("ana5", ""),
            "DEP": r.get("dep", ""),
            "INTERCO": r.get("interco", ""),
            "CT": r.get("ct", ""),
            "ST": r.get("st", ""),
            "T": r.get("t", ""),
        } for r in doc.rows])
        append_df_to_sheet("EVIO", summary_df)
        output_df = normalize_currency_df("EVIO_OUTPUT", output_df)
        append_df_to_sheet("EVIO_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"EVIO - Fatura {doc.invoice_number} - {period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.evio_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_pdf]
        )
        messagebox.showinfo("Sucesso", f"Fatura EVIO processada.\n{msg}", parent=self)
        self.load_all()


    def process_viaverde(self):
        doc = self.current_viaverde_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas VIA VERDE para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        for row in doc.rows:
            self._normalize_viaverde_hidden_defaults(row)
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("description", "")).strip() or not str(r.get("produit", "")).strip() or not str(r.get("unite", "")).strip() or not str(r.get("periode", "")).strip() or str(r.get("nombre", "")).strip() in ("", "None") or str(r.get("mnt", "")).strip() in ("", "None") or not str(r.get("project", "")).strip() or not str(r.get("resno", "")).strip() or not str(r.get("ana4", "")).strip() or not str(r.get("ana5", "")).strip() or not str(r.get("dep", "")).strip() or not str(r.get("ct", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas VIA VERDE por completar ou identificar manualmente antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_viaverde_doc_state()
        period = self.viaverde_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.viaverde_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da VIA VERDE.", parent=self)
            return

        self.set_admin_value(VIAVERDE_ADMIN_FILE, VIAVERDE_DEFAULT_ADMIN, "last_ca", ca)
        df_ca = read_viaverde_ca_mapping()
        mask = (df_ca.get("TIPO", pd.Series(dtype=str)).astype(str).str.upper() == "STANDARD") & (df_ca.get("PERIODO", pd.Series(dtype=str)).astype(str) == period)
        if df_ca.empty:
            df_ca = pd.DataFrame(columns=["TIPO","ANO","MES","PERIODO","CA","DESCRICAO","ACTIVE"])
        if mask.any():
            df_ca.loc[mask, "CA"] = ca
        else:
            df_ca.loc[len(df_ca)] = ["STANDARD", period[:4], period[4:6], period, ca, f"VIAVERDE - {period}", "1"]
        write_csv_df(VIAVERDE_CA_FILE, df_ca)

        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "VIA VERDE" / period
        ensure_dir(dest_dir)

        final_pdfs = []
        all_files = getattr(doc, "all_files", [doc.source_path])
        pdfs = getattr(doc, "pdf_files", [])
        for p in all_files:
            final_name = f"ViaVerde_{p.stem}_{ca_suffix}{p.suffix}" if p.suffix.lower()==".pdf" else p.name
            final_path = dest_dir / final_name
            if final_path.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
                return
        for p in all_files:
            final_name = f"ViaVerde_{p.stem}_{ca_suffix}{p.suffix}" if p.suffix.lower()==".pdf" else p.name
            final_path = dest_dir / final_name
            shutil.move(str(p), str(final_path))
            if p.suffix.lower()==".pdf":
                final_pdfs.append(final_path)
        register_processed_invoice("VIAVERDE", doc.invoice_key, doc.file_hash, doc.invoice_number, period, doc.doc_type, ca, doc.file_name, ", ".join(fp.name for fp in final_pdfs), self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "VIAVERDE", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": "standard", "RowsCount": len(doc.rows), "TotalValorHT": round_money(sum(float(r.get("mnt",0) or 0) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": ", ".join(p.name for p in pdfs), "FinalPdf": ", ".join(p.name for p in final_pdfs),
            "FinalXml": doc.source_path.name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }])
        append_df_to_sheet("VIAVERDE", summary_df)

        output_rows = []
        for row in doc.rows:
            output_rows.append({
                "Description": row.get("description",""), "Produit": row.get("produit",""), "ProdFourn": row.get("prodfourn",""),
                "Unité": row.get("unite",""), "Période": row.get("periode",""), "Nombre": row.get("nombre",""),
                "PrixUnit": row.get("prixunit",""), "Mnt HT en dev.": row.get("mnt",""), "Compte": row.get("compte",""),
                "Ana1": row.get("ana1",""), "PROJECT": row.get("project",""), "RESNO": row.get("resno",""), "Ana4": row.get("ana4",""),
                "Ana5": row.get("ana5",""), "DEP": row.get("dep",""), "INTERCO": row.get("interco",""), "CT": row.get("ct",""),
                "ST": row.get("st",""), "T": row.get("t","")
            })
        append_df_to_sheet("VIAVERDE_OUTPUT", normalize_currency_df("VIAVERDE_OUTPUT", pd.DataFrame(output_rows)))
        append_history_row({
            "Supplier":"VIAVERDE","CA":ca,"Periodo":period,"InvoiceNumber":doc.invoice_number,"DocType":"standard",
            "Estado":"Processado","PdfFile":", ".join(p.name for p in pdfs),"FinalFile":", ".join(p.name for p in final_pdfs),
            "ProcessedBy":self.current_user,"ProcessedAt":datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        ok, msg = create_outlook_draft(
            subject=f"VIA VERDE - {period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.viaverde_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=final_pdfs
        )
        messagebox.showinfo("Sucesso", f"Via Verde processada.\n{msg}", parent=self)
        self.load_all()

    def set_admin_value(self, path: Path, defaults: list[dict[str, str]], key: str, value: str):
        info, df = read_admin_info(path, defaults)
        mask = df["key"].astype(str).str.strip() == key
        if mask.any():
            df.loc[mask, "value"] = value
        else:
            df.loc[len(df)] = [key, value, ""]
        write_admin_info(path, df)


    def process_ayvens(self):
        doc = self.current_ayvens_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas AYVENS para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas AYVENS por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_ayvens_doc_state()
        period = self.ayvens_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.ayvens_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da AYVENS.", parent=self)
            return

        self.set_admin_value(AYVENS_ADMIN_FILE, AYVENS_DEFAULT_ADMIN, f"{'extra_ca_' + period[:4] if doc.doc_type == 'extra' else 'rent_ca_' + period}", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        subfolder = "EXTRAS" if doc.doc_type == "extra" else "RENDAS"
        dest_dir = BASE_DIR / "AYVENS" / period / subfolder
        ensure_dir(dest_dir)

        short_inv = re.sub(r"[^A-Za-z0-9]+", "", doc.invoice_number or "") or "SEMNUMERO"
        final_name = f"AYVENS_{doc.doc_type.upper()}_{short_inv}_{ca_suffix}{doc.source_path.suffix}"
        final_path = dest_dir / final_name
        if final_path.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
            return

        shutil.move(str(doc.source_path), str(final_path))
        register_processed_invoice("AYVENS", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, final_name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "AYVENS", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "FinalPdf": final_name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Type": r.get("type", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "PrixUnit": r.get("prixunit", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
            "Compte": r.get("compte", ""),
            "Ana1": r.get("ana1", ""),
            "PROJECT": r.get("project", ""),
            "RESNO": r.get("resno", ""),
            "Ana4": r.get("ana4", ""),
            "Ana5": r.get("ana5", ""),
            "DEP": r.get("dep", ""),
            "INTERCO": r.get("interco", ""),
            "CT": r.get("ct", ""),
            "ST": r.get("st", ""),
            "T": r.get("t", ""),
        } for r in doc.rows])
        append_df_to_sheet("AYVENS", summary_df)
        output_df = normalize_currency_df("AYVENS_OUTPUT", output_df)
        append_df_to_sheet("AYVENS_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"AYVENS - Fatura {doc.invoice_number} - {period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.ayvens_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_path]
        )
        messagebox.showinfo("Sucesso", f"Fatura AYVENS processada.\n{msg}", parent=self)
        self.load_all()

    def process_samsic(self):
        doc = self.current_samsic_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas SAMSIC para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas SAMSIC por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_samsic_doc_state()
        period = self.samsic_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.samsic_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da SAMSIC.", parent=self)
            return

        self.set_admin_value(SAMSIC_ADMIN_FILE, SAMSIC_DEFAULT_ADMIN, "current_annual_ca", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "SAMSIC" / period
        ensure_dir(dest_dir)

        short_inv = short_samsic_invoice_number(doc.invoice_number)
        final_name = f"SAMSIC_{short_inv}_{ca_suffix}{doc.source_path.suffix}"
        final_path = dest_dir / final_name
        if final_path.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
            return

        shutil.move(str(doc.source_path), str(final_path))
        register_processed_invoice("SAMSIC", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, final_name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "SAMSIC", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "FinalPdf": final_name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "PrixUnit": r.get("prixunit", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
        } for r in doc.rows])
        append_df_to_sheet("SAMSIC", summary_df)
        output_df = normalize_currency_df("SAMSIC_OUTPUT", output_df)
        append_df_to_sheet("SAMSIC_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"SAMSIC - Fatura {doc.invoice_number} - {doc.period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.samsic_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_path]
        )
        messagebox.showinfo("Sucesso", f"Fatura SAMSIC processada.\n{msg}", parent=self)
        self.load_all()

    # ---------------- logging
    def log(self, msg: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{stamp}] {msg}\n")
        self.log_text.see("end")




# --- AYVENS v1.4.1 overrides -------------------------------------------------

def find_matching_evio_excel(invoice_number: str) -> Optional[Path]:
    normalized = invoice_number.replace("/", "_").replace("  ", " ").strip()
    candidates = sorted(BASE_DIR.glob("EVIO*.xlsx"))
    target = normalize_text(normalized)
    for p in candidates:
        if target and target in normalize_text(p.stem):
            return p
    for p in candidates:
        if "resumo" in normalize_text(p.stem):
            return p
    return None

def parse_evio_document(pdf_path: Path) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    invoice_number = find_first(r"(FT\s*\d{4}[A-Z]\d+/\d+)", text, re.IGNORECASE)
    invoice_number = re.sub(r"\s+", " ", invoice_number).strip()
    invoice_date = find_first(r"(\d{4}-\d{2}-\d{2})", text, re.IGNORECASE)
    period = datetime.now().strftime("%Y%m")
    if invoice_date:
        try:
            period = datetime.strptime(invoice_date, "%Y-%m-%d").strftime("%Y%m")
        except Exception:
            pass
    aux_path = find_matching_evio_excel(invoice_number)
    file_hash = file_sha256(pdf_path)
    invoice_key = f"EVIO|{invoice_number}|{period}|charging"

    rec = InvoiceRecord(
        supplier="EVIO",
        source_path=pdf_path,
        file_name=pdf_path.name,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        period=period,
        doc_type="charging",
        aux_path=aux_path,
        file_hash=file_hash,
        invoice_key=invoice_key,
    )
    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
        return rec
    if not aux_path or not aux_path.exists():
        rec.errors.append("Excel auxiliar EVIO não encontrado.")
        rec.status = "Erro"
        return rec

    try:
        df = pd.read_excel(aux_path, dtype=str).fillna("")
        mapping = read_evio_mapping()
        cols = {normalize_text(c): c for c in df.columns}
        matricula_col = next((v for k,v in cols.items() if k == "matricula"), None)
        energia_col = next((v for k,v in cols.items() if k == "energia total"), None)
        total_col = next((v for k,v in cols.items() if "total excl. iva" in k), None)
        iva_col = next((v for k,v in cols.items() if "taxa de iva" in k), None)
        if not matricula_col or not energia_col or not total_col:
            raise RuntimeError("Colunas EVIO necessárias não detectadas no Excel auxiliar.")

        work = df[[matricula_col, energia_col, total_col] + ([iva_col] if iva_col else [])].copy()
        rename = {matricula_col: "matricula", energia_col: "energia", total_col: "valor"}
        if iva_col:
            rename[iva_col] = "iva"
        work = work.rename(columns=rename)
        work["matricula"] = work["matricula"].astype(str).str.strip()
        work = work[work["matricula"] != ""]
        work["energia"] = work["energia"].apply(safe_float)
        work["valor"] = work["valor"].apply(safe_float)
        if "iva" in work.columns:
            work["iva"] = work["iva"].apply(safe_float)
        else:
            work["iva"] = 23.0
        grouped = work.groupby("matricula", as_index=False).agg({"energia":"sum","valor":"sum","iva":"max"})

        rows=[]
        ignored=0
        mapping_norm = mapping.copy()
        if not mapping_norm.empty:
            mapping_norm["_plate_norm"] = mapping_norm["description"].astype(str).apply(normalize_plate)
        else:
            mapping_norm["_plate_norm"] = []

        for _, r in grouped.iterrows():
            matricula_raw = str(r["matricula"]).strip()
            plate_norm = normalize_plate(matricula_raw)
            match = mapping_norm[mapping_norm["_plate_norm"] == plate_norm] if not mapping_norm.empty else pd.DataFrame()
            mapped = not match.empty
            if mapped:
                m = match.iloc[0]
                if str(m.get("active", "1")).strip() == "0":
                    ignored += 1
                    continue
            else:
                ignored += 1
                m = pd.Series(dtype=str)

            iva = int(round(float(r["iva"] or 23)))
            code_iva = "BG" if iva == 23 else "BR" if iva == 6 else ""
            rows.append({
                "confirmed": False,
                "agresso": "☐",
                "description": matricula_raw,
                "produit": str(m.get("produit","")).strip(),
                "prodfourn": str(m.get("prodfourn","")).strip(),
                "unite": str(m.get("unite","")).strip() or "US",
                "periode": period,
                "f": "U",
                "nombre": round(float(r["energia"]), 2),
                "prixunit": round((float(r["valor"]) / float(r["energia"])) if float(r["energia"]) else float(r["valor"]), 4),
                "escompte": 0,
                "mnt": round_money(float(r["valor"])),
                "devise": "EUR",
                "s": "F",
                "compte": str(m.get("compte","")).strip(),
                "ana1": str(m.get("ana1","")).strip(),
                "project": str(m.get("project","")).strip(),
                "resno": str(m.get("resno","")).strip(),
                "ana4": str(m.get("ana4","")).strip(),
                "ana5": str(m.get("ana5","")).strip() or matricula_raw,
                "dep": str(m.get("dep","")).strip(),
                "interco": str(m.get("interco","")).strip() or "9",
                "ct": str(m.get("ct","")).strip() or code_iva,
                "st": str(m.get("st","")).strip(),
                "t": str(m.get("t","")).strip() or "D",
                "iva": iva,
                "code_iva": code_iva,
                "mapped": mapped,
            })
        rec.rows = rows
        # EVIO deve aparecer mesmo com matrículas sem mapping, para permitir correcção na gestão
        if not rows:
            rec.errors.append("Nenhuma linha EVIO detectada no Excel auxiliar.")
            rec.status = "Erro"
    except Exception as e:
        rec.errors.append(str(e))
        rec.status = "Erro"
    return rec

# ============================================================
# MANAGEMENT UI
# ============================================================

# ============================================================
# PARSERS - VIA VERDE
# ============================================================

def parse_viaverde_csv_document(csv_path: Path, admin: dict[str, str]) -> InvoiceRecord:
    period = parse_viaverde_period_from_xml(csv_path)
    file_hash = file_sha256(csv_path)
    invoice_key = f"VIAVERDE|{csv_path.stem}|{period}"
    rec = InvoiceRecord(supplier="VIAVERDE", source_path=csv_path, file_name=csv_path.name, invoice_number=csv_path.stem, period=period, doc_type="standard", file_hash=file_hash, invoice_key=invoice_key)
    if invoice_already_processed(invoice_key, file_hash):
        rec.status = "Duplicado"
        return rec
    pdfs, all_files = _find_viaverde_associated_files(csv_path)
    rec.pdf_files = pdfs
    rec.all_files = all_files
    relations = read_viaverde_relations()
    rel_map = {normalize_plate(str(r["description"])): r.to_dict() for _, r in relations.iterrows() if str(r.get("active","1"))!="0"}
    agg: dict[tuple[str, bool], dict[str, Any]] = {}
    try:
        df = pd.read_csv(csv_path, sep=';', encoding='latin1', skiprows=7, dtype=str, engine='python').fillna('')
        for _, row in df.iterrows():
            amount = safe_float(row.get('VALOR','0'))
            if amount <= 0:
                continue
            raw_plate = str(row.get('MATRÍCULA','')).strip()
            conta_digital = str(row.get('CONTA DIGITAL','')).strip().lower()
            service_type = str(row.get('SERVIÇO','')).strip()
            operador = str(row.get('OPERADOR','')).strip()
            entrada = str(row.get('ENTRADA','')).strip()
            saida = str(row.get('SAÍDA','')).strip()
            special = _viaverde_is_special(service_type, operador, entrada, saida, conta_digital)
            _add_viaverde_agg(agg, rel_map, period, raw_plate, amount, special=special, conta_digital=conta_digital, admin=admin, errors=rec.errors)
    except Exception as e:
        rec.errors.append(f"Erro CSV Via Verde: {e}")
    rec.rows = list(agg.values())
    rec.ca, year = resolve_viaverde_ca(period, admin)
    rec.year = year
    rec.status = "Erro" if rec.errors and not rec.rows else "Pendente"
    return rec


def parse_viaverde_detail_pdf_document(pdf_path: Path, admin: dict[str, str]) -> InvoiceRecord:
    period = parse_viaverde_period_from_name(pdf_path)
    file_hash = file_sha256(pdf_path)
    invoice_key = f"VIAVERDE|{pdf_path.stem}|{period}|DETAILPDF"
    rec = InvoiceRecord(supplier="VIAVERDE", source_path=pdf_path, file_name=pdf_path.name, invoice_number=pdf_path.stem, period=period, doc_type="standard", file_hash=file_hash, invoice_key=invoice_key)
    if invoice_already_processed(invoice_key, file_hash):
        rec.status = "Duplicado"
        return rec

    pdfs, all_files = _find_viaverde_associated_files_from_detail(pdf_path)
    rec.pdf_files = pdfs
    rec.all_files = all_files

    relations = read_viaverde_relations()
    rel_map = {normalize_plate(str(r["description"])): r.to_dict() for _, r in relations.iterrows() if str(r.get("active", "1")) != "0"}
    ov_df = read_viaverde_unknown_overrides()
    ov_map = {}
    for _, r in ov_df.iterrows():
        if str(r.get("active","1")) == "0":
            continue
        ov_map[(str(r.get("period","")).strip(), str(r.get("identifier","")).strip(), str(r.get("reference","")).strip(), str(r.get("special","0")).strip())] = r.to_dict()

    digital_map = _viaverde_digital_mapping(admin)
    agg: dict[tuple[str, bool], dict[str, Any]] = {}

    text = extract_pdf_text(pdf_path)
    lines = _iter_viaverde_page1_lines(text)
    in_digital = False
    for line in lines:
        if 'Serviços Digitais' in line or 'Servicos Digitais' in line:
            in_digital = True
            continue
        if any(x in line for x in ['PAGAMENTOS DE SERVIÇOS VIA VERDE', 'PAGAMENTOS DE SERVICOS VIA VERDE', 'VALORES TOTAIS', 'Identificador Matricula Referência', 'Identificador Matricula Referencia', 'Referência Total Alertas', 'Referencia Total Alertas', '(*) Sem alertas']):
            continue
        if in_digital:
            m = re.match(r'^([A-Za-z0-9._%+-]+@[A-Za-z0-9.-]+)\s+(\d+)\s+([\d,]+)$', line, re.I)
            if m:
                email = m.group(1).strip()
                reference = m.group(2).strip()
                amount = safe_float(m.group(3))
                mapped_plate = digital_map.get(_viaverde_email_key(email), '')
                if mapped_plate:
                    _add_viaverde_agg(agg, rel_map, period, mapped_plate, amount, special=True, conta_digital=email, admin=admin, errors=rec.errors, reference=reference, unknown_override_map=ov_map)
                else:
                    display = f"{_viaverde_email_key(email)}_est"
                    key = (normalize_plate(display), True)
                    if key not in agg:
                        note = f"Email Via Verde sem mapping: {email}. Edita o mapa de gestão."
                        agg[key] = _viaverde_unresolved_row(display, period, 0.0, note, reference=reference, special=True)
                        rec.errors.append(note)
                    agg[key]["mnt"] = f"{safe_float(agg[key].get('mnt', 0)) + amount:.2f}"
                    agg[key]["prixunit"] = agg[key]["mnt"]
            continue
        m = re.match(r'^(\d{9,})\s+([A-Z0-9-]+|Desconhecida)\s+(\d+)\s+([\d,]+)$', line, re.I)
        if m:
            identifier = m.group(1).strip()
            raw_plate = m.group(2).strip()
            reference = m.group(3).strip()
            amount = safe_float(m.group(4))
            _add_viaverde_agg(agg, rel_map, period, raw_plate, amount, special=False, conta_digital='', admin=admin, errors=rec.errors, identifier=identifier, reference=reference, unknown_override_map=ov_map)

    rec.rows = [row for row in agg.values() if safe_float(row.get('mnt', 0)) > 0]
    rec.rows = apply_viaverde_unknown_overrides(rec.rows, period)
    rec.ca, year = resolve_viaverde_ca(period, admin)
    rec.year = year
    rec.status = "Erro" if rec.errors and not rec.rows else "Pendente"
    return rec
def parse_viaverde_period_from_name(path: Path) -> str:
    m = re.search(r"_(20\d{2})_(\d{2})(?:\D|$)", path.stem)
    if m:
        return f"{m.group(1)}{m.group(2)}"
    return datetime.now().strftime("%Y%m")

def parse_viaverde_period_from_xml(xml_path: Path) -> str:
    try:
        root = ET.parse(xml_path).getroot()
        mes = (root.findtext("MES_EMISSAO") or "").strip()
        if len(mes) >= 6:
            return mes[:6]
    except Exception:
        pass
    m = re.search(r"(20\d{2})_(\d{2})", xml_path.stem)
    if m:
        return f"{m.group(1)}{m.group(2)}"
    return datetime.now().strftime("%Y%m")

def parse_viaverde_xml_document(xml_path: Path, admin: dict[str, str]) -> InvoiceRecord:
    period = parse_viaverde_period_from_xml(xml_path)
    file_hash = file_sha256(xml_path)
    invoice_key = f"VIAVERDE|{xml_path.stem}|{period}"
    rec = InvoiceRecord(supplier="VIAVERDE", source_path=xml_path, file_name=xml_path.name, invoice_number=xml_path.stem, period=period, doc_type="standard", file_hash=file_hash, invoice_key=invoice_key)
    if invoice_already_processed(invoice_key, file_hash):
        rec.status = "Duplicado"
        return rec
    pdfs, all_files = _find_viaverde_associated_files(xml_path)
    rec.pdf_files = pdfs
    rec.all_files = all_files
    relations = read_viaverde_relations()
    rel_map = {normalize_plate(str(r["description"])): r.to_dict() for _, r in relations.iterrows() if str(r.get("active","1"))!="0"}
    agg: dict[tuple[str, bool], dict[str, Any]] = {}
    try:
        root = ET.parse(xml_path).getroot()
        for ident in root.findall("IDENTIFICADOR"):
            conta_digital = (ident.findtext("CONTA_DIGITAL") or "").strip().lower()
            ident_total = safe_float(ident.findtext("TOTAL") or "0")
            trans_sum = 0.0
            for trans in ident.findall("TRANSACCAO"):
                raw_plate = (trans.findtext("MATRICULA") or "").strip()
                amount = safe_float(trans.findtext("IMPORTANCIA") or "0")
                if amount <= 0:
                    continue
                trans_sum += amount
                service_type = (trans.findtext("TIPO") or "").strip()
                operador = (trans.findtext("OPERADOR") or "").strip()
                entrada = (trans.findtext("ENTRADA") or "").strip()
                saida = (trans.findtext("SAIDA") or "").strip()
                special = _viaverde_is_special(service_type, operador, entrada, saida, conta_digital)
                _add_viaverde_agg(agg, rel_map, period, raw_plate, amount, special=special, conta_digital=conta_digital, admin=admin, errors=rec.errors)
            if trans_sum <= 0 and ident_total > 0:
                raw_plate = (ident.findtext("TRANSACCAO/MATRICULA") or ident.findtext("MATRICULA") or "").strip()
                _add_viaverde_agg(agg, rel_map, period, raw_plate, ident_total, special=bool(conta_digital), conta_digital=conta_digital, admin=admin, errors=rec.errors)
    except Exception as e:
        rec.errors.append(f"Erro XML Via Verde: {e}")
    rec.rows = list(agg.values())
    rec.ca, year = resolve_viaverde_ca(period, admin)
    rec.year = year
    rec.status = "Erro" if rec.errors and not rec.rows else "Pendente"
    return rec

def read_evio_mapping() -> pd.DataFrame:
    _bootstrap_evio_mapping_file()
    df = read_csv_df(EVIO_MAPPING_FILE, EVIO_MAPPING_HEADER)
    for col in EVIO_MAPPING_HEADER:
        if col not in df.columns:
            df[col] = ""
    df["active"] = df["active"].replace("", "1")
    return df[EVIO_MAPPING_HEADER].fillna("")

def read_viaverde_ca_mapping() -> pd.DataFrame:
    cols = ["TIPO","ANO","MES","PERIODO","CA","DESCRICAO","ACTIVE"]
    df = read_csv_df(VIAVERDE_CA_FILE, cols)
    if df.empty:
        return pd.DataFrame(columns=cols)
    for c in cols:
        if c not in df.columns:
            df[c] = ""
    return df[cols].fillna("")


def read_viaverde_relations() -> pd.DataFrame:
    existing = read_csv_df(VIAVERDE_RELATION_FILE, VIAVERDE_RELATION_HEADER)
    base = pd.DataFrame(VIAVERDE_DEFAULT_RELATIONS, columns=VIAVERDE_RELATION_HEADER)

    for col in VIAVERDE_RELATION_HEADER:
        if col not in existing.columns:
            existing[col] = ""
        if col not in base.columns:
            base[col] = ""

    # Importante:
    # o ficheiro gravado pelo utilizador tem de ganhar aos defaults.
    # Por isso os defaults entram primeiro e o existing entra por último.
    merged = pd.concat(
        [base[VIAVERDE_RELATION_HEADER], existing[VIAVERDE_RELATION_HEADER]],
        ignore_index=True
    ).fillna("")

    merged["__plate"] = merged["description"].astype(str).apply(normalize_plate)
    merged["__priority"] = range(len(merged))
    merged = merged.sort_values(["__plate", "__priority"]).drop_duplicates(["__plate"], keep="last")
    merged = merged.drop(columns=["__plate", "__priority"])
    merged["active"] = merged["active"].replace("", "1")

    write_csv_df(VIAVERDE_RELATION_FILE, merged[VIAVERDE_RELATION_HEADER])
    return merged[VIAVERDE_RELATION_HEADER].fillna("")


def read_viaverde_unknown_overrides() -> pd.DataFrame:
    existing = read_csv_df(VIAVERDE_UNKNOWN_OVERRIDE_FILE, VIAVERDE_UNKNOWN_OVERRIDE_HEADER)
    if existing.empty:
        existing = pd.DataFrame(columns=VIAVERDE_UNKNOWN_OVERRIDE_HEADER)
    for col in VIAVERDE_UNKNOWN_OVERRIDE_HEADER:
        if col not in existing.columns:
            existing[col] = ""
    existing["active"] = existing["active"].replace("", "1")
    write_csv_df(VIAVERDE_UNKNOWN_OVERRIDE_FILE, existing[VIAVERDE_UNKNOWN_OVERRIDE_HEADER])
    return existing[VIAVERDE_UNKNOWN_OVERRIDE_HEADER].fillna("")


def apply_viaverde_unknown_overrides(rows: list[dict[str, Any]], period: str) -> list[dict[str, Any]]:
    ov_df = read_viaverde_unknown_overrides()
    if ov_df.empty:
        return rows
    ov_map: dict[tuple[str, str, str, str], dict[str, Any]] = {}
    for _, r in ov_df.iterrows():
        if str(r.get("active", "1")).strip() == "0":
            continue
        key = (
            str(r.get("period", "")).strip(),
            str(r.get("identifier", "")).strip(),
            str(r.get("reference", "")).strip(),
            str(r.get("special", "0")).strip(),
        )
        ov_map[key] = r.to_dict()

    out: list[dict[str, Any]] = []
    for row in rows:
        identifier = str(row.get("vv_identifier", "")).strip()
        reference = str(row.get("vv_reference", "")).strip()
        special = "1" if str(row.get("vv_special", "0")).strip() == "1" else "0"
        key = (str(period).strip(), identifier, reference, special)
        override = ov_map.get(key)
        if override:
            display_desc = str(override.get("description", "")).strip() or ("DESCONHECIDA_est" if special == "1" else "DESCONHECIDA")
            new_row = dict(row)
            for fld in ("description","produit","prodfourn","unite","compte","ana1","project","resno","ana4","ana5","dep","interco","ct","st","t"):
                if str(override.get(fld, "")).strip():
                    new_row[fld] = str(override.get(fld, "")).strip()
            new_row["description"] = display_desc
            if not str(new_row.get("ana5", "")).strip():
                new_row["ana5"] = display_desc
            new_row["manual_required"] = False
            out.append(new_row)
        else:
            out.append(row)
    return out
def resolve_viaverde_ca(period: str, admin: dict[str, str]) -> tuple[str, str]:
    year = (period or "")[:4]
    df = read_viaverde_ca_mapping()
    if not df.empty:
        hit = df[(df["TIPO"].astype(str).str.upper()=="STANDARD") & (df["PERIODO"].astype(str)==str(period)) & (df["ACTIVE"].astype(str)!="0")]
        if not hit.empty:
            return str(hit.iloc[0]["CA"]).strip(), year
    return admin.get("last_ca",""), year

def short_evio_invoice_number(invoice_number: str) -> str:
    return re.sub(r"[^A-Za-z0-9]+", "", invoice_number or "")

def short_samsic_invoice_number(invoice_number: str) -> str:
    return re.sub(r"[^A-Za-z0-9]", "", invoice_number or "") or "SEMNUMERO"



# ============================================================
# PARSERS - EVIO
# ============================================================

def ensure_environment():
    ensure_dir(BIN_DIR)
    ensure_dir(BACKUP_DIR)
    init_db()
    ensure_excel()

    if not EDP_MAP_FILE.exists():
        pd.DataFrame(columns=["Suffix", "Piso", "CA"]).to_csv(EDP_MAP_FILE, index=False, encoding="utf-8-sig")
    if not EPAL_MAP_FILE.exists():
        pd.DataFrame(columns=["CL", "Piso", "CA"]).to_csv(EPAL_MAP_FILE, index=False, encoding="utf-8-sig")

    if not GALP_MAPPING_FILE.exists():
        pd.DataFrame(columns=GALP_MAPPING_HEADER).to_csv(GALP_MAPPING_FILE, index=False, encoding="utf-8-sig")
    if not GALP_ADMIN_FILE.exists():
        pd.DataFrame(GALP_DEFAULT_ADMIN).to_csv(GALP_ADMIN_FILE, index=False, encoding="utf-8-sig")

    if not DELTA_MAPPING_FILE.exists():
        pd.DataFrame(DELTA_DEFAULT_MAPPING).to_csv(DELTA_MAPPING_FILE, index=False, encoding="utf-8-sig")
    if not DELTA_ADMIN_FILE.exists():
        pd.DataFrame(DELTA_DEFAULT_ADMIN).to_csv(DELTA_ADMIN_FILE, index=False, encoding="utf-8-sig")
    if not SAMSIC_ADMIN_FILE.exists():
        pd.DataFrame(SAMSIC_DEFAULT_ADMIN).to_csv(SAMSIC_ADMIN_FILE, index=False, encoding="utf-8-sig")
    if not EVIO_ADMIN_FILE.exists():
        pd.DataFrame(EVIO_DEFAULT_ADMIN).to_csv(EVIO_ADMIN_FILE, index=False, encoding="utf-8-sig")
    if not EVIO_MAPPING_FILE.exists():
        pd.DataFrame(EVIO_DEFAULT_MAPPING, columns=EVIO_MAPPING_HEADER).to_csv(EVIO_MAPPING_FILE, index=False, encoding="utf-8-sig")
    else:
        _bootstrap_evio_mapping_file()
    if not VIAVERDE_ADMIN_FILE.exists():
        pd.DataFrame(VIAVERDE_DEFAULT_ADMIN).to_csv(VIAVERDE_ADMIN_FILE, index=False, encoding="utf-8-sig")
    else:
        _bootstrap_admin_defaults(VIAVERDE_ADMIN_FILE, VIAVERDE_DEFAULT_ADMIN)
    if not VIAVERDE_RELATION_FILE.exists():
        pd.DataFrame(VIAVERDE_DEFAULT_RELATIONS, columns=VIAVERDE_RELATION_HEADER).to_csv(VIAVERDE_RELATION_FILE, index=False, encoding="utf-8-sig")
    if not VIAVERDE_CA_FILE.exists():
        pd.DataFrame(columns=["TIPO","ANO","MES","PERIODO","CA","DESCRICAO","ACTIVE"]).to_csv(VIAVERDE_CA_FILE, index=False, encoding="utf-8-sig")
    if not AYVENS_RELATION_FILE.exists():
        pd.DataFrame(columns=_ayvens_relation_columns()).to_csv(AYVENS_RELATION_FILE, index=False, encoding="utf-8-sig")
    bootstrap_ayvens_from_excel()


# ============================================================
# PARSERS - EDP / EPAL
# ============================================================
def find_first(pattern: str, text: str, flags: int = 0) -> str:
    m = re.search(pattern, text, flags)
    return m.group(1).strip() if m else ""


def extract_kwh_total(text: str) -> str:
    matches = re.findall(
        r"Imposto sobre Cons\. Eletricidade\s+\d{2}/\d{2}/\d{4}\s+\d{2}/\d{2}/\d{4}\s+([\-\d\.]+,\d{4})",
        text, re.IGNORECASE
    )
    total = 0.0
    for value in matches:
        total += safe_float(value)
    return f"{total:.4f}".replace(".", ",")


def extract_edp_av_total(text: str) -> str:
    matches = re.findall(
        r"Contribuição Áudio-Visual.*?([\-\d\.]+,\d{2})\s*€",
        text, re.IGNORECASE | re.DOTALL
    )
    total = 0.0
    for value in matches:
        total += safe_float(value)
    return f"{total:.2f}".replace(".", ",")


def extract_edp_total_before_iva_23(text: str) -> str:
    match = re.search(
        r"Total\s*\(antes de IVA a 23%\)\s+23%\s+([\d\.]+,\d{2})\s*€",
        text, re.IGNORECASE | re.DOTALL
    )
    return match.group(1).strip() if match else ""


def parse_edp_pdf(pdf_path: Path, edp_map: dict[str, dict[str, str]]) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    invoice_date = find_first(r"DATA DE EMISSÃO\s+(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    cpe = find_first(r"CÓDIGO PONTO ENTREGA ELETRICIDADE\s+([A-Z0-9]+)", text, re.IGNORECASE)
    billing_period = find_first(r"PERÍODO DE FATURAÇÃO\s+(\d{2}/\d{2}/\d{4}\s+a\s+\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    invoice_number = find_first(r"NÚMERO\s+(?:DA\s+FATURA|NOTA\s+CRÉDITO)\s+((?:FT|NC)\s+[A-Z0-9/]+)", text, re.IGNORECASE)
    av = extract_edp_av_total(text)
    total_before = extract_edp_total_before_iva_23(text)
    kwh = extract_kwh_total(text)
    suffix = cpe[-2:] if len(cpe) >= 2 else ""

    mapping = edp_map.get(suffix, {})
    ca = mapping.get("CA", "")
    piso = mapping.get("Piso", "")

    period = datetime.now().strftime("%Y%m")
    if billing_period:
        m = re.search(r"a\s+(\d{2})/(\d{2})/(\d{4})", billing_period)
        if m:
            period = f"{m.group(3)}{m.group(2)}"

    file_hash = file_sha256(pdf_path)
    invoice_key = f"EDP|{invoice_number}|{cpe}|{period}"
    final_name = f"EDP_{suffix}_{pdf_path.stem}_CA{ca}.pdf" if suffix and ca else ""

    rec = InvoiceRecord(
        supplier="EDP",
        source_path=pdf_path,
        file_name=pdf_path.name,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        billing_period=billing_period,
        period=period,
        ca=ca,
        piso=piso,
        cpe=cpe,
        cpe_suffix=suffix,
        av=pt_to_float(av),
        total_before_iva_23=pt_to_float(total_before),
        kwh=pt_to_float(kwh),
        file_hash=file_hash,
        invoice_key=invoice_key,
        final_name=final_name,
    )
    if not suffix:
        rec.errors.append("Não foi possível determinar o suffix do CPE.")
    if suffix and not ca:
        rec.errors.append(f"CA não encontrado para suffix {suffix}.")
    if suffix and not piso:
        rec.errors.append(f"Piso não encontrado para suffix {suffix}.")
    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
    elif rec.errors:
        rec.status = "Erro"
    return rec


def extract_invoice_number_digits(invoice_text: str) -> str:
    return re.sub(r"\D", "", invoice_text or "")


def parse_epal_pdf(pdf_path: Path, epal_map: dict[str, dict[str, str]]) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    cl = find_first(r"LOCAL\s*N[.ºo°]*\s*-\s*(\d+)", text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
    invoice = ""
    for pattern in [
        r"FATURA\s*N[.ºo°]*\s*((?:FT)\s*\d+/\d+)",
        r"NOTA\s+DE\s+CRÉDITO\s*N[.ºo°]*\s*((?:NC)\s*\d+/\d+)",
        r"RESUMO\s+DA\s+FATURA\s*N[.ºo°]*\s*((?:FT)\s*\d+/\d+)",
        r"RESUMO\s+DA\s+NOTA\s+DE\s+CRÉDITO\s*N[.ºo°]*\s*((?:NC)\s*\d+/\d+)",
    ]:
        invoice = find_first(pattern, text, re.IGNORECASE | re.MULTILINE | re.DOTALL)
        if invoice:
            invoice = re.sub(r"\s+", " ", invoice).strip()
            break

    invoice_digits = extract_invoice_number_digits(invoice)
    invoice_date = find_first(r"emitida\s+em\s+(\d{2}-\d{2}-\d{4})", text, re.IGNORECASE).replace("-", "/")
    billing_period = ""
    m = re.search(r"Período\s+de\s+Faturação.*?:\s*(\d{2}-\d{2}-\d{4})\s*a\s*(\d{2}-\d{2}-\d{4})", text, re.IGNORECASE | re.DOTALL)
    if m:
        billing_period = f"{m.group(1).replace('-', '/')} a {m.group(2).replace('-', '/')}"

    total = 0.0
    for pat in [r"Valor\s+a\s+Receber\s+(-?[\d,]+)\s*€?", r"Valor\s+a\s+pagar\s+(-?[\d,]+)\s*€?"]:
        v = find_first(pat, text, re.IGNORECASE | re.DOTALL)
        if v:
            total = safe_float(v)
            break

    abastecimento = safe_float(find_first(r"EPAL\s*-\s*Abastecimento\s+de\s+Água\s+(-?[\d,]+)", text, re.IGNORECASE))
    saneamento = safe_float(find_first(r"CMLisboa\s*-\s*Saneamento\s+(-?[\d,]+)", text, re.IGNORECASE))
    residuos = safe_float(find_first(r"CMLisboa\s*-\s*Resíduos\s+Urbanos\s+(-?[\d,]+)", text, re.IGNORECASE))
    adicional = safe_float(find_first(r"CMLisboa\s*-\s*Adicional\s+(-?[\d,]+)", text, re.IGNORECASE))
    taxas = safe_float(find_first(r"Taxas\s+(-?[\d,]+)", text, re.IGNORECASE))

    def _extract_abastecimento_quantities(full_text: str) -> tuple[float, float]:
        a_faturar_local = 0.0
        a_deduzir_local = 0.0

        detail_match = re.search(
            r"DETALHE\s+DA\s+FATURA.*?EPAL\s*-\s*Abastecimento\s+de\s+Água(?:\(Escalão/30 dias\))?(.*?)(?=QUOTA\s+SERVIÇO)",
            full_text,
            re.IGNORECASE | re.DOTALL,
        )
        if not detail_match:
            return 0.0, 0.0

        block = detail_match.group(1)
        lines_local = [ln.strip() for ln in block.splitlines() if ln.strip()]
        date_re = re.compile(r"\d{2}-\d{2}-\d{4}\s+a\s+\d{2}-\d{2}-\d{4}$")
        qty_re = re.compile(r"-?\d+,\d{3}$")
        money_re = re.compile(r"-?\d+,\d{4}$")

        i = 0
        while i < len(lines_local):
            if not date_re.fullmatch(lines_local[i]):
                i += 1
                continue

            nums: list[str] = []
            j = i + 1
            while j < len(lines_local) and len(nums) < 4:
                candidate = lines_local[j]
                if qty_re.fullmatch(candidate) or money_re.fullmatch(candidate):
                    nums.append(candidate)
                j += 1

            if len(nums) >= 4 and qty_re.fullmatch(nums[0]) and money_re.fullmatch(nums[1]) and money_re.fullmatch(nums[2]) and money_re.fullmatch(nums[3]):
                qty = safe_float(nums[0])
                payable = safe_float(nums[3])
                if payable < 0:
                    a_deduzir_local += abs(qty)
                else:
                    a_faturar_local += qty
                i = j
                continue

            i += 1

        return a_faturar_local, a_deduzir_local

    a_faturar, a_deduzir = _extract_abastecimento_quantities(text)

    if a_faturar == 0.0 and a_deduzir == 0.0:
        cons = find_first(r"CONSUMO\s+FATURADO\s+.*?([\d\s,]+)\s+litros", text, re.IGNORECASE | re.DOTALL)
        if cons:
            a_faturar = safe_float(cons)

    m3 = a_faturar - a_deduzir

    period = datetime.now().strftime("%Y%m")
    if billing_period:
        m2 = re.search(r"a\s+(\d{2})/(\d{2})/(\d{4})", billing_period)
        if m2:
            period = f"{m2.group(3)}{m2.group(2)}"

    mapping = epal_map.get(cl, {})
    ca = mapping.get("CA", "")
    piso = mapping.get("Piso", "")
    file_hash = file_sha256(pdf_path)
    invoice_key = f"EPAL|{invoice}|{cl}|{period}"
    final_name = f"EPAL_{cl}_{invoice_digits}_CA{ca}.pdf" if cl and invoice_digits and ca else ""

    rec = InvoiceRecord(
        supplier="EPAL",
        source_path=pdf_path,
        file_name=pdf_path.name,
        invoice_number=invoice,
        invoice_digits=invoice_digits,
        invoice_date=invoice_date,
        billing_period=billing_period,
        period=period,
        cl=cl,
        ca=ca,
        piso=piso,
        total=total,
        a_faturar=a_faturar,
        a_deduzir=a_deduzir,
        m3=m3,
        abastecimento=abastecimento,
        saneamento=saneamento,
        residuos=residuos,
        adicional=adicional,
        taxas=taxas,
        file_hash=file_hash,
        invoice_key=invoice_key,
        final_name=final_name,
    )
    if not cl:
        rec.errors.append("CL não encontrada.")
    if cl and not ca:
        rec.errors.append(f"CA não encontrado para CL {cl}.")
    if cl and not piso:
        rec.errors.append(f"Piso não encontrado para CL {cl}.")
    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
    elif rec.errors:
        rec.status = "Erro"
    return rec


# ============================================================
# PARSERS - GALP
# ============================================================
def read_galp_mapping() -> pd.DataFrame:
    df = read_csv_df(GALP_MAPPING_FILE, GALP_MAPPING_HEADER)
    for col in GALP_MAPPING_HEADER:
        if col not in df.columns:
            df[col] = ""
    df["active"] = df["active"].replace("", "1")
    return df[GALP_MAPPING_HEADER].fillna("")


def find_matching_aux_file(pdf_path: Path) -> Optional[Path]:
    base = pdf_path.stem
    for ext in [".xlsx", ".xls"]:
        candidate = BASE_DIR / f"{base}{ext}"
        if candidate.exists():
            return candidate
    return None


def extract_period_from_galp_text(text: str) -> str:
    m = re.search(r"At[eé]\s+(\d{1,2})\s+([A-ZÇ]{3})\s+(\d{4})", text, re.IGNORECASE)
    if m:
        month = MONTHS_PT.get(m.group(2).upper())
        if month:
            return f"{m.group(3)}{month}"
    m2 = re.search(r"(\d{1,2})\s+([A-ZÇ]{3})\s+(\d{4})", text, re.IGNORECASE)
    if m2:
        month = MONTHS_PT.get(m2.group(2).upper())
        if month:
            return f"{m2.group(3)}{month}"
    return datetime.now().strftime("%Y%m")


def extract_galp_invoice_number(text: str) -> str:
    m = re.search(r"Fatura:\s*([0-9]+)", text, re.IGNORECASE)
    return m.group(1) if m else ""


def extract_galp_total_incl_iva(text: str) -> float:
    patterns = [
        r"VALOR TOTAL DESTA FATURA\s+[0-9]+,[0-9]+\s*\+\s*[0-9]+,[0-9]+\s*=\s*([0-9]+,[0-9]+)",
        r"TOTAL\s+([0-9]+,[0-9]+)\s*EUR",
        r"Total c/IVA\s*([0-9]+,[0-9]+)\s*EUR",
    ]
    for pat in patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            return safe_float(m.group(1))
    return 0.0


def detect_galp_doc_type(text: str) -> str:
    t = text.lower()
    if "serviços galp" in t and "anuidade" in t:
        return "annual"
    if "gasolina" in t or "gasóleo" in t or "gasoleo" in t:
        return "fuel"
    return "unknown"


def detect_excel_columns(df: pd.DataFrame) -> dict[str, Optional[str]]:
    cols = {"matricula": None, "litros": None, "valor": None}
    for col in df.columns:
        l = normalize_text(str(col))
        if l in ("description", "descricao", "matricula"):
            cols["matricula"] = col
        elif l in ("qt litros abast.", "qt litros abast", "qt litros", "qt. litros abast.", "nombre"):
            cols["litros"] = col
        elif l in ("valor liq. tot. c/iva (euros)", "valor liq. tot. c/iva", "valor liq tot c/iva (euros)", "mnt ht en dev."):
            cols["valor"] = col
    return cols


def parse_galp_document(pdf_path: Path) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    period = extract_period_from_galp_text(text)
    invoice_number = extract_galp_invoice_number(text) or pdf_path.stem
    doc_type = detect_galp_doc_type(text)
    aux_path = find_matching_aux_file(pdf_path)
    file_hash = file_sha256(pdf_path)
    invoice_key = f"GALP|{invoice_number}|{period}|{doc_type}"

    rec = InvoiceRecord(
        supplier="GALP",
        source_path=pdf_path,
        file_name=pdf_path.name,
        invoice_number=invoice_number,
        period=period,
        doc_type=doc_type,
        aux_path=aux_path,
        file_hash=file_hash,
        invoice_key=invoice_key,
    )
    if not aux_path:
        rec.errors.append("Excel auxiliar não encontrado.")
        rec.status = "Erro"
        return rec
    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
        return rec

    try:
        df = pd.read_excel(aux_path, dtype=str).fillna("")
        cols = detect_excel_columns(df)
        if not cols["matricula"] or not cols["litros"] or not cols["valor"]:
            raise RuntimeError("Não foi possível detectar as colunas do Excel auxiliar.")

        work = df[[cols["matricula"], cols["litros"], cols["valor"]]].copy()
        work.columns = ["matricula", "litros", "valor"]
        work["matricula"] = work["matricula"].astype(str).str.strip()
        work = work[work["matricula"] != ""]
        work["litros"] = work["litros"].apply(safe_float)
        work["valor"] = work["valor"].apply(safe_float)

        grouped = work.groupby("matricula", as_index=False).agg({"litros": "sum", "valor": "sum"})
        mapping = read_galp_mapping()
        annual_total_incl_iva = extract_galp_total_incl_iva(text) if doc_type == "annual" else 0.0
        annual_entretien = {"BC-35-EJ", "AX-06-SZ", "AA-75-AJ"}

        rows: list[dict[str, Any]] = []
        ignored = 0
        for _, r in grouped.iterrows():
            matricula = str(r["matricula"]).strip()
            match = mapping[mapping["description"].astype(str).str.strip().str.upper() == matricula.upper()]
            if match.empty:
                ignored += 1
                continue
            m = match.iloc[0]
            if str(m.get("active", "1")).strip() == "0":
                ignored += 1
                continue
            if normalize_text(str(m.get("fuel_type", ""))) == "anuidade" and doc_type == "fuel":
                ignored += 1
                continue

            ct = str(m.get("ct", "")).strip().upper()
            produit = str(m.get("produit", "")).strip()
            prodfourn = str(m.get("prodfourn", "")).strip()
            nombre = round(float(r["litros"]), 2)
            valor_final = float(r["valor"])

            if doc_type == "annual":
                produit = "ENTRETIEN" if matricula.upper() in annual_entretien else "ENTRETIEN_PERSO"
                prodfourn = produit
                valor_final = annual_total_incl_iva if annual_total_incl_iva > 0 else round_money(float(r["valor"]))
                ct = "BG"
            else:
                if ct == "GA":
                    valor_final = valor_final / 1.23

            rows.append({
                "confirmed": False,
                "agresso": "☐",
                "description": matricula,
                "tipo": str(m.get("fuel_type", "")).strip(),
                "produit": produit,
                "prodfourn": prodfourn,
                "unite": str(m.get("unite", "")).strip(),
                "periode": period,
                "nombre": 1.0 if doc_type == "annual" else nombre,
                "compte": str(m.get("compte", "")).strip(),
                "ana1": str(m.get("ana1", "")).strip(),
                "project": str(m.get("project", "")).strip(),
                "resno": str(m.get("resno", "")).strip(),
                "ana4": str(m.get("ana4", "")).strip(),
                "ana5": str(m.get("ana5", "")).strip(),
                "dep": str(m.get("dep", "")).strip(),
                "ct": ct,
                "mnt": round_money(valor_final),
            })
        rec.rows = rows
        if not rows:
            rec.errors.append("Nenhuma linha GALP válida após mapeamento.")
            rec.status = "Erro"
    except Exception as e:
        rec.errors.append(str(e))
        rec.status = "Erro"
    return rec


# ============================================================
# PARSERS - DELTA
# ============================================================
def read_delta_mapping() -> pd.DataFrame:
    df = read_csv_df(DELTA_MAPPING_FILE, ["material", "produto_agresso"])
    if "material" not in df.columns:
        df["material"] = ""
    if "produto_agresso" not in df.columns:
        df["produto_agresso"] = ""
    return df[["material", "produto_agresso"]].fillna("")


def get_delta_product(description: str, mapping_df: pd.DataFrame) -> str:
    norm_desc = normalize_text(description)
    best = ""
    best_len = -1
    for _, row in mapping_df.iterrows():
        material = normalize_text(str(row.get("material", "")))
        product = str(row.get("produto_agresso", "")).strip()
        if not material or not product:
            continue
        if material in norm_desc and len(material) > best_len:
            best = product
            best_len = len(material)
    return best


def parse_delta_pdf(pdf_path: Path) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    invoice_number = find_first(r"Número/Data\s+(\d+)\s*/\s*\d{2}\.\d{2}\.\d{4}", text, re.IGNORECASE)
    invoice_date = find_first(r"Número/Data\s+\d+\s*/\s*(\d{2}\.\d{2}\.\d{4})", text, re.IGNORECASE)
    period = datetime.now().strftime("%Y%m")
    if invoice_date:
        try:
            dt = datetime.strptime(invoice_date, "%d.%m.%Y")
            period = dt.strftime("%Y%m")
        except Exception:
            pass

    file_hash = file_sha256(pdf_path)
    invoice_key = f"DELTA|{invoice_number}|{period}|goods"
    rec = InvoiceRecord(
        supplier="DELTA",
        source_path=pdf_path,
        file_name=pdf_path.name,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        period=period,
        doc_type="goods",
        file_hash=file_hash,
        invoice_key=invoice_key,
    )

    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
        return rec

    mapping_df = read_delta_mapping()
    lines = []
    pattern = re.compile(
        r"(?m)^(\d+)\s+([\d,]+)\s+([A-Z]{2})\s+[\d,]+\s+EUR\s+\d+\s+[A-Z]{2}\s+[\d\.,]+\s+(6|23)%\s*$"
        r"\n(.*?)\s+([\d\.,]+)\s*$"
    )
    for m in pattern.finditer(text):
        qty = safe_float(m.group(2))
        unit = m.group(3)
        iva = int(m.group(4))
        description = " ".join(m.group(5).split())
        valor_liq = safe_float(m.group(6))
        if description.lower().startswith("montante do desco"):
            continue
        if valor_liq <= 0:
            continue
        code_iva = "BG" if iva == 23 else "BR" if iva == 6 else ""
        produit = get_delta_product(description, mapping_df)
        lines.append({
            "confirmed": False,
            "agresso": "☐",
            "description": description,
            "produit": produit,
            "periode": period,
            "nombre": qty,
            "mnt": round_money(valor_liq),
            "iva": iva,
            "code_iva": code_iva,
        })
    rec.rows = lines
    if not lines:
        rec.errors.append("Nenhuma linha DELTA detectada.")
        rec.status = "Erro"
    return rec



# ============================================================
# PARSERS - SAMSIC
# ============================================================
def parse_samsic_pdf(pdf_path: Path) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    invoice_number = find_first(r"FATURA\s+No\.\:\s*([A-Z]+\s*\d+[A-Z]?/\d+)", text, re.IGNORECASE)
    invoice_number = re.sub(r"\s+", " ", invoice_number).strip()
    invoice_date = find_first(r"(\d{2}/\d{2}/\d{4})\s*DATA\s*:", text, re.IGNORECASE)
    if not invoice_date:
        invoice_date = find_first(r"DATA\s*:\s*(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)

    period_name = find_first(r"Periodo:\s*([A-Za-zçÇéÉãõÕúÚíÍâÂêÊôÔàÀ0-9\s]+?)OBS\s+UN", text, re.IGNORECASE)
    if not period_name:
        period_name = find_first(r"OBS\s+Periodo:\s*([A-Za-zçÇéÉãõÕúÚíÍâÂêÊôÔàÀ0-9\s]+?)\s+UN", text, re.IGNORECASE)

    period = datetime.now().strftime("%Y%m")
    if period_name:
        norm = normalize_text(period_name)
        year_match = re.search(r"(20\d{2})", norm)
        month_map = {
            "janeiro": "01", "fevereiro": "02", "marco": "03", "março": "03", "abril": "04",
            "maio": "05", "junho": "06", "julho": "07", "agosto": "08", "setembro": "09",
            "outubro": "10", "novembro": "11", "dezembro": "12"
        }
        month = ""
        for k, v in month_map.items():
            if k in norm:
                month = v
                break
        if year_match and month:
            period = f"{year_match.group(1)}{month}"
        elif month:
            year = datetime.strptime(invoice_date, "%d/%m/%Y").strftime("%Y") if invoice_date else datetime.now().strftime("%Y")
            period = f"{year}{month}"
    elif invoice_date:
        try:
            period = datetime.strptime(invoice_date, "%d/%m/%Y").strftime("%Y%m")
        except Exception:
            pass

    file_hash = file_sha256(pdf_path)
    invoice_key = f"SAMSIC|{invoice_number}|{period}|services"
    rec = InvoiceRecord(
        supplier="SAMSIC",
        source_path=pdf_path,
        file_name=pdf_path.name,
        invoice_number=invoice_number,
        invoice_date=invoice_date,
        period=period,
        doc_type="services",
        file_hash=file_hash,
        invoice_key=invoice_key,
    )

    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
        return rec

    lines = []

    def _mk_line(description: str, produit: str, code: str, unit: str, qty: float, unit_price: float,
                 total: float, iva: int, cost_center: str) -> dict[str, Any]:
        code_iva = "BG" if iva == 23 else "BR" if iva == 6 else ""
        return {
            "confirmed": False,
            "agresso": "☐",
            "description": description,
            "produit": produit,
            "prodfourn": produit,
            "unite": unit or "UN",
            "periode": period,
            "f": "U",
            "nombre": qty if qty else 1,
            "prixunit": round(unit_price, 4),
            "escompte": 0,
            "mnt": round_money(total),
            "iva": iva,
            "code_iva": code_iva,
            "devise": "EUR",
            "s": "F",
            "compte": "62670100",
            "ana1": "9",
            "project": "D05015",
            "resno": "DEP5015",
            "ana4": "610005",
            "ana5": produit,
            "dep": "5015",
            "interco": "9",
            "ct": code_iva,
            "st": "",
            "t": "D",
            "source_code": code,
            "cost_center": cost_center,
            "manual_required": False,
        }

    # 1) Primeiro: faturas de acerto / atualização de preço.
    # O texto extraído real vem em ordem:
    # SERVICONTR + descrição + C.Custo + UN + Quant. + Preço Un. + IVA + Total
    flat = " ".join(text.split())
    if re.search(r"Acerto de faturação\s*-\s*Atualização de preço", flat, re.IGNORECASE):
        acerto_re = re.compile(
            r"(?P<code>SERVICONTR)\s+"
            r"(?P<desc>Acerto de faturação - .*?)\s+"
            r"(?P<cc>\d{9})\s+UN\s+"
            r"(?P<qty>\d+\.\d{3})\s+"
            r"(?P<unit>[\d ]+,\d{2})\s+"
            r"(?P<iva>\d{1,2},\d{2})%\s+"
            r"(?P<total>[\d ]+,\d{2})",
            re.IGNORECASE
        )

        for m in acerto_re.finditer(flat):
            description = " ".join(m.group("desc").split()).strip(" -")
            qty = safe_float(m.group("qty"))
            unit_price = safe_float(m.group("unit"))
            total = safe_float(m.group("total"))
            iva = int(round(safe_float(m.group("iva"))))
            cost_center = m.group("cc").strip()
            code = m.group("code").strip().upper()

            norm_desc = normalize_text(description)
            if "limpeza" in norm_desc:
                produit = "NETTOYAGE"
            elif "consum" in norm_desc or "higiene" in norm_desc or "cwc" in norm_desc:
                produit = "HYGIENE_BAT"
            else:
                produit = code or "SERVICO"

            lines.append(_mk_line(description, produit, code, "UN", qty, unit_price, total, iva, cost_center))

        if lines:
            rec.rows = lines
            return rec

    # 2) Faturação mensal normal
    block_re = re.compile(
        r"([\d ]+,\d{2})\s+([\d\.]+)(.+?)([A-Z0-9_]+)\s+(\d{1,2},\d{2})%\s+([\d ]+,\d{2})([A-Z]{2})(\d{9})",
        re.S,
    )

    for m in block_re.finditer(text):
        total_1 = safe_float(m.group(1))
        qty = safe_float(m.group(2))
        raw_desc = " ".join(m.group(3).split())
        raw_code = m.group(4).strip()
        iva = int(round(safe_float(m.group(5))))
        total_2 = safe_float(m.group(6))
        unit = m.group(7).strip()
        cost_center = m.group(8).strip()

        total = total_2 if total_2 else total_1
        if qty == 0 and total == 0:
            continue

        code = raw_code
        description = raw_desc

        if raw_code.endswith("SERVICONTR") and raw_desc:
            description = raw_desc
            code = "SERVICONTR"

        if len(description) < 4:
            continue

        norm_desc = normalize_text(description)
        if "limpeza" in norm_desc:
            produit = "NETTOYAGE"
        elif "consum" in norm_desc or "higiene" in norm_desc or "cwc" in norm_desc:
            produit = "HYGIENE_BAT"
        else:
            produit = code or "SERVICO"

        prix_unit = round(total / qty, 4) if qty else round(total, 4)
        lines.append(_mk_line(description, produit, code, unit or "UN", qty, prix_unit, total, iva, cost_center))

    rec.rows = lines
    if not lines:
        rec.errors.append("Nenhuma linha SAMSIC detectada.")
        rec.status = "Erro"
    return rec

class CsvEditorWindow(tk.Toplevel):
    def __init__(self, master, title: str, file_path: Path, columns: list[str]):
        super().__init__(master)
        self.title(title)
        self.geometry("1200x700")
        self.file_path = file_path
        self.columns = columns
        self.df = read_csv_df(file_path, columns)
        if self.df.empty:
            self.df = pd.DataFrame(columns=columns)
        for col in columns:
            if col not in self.df.columns:
                self.df[col] = ""

        self.tree = ttk.Treeview(self, columns=columns, show="headings")
        for col in columns:
            self.tree.heading(col, text=col)
            self.tree.column(col, width=150, anchor="center")
        self.tree.pack(fill="both", expand=True, padx=10, pady=10)

        btns = ttk.Frame(self)
        btns.pack(fill="x", padx=10, pady=10)
        ttk.Button(btns, text="Adicionar", command=self.add_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Editar", command=self.edit_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Remover", command=self.remove_row).pack(side="left", padx=4)
        ttk.Button(btns, text="Guardar", command=self.save).pack(side="right", padx=4)

        self.refresh()

    def refresh(self):
        for item in self.tree.get_children():
            self.tree.delete(item)
        for idx, (_, row) in enumerate(self.df.iterrows()):
            self.tree.insert("", "end", iid=str(idx), values=[row.get(c, "") for c in self.columns])

    def ask_values(self, initial: Optional[dict[str, str]] = None) -> Optional[dict[str, str]]:
        initial = initial or {}
        values: dict[str, str] = {}
        for col in self.columns:
            v = simpledialog.askstring(self.title(), col, initialvalue=initial.get(col, ""), parent=self)
            if v is None:
                return None
            values[col] = v.strip()
        return values

    def add_row(self):
        values = self.ask_values()
        if values is None:
            return
        self.df.loc[len(self.df)] = values
        self.refresh()

    def edit_row(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return
        idx = int(sel[0])
        values = self.ask_values({c: str(self.df.iloc[idx].get(c, "")) for c in self.columns})
        if values is None:
            return
        for col, val in values.items():
            self.df.at[idx, col] = val
        self.refresh()

    def remove_row(self):
        sel = self.tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return
        idx = int(sel[0])
        self.df = self.df.drop(self.df.index[idx]).reset_index(drop=True)
        self.refresh()

    def save(self):
        write_csv_df(self.file_path, self.df[self.columns])
        messagebox.showinfo("Sucesso", f"Guardado: {self.file_path.name}", parent=self)


# ============================================================
# APP
# ============================================================
class FaturasFacilitiesV12(tk.Tk):
    def __init__(self):
        super().__init__()
        self.title(f"{APP_DISPLAY_NAME} v{APP_VERSION}")
        self.geometry("1860x980")

        ensure_environment()
        self.email_to = tk.StringVar(value=EMAIL_TO_DEFAULT)
        self.email_cc = tk.StringVar(value=EMAIL_CC_DEFAULT)

        self.edp_map: dict[str, dict[str, str]] = {}
        self.epal_map: dict[str, dict[str, str]] = {}
        self.galp_admin, self.galp_admin_df = read_admin_info(GALP_ADMIN_FILE, GALP_DEFAULT_ADMIN)
        self.delta_admin, self.delta_admin_df = read_admin_info(DELTA_ADMIN_FILE, DELTA_DEFAULT_ADMIN)
        self.samsic_admin, self.samsic_admin_df = read_admin_info(SAMSIC_ADMIN_FILE, SAMSIC_DEFAULT_ADMIN)
        self.evio_admin, self.evio_admin_df = read_admin_info(EVIO_ADMIN_FILE, EVIO_DEFAULT_ADMIN)
        self.viaverde_admin, self.viaverde_admin_df = read_admin_info(VIAVERDE_ADMIN_FILE, VIAVERDE_DEFAULT_ADMIN)
        self.ayvens_admin, self.ayvens_admin_df = read_admin_info(AYVENS_ADMIN_FILE, AYVENS_DEFAULT_ADMIN)

        self.pending_edp: list[InvoiceRecord] = []
        self.pending_epal: list[InvoiceRecord] = []
        self.pending_galp: list[InvoiceRecord] = []
        self.pending_delta: list[InvoiceRecord] = []
        self.pending_samsic: list[InvoiceRecord] = []
        self.pending_evio: list[InvoiceRecord] = []
        self.pending_viaverde: list[InvoiceRecord] = []
        self.pending_ayvens: list[InvoiceRecord] = []
        self.error_records: list[InvoiceRecord] = []

        self.galp_index = 0
        self.delta_index = 0
        self.samsic_index = 0
        self.evio_index = 0
        self.viaverde_index = 0
        self.ayvens_index = 0

        self.galp_previous_ca_var = tk.StringVar()
        self.galp_ca_var = tk.StringVar()
        self.galp_period_var = tk.StringVar()
        self.delta_ca_var = tk.StringVar()
        self.delta_period_var = tk.StringVar()
        self.samsic_ca_var = tk.StringVar()
        self.samsic_period_var = tk.StringVar()
        self.evio_ca_var = tk.StringVar()
        self.evio_period_var = tk.StringVar()
        self.viaverde_previous_ca_var = tk.StringVar()
        self.viaverde_ca_var = tk.StringVar()
        self.viaverde_period_var = tk.StringVar()
        self.ayvens_ca_var = tk.StringVar()
        self.ayvens_period_var = tk.StringVar()
        self.galp_total_var = tk.StringVar(value="Resumo total: -")
        self.delta_total_var = tk.StringVar(value="Resumo total: -")
        self.samsic_total_var = tk.StringVar(value="Resumo total: -")
        self.evio_total_var = tk.StringVar(value="Resumo total: -")
        self.viaverde_total_var = tk.StringVar(value="Resumo total: -")
        self.ayvens_total_var = tk.StringVar(value="Resumo total: -")

        self._build_ui()
        self._configure_treeview_style()
        self.load_all()


    def _configure_treeview_style(self):
        style = ttk.Style(self)
        style.configure("Treeview", rowheight=24)
        style.map("Treeview", background=[("selected", "#d9d9d9")], foreground=[("selected", "black")])

    def _pack_tree(self, parent, tree):
        # IMPORTANT: do not mix grid and pack inside notebook tabs.
        # The tree is usually created with `parent` as its master, so keep all
        # widgets in the same container and manage them only with `pack`.
        yscroll = ttk.Scrollbar(parent, orient="vertical", command=tree.yview)
        xscroll = ttk.Scrollbar(parent, orient="horizontal", command=tree.xview)
        tree.configure(yscrollcommand=yscroll.set, xscrollcommand=xscroll.set)

        yscroll.pack(side="right", fill="y", padx=(0, 8), pady=(8, 0))
        xscroll.pack(side="bottom", fill="x", padx=8, pady=(0, 8))
        tree.pack(fill="both", expand=True, padx=8, pady=8)
        return parent

    def _apply_tree_tags(self, tree):
        for idx, item in enumerate(tree.get_children()):
            tag = "evenrow" if idx % 2 == 0 else "oddrow"
            tree.item(item, tags=(tag,))
        tree.tag_configure("evenrow", background="#f7f7f7")
        tree.tag_configure("oddrow", background="#ffffff")

    # ---------------- UI
    def _build_ui(self):
        top = ttk.Frame(self, padding=8)
        top.pack(fill="x")

        self.base_var = tk.StringVar(value=str(BASE_DIR))
        self.bin_var = tk.StringVar(value=str(BIN_DIR))
        self.excel_var = tk.StringVar(value=str(EXCEL_FILE))

        ttk.Label(top, text="Pasta base:").grid(row=0, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.base_var, width=100).grid(row=0, column=1, sticky="ew", padx=4)
        ttk.Button(top, text="Escolher", command=self.pick_base).grid(row=0, column=2, padx=4)

        ttk.Label(top, text="Pasta .bin:").grid(row=1, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.bin_var, width=100).grid(row=1, column=1, sticky="ew", padx=4)

        ttk.Label(top, text="Excel histórico:").grid(row=2, column=0, sticky="w")
        ttk.Entry(top, textvariable=self.excel_var, width=100).grid(row=2, column=1, sticky="ew", padx=4)

        ttk.Label(top, text="Email To:").grid(row=0, column=3, sticky="e", padx=(25, 4))
        ttk.Entry(top, textvariable=self.email_to, width=35).grid(row=0, column=4, sticky="ew")
        ttk.Label(top, text="CC:").grid(row=1, column=3, sticky="e", padx=(25, 4))
        ttk.Entry(top, textvariable=self.email_cc, width=35).grid(row=1, column=4, sticky="ew")

        btns = ttk.Frame(top)
        btns.grid(row=2, column=3, columnspan=2, sticky="e", pady=4)
        ttk.Button(btns, text="Actualizar / Ler PDFs", command=self.load_all).pack(side="left", padx=4)
        ttk.Button(btns, text="Marcar todos", command=self.mark_all_current).pack(side="left", padx=4)
        ttk.Button(btns, text="Desmarcar todos", command=self.unmark_all_current).pack(side="left", padx=4)
        ttk.Button(btns, text="Editar linha", command=self.edit_current_line).pack(side="left", padx=4)
        ttk.Button(btns, text="Informação de Gestão", command=self.open_management).pack(side="left", padx=4)
        ttk.Button(btns, text="Processar tab actual", command=self.process_current_tab).pack(side="left", padx=4)

        top.columnconfigure(1, weight=1)
        top.columnconfigure(4, weight=1)

        self.notebook = ttk.Notebook(self)
        self.notebook.pack(fill="both", expand=True, padx=8, pady=8)

        self.tab_history = ttk.Frame(self.notebook)
        self.tab_maps = ttk.Frame(self.notebook)
        self.tab_errors = ttk.Frame(self.notebook)

        self._build_history_tab()
        self._build_maps_tab()
        self._build_errors_tab()

        log_frame = ttk.LabelFrame(self, text="Log")
        log_frame.pack(fill="x", padx=8, pady=(0, 8))
        self.log_text = tk.Text(log_frame, height=4)
        self.log_text.pack(fill="both", expand=True)

    def _build_history_tab(self):
        cols = ("supplier", "invoice", "period", "ca", "processed_by", "final", "processed")
        body = ttk.Frame(self.tab_history)
        body.pack(fill="both", expand=True)
        self.tree_history = ttk.Treeview(body, columns=cols, show="headings")
        headers = {
            "supplier": "Supplier", "invoice": "Invoice", "period": "Período",
            "ca": "CA", "processed_by": "Utilizador", "final": "Ficheiro Final", "processed": "Processado em"
        }
        for c in cols:
            self.tree_history.heading(c, text=headers[c])
            self.tree_history.column(c, width=220 if c == "final" else 140, anchor="center")
        self._pack_tree(body, self.tree_history)

    def _build_maps_tab(self):
        frame = ttk.Frame(self.tab_maps, padding=10)
        frame.pack(fill="both", expand=True)
        ttk.Label(frame, text="Ficheiros de gestão activos na pasta .bin").pack(anchor="center")
        self.map_list = tk.Listbox(frame, height=10)
        self.map_list.pack(fill="both", expand=True, pady=10)
        for p in [EDP_MAP_FILE, EPAL_MAP_FILE, GALP_ADMIN_FILE, GALP_MAPPING_FILE, DELTA_ADMIN_FILE, DELTA_MAPPING_FILE, SAMSIC_ADMIN_FILE, EVIO_ADMIN_FILE, EVIO_MAPPING_FILE, VIAVERDE_ADMIN_FILE, VIAVERDE_CA_FILE, VIAVERDE_RELATION_FILE, AYVENS_ADMIN_FILE, AYVENS_TEMPLATE_FILE]:
            self.map_list.insert("end", p.name)

    def _build_errors_tab(self):
        cols = ("supplier", "file", "error")
        body = ttk.Frame(self.tab_errors)
        body.pack(fill="both", expand=True)
        self.tree_errors = ttk.Treeview(body, columns=cols, show="headings")
        for c, h, w in [("supplier", "Supplier", 120), ("file", "Ficheiro", 260), ("error", "Erro", 1000)]:
            self.tree_errors.heading(c, text=h)
            self.tree_errors.column(c, width=w, anchor="center")
        self._pack_tree(body, self.tree_errors)

    def _build_edp_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="EDP pendentes").pack(side="left")
        cols = ("agresso", "invoice", "period", "suffix", "ca", "kwh", "valor", "av", "file", "piso")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "invoice": "Fatura", "period": "Período", "suffix": "Suffix",
            "ca": "CA", "kwh": "kWh", "valor": "Valor", "av": "AV (€)", "file": "Ficheiro", "piso": "Piso"
        }
        widths = {
            "agresso": 70, "invoice": 170, "period": 90, "suffix": 80, "ca": 110,
            "kwh": 110, "valor": 90, "av": 85, "file": 240, "piso": 90
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=widths.get(c, 120), anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_pending_tree_click(e, "EDP"))
        self.tree_edp = tree
        self.tab_edp = tab

    def _build_epal_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Label(top, text="EPAL pendentes").pack(side="left")
        cols = (
            "agresso", "period", "ca", "a_faturar", "a_deduzir", "m3", "abastecimento",
            "saneamento", "residuos", "adicional", "taxas", "valor",
            "file", "invoice", "cl", "piso"
        )
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "period": "Período", "ca": "CA", "a_faturar": "A Faturar",
            "a_deduzir": "A Deduzir", "m3": "M3", "abastecimento": "Abastec.",
            "saneamento": "Saneam.", "residuos": "Resíduos", "adicional": "Adicional",
            "taxas": "Taxas", "valor": "Valor", "file": "Ficheiro",
            "invoice": "Fatura", "cl": "CL", "piso": "Piso"
        }
        widths = {
            "agresso": 70, "period": 90, "ca": 100, "a_faturar": 95, "a_deduzir": 95, "m3": 85,
            "abastecimento": 95, "saneamento": 95, "residuos": 95, "adicional": 95,
            "taxas": 85, "valor": 85, "file": 220, "invoice": 160, "cl": 90, "piso": 90
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=widths.get(c, 120), anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_pending_tree_click(e, "EPAL"))
        self.tree_epal = tree
        self.tab_epal = tab

    def _build_galp_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_galp).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_galp).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("GALP", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("GALP", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA anterior:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.galp_previous_ca_var, width=18, state="readonly").pack(side="left")
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.galp_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.galp_period_var, width=10).pack(side="left")
        self.galp_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.galp_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.galp_total_var).pack(side="left", padx=16)

        self.galp_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.galp_status).pack(anchor="center", padx=10)

        cols = ("agresso", "description", "tipo", "produit", "prodfourn", "unite", "periode",
                "nombre", "mnt", "compte", "ana1", "project", "resno", "ana4", "ana5", "dep", "ct")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "tipo": "Tipo", "produit": "Produit",
            "prodfourn": "ProdFourn", "unite": "Unité", "periode": "Période", "nombre": "Nombre",
            "compte": "Compte", "ana1": "Ana1", "project": "PROJECT", "resno": "RESNO",
            "ana4": "Ana4", "ana5": "Ana5", "dep": "DEP", "ct": "CT", "mnt": "Mnt HT en dev."
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=120, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "GALP"))
        self.tree_galp = tree
        self.tab_galp = tab

    def _build_delta_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_delta).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_delta).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("DELTA", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("DELTA", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.delta_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.delta_period_var, width=10).pack(side="left")
        self.delta_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.delta_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.delta_total_var).pack(side="left", padx=16)

        self.delta_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.delta_status).pack(anchor="center", padx=10)

        cols = ("agresso", "description", "produit", "periode", "nombre", "mnt", "iva", "code_iva")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "produit": "Produit", "periode": "Période",
            "nombre": "Nombre", "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=150 if c == "description" else 120, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "DELTA"))
        self.tree_delta = tree
        self.tab_delta = tab


    def _build_samsic_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_samsic).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_samsic).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("SAMSIC", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("SAMSIC", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.samsic_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.samsic_period_var, width=10).pack(side="left")
        self.samsic_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.samsic_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.samsic_total_var).pack(side="left", padx=16)

        self.samsic_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.samsic_status).pack(anchor="center", padx=10)

        cols = VISIBLE_FIELDS_COMMON_REDUCED
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "produit": "Produit", "prodfourn": "ProdFourn",
            "unite": "Unité", "periode": "Période", "nombre": "Nombre", "prixunit": "PrixUnit",
            "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=130 if c == "description" else 100, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "SAMSIC"))
        self.tree_samsic = tree
        self.tab_samsic = tab


    def _build_evio_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_evio).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_evio).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("EVIO", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("EVIO", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.evio_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.evio_period_var, width=10).pack(side="left")
        self.evio_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.evio_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.evio_total_var).pack(side="left", padx=16)

        self.evio_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.evio_status).pack(anchor="center", padx=10)

        cols = VISIBLE_FIELDS_COMMON_REDUCED
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "produit": "Produit", "prodfourn": "ProdFourn",
            "unite": "Unité", "periode": "Période", "nombre": "Nombre", "prixunit": "PrixUnit",
            "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA", "compte": "Compte",
            "ana1": "Ana1", "project": "PROJECT", "resno": "RESNO", "ana4": "Ana4", "ana5": "Ana5", "dep": "DEP"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=160 if c == "description" else 110, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "EVIO"))
        self.tree_evio = tree
        self.tab_evio = tab


    def _build_viaverde_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_viaverde).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_viaverde).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("VIAVERDE", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("VIAVERDE", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA anterior:").pack(side="left", padx=(20,4))
        ttk.Entry(top, textvariable=self.viaverde_previous_ca_var, width=18, state="readonly").pack(side="left")
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(12,4))
        ttk.Entry(top, textvariable=self.viaverde_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12,4))
        ttk.Entry(top, textvariable=self.viaverde_period_var, width=10).pack(side="left")
        self.viaverde_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.viaverde_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.viaverde_total_var).pack(side="left", padx=16)

        self.viaverde_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.viaverde_status).pack(anchor="center", padx=10)

        cols = VISIBLE_FIELDS_VIAVERDE
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {"agresso":"Agresso","description":"Description","produit":"Produit","prodfourn":"ProdFourn","unite":"Unité","periode":"Période","nombre":"Nombre","prixunit":"PrixUnit","mnt":"Mnt HT en dev.","compte":"Compte","ana1":"Ana1","project":"PROJECT","resno":"RESNO","ana4":"Ana4","ana5":"Ana5","dep":"DEP","interco":"INTERCO","ct":"CT","st":"ST","t":"T"}
        for c in cols:
            tree.heading(c, text=hdr.get(c,c))
            tree.column(c, width=120, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "VIAVERDE"))
        self.tree_viaverde = tree
        self.tab_viaverde = tab

    def _build_ayvens_tab(self):
        tab = ttk.Frame(self.notebook)
        top = ttk.Frame(tab, padding=8)
        top.pack(fill="x")
        ttk.Button(top, text="Anterior", command=self.prev_ayvens).pack(side="left", padx=3)
        ttk.Button(top, text="Seguinte", command=self.next_ayvens).pack(side="left", padx=3)
        ttk.Button(top, text="Marcar todos", command=lambda: self.mark_rows("AYVENS", True)).pack(side="left", padx=10)
        ttk.Button(top, text="Desmarcar todos", command=lambda: self.mark_rows("AYVENS", False)).pack(side="left", padx=3)
        ttk.Label(top, text="CA actual:").pack(side="left", padx=(20, 4))
        ttk.Entry(top, textvariable=self.ayvens_ca_var, width=18).pack(side="left")
        ttk.Label(top, text="Período:").pack(side="left", padx=(12, 4))
        ttk.Entry(top, textvariable=self.ayvens_period_var, width=10).pack(side="left")
        self.ayvens_doc_info = tk.StringVar(value="-")
        ttk.Label(top, textvariable=self.ayvens_doc_info).pack(side="left", padx=16)
        ttk.Label(top, textvariable=self.ayvens_total_var).pack(side="left", padx=16)

        self.ayvens_status = tk.StringVar(value="")
        ttk.Label(tab, textvariable=self.ayvens_status).pack(anchor="center", padx=10)

        cols = ("agresso", "description", "type", "produit", "prodfourn", "unite", "periode", "nombre", "prixunit", "mnt", "iva", "code_iva")
        body = ttk.Frame(tab)
        body.pack(fill="both", expand=True)
        tree = ttk.Treeview(body, columns=cols, show="headings")
        hdr = {
            "agresso": "Agresso", "description": "Description", "type": "Type", "produit": "Produit",
            "prodfourn": "ProdFourn", "unite": "Unité", "periode": "Période", "nombre": "Nombre",
            "prixunit": "PrixUnit", "mnt": "Mnt HT en dev.", "iva": "IVA %", "code_iva": "Code IVA"
        }
        for c in cols:
            tree.heading(c, text=hdr.get(c, c))
            tree.column(c, width=150 if c in ("description","produit","prodfourn") else 110, anchor="center")
        self._pack_tree(body, tree)
        tree.bind("<Button-1>", lambda e: self.on_row_toggle(e, "AYVENS"))
        self.tree_ayvens = tree
        self.tab_ayvens = tab

    # ---------------- management
    def open_management(self):
        menu = tk.Menu(self, tearoff=False)
        menu.add_command(label="EDP Mapping", command=lambda: CsvEditorWindow(self, "EDP Mapping", EDP_MAP_FILE, ["Suffix", "Piso", "CA"]))
        menu.add_command(label="EPAL Mapping", command=lambda: CsvEditorWindow(self, "EPAL Mapping", EPAL_MAP_FILE, ["CL", "Piso", "CA"]))
        menu.add_separator()
        menu.add_command(label="GALP Admin", command=lambda: CsvEditorWindow(self, "GALP Admin", GALP_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="GALP Vehicle Mapping", command=lambda: CsvEditorWindow(self, "GALP Vehicle Mapping", GALP_MAPPING_FILE, GALP_MAPPING_HEADER))
        menu.add_separator()
        menu.add_command(label="DELTA Admin", command=lambda: CsvEditorWindow(self, "DELTA Admin", DELTA_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="DELTA Product Mapping", command=lambda: CsvEditorWindow(self, "DELTA Product Mapping", DELTA_MAPPING_FILE, ["material", "produto_agresso"]))
        menu.add_separator()
        menu.add_command(label="SAMSIC Admin", command=lambda: CsvEditorWindow(self, "SAMSIC Admin", SAMSIC_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_separator()
        menu.add_command(label="EVIO Admin", command=lambda: CsvEditorWindow(self, "EVIO Admin", EVIO_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="EVIO Vehicle Mapping", command=lambda: CsvEditorWindow(self, "EVIO Vehicle Mapping", EVIO_MAPPING_FILE, EVIO_MAPPING_HEADER))
        menu.add_separator()
        menu.add_command(label="VIA VERDE Admin", command=lambda: CsvEditorWindow(self, "VIA VERDE Admin", VIAVERDE_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="VIA VERDE CA Mapping", command=lambda: CsvEditorWindow(self, "VIA VERDE CA Mapping", VIAVERDE_CA_FILE, ["TIPO","ANO","MES","PERIODO","CA","DESCRICAO","ACTIVE"]))
        menu.add_command(label="VIA VERDE Relation Map", command=lambda: CsvEditorWindow(self, "VIA VERDE Relation Map", VIAVERDE_RELATION_FILE, VIAVERDE_RELATION_HEADER))
        menu.add_separator()
        menu.add_command(label="AYVENS Admin", command=lambda: CsvEditorWindow(self, "AYVENS Admin", AYVENS_ADMIN_FILE, ["key", "value", "notes"]))
        menu.add_command(label="AYVENS Monthly Template", command=lambda: CsvEditorWindow(self, "AYVENS Monthly Template", AYVENS_TEMPLATE_FILE, list(read_ayvens_template().columns) if not read_ayvens_template().empty else ["Pos", "Produit", "Description", "ProdFourn", "Unité", "Période", "F", "Nombre", "PrixUnit", "Escompte", "Mnt HT en dev.", "Devise", "S", "Compte", "Ana1", "PROJECT", "RESNO", "Ana4", "Ana5", "DEP", "INTERCO", "CT", "ST", "T", "Lock"]))
        menu.add_command(label="AYVENS Relations", command=lambda: CsvEditorWindow(self, "AYVENS Relations", AYVENS_RELATION_FILE, _ayvens_relation_columns()))
        try:
            menu.tk_popup(self.winfo_pointerx(), self.winfo_pointery())
        finally:
            menu.grab_release()

    # ---------------- pickers
    def pick_base(self):
        path = filedialog.askdirectory(initialdir=str(BASE_DIR), title="Seleccionar pasta base")
        if path:
            messagebox.showinfo("Informação", "A Prod 1.0 usa a pasta do executável/script como base. Move o EXE para a pasta pretendida.", parent=self)

    # ---------------- load
    def load_all(self):
        self.edp_map = self.load_map_file(EDP_MAP_FILE, "Suffix")
        self.epal_map = self.load_map_file(EPAL_MAP_FILE, "CL")
        self.galp_admin, self.galp_admin_df = read_admin_info(GALP_ADMIN_FILE, GALP_DEFAULT_ADMIN)
        self.delta_admin, self.delta_admin_df = read_admin_info(DELTA_ADMIN_FILE, DELTA_DEFAULT_ADMIN)
        self.samsic_admin, self.samsic_admin_df = read_admin_info(SAMSIC_ADMIN_FILE, SAMSIC_DEFAULT_ADMIN)
        self.evio_admin, self.evio_admin_df = read_admin_info(EVIO_ADMIN_FILE, EVIO_DEFAULT_ADMIN)
        self.viaverde_admin, self.viaverde_admin_df = read_admin_info(VIAVERDE_ADMIN_FILE, VIAVERDE_DEFAULT_ADMIN)
        self.ayvens_admin, self.ayvens_admin_df = read_admin_info(AYVENS_ADMIN_FILE, AYVENS_DEFAULT_ADMIN)

        self.pending_edp = []
        self.pending_epal = []
        self.pending_galp = []
        self.pending_delta = []
        self.pending_samsic = []
        self.pending_evio = []
        self.pending_viaverde = []
        self.pending_ayvens = []
        self.error_records = []

        for pdf in sorted(BASE_DIR.glob("*.pdf")):
            name = pdf.name.upper()
            try:
                if name.startswith("EDP_"):
                    rec = parse_edp_pdf(pdf, self.edp_map)
                    self.route_record(rec, self.pending_edp)
                elif name.startswith("EPAL_"):
                    rec = parse_epal_pdf(pdf, self.epal_map)
                    self.route_record(rec, self.pending_epal)
                elif name.startswith("GALP_"):
                    rec = parse_galp_document(pdf)
                    self.route_record(rec, self.pending_galp)
                elif name.startswith("DELTA_"):
                    rec = parse_delta_pdf(pdf)
                    self.route_record(rec, self.pending_delta)
                elif name.startswith("SAMSIC"):
                    rec = parse_samsic_pdf(pdf)
                    self.route_record(rec, self.pending_samsic)
                elif name.startswith("EVIO"):
                    rec = parse_evio_document(pdf)
                    self.route_record(rec, self.pending_evio)
                elif name.startswith("VIAVERDE"):
                    pass
                else:
                    is_ayvens = any(k in name for k in ["AYVENS", "LEASEPLAN", "LEASE_PLAN", "LEASE PLAN"])
                    if not is_ayvens:
                        sample = normalize_text(extract_pdf_text(pdf)[:4000])
                        is_ayvens = any(k in sample for k in ["lease plan portugal", "ayvens", "lpptft003", "lpptft010"])
                    if is_ayvens:
                        rec = parse_ayvens_document(pdf, self.ayvens_admin)
                        self.route_record(rec, self.pending_ayvens)
            except Exception as e:
                rec = InvoiceRecord(supplier="UNKNOWN", source_path=pdf, file_name=pdf.name, status="Erro", errors=[str(e)])
                self.error_records.append(rec)

        for detail_pdf in sorted(BASE_DIR.glob("ViaVerde_Detalhe_*_*.pdf")):
            try:
                rec = parse_viaverde_detail_pdf_document(detail_pdf, self.viaverde_admin)
                self.route_record(rec, self.pending_viaverde)
            except Exception as e:
                rec = InvoiceRecord(supplier="VIAVERDE", source_path=detail_pdf, file_name=detail_pdf.name, status="Erro", errors=[str(e)])
                self.error_records.append(rec)

        self.pending_viaverde = combine_viaverde_records(self.pending_viaverde, self.viaverde_admin)

        self.rebuild_notebook()
        self.populate_pending_trees()
        self.populate_history_tree()
        self.populate_error_tree()
        self.log(
            f"Pendentes EDP: {len(self.pending_edp)} | Pendentes EPAL: {len(self.pending_epal)} | "
            f"GALP: {len(self.pending_galp)} | DELTA: {len(self.pending_delta)} | SAMSIC: {len(self.pending_samsic)} | EVIO: {len(self.pending_evio)} | VIA VERDE: {len(self.pending_viaverde)} | AYVENS: {len(self.pending_ayvens)} | Erros: {len(self.error_records)}"
        )
        self.log("Leitura concluída.")

    def route_record(self, rec: InvoiceRecord, bucket: list[InvoiceRecord]):
        if rec.status == "Erro":
            self.error_records.append(rec)
        elif rec.status == "Duplicado":
            self.log(f"Ignorado duplicado: {rec.file_name}")
        else:
            if rec.supplier in ("EDP", "EPAL"):
                rec.selected = get_saved_selection(rec.supplier, rec.invoice_key, default=False)
            bucket.append(rec)

    def rebuild_notebook(self):
        for tab_id in self.notebook.tabs():
            self.notebook.forget(tab_id)

        if self.pending_edp:
            self._build_edp_tab()
            self.notebook.add(self.tab_edp, text="EDP")
        if self.pending_epal:
            self._build_epal_tab()
            self.notebook.add(self.tab_epal, text="EPAL")
        if self.pending_galp:
            self._build_galp_tab()
            self.notebook.add(self.tab_galp, text="GALP")
        if self.pending_delta:
            self._build_delta_tab()
            self.notebook.add(self.tab_delta, text="DELTA")
        if self.pending_samsic:
            self._build_samsic_tab()
            self.notebook.add(self.tab_samsic, text="SAMSIC")
        if self.pending_evio:
            self._build_evio_tab()
            self.notebook.add(self.tab_evio, text="EVIO")
        if self.pending_viaverde:
            self._build_viaverde_tab()
            self.notebook.add(self.tab_viaverde, text="VIA VERDE")
        if self.pending_ayvens:
            self._build_ayvens_tab()
            self.notebook.add(self.tab_ayvens, text="AYVENS")

        self.notebook.add(self.tab_history, text="Histórico")
        self.notebook.add(self.tab_maps, text="Mapeamentos")
        self.notebook.add(self.tab_errors, text="Erros")

    def load_map_file(self, path: Path, key_field: str) -> dict[str, dict[str, str]]:
        out: dict[str, dict[str, str]] = {}
        df = read_csv_df(path)
        if df.empty:
            return out
        for _, row in df.iterrows():
            key = str(row.get(key_field, "")).strip()
            if key:
                out[key] = {"Piso": str(row.get("Piso", "")).strip(), "CA": str(row.get("CA", "")).strip()}
        return out

    # ---------------- populate
    def populate_pending_trees(self):
        if hasattr(self, "tree_edp"):
            for i in self.tree_edp.get_children():
                self.tree_edp.delete(i)
            for idx, rec in enumerate(self.pending_edp):
                self.tree_edp.insert("", "end", iid=f"EDP::{idx}", values=(
                    "☑" if rec.selected else "☐",
                    rec.invoice_number, rec.period, rec.cpe_suffix, rec.ca,
                    f"{rec.kwh:.4f}", money_str(rec.total_before_iva_23), money_str(rec.av),
                    rec.file_name, rec.piso
                ))
            self._apply_tree_tags(self.tree_edp)

        if hasattr(self, "tree_epal"):
            for i in self.tree_epal.get_children():
                self.tree_epal.delete(i)
            for idx, rec in enumerate(self.pending_epal):
                self.tree_epal.insert("", "end", iid=f"EPAL::{idx}", values=(
                    "☑" if rec.selected else "☐",
                    rec.period,
                    rec.ca,
                    f"{rec.a_faturar:.3f}",
                    f"{rec.a_deduzir:.3f}",
                    f"{rec.m3:.3f}",
                    f"{rec.abastecimento:.4f}",
                    f"{rec.saneamento:.4f}",
                    f"{rec.residuos:.4f}",
                    f"{rec.adicional:.4f}",
                    f"{rec.taxas:.4f}",
                    money_str(rec.total),
                    rec.file_name,
                    rec.invoice_number,
                    rec.cl,
                    rec.piso,
                ))
            self._apply_tree_tags(self.tree_epal)

        if hasattr(self, "tree_galp"):
            self.load_galp_current(refresh_only=True)

        if hasattr(self, "tree_delta"):
            self.load_delta_current(refresh_only=True)
        if hasattr(self, "tree_samsic"):
            self.load_samsic_current(refresh_only=True)

        if hasattr(self, "tree_evio"):
            self.load_evio_current(refresh_only=True)
        if hasattr(self, "tree_viaverde"):
            self.load_viaverde_current(refresh_only=True)
        if hasattr(self, "tree_ayvens"):
            self.load_ayvens_current(refresh_only=True)

    def populate_history_tree(self):
        for i in self.tree_history.get_children():
            self.tree_history.delete(i)
        conn = sqlite3.connect(DB_FILE)
        cur = conn.cursor()
        cur.execute("""
            SELECT supplier, invoice_number, period, ca_used, processed_by, final_filename, processed_at
            FROM processed_invoices ORDER BY id DESC
        """)
        for idx, row in enumerate(cur.fetchall()):
            self.tree_history.insert("", "end", iid=str(idx), values=row)
        conn.close()
        self._apply_tree_tags(self.tree_history)

    def populate_error_tree(self):
        for i in self.tree_errors.get_children():
            self.tree_errors.delete(i)
        all_rows = []
        for rec in self.error_records:
            for err in rec.errors:
                all_rows.append((rec.supplier, rec.file_name, err))
        for idx, row in enumerate(all_rows):
            self.tree_errors.insert("", "end", iid=str(idx), values=row)
        self._apply_tree_tags(self.tree_errors)

    # ---------------- current tab helpers
    def current_tab_name(self) -> str:
        try:
            tab_id = self.notebook.select()
            return self.notebook.tab(tab_id, "text")
        except Exception:
            return ""

    def mark_all_current(self):
        tab = self.current_tab_name()
        if tab == "EDP":
            for rec in self.pending_edp:
                rec.selected = True
                save_pending_selection(rec.supplier, rec.invoice_key, True)
        elif tab == "EPAL":
            for rec in self.pending_epal:
                rec.selected = True
                save_pending_selection(rec.supplier, rec.invoice_key, True)
        elif tab == "GALP":
            self.mark_rows("GALP", True)
            return
        elif tab == "DELTA":
            self.mark_rows("DELTA", True)
            return
        elif tab == "SAMSIC":
            self.mark_rows("SAMSIC", True)
            return
        elif tab == "EVIO":
            self.mark_rows("EVIO", True)
            return
        elif tab == "VIA VERDE":
            self.mark_rows("VIAVERDE", True)
            return
        elif tab == "AYVENS":
            self.mark_rows("AYVENS", True)
            return
        self.populate_pending_trees()

    def unmark_all_current(self):
        tab = self.current_tab_name()
        if tab == "EDP":
            for rec in self.pending_edp:
                rec.selected = False
                save_pending_selection(rec.supplier, rec.invoice_key, False)
        elif tab == "EPAL":
            for rec in self.pending_epal:
                rec.selected = False
                save_pending_selection(rec.supplier, rec.invoice_key, False)
        elif tab == "GALP":
            self.mark_rows("GALP", False)
            return
        elif tab == "DELTA":
            self.mark_rows("DELTA", False)
            return
        elif tab == "SAMSIC":
            self.mark_rows("SAMSIC", False)
            return
        elif tab == "EVIO":
            self.mark_rows("EVIO", False)
            return
        elif tab == "VIA VERDE":
            self.mark_rows("VIAVERDE", False)
            return
        elif tab == "AYVENS":
            self.mark_rows("AYVENS", False)
            return
        self.populate_pending_trees()

    def edit_current_line(self):
        tab = self.current_tab_name()
        if tab in ("EDP", "EPAL"):
            self.edit_simple_record(tab)
        elif tab == "GALP":
            self.edit_doc_row("GALP")
        elif tab == "DELTA":
            self.edit_doc_row("DELTA")
        elif tab == "SAMSIC":
            self.edit_doc_row("SAMSIC")
        elif tab == "EVIO":
            self.edit_doc_row("EVIO")
        elif tab == "VIA VERDE":
            self.edit_doc_row("VIAVERDE")
        elif tab == "AYVENS":
            self.edit_doc_row("AYVENS")
        else:
            messagebox.showinfo("Informação", "Escolhe uma tab com linhas editáveis.", parent=self)

    def edit_simple_record(self, supplier: str):
        tree = self.tree_edp if supplier == "EDP" else self.tree_epal
        records = self.pending_edp if supplier == "EDP" else self.pending_epal
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return
        idx = int(sel[0].split("::", 1)[1])
        rec = records[idx]
        new_piso = simpledialog.askstring("Piso", "Piso:", initialvalue=rec.piso, parent=self)
        if new_piso is not None:
            rec.piso = new_piso.strip()
        new_ca = simpledialog.askstring("CA", "CA:", initialvalue=rec.ca, parent=self)
        if new_ca is not None:
            rec.ca = new_ca.strip()
        new_period = simpledialog.askstring("Período", "Período (YYYYMM):", initialvalue=rec.period, parent=self)
        if new_period is not None and new_period.strip():
            rec.period = new_period.strip()
        if supplier == "EDP" and rec.cpe_suffix and rec.ca:
            rec.final_name = f"EDP_{rec.cpe_suffix}_{rec.source_path.stem}_CA{rec.ca}.pdf"
        if supplier == "EPAL" and rec.cl and rec.invoice_digits and rec.ca:
            rec.final_name = f"EPAL_{rec.cl}_{rec.invoice_digits}_CA{rec.ca}.pdf"
        self.populate_pending_trees()

    def edit_doc_row(self, supplier: str):
        if supplier == "GALP":
            doc = self.current_galp_doc()
            tree = self.tree_galp
        elif supplier == "DELTA":
            doc = self.current_delta_doc()
            tree = self.tree_delta
        elif supplier == "SAMSIC":
            doc = self.current_samsic_doc()
            tree = self.tree_samsic
        elif supplier == "EVIO":
            doc = self.current_evio_doc()
            tree = self.tree_evio
        elif supplier == "VIAVERDE":
            doc = self.current_viaverde_doc()
            tree = self.tree_viaverde
        else:
            doc = self.current_ayvens_doc()
            tree = self.tree_ayvens
        if not doc:
            return
        sel = tree.selection()
        if not sel:
            messagebox.showwarning("Aviso", "Selecciona primeiro uma linha.", parent=self)
            return
        idx = int(sel[0])
        row = doc.rows[idx]
        for field in [k for k in row.keys() if k not in ("confirmed", "agresso")]:
            new_val = simpledialog.askstring("Editar linha", field, initialvalue=str(row.get(field, "")), parent=self)
            if new_val is None:
                return
            row[field] = new_val.strip()
        if supplier == "VIAVERDE":
            desc = str(row.get("description", "")).strip()
            is_special = desc.lower().endswith("_est")
            base_plate = desc[:-4] if is_special else desc
            rel_df = read_viaverde_relations()
            rel_map = {normalize_plate(str(r["description"])): r.to_dict() for _, r in rel_df.iterrows() if str(r.get("active","1"))!="0"}
            rel = rel_map.get(normalize_plate(base_plate))
            if rel is not None:
                row["description"] = f"{base_plate}_est" if is_special else base_plate
                row["produit"] = rel.get("produit", row.get("produit", ""))
                row["prodfourn"] = rel.get("prodfourn", row.get("prodfourn", ""))
                row["unite"] = rel.get("unite", row.get("unite", "US"))
                row["compte"] = rel.get("compte", row.get("compte", ""))
                row["ana1"] = rel.get("ana1", row.get("ana1", ""))
                row["project"] = rel.get("project", row.get("project", ""))
                row["resno"] = rel.get("resno", row.get("resno", ""))
                row["ana4"] = rel.get("ana4", row.get("ana4", ""))
                row["ana5"] = row["description"]
                row["dep"] = rel.get("dep", row.get("dep", ""))
                row["interco"] = rel.get("interco", row.get("interco", ""))
                row["ct"] = rel.get("ct", row.get("ct", ""))
                row["st"] = rel.get("st", row.get("st", ""))
                row["t"] = rel.get("t", row.get("t", "D"))
                row["manual_required"] = False
        self.populate_pending_trees()

    # ---------------- tree clicks
    def on_pending_tree_click(self, event, supplier: str):
        tree = self.tree_edp if supplier == "EDP" else self.tree_epal
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item_id = tree.identify_row(event.y)
        column_id = tree.identify_column(event.x)
        allowed_columns = ("#1",)
        if not item_id or column_id not in allowed_columns:
            return
        idx = int(item_id.split("::", 1)[1])
        records = self.pending_edp if supplier == "EDP" else self.pending_epal
        rec = records[idx]
        if column_id == "#1":
            rec.selected = not rec.selected
            save_pending_selection(rec.supplier, rec.invoice_key, rec.selected)

        self.populate_pending_trees()
        return "break"

    def on_row_toggle(self, event, supplier: str):
        tree = self.tree_galp if supplier == "GALP" else self.tree_delta if supplier == "DELTA" else self.tree_samsic if supplier == "SAMSIC" else self.tree_evio if supplier == "EVIO" else self.tree_viaverde if supplier == "VIAVERDE" else self.tree_ayvens
        region = tree.identify("region", event.x, event.y)
        if region != "cell":
            return
        item = tree.identify_row(event.y)
        col = tree.identify_column(event.x)
        if not item or col != "#1":
            return
        idx = int(item)
        doc = self.current_galp_doc() if supplier == "GALP" else self.current_delta_doc() if supplier == "DELTA" else self.current_samsic_doc() if supplier == "SAMSIC" else self.current_evio_doc() if supplier == "EVIO" else self.current_viaverde_doc() if supplier == "VIAVERDE" else self.current_ayvens_doc()
        if not doc or idx >= len(doc.rows):
            return "break"
        doc.rows[idx]["confirmed"] = not doc.rows[idx]["confirmed"]
        doc.rows[idx]["agresso"] = "☑" if doc.rows[idx]["confirmed"] else "☐"
        self.populate_pending_trees()
        return "break"

    # ---------------- doc navigation
    def save_current_delta_doc_state(self):
        doc = self.current_delta_doc()
        if doc:
            doc.ca = self.delta_ca_var.get().strip()
            doc.period = self.delta_period_var.get().strip() or doc.period

    def save_current_samsic_doc_state(self):
        doc = self.current_samsic_doc()
        if doc:
            doc.ca = self.samsic_ca_var.get().strip()
            doc.period = self.samsic_period_var.get().strip() or doc.period

    def save_current_evio_doc_state(self):
        doc = self.current_evio_doc()
        if doc:
            doc.ca = self.evio_ca_var.get().strip()
            doc.period = self.evio_period_var.get().strip() or doc.period

    def save_current_galp_ca_to_doc(self):
        doc = self.current_galp_doc()
        if doc:
            doc.ca = self.galp_ca_var.get().strip()
            doc.period = self.galp_period_var.get().strip() or doc.period

    def current_galp_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_galp:
            return None
        if self.galp_index >= len(self.pending_galp):
            self.galp_index = 0
        return self.pending_galp[self.galp_index]

    def current_delta_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_delta:
            return None
        if self.delta_index >= len(self.pending_delta):
            self.delta_index = 0
        return self.pending_delta[self.delta_index]

    def prev_galp(self):
        if self.pending_galp:
            self.save_current_galp_ca_to_doc()
            self.galp_index = (self.galp_index - 1) % len(self.pending_galp)
            self.load_galp_current(refresh_only=True)

    def next_galp(self):
        if self.pending_galp:
            self.save_current_galp_ca_to_doc()
            self.galp_index = (self.galp_index + 1) % len(self.pending_galp)
            self.load_galp_current(refresh_only=True)

    def prev_delta(self):
        if self.pending_delta:
            self.save_current_delta_doc_state()
            self.delta_index = (self.delta_index - 1) % len(self.pending_delta)
            self.load_delta_current(refresh_only=True)

    def next_delta(self):
        if self.pending_delta:
            self.save_current_delta_doc_state()
            self.delta_index = (self.delta_index + 1) % len(self.pending_delta)
            self.load_delta_current(refresh_only=True)


    def current_samsic_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_samsic:
            return None
        if self.samsic_index >= len(self.pending_samsic):
            self.samsic_index = 0
        return self.pending_samsic[self.samsic_index]

    def current_evio_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_evio:
            return None
        if self.evio_index >= len(self.pending_evio):
            self.evio_index = 0
        return self.pending_evio[self.evio_index]

    def prev_samsic(self):
        if self.pending_samsic:
            self.save_current_samsic_doc_state()
            self.samsic_index = (self.samsic_index - 1) % len(self.pending_samsic)
            self.load_samsic_current(refresh_only=True)

    def next_samsic(self):
        if self.pending_samsic:
            self.save_current_samsic_doc_state()
            self.samsic_index = (self.samsic_index + 1) % len(self.pending_samsic)
            self.load_samsic_current(refresh_only=True)

    def prev_evio(self):
        if self.pending_evio:
            self.save_current_evio_doc_state()
            self.evio_index = (self.evio_index - 1) % len(self.pending_evio)
            self.load_evio_current(refresh_only=True)

    def next_evio(self):
        if self.pending_evio:
            self.save_current_evio_doc_state()
            self.evio_index = (self.evio_index + 1) % len(self.pending_evio)
            self.load_evio_current(refresh_only=True)


    def save_current_viaverde_doc_state(self):
        doc = self.current_viaverde_doc()
        if doc:
            doc.ca = self.viaverde_ca_var.get().strip()
            doc.period = self.viaverde_period_var.get().strip() or doc.period

    def current_viaverde_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_viaverde:
            return None
        if self.viaverde_index >= len(self.pending_viaverde):
            self.viaverde_index = 0
        return self.pending_viaverde[self.viaverde_index]

    def prev_viaverde(self):
        if self.pending_viaverde:
            self.save_current_viaverde_doc_state()
            self.viaverde_index = (self.viaverde_index - 1) % len(self.pending_viaverde)
            self.load_viaverde_current(refresh_only=True)

    def next_viaverde(self):
        if self.pending_viaverde:
            self.save_current_viaverde_doc_state()
            self.viaverde_index = (self.viaverde_index + 1) % len(self.pending_viaverde)
            self.load_viaverde_current(refresh_only=True)

    def save_current_ayvens_doc_state(self):
        doc = self.current_ayvens_doc()
        if doc:
            doc.ca = self.ayvens_ca_var.get().strip()
            doc.period = self.ayvens_period_var.get().strip() or doc.period

    def current_ayvens_doc(self) -> Optional[InvoiceRecord]:
        if not self.pending_ayvens:
            return None
        if self.ayvens_index >= len(self.pending_ayvens):
            self.ayvens_index = 0
        return self.pending_ayvens[self.ayvens_index]

    def prev_ayvens(self):
        if self.pending_ayvens:
            self.save_current_ayvens_doc_state()
            self.ayvens_index = (self.ayvens_index - 1) % len(self.pending_ayvens)
            self.load_ayvens_current(refresh_only=True)

    def next_ayvens(self):
        if self.pending_ayvens:
            self.save_current_ayvens_doc_state()
            self.ayvens_index = (self.ayvens_index + 1) % len(self.pending_ayvens)
            self.load_ayvens_current(refresh_only=True)

    def mark_rows(self, supplier: str, value: bool):
        doc = self.current_galp_doc() if supplier == "GALP" else self.current_delta_doc() if supplier == "DELTA" else self.current_samsic_doc() if supplier == "SAMSIC" else self.current_evio_doc() if supplier == "EVIO" else self.current_viaverde_doc() if supplier == "VIAVERDE" else self.current_ayvens_doc()
        if not doc:
            return
        for row in doc.rows:
            row["confirmed"] = value
            row["agresso"] = "☑" if value else "☐"
        self.populate_pending_trees()

    def load_galp_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_galp"):
            return
        for i in self.tree_galp.get_children():
            self.tree_galp.delete(i)
        doc = self.current_galp_doc()
        if not doc:
            self.galp_doc_info.set("-")
            self.galp_status.set("")
            self.galp_total_var.set("Resumo total: -")
            return

        if doc.doc_type == "annual":
            annual_ca = self.galp_admin.get("annual_card_ca", "")
            self.galp_previous_ca_var.set("")
            if not doc.ca:
                doc.ca = annual_ca
        elif doc.doc_type == "fuel":
            previous_ca = self.galp_admin.get("last_fuel_ca", "") or get_last_processed_ca("GALP", "fuel")
            self.galp_previous_ca_var.set(previous_ca)
            if not getattr(doc, "ca", ""):
                doc.ca = ""
        else:
            self.galp_previous_ca_var.set("")
        self.galp_ca_var.set(getattr(doc, "ca", "") or "")
        self.galp_period_var.set(getattr(doc, "period", "") or "")

        self.galp_doc_info.set(
            f"Documento {self.galp_index + 1}/{len(self.pending_galp)} | Nº: {doc.invoice_number} | "
            f"Período: {doc.period} | Tipo: {doc.doc_type}"
        )
        extra = " CA anual fixo." if doc.doc_type == "annual" else " Introduz o novo CA actual para renomear e enviar para faturas." if doc.doc_type == "fuel" else ""
        self.galp_status.set(f"Foram carregadas {len(doc.rows)} linhas sumarizadas por matrícula.{extra}")
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.galp_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        for idx, row in enumerate(doc.rows):
            self.tree_galp.insert("", "end", iid=str(idx), values=(
                row.get("agresso", "☐"), row.get("description", ""), row.get("tipo", ""),
                row.get("produit", ""), row.get("prodfourn", ""), row.get("unite", ""),
                row.get("periode", ""), row.get("nombre", ""), row.get("mnt", ""),
                row.get("compte", ""), row.get("ana1", ""), row.get("project", ""), row.get("resno", ""),
                row.get("ana4", ""), row.get("ana5", ""), row.get("dep", ""),
                row.get("ct", "")
            ))
        self._apply_tree_tags(self.tree_galp)

    def load_delta_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_delta"):
            return
        for i in self.tree_delta.get_children():
            self.tree_delta.delete(i)
        doc = self.current_delta_doc()
        if not doc:
            self.delta_doc_info.set("-")
            self.delta_status.set("")
            self.delta_total_var.set("Resumo total: -")
            return

        if not self.delta_ca_var.get().strip():
            self.delta_ca_var.set(self.delta_admin.get("last_ca", "") or get_last_processed_ca("DELTA"))

        self.delta_period_var.set(getattr(doc, "period", "") or "")
        self.delta_doc_info.set(
            f"Documento {self.delta_index + 1}/{len(self.pending_delta)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.delta_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        sem_mapping = sum(1 for r in doc.rows if not str(r.get("produit", "")).strip())
        self.delta_status.set(f"DELTA carregada com {len(doc.rows)} linhas úteis. Sem mapping: {sem_mapping}.")
        for idx, row in enumerate(doc.rows):
            self.tree_delta.insert("", "end", iid=str(idx), values=(
                row.get("agresso", "☐"), row.get("description", ""), row.get("produit", ""),
                row.get("periode", ""), row.get("nombre", ""), row.get("mnt", ""),
                row.get("iva", ""), row.get("code_iva", "")
            ))
        self._apply_tree_tags(self.tree_delta)


    def load_samsic_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_samsic"):
            return
        for i in self.tree_samsic.get_children():
            self.tree_samsic.delete(i)
        doc = self.current_samsic_doc()
        if not doc:
            self.samsic_doc_info.set("-")
            self.samsic_status.set("")
            self.samsic_total_var.set("Resumo total: -")
            return

        if not self.samsic_ca_var.get().strip():
            self.samsic_ca_var.set(self.samsic_admin.get("current_annual_ca", ""))

        self.samsic_period_var.set(getattr(doc, "period", "") or "")
        self.samsic_doc_info.set(
            f"Documento {self.samsic_index + 1}/{len(self.pending_samsic)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.samsic_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        self.samsic_status.set("Confirma os valores e, se necessário, actualiza o CA anual antes de processar.")
        for idx, row in enumerate(doc.rows):
            self.tree_samsic.insert("", "end", iid=str(idx), values=tuple(row.get(c, "☐" if c=="agresso" else "") for c in VISIBLE_FIELDS_COMMON_REDUCED))
        self._apply_tree_tags(self.tree_samsic)


    def load_evio_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_evio"):
            return
        for i in self.tree_evio.get_children():
            self.tree_evio.delete(i)
        doc = self.current_evio_doc()
        if not doc:
            self.evio_doc_info.set("-")
            self.evio_status.set("")
            self.evio_total_var.set("Resumo total: -")
            return

        if not getattr(doc, "ca", ""):
            doc.ca = self.evio_admin.get("last_ca", "") or get_last_processed_ca("EVIO")
        self.evio_ca_var.set(getattr(doc, "ca", "") or "")
        self.evio_period_var.set(getattr(doc, "period", "") or "")
        self.evio_doc_info.set(
            f"Documento {self.evio_index + 1}/{len(self.pending_evio)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.evio_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        sem_mapping = sum(1 for r in doc.rows if not str(r.get("produit", "")).strip())
        self.evio_status.set(f"EVIO carregada com {len(doc.rows)} linhas sumarizadas por matrícula. Sem mapping: {sem_mapping}.")
        for idx, row in enumerate(doc.rows):
            self.tree_evio.insert("", "end", iid=str(idx), values=tuple(row.get(c, "☐" if c=="agresso" else "") for c in VISIBLE_FIELDS_COMMON_REDUCED))
        self._apply_tree_tags(self.tree_evio)


    def load_viaverde_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_viaverde"):
            return
        for i in self.tree_viaverde.get_children():
            self.tree_viaverde.delete(i)
        doc = self.current_viaverde_doc()
        if not doc:
            self.viaverde_doc_info.set("-")
            self.viaverde_status.set("")
            self.viaverde_total_var.set("Resumo total: -")
            return

        prev_ca = self.viaverde_admin.get("last_ca", "") or get_last_processed_ca("VIAVERDE")
        self.viaverde_previous_ca_var.set(prev_ca)
        if not getattr(doc, "ca", ""):
            doc.ca = prev_ca
        self.viaverde_ca_var.set(getattr(doc, "ca", "") or "")
        self.viaverde_period_var.set(getattr(doc, "period", "") or "")
        self.viaverde_doc_info.set(
            f"Documento {self.viaverde_index + 1}/{len(self.pending_viaverde)} | Nº: {doc.invoice_number} | Período: {doc.period}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.viaverde_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        sem_mapping = sum(1 for r in doc.rows if not str(r.get("produit", "")).strip())
        manual_req = sum(1 for r in doc.rows if bool(r.get("manual_required", False)))
        ficheiros_origem = len(getattr(doc, "source_names", []) or [])
        self.viaverde_status.set(
            f"VIA VERDE mensal consolidada com {len(doc.rows)} linhas com consumo > 0, a partir de {ficheiros_origem or 1} ficheiro(s) de suporte. "
            f"Sem mapping: {sem_mapping}. Identificação manual obrigatória: {manual_req}. Serviços especiais surgem como linha separada com sufixo _est."
        )
        for idx, row in enumerate(doc.rows):
            self.tree_viaverde.insert("", "end", iid=str(idx), values=tuple(row.get(c, "☐" if c=="agresso" else "") for c in VISIBLE_FIELDS_VIAVERDE))
        self._apply_tree_tags(self.tree_viaverde)

    def load_ayvens_current(self, refresh_only: bool = False):
        if not hasattr(self, "tree_ayvens"):
            return
        for i in self.tree_ayvens.get_children():
            self.tree_ayvens.delete(i)
        doc = self.current_ayvens_doc()
        if not doc:
            self.ayvens_doc_info.set("-")
            self.ayvens_status.set("")
            self.ayvens_total_var.set("Resumo total: -")
            return

        if not getattr(doc, "ca", ""):
            doc.ca = resolve_ayvens_ca(doc.period, doc.doc_type, self.ayvens_admin)
        self.ayvens_ca_var.set(getattr(doc, "ca", "") or "")
        self.ayvens_period_var.set(getattr(doc, "period", "") or "")
        self.ayvens_doc_info.set(
            f"Documento {self.ayvens_index + 1}/{len(self.pending_ayvens)} | Nº: {doc.invoice_number} | Período: {doc.period} | Tipo: {doc.doc_type}"
        )
        total_doc = round_money(sum(float(r.get("mnt", 0) or 0) for r in doc.rows))
        self.ayvens_total_var.set(f"Resumo total: {format_amount_pt(total_doc)} €")
        self.ayvens_status.set("AYVENS carregada. Rendas: locação sem IVA, excepto AX-06-SZ e BA-21-FV; serviços sujeitos x1,23; isentos directos da fatura.")
        for idx, row in enumerate(doc.rows):
            self.tree_ayvens.insert("", "end", iid=str(idx), values=(
                row.get("agresso", "☐"), row.get("description", ""), row.get("type", ""),
                row.get("produit", ""), row.get("prodfourn", ""), row.get("unite", ""),
                row.get("periode", ""), row.get("nombre", ""), row.get("prixunit", ""),
                row.get("mnt", ""), row.get("iva", ""), row.get("code_iva", "")
            ))
        self._apply_tree_tags(self.tree_ayvens)

    # ---------------- processing

    def validate_period_value(self, period: str) -> bool:
        return bool(re.match(r"^\d{6}$", period or ""))

    def process_current_tab(self):
        tab = self.current_tab_name()
        if tab == "EDP":
            self.process_simple_supplier("EDP", self.pending_edp)
        elif tab == "EPAL":
            self.process_simple_supplier("EPAL", self.pending_epal)
        elif tab == "GALP":
            self.process_galp()
        elif tab == "DELTA":
            self.process_delta()
        elif tab == "SAMSIC":
            self.process_samsic()
        elif tab == "EVIO":
            self.process_evio()
        elif tab == "VIA VERDE":
            self.process_viaverde()
        elif tab == "AYVENS":
            self.process_ayvens()
        else:
            messagebox.showinfo("Informação", "Não há nada para processar nesta tab.", parent=self)

    def process_simple_supplier(self, supplier: str, records: list[InvoiceRecord]):
        to_process = [r for r in records if r.selected]
        if not to_process:
            messagebox.showwarning("Aviso", "Não existem linhas seleccionadas para processar.", parent=self)
            return

        processed = 0
        for rec in to_process:
            if not rec.final_name:
                messagebox.showerror("Erro", f"Nome final não gerado para {rec.file_name}", parent=self)
                return
            if not rec.ca:
                messagebox.showerror("Erro", f"CA vazio para {rec.file_name}", parent=self)
                return

        for rec in to_process:
            destination_dir = BASE_DIR / supplier / rec.period
            ensure_dir(destination_dir)
            final_path = destination_dir / rec.final_name
            if final_path.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro final: {final_path.name}", parent=self)
                return

        for rec in to_process:
            destination_dir = BASE_DIR / supplier / rec.period
            final_path = destination_dir / rec.final_name
            shutil.move(str(rec.source_path), str(final_path))
            register_processed_invoice(supplier, rec.invoice_key, rec.file_hash, rec.invoice_number, rec.period,
                                       rec.doc_type, rec.ca, rec.file_name, rec.final_name, self.current_user)

            supplier_row = {
                "InvoiceNumber": rec.invoice_number,
                "CA": rec.ca,
                "Estado": "Processado",
                "Periodo": rec.period,
                "DocType": rec.doc_type or "standard",
                "PdfFile": rec.file_name,
                "FinalFile": rec.final_name,
                "ProcessedBy": self.current_user,
                "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            }
            if supplier == "EDP":
                supplier_row = {
                    "InvoiceNumber": rec.invoice_number,
                    "CA": rec.ca,
                    "Estado": "Processado",
                    "kWh": f"{rec.kwh:.4f}",
                    "AV (€)": money_str(rec.av),
                    "Valor(€)": money_str(rec.total_before_iva_23),
                    "Periodo": rec.period,
                    "DocType": rec.doc_type or "standard",
                    "PdfFile": rec.file_name,
                    "FinalFile": rec.final_name,
                    "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            elif supplier == "EPAL":
                supplier_row = {
                    "Periodo": rec.period,
                    "CA": rec.ca,
                    "A_Faturar": f"{rec.a_faturar:.3f}",
                    "A_Deduzir": f"{rec.a_deduzir:.3f}",
                    "M3": f"{rec.m3:.3f}",
                    "Abastecimento": f"{rec.abastecimento:.4f}",
                    "Saneamento": f"{rec.saneamento:.4f}",
                    "Residuos": f"{rec.residuos:.4f}",
                    "Adicional": f"{rec.adicional:.4f}",
                    "Taxas": f"{rec.taxas:.4f}",
                    "Valor": money_str(rec.total),
                    "CL": rec.cl,
                    "Piso": rec.piso,
                    "InvoiceNumber": rec.invoice_number,
                    "DocType": rec.doc_type or "standard",
                    "Estado": "Processado",
                    "PdfFile": rec.file_name,
                    "FinalFile": rec.final_name,
                    "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
                }
            append_df_to_sheet(supplier, pd.DataFrame([supplier_row]))
            clear_pending_selection(rec.supplier, rec.invoice_key)

            append_history_row({
                "Supplier": supplier,
                "CA": rec.ca,
                "Periodo": rec.period,
                "InvoiceNumber": rec.invoice_number,
                "DocType": rec.doc_type or "standard",
                "Estado": "Processado",
                "PdfFile": rec.file_name,
                "FinalFile": rec.final_name,
                "ProcessedBy": self.current_user,
                "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
            })
            processed += 1
        messagebox.showinfo("Sucesso", f"{supplier}: {processed} fatura(s) processada(s).", parent=self)
        self.load_all()

    def process_galp(self):
        doc = self.current_galp_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas GALP para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas GALP por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_galp_ca_to_doc()
        period = self.galp_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.galp_ca_var.get().strip()
        if not ca:
            if doc.doc_type == "fuel":
                messagebox.showerror("Erro", "Indica o novo CA actual para o combustível. O CA anterior é apenas referência.", parent=self)
            else:
                messagebox.showerror("Erro", "Indica o CA actual.", parent=self)
            return
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "GALP" / period
        ensure_dir(dest_dir)

        pdf_new = dest_dir / f"GALP_{doc.invoice_number}_{period}_{ca_suffix}{doc.source_path.suffix}"
        aux_new = None
        if pdf_new.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {pdf_new.name}.", parent=self)
            return

        if doc.aux_path:
            aux_new = dest_dir / f"GALP_{doc.invoice_number}_{period}_{ca_suffix}{doc.aux_path.suffix}"
            if aux_new.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro {aux_new.name}.", parent=self)
                return

        shutil.move(str(doc.source_path), str(pdf_new))
        if doc.aux_path and doc.aux_path.exists():
            shutil.move(str(doc.aux_path), str(aux_new))

        if doc.doc_type == "fuel":
            self.set_admin_value(GALP_ADMIN_FILE, GALP_DEFAULT_ADMIN, "last_fuel_ca", ca)

        register_processed_invoice("GALP", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, pdf_new.name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "GALP", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalLitros": round(sum(float(r.get("nombre", 0)) for r in doc.rows), 2),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "AuxFile": doc.aux_path.name if doc.aux_path else "",
            "FinalPdf": pdf_new.name, "FinalAux": aux_new.name if aux_new else "",
            "ProcessedBy": self.current_user,
            "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Tipo": r.get("tipo", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "Compte": r.get("compte", ""),
            "Ana1": r.get("ana1", ""),
            "PROJECT": r.get("project", ""),
            "RESNO": r.get("resno", ""),
            "Ana4": r.get("ana4", ""),
            "Ana5": r.get("ana5", ""),
            "DEP": r.get("dep", ""),
            "CT": r.get("ct", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
        } for r in doc.rows])
        append_df_to_sheet("GALP", summary_df)
        append_df_to_sheet("GALP_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"GALP - Fatura {doc.invoice_number} - {doc.period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.galp_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[pdf_new]
        )
        messagebox.showinfo("Sucesso", f"Fatura GALP processada.\n{msg}", parent=self)
        self.load_all()


    def process_delta(self):
        doc = self.current_delta_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas DELTA para processar.", parent=self)
            return

        # Sincroniza primeiro o estado visível do ecrã para doc.rows
        self.save_current_delta_doc_state()

        rows_to_process = [r for r in doc.rows if bool(r.get("confirmed", False))]
        if not rows_to_process:
            messagebox.showerror("Erro", "Marca pelo menos uma linha DELTA antes de processar.", parent=self)
            return

        def _is_missing(v: object) -> bool:
            s = str(v).strip()
            return s in ("", "None", "nan", "NaN")

        unresolved = [
            r.get("description", "") or "(sem descrição)"
            for r in rows_to_process
            if bool(r.get("manual_required", False))
            or _is_missing(r.get("description", ""))
            or _is_missing(r.get("produit", ""))
            or _is_missing(r.get("periode", ""))
            or _is_missing(r.get("nombre", ""))
            or _is_missing(r.get("mnt", ""))
            or _is_missing(r.get("iva", ""))
            or _is_missing(r.get("code_iva", ""))
        ]
        if unresolved:
            messagebox.showerror(
                "Erro",
                "Existem linhas DELTA por completar ou rever antes de processar:\n- "
                + "\n- ".join(map(str, unresolved[:10])),
                parent=self,
            )
            return

        period = self.delta_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in rows_to_process:
            row["periode"] = period

        ca = self.delta_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual para a DELTA.", parent=self)
            return

        self.set_admin_value(DELTA_ADMIN_FILE, DELTA_DEFAULT_ADMIN, "last_ca", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "DELTA" / period
        ensure_dir(dest_dir)

        final_name = f"DELTA_{doc.source_path.stem}_{ca_suffix}{doc.source_path.suffix}"
        final_path = dest_dir / final_name
        if final_path.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
            return

        shutil.move(str(doc.source_path), str(final_path))
        register_processed_invoice(
            "DELTA",
            doc.invoice_key,
            doc.file_hash,
            doc.invoice_number,
            period,
            doc.doc_type,
            ca,
            doc.file_name,
            final_name,
            self.current_user,
        )

        summary_df = pd.DataFrame([{
            "Supplier": "DELTA", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(rows_to_process),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in rows_to_process)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "FinalPdf": final_name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Produit": r.get("produit", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
        } for r in rows_to_process])
        append_df_to_sheet("DELTA", summary_df)
        append_df_to_sheet("DELTA_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"DELTA - Fatura {doc.invoice_number} - {doc.period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.delta_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_path]
        )
        messagebox.showinfo("Sucesso", f"Fatura DELTA processada.\n{msg}", parent=self)
        self.load_all()


    def process_evio(self):
        doc = self.current_evio_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas EVIO para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas EVIO por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_evio_doc_state()
        period = self.evio_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.evio_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da EVIO.", parent=self)
            return

        self.set_admin_value(EVIO_ADMIN_FILE, EVIO_DEFAULT_ADMIN, "last_ca", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "EVIO" / period
        ensure_dir(dest_dir)

        short_inv = short_evio_invoice_number(doc.invoice_number)
        final_pdf_name = f"EVIO_{short_inv}_{ca_suffix}{doc.source_path.suffix}"
        final_pdf = dest_dir / final_pdf_name
        if final_pdf.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_pdf_name}.", parent=self)
            return

        final_aux = None
        if doc.aux_path and doc.aux_path.exists():
            final_aux = dest_dir / f"EVIO_{short_inv}_{ca_suffix}{doc.aux_path.suffix}"
            if final_aux.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro {final_aux.name}.", parent=self)
                return

        shutil.move(str(doc.source_path), str(final_pdf))
        if doc.aux_path and doc.aux_path.exists():
            shutil.move(str(doc.aux_path), str(final_aux))

        register_processed_invoice("EVIO", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, final_pdf_name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "EVIO", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "AuxFile": doc.aux_path.name if doc.aux_path else "",
            "FinalPdf": final_pdf_name, "FinalAux": final_aux.name if final_aux else "",
            "ProcessedBy": self.current_user,
            "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "PrixUnit": r.get("prixunit", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
            "Compte": r.get("compte", ""),
            "Ana1": r.get("ana1", ""),
            "PROJECT": r.get("project", ""),
            "RESNO": r.get("resno", ""),
            "Ana4": r.get("ana4", ""),
            "Ana5": r.get("ana5", ""),
            "DEP": r.get("dep", ""),
            "INTERCO": r.get("interco", ""),
            "CT": r.get("ct", ""),
            "ST": r.get("st", ""),
            "T": r.get("t", ""),
        } for r in doc.rows])
        append_df_to_sheet("EVIO", summary_df)
        append_df_to_sheet("EVIO_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"EVIO - Fatura {doc.invoice_number} - {period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.evio_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_pdf]
        )
        messagebox.showinfo("Sucesso", f"Fatura EVIO processada.\n{msg}", parent=self)
        self.load_all()


    def process_viaverde(self):
        doc = self.current_viaverde_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas VIA VERDE para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas VIA VERDE por completar ou identificar manualmente antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_viaverde_doc_state()
        period = self.viaverde_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.viaverde_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da VIA VERDE.", parent=self)
            return

        self.set_admin_value(VIAVERDE_ADMIN_FILE, VIAVERDE_DEFAULT_ADMIN, "last_ca", ca)
        df_ca = read_viaverde_ca_mapping()
        mask = (df_ca.get("TIPO", pd.Series(dtype=str)).astype(str).str.upper() == "STANDARD") & (df_ca.get("PERIODO", pd.Series(dtype=str)).astype(str) == period)
        if df_ca.empty:
            df_ca = pd.DataFrame(columns=["TIPO","ANO","MES","PERIODO","CA","DESCRICAO","ACTIVE"])
        if mask.any():
            df_ca.loc[mask, "CA"] = ca
        else:
            df_ca.loc[len(df_ca)] = ["STANDARD", period[:4], period[4:6], period, ca, f"VIAVERDE - {period}", "1"]
        write_csv_df(VIAVERDE_CA_FILE, df_ca)

        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "VIA VERDE" / period
        ensure_dir(dest_dir)

        final_pdfs = []
        all_files = getattr(doc, "all_files", [doc.source_path])
        pdfs = getattr(doc, "pdf_files", [])
        for p in all_files:
            final_name = f"ViaVerde_{p.stem}_{ca_suffix}{p.suffix}" if p.suffix.lower()==".pdf" else p.name
            final_path = dest_dir / final_name
            if final_path.exists():
                messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
                return
        for p in all_files:
            final_name = f"ViaVerde_{p.stem}_{ca_suffix}{p.suffix}" if p.suffix.lower()==".pdf" else p.name
            final_path = dest_dir / final_name
            shutil.move(str(p), str(final_path))
            if p.suffix.lower()==".pdf":
                final_pdfs.append(final_path)
        register_processed_invoice("VIAVERDE", doc.invoice_key, doc.file_hash, doc.invoice_number, period, doc.doc_type, ca, doc.file_name, ", ".join(fp.name for fp in final_pdfs), self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "VIAVERDE", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": "standard", "RowsCount": len(doc.rows), "TotalValorHT": round_money(sum(float(r.get("mnt",0) or 0) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": ", ".join(p.name for p in pdfs), "FinalPdf": ", ".join(p.name for p in final_pdfs),
            "FinalXml": doc.source_path.name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        }])
        append_df_to_sheet("VIAVERDE", summary_df)

        output_rows = []
        for row in doc.rows:
            output_rows.append({
                "Description": row.get("description",""), "Produit": row.get("produit",""), "ProdFourn": row.get("prodfourn",""),
                "Unité": row.get("unite",""), "Période": row.get("periode",""), "Nombre": row.get("nombre",""),
                "PrixUnit": row.get("prixunit",""), "Mnt HT en dev.": row.get("mnt",""), "Compte": row.get("compte",""),
                "Ana1": row.get("ana1",""), "PROJECT": row.get("project",""), "RESNO": row.get("resno",""), "Ana4": row.get("ana4",""),
                "Ana5": row.get("ana5",""), "DEP": row.get("dep",""), "INTERCO": row.get("interco",""), "CT": row.get("ct",""),
                "ST": row.get("st",""), "T": row.get("t","")
            })
        append_df_to_sheet("VIAVERDE_OUTPUT", pd.DataFrame(output_rows))
        append_history_row({
            "Supplier":"VIAVERDE","CA":ca,"Periodo":period,"InvoiceNumber":doc.invoice_number,"DocType":"standard",
            "Estado":"Processado","PdfFile":", ".join(p.name for p in pdfs),"FinalFile":", ".join(p.name for p in final_pdfs),
            "ProcessedBy":self.current_user,"ProcessedAt":datetime.now().strftime("%Y-%m-%d %H:%M:%S")
        })
        ok, msg = create_outlook_draft(
            subject=f"VIA VERDE - {period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.viaverde_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=final_pdfs
        )
        messagebox.showinfo("Sucesso", f"Via Verde processada.\n{msg}", parent=self)
        self.load_all()

    def set_admin_value(self, path: Path, defaults: list[dict[str, str]], key: str, value: str):
        info, df = read_admin_info(path, defaults)
        mask = df["key"].astype(str).str.strip() == key
        if mask.any():
            df.loc[mask, "value"] = value
        else:
            df.loc[len(df)] = [key, value, ""]
        write_admin_info(path, df)


    def process_ayvens(self):
        doc = self.current_ayvens_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas AYVENS para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas AYVENS por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_ayvens_doc_state()
        period = self.ayvens_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.ayvens_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da AYVENS.", parent=self)
            return

        self.set_admin_value(AYVENS_ADMIN_FILE, AYVENS_DEFAULT_ADMIN, f"{'extra_ca_' + period[:4] if doc.doc_type == 'extra' else 'rent_ca_' + period}", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        subfolder = "EXTRAS" if doc.doc_type == "extra" else "RENDAS"
        dest_dir = BASE_DIR / "AYVENS" / period / subfolder
        ensure_dir(dest_dir)

        short_inv = re.sub(r"[^A-Za-z0-9]+", "", doc.invoice_number or "") or "SEMNUMERO"
        final_name = f"AYVENS_{doc.doc_type.upper()}_{short_inv}_{ca_suffix}{doc.source_path.suffix}"
        final_path = dest_dir / final_name
        if final_path.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
            return

        shutil.move(str(doc.source_path), str(final_path))
        register_processed_invoice("AYVENS", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, final_name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "AYVENS", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "FinalPdf": final_name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Type": r.get("type", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "PrixUnit": r.get("prixunit", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
            "Compte": r.get("compte", ""),
            "Ana1": r.get("ana1", ""),
            "PROJECT": r.get("project", ""),
            "RESNO": r.get("resno", ""),
            "Ana4": r.get("ana4", ""),
            "Ana5": r.get("ana5", ""),
            "DEP": r.get("dep", ""),
            "INTERCO": r.get("interco", ""),
            "CT": r.get("ct", ""),
            "ST": r.get("st", ""),
            "T": r.get("t", ""),
        } for r in doc.rows])
        append_df_to_sheet("AYVENS", summary_df)
        append_df_to_sheet("AYVENS_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"AYVENS - Fatura {doc.invoice_number} - {period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.ayvens_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_path]
        )
        messagebox.showinfo("Sucesso", f"Fatura AYVENS processada.\n{msg}", parent=self)
        self.load_all()

    def process_samsic(self):
        doc = self.current_samsic_doc()
        if not doc:
            return
        if not doc.rows:
            messagebox.showwarning("Aviso", "Não existem linhas SAMSIC para processar.", parent=self)
            return
        if not all(r.get("confirmed", False) for r in doc.rows):
            messagebox.showerror("Erro", "Marca todas as linhas do Agresso antes de processar.", parent=self)
            return
        unresolved = [r.get("description", "") for r in doc.rows if bool(r.get("manual_required", False)) or not str(r.get("produit", "")).strip() or not str(r.get("compte", "")).strip()]
        if unresolved:
            messagebox.showerror("Erro", "Existem linhas SAMSIC por completar ou rever antes de processar:\n- " + "\n- ".join(map(str, unresolved[:10])), parent=self)
            return

        self.save_current_samsic_doc_state()
        period = self.samsic_period_var.get().strip()
        if not self.validate_period_value(period):
            messagebox.showerror("Erro", "Período inválido. Usa YYYYMM, por exemplo 202603.", parent=self)
            return
        doc.period = period
        for row in doc.rows:
            row["periode"] = period

        ca = self.samsic_ca_var.get().strip()
        if not ca:
            messagebox.showerror("Erro", "Indica o CA actual da SAMSIC.", parent=self)
            return

        self.set_admin_value(SAMSIC_ADMIN_FILE, SAMSIC_DEFAULT_ADMIN, "current_annual_ca", ca)
        ca_suffix = ca.upper() if ca.upper().startswith("CA") else f"CA{ca}"
        dest_dir = BASE_DIR / "SAMSIC" / period
        ensure_dir(dest_dir)

        short_inv = short_samsic_invoice_number(doc.invoice_number)
        final_name = f"SAMSIC_{short_inv}_{ca_suffix}{doc.source_path.suffix}"
        final_path = dest_dir / final_name
        if final_path.exists():
            messagebox.showerror("Erro", f"Já existe o ficheiro {final_name}.", parent=self)
            return

        shutil.move(str(doc.source_path), str(final_path))
        register_processed_invoice("SAMSIC", doc.invoice_key, doc.file_hash, doc.invoice_number, period,
                                   doc.doc_type, ca, doc.file_name, final_name, self.current_user)

        summary_df = pd.DataFrame([{
            "Supplier": "SAMSIC", "CA": ca, "Periodo": period, "InvoiceNumber": doc.invoice_number,
            "DocType": doc.doc_type, "RowsCount": len(doc.rows),
            "TotalValorHT": round_money(sum(float(r.get("mnt", 0)) for r in doc.rows)),
            "Estado": "Processado", "PdfFile": doc.file_name,
            "FinalPdf": final_name, "ProcessedBy": self.current_user, "ProcessedAt": datetime.now().strftime("%Y-%m-%d %H:%M:%S"),
        }])
        output_df = pd.DataFrame([{
            "Description": r.get("description", ""),
            "Produit": r.get("produit", ""),
            "ProdFourn": r.get("prodfourn", ""),
            "Unité": r.get("unite", ""),
            "Période": r.get("periode", ""),
            "Nombre": r.get("nombre", ""),
            "PrixUnit": r.get("prixunit", ""),
            "Mnt HT en dev.": r.get("mnt", ""),
            "IVA %": r.get("iva", ""),
            "Code IVA": r.get("code_iva", ""),
        } for r in doc.rows])
        append_df_to_sheet("SAMSIC", summary_df)
        append_df_to_sheet("SAMSIC_OUTPUT", output_df)

        ok, msg = create_outlook_draft(
            subject=f"SAMSIC - Fatura {doc.invoice_number} - {doc.period} - {ca_suffix}",
            body=build_standard_email_body(),
            to_addr=self.email_to.get().strip() or self.samsic_admin.get("email_to", ""),
            cc_addr=self.email_cc.get().strip(),
            attachments=[final_path]
        )
        messagebox.showinfo("Sucesso", f"Fatura SAMSIC processada.\n{msg}", parent=self)
        self.load_all()

    # ---------------- logging
    def log(self, msg: str):
        stamp = datetime.now().strftime("%H:%M:%S")
        self.log_text.insert("end", f"[{stamp}] {msg}\n")
        self.log_text.see("end")




# --- AYVENS v1.4.1 overrides -------------------------------------------------
def _ayvens_ca_columns() -> list[str]:
    return ["TIPO", "ANO", "MES", "PERIODO", "CA", "DESCRICAO", "ACTIVE"]


def _normalize_ayvens_ca_df(df: pd.DataFrame) -> pd.DataFrame:
    cols = _ayvens_ca_columns()
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    df = df.rename(columns={c: c.upper() for c in df.columns}).fillna("")
    for col in cols:
        if col not in df.columns:
            df[col] = ""
    df = df[cols].copy()
    for col in cols:
        df[col] = df[col].astype(str).fillna("").str.strip()
    df["TIPO"] = df["TIPO"].str.upper()
    df["ACTIVE"] = df["ACTIVE"].replace("", "1")
    return df


def read_ayvens_ca_mapping() -> pd.DataFrame:
    bootstrap_ayvens_from_excel()
    return _normalize_ayvens_ca_df(read_csv_df(AYVENS_CA_FILE, _ayvens_ca_columns()))


def _build_ayvens_ca_mapping_from_admin_xlsx(admin_xlsx: Optional[Path]) -> pd.DataFrame:
    rows = []
    if not admin_xlsx or not admin_xlsx.exists():
        return pd.DataFrame(columns=_ayvens_ca_columns())
    try:
        ca_lista = pd.read_excel(admin_xlsx, sheet_name="CA_lista", dtype=str).fillna("")
    except Exception:
        return pd.DataFrame(columns=_ayvens_ca_columns())
    for _, row in ca_lista.iterrows():
        ca = str(row.get("No.commande", "")).strip()
        ref_ext = str(row.get("Ref. Ext", "")).strip()
        norm = normalize_text(ref_ext)
        if not ca or not ref_ext:
            continue
        year_match = re.search(r"(20\d{2})", ref_ext)
        year = year_match.group(1) if year_match else ""
        if "extras" in norm and year:
            rows.append({"TIPO": "EXTRA", "ANO": year, "MES": "", "PERIODO": "", "CA": ca, "DESCRICAO": ref_ext, "ACTIVE": "1"})
            continue
        month_num = ""
        month_abbr = ""
        for abbr, num in MONTHS_PT.items():
            if normalize_text(abbr) in norm:
                month_abbr = abbr
                month_num = num
                break
        if year and month_num:
            rows.append({"TIPO": "RENT", "ANO": year, "MES": month_num, "PERIODO": f"{year}{month_num}", "CA": ca, "DESCRICAO": ref_ext, "ACTIVE": "1"})
    return _normalize_ayvens_ca_df(pd.DataFrame(rows, columns=_ayvens_ca_columns()))


def _build_ayvens_relations_from_agresso_df(df: pd.DataFrame) -> pd.DataFrame:
    cols = _ayvens_relation_columns()
    if df is None or df.empty:
        return pd.DataFrame(columns=cols)
    rename_map = {
        "Ayvens_MATRÍCULA": "matricula", "Ayvens_MATRICULA": "matricula",
        "Ayvens_Produto": "ayvens_produto", "Ayvens_PRODUTO": "ayvens_produto",
        "Description": "description", "Produit": "produit", "ProdFourn": "prodfourn",
        "Unité": "unite", "Compte": "compte", "Ana1": "ana1", "PROJECT": "project",
        "RESNO": "resno", "Ana4": "ana4", "Ana5": "ana5", "DEP": "dep",
        "INTERCO": "interco", "CT": "ct", "ST": "st", "T": "t",
    }
    work = df.rename(columns=rename_map).fillna("")
    if "matricula" not in work.columns or "ayvens_produto" not in work.columns:
        return pd.DataFrame(columns=cols)
    out = []
    for _, row in work.iterrows():
        rec = {str(k): str(v) for k, v in row.to_dict().items()}
        plate_raw = rec.get("matricula", "").strip()
        plate = normalize_plate(plate_raw)
        produto_ayv = rec.get("ayvens_produto", "").strip()
        if not plate or not produto_ayv:
            continue
        normp = normalize_text(produto_ayv)
        if "locacao" in normp:
            line_type = "RENT"
        elif "sujeito" in normp:
            line_type = "SERVICE_VAT"
        elif "isento" in normp:
            line_type = "SERVICE_EXEMPT"
        else:
            continue
        out.append({
            "matricula": plate,
            "line_type": line_type,
            "description": plate_raw or rec.get("description", ""),
            "produit": rec.get("produit", ""),
            "prodfourn": rec.get("prodfourn", ""),
            "unite": rec.get("unite", "US"),
            "compte": rec.get("compte", ""),
            "ana1": rec.get("ana1", ""),
            "project": rec.get("project", ""),
            "resno": rec.get("resno", ""),
            "ana4": rec.get("ana4", ""),
            "ana5": rec.get("ana5", plate_raw),
            "dep": rec.get("dep", ""),
            "interco": rec.get("interco", ""),
            "ct": rec.get("ct", ""),
            "st": rec.get("st", ""),
            "t": rec.get("t", "D"),
            "active": "1",
        })
    # generic extra fallback
    out.append({
        "matricula": "EXTRA_DEFAULT", "line_type": "EXTRA", "description": "Serviço não contratado",
        "produit": "PORTAGENS", "prodfourn": "PORTAGENS", "unite": "US",
        "compte": "62260300", "ana1": "9", "project": "DO5015", "resno": "DEP5015",
        "ana4": "616220", "ana5": "AYVENS_EXTRA", "dep": "5015", "interco": "9",
        "ct": "BR", "st": "", "t": "D", "active": "1"
    })
    return _normalize_ayvens_relation_df(pd.DataFrame(out, columns=cols))


def _load_ayvens_relation_sources() -> pd.DataFrame:
    frames: list[pd.DataFrame] = []
    admin_xlsx = _ayvens_candidate_file("Ayvens_admin.xlsx")
    example_xlsx = _ayvens_candidate_file("ayvens_exemplo.xlsx")
    if admin_xlsx:
        try:
            frames.append(_build_ayvens_relations_from_agresso_df(pd.read_excel(admin_xlsx, sheet_name="CA_mensal", dtype=str)))
        except Exception:
            pass
    if example_xlsx:
        try:
            frames.append(_build_ayvens_relations_from_agresso_df(pd.read_excel(example_xlsx, sheet_name="agresso_ca", dtype=str)))
        except Exception:
            pass
    if AYVENS_RELATION_FILE.exists():
        try:
            frames.append(_normalize_ayvens_relation_df(read_csv_df(AYVENS_RELATION_FILE)))
        except Exception:
            pass
    if not frames:
        return pd.DataFrame(columns=_ayvens_relation_columns())
    merged = pd.concat(frames, ignore_index=True).fillna("")
    merged["__matricula"] = merged["matricula"].apply(normalize_plate)
    merged["__line_type"] = merged["line_type"].astype(str).str.upper().str.strip()
    merged["__priority"] = range(len(merged))
    merged = merged.sort_values(["__matricula", "__line_type", "__priority"]).drop_duplicates(["__matricula", "__line_type"], keep="last")
    merged = merged.drop(columns=["__matricula", "__line_type", "__priority"])
    return _normalize_ayvens_relation_df(merged)


def bootstrap_ayvens_from_excel():
    ensure_dir(BIN_DIR)
    if not AYVENS_ADMIN_FILE.exists():
        pd.DataFrame(AYVENS_DEFAULT_ADMIN).to_csv(AYVENS_ADMIN_FILE, index=False, encoding="utf-8-sig")
    admin_xlsx = _ayvens_candidate_file("Ayvens_admin.xlsx")
    example_xlsx = _ayvens_candidate_file("ayvens_exemplo.xlsx")
    ca_df = _build_ayvens_ca_mapping_from_admin_xlsx(admin_xlsx)
    if not ca_df.empty:
        write_csv_df(AYVENS_CA_FILE, ca_df)
    elif not AYVENS_CA_FILE.exists():
        pd.DataFrame(columns=_ayvens_ca_columns()).to_csv(AYVENS_CA_FILE, index=False, encoding="utf-8-sig")
    rel_df = _load_ayvens_relation_sources()
    if not rel_df.empty:
        write_csv_df(AYVENS_RELATION_FILE, rel_df)
    elif not AYVENS_RELATION_FILE.exists():
        pd.DataFrame(columns=_ayvens_relation_columns()).to_csv(AYVENS_RELATION_FILE, index=False, encoding="utf-8-sig")


def resolve_ayvens_ca(period: str, doc_type: str, admin: dict[str, str]) -> str:
    df = read_ayvens_ca_mapping()
    if not df.empty:
        if doc_type == "extra":
            year = str(period)[:4]
            m = df[(df["TIPO"] == "EXTRA") & (df["ANO"] == year) & (df["ACTIVE"] != "0")]
            if not m.empty:
                return str(m.iloc[0]["CA"]).strip()
        else:
            m = df[(df["TIPO"] == "RENT") & (df["PERIODO"] == str(period)) & (df["ACTIVE"] != "0")]
            if not m.empty:
                return str(m.iloc[0]["CA"]).strip()
    if doc_type == "extra":
        return admin.get(f"extra_ca_{str(period)[:4]}", "") or admin.get("extra_ca", "")
    return admin.get(f"rent_ca_{period}", "") or admin.get("rent_ca_default", "")


def _period_from_billing_range(billing_range: str, invoice_date: str) -> str:
    m = re.search(r"/(\d{2})/(20\d{2})$", str(billing_range or "").strip())
    if m:
        return f"{m.group(2)}{m.group(1)}"
    return parse_ayvens_period(invoice_date)


def parse_ayvens_document(pdf_path: Path, ayvens_admin: dict[str, str]) -> InvoiceRecord:
    text = extract_pdf_text(pdf_path)
    invoice_number = re.sub(r"\s+", " ", find_first(r"(FT\s*\d+/\d+)", text, re.IGNORECASE)).strip()
    invoice_date = find_first(r"(20\d{2}/\d{2}/\d{2})", text)
    norm_text = normalize_text(text)
    doc_type = "extra" if "servicos nao contratados" in norm_text else "rent"
    parsed_lines_rent = parse_ayvens_rent_lines(text) if doc_type == "rent" else []
    period = _period_from_billing_range(parsed_lines_rent[0]["billing_range"], invoice_date) if parsed_lines_rent else parse_ayvens_period(invoice_date)
    file_hash = file_sha256(pdf_path)
    invoice_key = f"AYVENS|{invoice_number}|{period}|{doc_type}"
    rec = InvoiceRecord(supplier="AYVENS", source_path=pdf_path, file_name=pdf_path.name, invoice_number=invoice_number, invoice_date=invoice_date, period=period, doc_type=doc_type, file_hash=file_hash, invoice_key=invoice_key)
    if invoice_already_processed(rec.invoice_key, rec.file_hash):
        rec.status = "Duplicado"
        return rec
    rec.ca = resolve_ayvens_ca(period, doc_type, ayvens_admin)
    template_index, generic_index = build_ayvens_template_index()
    galp_defaults = read_galp_plate_defaults()
    rows = []
    missing = []
    if doc_type == "rent":
        if not parsed_lines_rent:
            rec.errors.append("Nenhuma linha de renda AYVENS detectada. Verifica o layout extraído do PDF.")
            rec.status = "Erro"
            return rec
        for item in parsed_lines_rent:
            plate = item["plate"]
            plate_norm = normalize_plate(plate)
            amt_map = {
                "RENT": item["locacao"] * 1.23 if plate_norm in AYVENS_SPECIAL_FULL_VAT_RENT_PLATES else item["locacao"],
                "SERVICE_VAT": item["exploracao_sujeita"] if plate_norm == "AA75AJ" else (item["exploracao_sujeita"] * 1.23 if item["exploracao_sujeita"] else 0.0),
                "SERVICE_EXEMPT": item["isento"],
            }
            for row_type, amount in amt_map.items():
                if amount <= 0:
                    continue
                tpl = template_index.get((plate_norm, row_type)) or build_ayvens_fallback_template(plate, row_type, generic_index, galp_defaults)
                if tpl:
                    rows.append(build_ayvens_row_from_template(tpl, period, amount, row_type))
                else:
                    missing.append(f"Linha {row_type} sem template para {plate}")
    else:
        parsed_lines = parse_ayvens_extra_lines(text)
        if not parsed_lines:
            rec.errors.append("Nenhuma linha de extra AYVENS detectada. Verifica o layout extraído do PDF.")
            rec.status = "Erro"
            return rec
        for item in parsed_lines:
            tpl = template_index.get(("EXTRA_DEFAULT", "EXTRA")) or generic_index.get("EXTRA") or build_ayvens_fallback_template(item["plate"], "EXTRA", generic_index, galp_defaults)
            if tpl:
                amount = item["total"] if item["total"] else item["net"]
                row = build_ayvens_row_from_template(tpl, period, amount, "EXTRA")
                row["description"] = item["description"]
                row["iva"] = str(item.get("iva_pct", 0))
                row["ct"] = resolve_ayvens_ct(tpl, "EXTRA", item["description"])
                row["code_iva"] = row["ct"]
                rows.append(row)
            else:
                missing.append(f"Linha EXTRA sem template para {item['plate']}")
    rec.rows = rows
    if missing:
        rec.errors.extend(missing)
        rec.status = "Erro"
    if not rec.ca:
        # permissivo na leitura; o processamento pedirá/confirmará o CA
        rec.status = rec.status if rec.status == "Erro" else "Pendente"
    elif not rec.errors:
        rec.status = "Pendente"
    return rec

def build_main_app():
    return FaturasFacilitiesV12()


def main():
    app = build_main_app()
    show_splash_and_start(app, delay_ms=SPLASH_DELAY_MS)


if __name__ == "__main__":
    main()