"""
Google Sheets storage — sostituisce Excel come database.

Il Google Sheet ha questi fogli:
  - elenco       → tutte le righe (prenotazioni + costi)
  - prenotazioni → vista filtrata solo prenotazioni (auto-generata)

Autenticazione via Service Account (credenziali in Streamlit secrets).

Setup una tantum (vedi README_SETUP.md):
  1. Crea Service Account su Google Cloud
  2. Condividi il Google Sheet con l'email del service account
  3. Metti le credenziali in .streamlit/secrets.toml
"""

import gspread
import pandas as pd
from datetime import date, datetime
from typing import List, Tuple
import streamlit as st

from core.models import Booking, Cost


# Nomi colonne nel Google Sheet (foglio "elenco")
SHEET_COLUMNS = [
    "caldiero", "dal", "al", "mese", "tax", "importo",
    "tipo", "causale", "ente", "nominativo", "documento", "nr",
    "data", "periodo", "intestata_a", "giorni",
    "inviato_1k", "ritenuta", "incassato", "lordo",
    "commission", "payment_charge", "vat", "euro_gg",
    "piattaforma_raw", "source_file",  # colonne extra per il report
]


@st.cache_resource
def get_gspread_client():
    """
    Restituisce client gspread autenticato via Service Account.
    Le credenziali vengono da st.secrets (Streamlit Cloud) o da
    .streamlit/secrets.toml in locale.
    """
    creds_dict = dict(st.secrets["gcp_service_account"])
    gc = gspread.service_account_from_dict(creds_dict)
    return gc


def get_sheet(sheet_name: str = "elenco"):
    """Apre il foglio specificato nel Google Sheet configurato."""
    gc = get_gspread_client()
    spreadsheet_id = st.secrets["google_sheets"]["spreadsheet_id"]
    sh = gc.open_by_key(spreadsheet_id)
    try:
        return sh.worksheet(sheet_name)
    except gspread.WorksheetNotFound:
        # Crea il foglio se non esiste
        ws = sh.add_worksheet(title=sheet_name, rows=1000, cols=30)
        if sheet_name == "elenco":
            ws.append_row(SHEET_COLUMNS)
        return ws


def load_elenco_from_sheets() -> pd.DataFrame:
    """
    Carica tutti i dati dal foglio 'elenco'.
    Restituisce DataFrame con colonne tipizzate.
    """
    ws = get_sheet("elenco")
    data = ws.get_all_records(expected_headers=SHEET_COLUMNS[:12])  # prime 12 colonne obbligatorie
    if not data:
        return pd.DataFrame(columns=SHEET_COLUMNS)
    df = pd.DataFrame(data)
    return df


def get_existing_codes() -> Tuple[set, set]:
    """
    Legge i codici già presenti per deduplicazione.
    Returns: (booking_codes, invoice_keys)
    """
    try:
        ws = get_sheet("elenco")
        # Legge solo colonne nr (L) e documento (K) per efficienza
        all_values = ws.get_all_values()
        if len(all_values) <= 1:
            return set(), set()

        headers = all_values[0]
        nr_idx = headers.index("nr") if "nr" in headers else 11
        doc_idx = headers.index("documento") if "documento" in headers else 10
        prop_idx = headers.index("caldiero") if "caldiero" in headers else 0

        booking_codes = set()
        invoice_keys = set()

        for row in all_values[1:]:
            if len(row) > nr_idx and row[nr_idx].strip():
                booking_codes.add(row[nr_idx].strip())
            if len(row) > doc_idx and row[doc_idx].strip() not in ("", "fattura", "bolletta", "prenotazione"):
                prop = row[prop_idx].strip() if len(row) > prop_idx else ""
                invoice_keys.add(f"{row[doc_idx].strip()}_{prop}")

        return booking_codes, invoice_keys
    except Exception:
        return set(), set()


def _booking_to_row(b: Booking) -> list:
    """Converte un Booking in lista di valori per il Sheet."""
    def fmt_date(d):
        return d.strftime("%Y-%m-%d") if d else ""

    nights = b.nights or 0
    euro_gg = round(b.gross_amount / nights, 2) if nights > 0 else ""

    return [
        b.property_num,                          # caldiero
        fmt_date(b.check_in),                    # dal
        fmt_date(b.check_out),                   # al
        b.check_in.month if b.check_in else "",  # mese
        "T",                                      # tax
        round(b.net_amount, 2),                  # importo
        "incasso",                               # tipo
        "da clienti",                            # causale
        b.platform.capitalize(),                 # ente
        b.guest_name,                            # nominativo
        "prenotazione",                          # documento
        b.confirm_code,                          # nr
        fmt_date(b.check_in),                   # data
        "",                                      # periodo
        "",                                      # intestata_a
        nights,                                  # giorni
        "",                                      # inviato_1k
        round(b.withholding_tax, 2) if b.withholding_tax else "", # ritenuta
        round(b.net_amount, 2),                  # incassato
        round(b.gross_amount, 2),               # lordo
        round(b.commission, 2) if b.commission else "",       # commission
        round(b.payment_charge, 2) if b.payment_charge else "", # payment_charge
        round(b.vat, 2) if b.vat else "",       # vat
        euro_gg,                                 # euro_gg
        b.platform,                              # piattaforma_raw
        b.source_file,                           # source_file
    ]


def _cost_to_row(c: Cost) -> list:
    """Converte un Cost in lista di valori per il Sheet."""
    def fmt_date(d):
        return d.strftime("%Y-%m-%d") if d else ""

    return [
        c.property_num,                          # caldiero
        fmt_date(c.date),                        # dal
        "",                                      # al
        c.date.month if c.date else "",          # mese
        "T",                                     # tax
        round(c.amount, 2),                      # importo
        "ordinarie",                             # tipo
        c.category,                              # causale
        c.supplier,                              # ente
        "bolletta" if c.category in ("acqua", "energia elettrica", "gas") else "fattura",  # nominativo
        c.invoice_num or "",                     # documento
        "",                                      # nr
        fmt_date(c.invoice_date or c.date),     # data
        "",                                      # periodo
        "si",                                    # intestata_a
        "",                                      # giorni
        "", "", "", "", "", "", "",              # inviato_1k ... vat
        "",                                      # euro_gg
        "costo",                                 # piattaforma_raw
        c.source_file,                           # source_file
    ]


def save_to_sheets(
    bookings: List[Booking],
    costs: List[Cost],
    dry_run: bool = False,
) -> Tuple[int, int, List[str]]:
    """
    Salva prenotazioni e costi su Google Sheets.
    Gestisce deduplicazione.

    Returns: (added_bookings, added_costs, skipped_codes)
    """
    existing_booking_codes, existing_invoice_keys = get_existing_codes()

    # Filtra duplicati
    new_bookings = []
    skipped = []
    for b in bookings:
        if b.confirm_code in existing_booking_codes:
            skipped.append(b.confirm_code)
        else:
            new_bookings.append(b)
            existing_booking_codes.add(b.confirm_code)

    new_costs = []
    for c in costs:
        key = f"{c.invoice_num}_{c.property_num}" if c.invoice_num else None
        if key and key in existing_invoice_keys:
            skipped.append(c.invoice_num)
        else:
            new_costs.append(c)
            if key:
                existing_invoice_keys.add(key)

    if dry_run:
        return len(new_bookings), len(new_costs), skipped

    if not new_bookings and not new_costs:
        return 0, 0, skipped

    ws = get_sheet("elenco")

    # Batch append per efficienza (una sola chiamata API)
    rows = []
    for b in new_bookings:
        rows.append(_booking_to_row(b))
    for c in new_costs:
        rows.append(_cost_to_row(c))

    if rows:
        ws.append_rows(rows, value_input_option="USER_ENTERED")

    return len(new_bookings), len(new_costs), skipped
