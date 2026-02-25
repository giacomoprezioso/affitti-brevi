"""
Parser per il file XLSX esportato da Booking.com (Payout report).

Come esportare da Booking:
  Extranet → Finance → Pagamenti → Esporta (seleziona periodo)
  Nome file: Payout_from_YYYY-MM-DD_until_YYYY-MM-DD.xlsx

Struttura: 31 colonne. Righe raggruppate per numero riferimento (col C):
  - Tipo 'Prenotazione' → dati principali
  - Tipo 'Ritenuta per locazione breve' → withholding_tax
  - Tipo Payout (vuoto o altro) → ignorato
"""

import pandas as pd
from datetime import date, datetime
from collections import defaultdict
from typing import List

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.models import Booking
from config import BOOKING_PROPERTY_MAP


def _detect_property(name: str) -> int:
    """Mappa il nome struttura Booking al numero di proprietà."""
    name_lower = name.lower()
    for keyword, num in BOOKING_PROPERTY_MAP.items():
        if keyword in name_lower:
            return num
    if "caldiero 5" in name_lower:
        return 5
    if "caldiero 7" in name_lower:
        return 7
    return 0


def _to_float(val) -> float:
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return 0.0
    try:
        return float(str(val).replace(",", ".").replace(" ", "").replace("-", "0") if str(val).strip() == "-" else str(val).replace(",", "."))
    except (ValueError, TypeError):
        return 0.0


def _to_date(val) -> date:
    """Converte valore pandas (datetime o serial) in date."""
    if val is None or (isinstance(val, float) and pd.isna(val)):
        return None
    if isinstance(val, (datetime,)):
        return val.date()
    if isinstance(val, date):
        return val
    if isinstance(val, pd.Timestamp):
        return val.date()
    # Prova come stringa
    s = str(val).strip()
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def parse_booking_xlsx(filepath: str) -> List[Booking]:
    """
    Legge il file XLSX Booking Payout e restituisce lista di Booking.
    """
    try:
        df = pd.read_excel(filepath, engine="openpyxl", header=0, dtype=str, na_filter=False)
    except Exception as e:
        raise ValueError(f"Errore lettura XLSX Booking: {e}")

    # Le colonne sono senza nome nell'header se Booking usa indici
    # Proviamo a leggere con header e gestire i nomi reali
    # Rileggiamo con types corretti (date come datetime)
    df_typed = pd.read_excel(filepath, engine="openpyxl", header=0)

    # Colonne per posizione (0-indexed):
    # 0=Tipo, 1=Descrizione, 2=Numero riferimento, 3=Check-in, 4=Check-out
    # 5=Data emissione, 6=Stato, 7=Camere, 8=Pernottamenti, 9=ID struttura
    # 10=Nome struttura, ..., 15=Importo lordo, 16=Commissione, 18=Costo transazione
    # 20=IVA, 22=Importo netto, 25=Importo dovuto, 26=Importo pagato, ...

    col_names = list(df_typed.columns)

    # Mappa colonne per nome (se presenti) o per posizione
    def get_col(row, name_substr, fallback_idx):
        for i, c in enumerate(col_names):
            if name_substr.lower() in str(c).lower():
                return row.iloc[i]
        if fallback_idx < len(row):
            return row.iloc[fallback_idx]
        return None

    # Raggruppa per colonna 1 (Descrizione = ID payout alfanumerico, es. XOG8WLipyJpndWAr)
    # Le righe Prenotazione, Ritenuta condividono lo stesso ID descrizione
    # La riga Payout ha tipo vuoto o NaN → skippiamo
    by_ref = defaultdict(list)
    for _, row in df_typed.iterrows():
        tipo = str(row.iloc[0]).strip() if pd.notna(row.iloc[0]) else ""
        desc = str(row.iloc[1]).strip() if pd.notna(row.iloc[1]) else ""  # ID payout alfanumerico

        if not desc or desc in ("nan", "", "-"):
            continue

        # Skip righe payout (tipo vuoto/NaN = righe sommario del payout)
        if tipo == "" or tipo.lower() == "nan":
            continue

        by_ref[desc].append(row)

    bookings = []
    for ref, rows in by_ref.items():
        # Trova riga prenotazione
        booking_row = None
        for r in rows:
            tipo = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ""
            if "prenotazione" in tipo.lower():
                booking_row = r
                break
        if booking_row is None:
            continue

        # Ritenuta: riga tipo "Ritenuta per locazione breve", importo in col 15 (Importo lordo)
        withholding = 0.0
        for r in rows:
            tipo = str(r.iloc[0]).strip() if pd.notna(r.iloc[0]) else ""
            if "ritenuta" in tipo.lower():
                val = r.iloc[15] if len(r) > 15 else 0  # Importo lordo col 15 (negativo)
                withholding += _to_float(val)

        property_name = str(booking_row.iloc[10]).strip() if len(booking_row) > 10 and pd.notna(booking_row.iloc[10]) else ""
        property_num = _detect_property(property_name)

        check_in = _to_date(booking_row.iloc[3]) if len(booking_row) > 3 else None
        check_out = _to_date(booking_row.iloc[4]) if len(booking_row) > 4 else None

        nights_val = booking_row.iloc[8] if len(booking_row) > 8 else 0
        try:
            nights = int(float(str(nights_val))) if pd.notna(nights_val) and str(nights_val) not in ("", "nan") else 0
        except (ValueError, TypeError):
            nights = 0

        gross = _to_float(booking_row.iloc[15]) if len(booking_row) > 15 else 0.0
        commission = _to_float(booking_row.iloc[16]) if len(booking_row) > 16 else 0.0
        payment_charge = _to_float(booking_row.iloc[18]) if len(booking_row) > 18 else 0.0
        vat = _to_float(booking_row.iloc[20]) if len(booking_row) > 20 else 0.0
        net = _to_float(booking_row.iloc[22]) if len(booking_row) > 22 else 0.0

        # confirm_code: numero prenotazione numerico (col 2), più leggibile dell'ID payout
        confirm_code = str(booking_row.iloc[2]).strip() if pd.notna(booking_row.iloc[2]) else ref

        # Nome ospite: non presente nel Payout export di Booking
        guest_name = confirm_code  # usiamo il numero prenotazione come identificativo

        bookings.append(Booking(
            platform="booking",
            property_num=property_num,
            guest_name=guest_name,
            check_in=check_in,
            check_out=check_out,
            nights=nights,
            gross_amount=gross,
            commission=commission,
            payment_charge=payment_charge,
            vat=vat,
            net_amount=net,
            withholding_tax=withholding,
            confirm_code=confirm_code,
            source_file=filepath,
        ))

    return bookings
