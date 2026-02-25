"""
Parser per il CSV esportato da Airbnb.

Come esportare da Airbnb:
  Account → Transazioni → Esporta CSV (seleziona periodo)

Struttura CSV: 21 colonne, encoding UTF-8 BOM.
Righe di tipo: Prenotazione / Payout / Ritenuta fiscale per il reddito italiano
"""

import pandas as pd
from datetime import datetime
from collections import defaultdict
from typing import List

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.models import Booking
from config import AIRBNB_LISTING_MAP


def _detect_property(annuncio: str) -> int:
    """Mappa il nome dell'annuncio al numero di proprietà."""
    annuncio_lower = annuncio.lower()
    for keyword, num in AIRBNB_LISTING_MAP.items():
        if keyword in annuncio_lower:
            return num
    # Fallback: prova a trovare "caldiero 5" o "caldiero 7" nel testo
    if "caldiero 5" in annuncio_lower or "5 person" in annuncio_lower:
        return 5
    if "caldiero 7" in annuncio_lower or "6 person" in annuncio_lower:
        return 7
    return 0  # proprietà sconosciuta


def _parse_date(s: str):
    """Converte data MM/DD/YYYY (formato americano Airbnb) in date."""
    if not s or str(s).strip() == "" or str(s) == "nan":
        return None
    s = str(s).strip()
    for fmt in ("%m/%d/%Y", "%Y-%m-%d", "%d/%m/%Y"):
        try:
            return datetime.strptime(s, fmt).date()
        except ValueError:
            continue
    return None


def _to_float(val) -> float:
    """Converte un valore in float, gestendo None e stringhe."""
    if val is None or str(val).strip() in ("", "nan", "-"):
        return 0.0
    try:
        return float(str(val).replace(",", ".").replace(" ", ""))
    except (ValueError, TypeError):
        return 0.0


def parse_airbnb_csv(filepath: str) -> List[Booking]:
    """
    Legge il CSV Airbnb e restituisce lista di Booking.
    Raggruppa righe per Codice di Conferma:
      - Riga 'Prenotazione' → dati principali
      - Righe 'Ritenuta fiscale...' → sommate
      - Righe 'Payout' → ignorate
    """
    try:
        df = pd.read_csv(filepath, encoding="utf-8-sig", dtype=str, na_filter=False)
    except Exception as e:
        raise ValueError(f"Errore lettura CSV Airbnb: {e}")

    # Normalizza nomi colonne
    df.columns = df.columns.str.strip()

    # Raggruppa per codice conferma (ignora righe senza codice = Payout)
    by_code = defaultdict(list)
    for _, row in df.iterrows():
        code = row.get("Codice di Conferma", "").strip()
        tipo = row.get("Tipo", "").strip()
        if code and tipo != "Payout" and "payout" not in tipo.lower():
            by_code[code].append(row)

    bookings = []
    for code, rows in by_code.items():
        # Trova la riga principale "Prenotazione"
        booking_row = None
        for r in rows:
            if r.get("Tipo", "").strip().lower() == "prenotazione":
                booking_row = r
                break
        if booking_row is None:
            continue

        # Somma ritenute (possono essere multiple per split pagamenti)
        withholding = 0.0
        for r in rows:
            if "ritenuta" in r.get("Tipo", "").lower():
                withholding += _to_float(r.get("Importo", 0))

        annuncio = booking_row.get("Annuncio", "")
        property_num = _detect_property(annuncio)

        check_in = _parse_date(booking_row.get("Data di inizio", ""))
        check_out = _parse_date(booking_row.get("Data di fine", ""))
        nights_str = booking_row.get("Notti", "0")
        try:
            nights = int(str(nights_str).strip()) if str(nights_str).strip() not in ("", "nan") else 0
        except ValueError:
            nights = 0

        bookings.append(Booking(
            platform="airbnb",
            property_num=property_num,
            guest_name=booking_row.get("Ospite", "").strip(),
            check_in=check_in,
            check_out=check_out,
            nights=nights,
            gross_amount=_to_float(booking_row.get("Guadagni lordi", 0)),
            commission=_to_float(booking_row.get("Costi del servizio", 0)),
            payment_charge=_to_float(booking_row.get("Commissione per Pagamento rapido", 0)),
            vat=0.0,  # non presente nel CSV Airbnb IT
            net_amount=_to_float(booking_row.get("Importo", 0)),
            withholding_tax=withholding,
            confirm_code=code,
            source_file=filepath,
        ))

    return bookings
