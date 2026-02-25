"""
Parser per il file CSV esportato da Booking.com (Payout report).

Come esportare da Booking:
  Extranet → Finance → Pagamenti → Esporta CSV
  Nome file: Payout_from_YYYY-MM-DD_until_YYYY-MM-DD.csv

Struttura CSV: 31 colonne, separatore virgola, encoding utf-8-sig.
Tipi di righe:
  - '(Payout)'                    → riga sommario pagamento (ignorata)
  - 'Prenotazione'                → dati principali prenotazione
  - 'Ritenuta per locazione breve'→ ritenuta fiscale (importo negativo in col 15)
  - 'credit_note'                 → nota di credito (eventuale rimborso/rettifica)

Colonne (0-indexed):
  0  = Tipo/tipo di transazione
  1  = Descrizione (payout ID alfanumerico, es. 'qatF6PVrfEzGNjsX')
  2  = Numero di riferimento (numero prenotazione numerico per 'Prenotazione',
                              numero ritenuta per 'Ritenuta per locazione breve')
  3  = Data check-in (YYYY-MM-DD)
  4  = Data check-out (YYYY-MM-DD)
  5  = Data emissione
  6  = Stato prenotazione
  7  = Camere
  8  = Pernottamenti
  9  = ID struttura
  10 = Nome struttura
  15 = Importo lordo
  16 = Commissione (negativo)
  18 = Costo transazione (negativo)
  20 = IVA (negativo)
  22 = Importo della transazione (netto = lordo + commissione + costi)
  25 = Importo dovuto (netto dopo ritenuta)

Nota: le righe Ritenuta usano col 15 per l'importo (negativo) e condividono
il payout ID (col 1) con la prenotazione corrispondente.
"""

import csv
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


def _to_float(val: str) -> float:
    """Converte stringa CSV in float. '-' e valori mancanti → 0.0."""
    if not val or val.strip() in ("-", "", "nan"):
        return 0.0
    try:
        return float(val.strip().replace(",", "."))
    except (ValueError, TypeError):
        return 0.0


def _to_date(val: str) -> date:
    """Converte stringa data (YYYY-MM-DD) in date."""
    if not val or val.strip() in ("-", "", "nan"):
        return None
    for fmt in ("%Y-%m-%d", "%d/%m/%Y", "%m/%d/%Y"):
        try:
            return datetime.strptime(val.strip(), fmt).date()
        except ValueError:
            continue
    return None


def parse_booking_csv(filepath: str) -> List[Booking]:
    """
    Legge il file CSV Booking Payout e restituisce lista di Booking.

    Gestisce:
    - Prenotazioni con ritenuta fiscale (abbinate per payout ID, col 1)
    - Credit note (ignorate nel conteggio prenotazioni)
    - Righe payout sommario (ignorate)
    """
    try:
        with open(filepath, encoding="utf-8-sig", newline="") as f:
            reader = csv.reader(f)
            rows = list(reader)
    except UnicodeDecodeError:
        try:
            with open(filepath, encoding="latin-1", newline="") as f:
                reader = csv.reader(f)
                rows = list(reader)
        except Exception as e:
            raise ValueError(f"Errore lettura CSV Booking: {e}")
    except Exception as e:
        raise ValueError(f"Errore lettura CSV Booking: {e}")

    if len(rows) < 2:
        return []

    # Salta header (riga 0)
    data_rows = rows[1:]

    # Raggruppa per payout ID (col 1 = Descrizione)
    # Tutte le righe con stesso payout ID (Prenotazione + Ritenute) vengono raggruppate
    by_payout_id = defaultdict(list)
    for row in data_rows:
        if len(row) < 3:
            continue
        tipo = row[0].strip()
        payout_id = row[1].strip()

        # Skip righe payout sommario e righe vuote
        if tipo in ("(Payout)", "") or not payout_id or payout_id == "-":
            continue

        by_payout_id[payout_id].append(row)

    bookings = []
    for payout_id, group_rows in by_payout_id.items():
        # Trova riga principale Prenotazione
        booking_row = None
        for r in group_rows:
            if "prenotazione" in r[0].strip().lower():
                booking_row = r
                break

        if booking_row is None:
            # Nessuna prenotazione in questo gruppo (solo ritenuta o credit_note standalone)
            continue

        # Somma ritenute fiscali (importo negativo in col 15)
        withholding = 0.0
        for r in group_rows:
            if "ritenuta" in r[0].strip().lower():
                if len(r) > 15:
                    withholding += _to_float(r[15])  # già negativo nel CSV

        # Dati dalla riga Prenotazione
        property_name = booking_row[10].strip() if len(booking_row) > 10 else ""
        property_num = _detect_property(property_name)

        check_in = _to_date(booking_row[3]) if len(booking_row) > 3 else None
        check_out = _to_date(booking_row[4]) if len(booking_row) > 4 else None

        nights = 0
        if len(booking_row) > 8 and booking_row[8].strip() not in ("-", ""):
            try:
                nights = int(float(booking_row[8].strip()))
            except (ValueError, TypeError):
                nights = 0

        gross = _to_float(booking_row[15]) if len(booking_row) > 15 else 0.0
        commission = _to_float(booking_row[16]) if len(booking_row) > 16 else 0.0
        payment_charge = _to_float(booking_row[18]) if len(booking_row) > 18 else 0.0
        vat = _to_float(booking_row[20]) if len(booking_row) > 20 else 0.0
        net = _to_float(booking_row[22]) if len(booking_row) > 22 else 0.0

        # confirm_code = numero prenotazione (col 2) o payout_id se non disponibile
        confirm_code = booking_row[2].strip() if len(booking_row) > 2 and booking_row[2].strip() not in ("-", "") else payout_id
        # Nome ospite non disponibile nel payout export
        guest_name = confirm_code

        # Gli importi negativi nelle prenotazioni indicano cancellazioni/rimborsi
        # Le lasciamo col segno originale (gestite dal deduplicator)
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
            source_file=os.path.basename(filepath),
        ))

    return bookings
