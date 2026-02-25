"""
Controllo duplicati: evita di inserire in Excel record già presenti.

Legge il foglio 'elenco':
  - Colonna L (nr.) per codici prenotazione Airbnb/Booking
  - Colonna K (documento) per numeri fattura PDF
"""

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from config import SHEET_ELENCO, COL_MAP


def load_existing_codes(wb) -> tuple[set, set]:
    """
    Legge il foglio elenco e restituisce:
      (booking_codes, invoice_numbers)
    Entrambi come set di stringhe normalizzate.
    """
    ws = wb[SHEET_ELENCO]
    booking_codes = set()
    invoice_numbers = set()

    nr_col = COL_MAP["nr"]         # colonna L (1-indexed)
    doc_col = COL_MAP["documento"]  # colonna K (1-indexed)

    for row in ws.iter_rows(min_row=2, values_only=True):
        # Colonna L (nr.) - codici prenotazione
        val_nr = row[nr_col - 1] if len(row) >= nr_col else None
        if val_nr is not None and str(val_nr).strip():
            booking_codes.add(str(val_nr).strip())

        # Colonna K (documento) - numeri fattura
        val_doc = row[doc_col - 1] if len(row) >= doc_col else None
        if val_doc is not None and str(val_doc).strip() not in ("", "fattura", "bolletta", "nan"):
            invoice_numbers.add(str(val_doc).strip())

    return booking_codes, invoice_numbers


def is_booking_duplicate(confirm_code: str, existing_codes: set) -> bool:
    return str(confirm_code).strip() in existing_codes


def is_invoice_duplicate(invoice_num: str, existing_invoices: set, property_num: int = None) -> bool:
    if not invoice_num or invoice_num.strip() == "":
        return False
    # Per fatture split (stesso invoice_num su più proprietà), aggiungi property_num come chiave
    key = str(invoice_num).strip()
    if property_num is not None:
        key = f"{key}_{property_num}"
    return key in existing_invoices
