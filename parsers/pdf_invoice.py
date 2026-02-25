"""
Parser per fatture PDF (bollette e fatture servizi).

Fornitori supportati (riconosciuti automaticamente dal testo):
  - Acque Veronesi (bolletta acqua)
  - Vivi Energia (bolletta luce)
  - Reshma (pulizie) - copre entrambe le proprietà → 2 Cost con importo 50/50

Il fornitore viene rilevato dal testo del PDF, non dal nome del file.
La proprietà viene rilevata dall'indirizzo di fornitura nel testo.
"""

import re
from datetime import datetime, date
from typing import List

import sys
import os
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.models import Cost


def _extract_amount(text: str, supplier: str) -> float:
    """Estrae l'importo totale dalla fattura in base al fornitore."""
    patterns_by_supplier = {
        "acque veronesi": [
            r"Totale\s+Bolletta\s*[:\s]+(\d{1,4}[,\.]\d{2})\s*EUR",
            r"TOTALE\s+(?:DA\s+PAGARE|BOLLETTA)\s*[:\s]+(\d{1,4}[,\.]\d{2})",
        ],
        "vivi": [
            r"TOTALE\s+DA\s+PAGARE\s*[:\s]+(\d{1,4}[,\.]\d{2})\s*EUR",
            r"Totale\s+da\s+pagare\s*[:\s]+(\d{1,4}[,\.]\d{2})",
        ],
        "reshma": [
            r"[Nn]etto\s+a\s+pagare\s+(\d{1,4}[,\.]\d{2})\s*[€EUR]",
            r"[Tt]otale\s+documento\s+(\d{1,4}[,\.]\d{2})\s*[€EUR]",
            r"[Ii]mporto\s+prodotti\s+o\s+servizi\s+(\d{1,4}[,\.]\d{2})\s*[€EUR]",
            r"TOTALE\s*[:\s]+(\d{1,4}[,\.]\d{2})\s*EUR",
        ],
    }

    # Pattern specifici per fornitore
    for key, patterns in patterns_by_supplier.items():
        if key in supplier.lower():
            for pattern in patterns:
                m = re.search(pattern, text, re.IGNORECASE)
                if m:
                    return float(m.group(1).replace(",", "."))

    # Fallback generico
    fallback_patterns = [
        r"TOTALE\s+DA\s+PAGARE\s*[:\s]+(\d{1,4}[,\.]\d{2})",
        r"Totale\s+[Ff]attura\s*[:\s]+(\d{1,4}[,\.]\d{2})",
        r"Importo\s+[Tt]otale\s*[:\s]+(\d{1,4}[,\.]\d{2})",
        r"[Tt]otale\s*[:\s]+(\d{1,4}[,\.]\d{2})\s*EUR",
        r"€\s*(\d{1,4}[,\.]\d{2})\b",
    ]
    for pattern in fallback_patterns:
        m = re.search(pattern, text)
        if m:
            return float(m.group(1).replace(",", "."))

    return 0.0


def _extract_date(text: str) -> date:
    """Estrae la data della fattura."""
    patterns = [
        r"[Dd]ata\s+di\s+emissione\s*[:\s]+(\d{2}/\d{2}/\d{4})",
        r"[Dd]ata\s+[Ff]attura\s*[:\s]+(\d{2}/\d{2}/\d{4})",
        r"[Ff]attura\s+del\s+(\d{2}/\d{2}/\d{4})",
        r"del\s+(\d{2}/\d{2}/\d{4})",
        r"[Dd]ata\s+emissione\s*[:\s]+(\d{2}/\d{2}/\d{4})",
        r"FATTURA\s+NR\.?\s+\S+\s+del\s+(\d{2}/\d{2}/\d{4})",
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            try:
                return datetime.strptime(m.group(1), "%d/%m/%Y").date()
            except ValueError:
                continue
    # Fallback: cerca tutte le date e prende la più recente (probabile data fattura)
    all_dates = re.findall(r"(\d{2}/\d{2}/\d{4})", text)
    parsed_dates = []
    for d in all_dates:
        try:
            parsed_dates.append(datetime.strptime(d, "%d/%m/%Y").date())
        except ValueError:
            continue
    if parsed_dates:
        return max(parsed_dates)  # la data più recente è probabilmente quella della fattura
    return None


def _extract_invoice_number(text: str) -> str:
    """Estrae il numero della fattura."""
    patterns = [
        r"FATTURA\s+NR\.?\s+(\S+)",
        r"[Ff]attura\s+[Nn]\.?\s+(\S+)",
        r"[Ff]attura\s+[Nn]r\.?\s+(\S+)",
        r"[Bb]olletta\s+[Nn]\.?\s+(\S+)",
        r"[Nn]umero\s+[Ff]attura\s*[:\s]+(\S+)",
        r"[Nn]\.?\s*(\d{4}/\d{4}/\d+)",  # formato Acque Veronesi
    ]
    for pattern in patterns:
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()
    return ""


def _detect_supplier(text: str, filepath: str) -> tuple[str, str]:
    """
    Riconosce il fornitore dal testo PDF.
    Restituisce (supplier_name, category).
    """
    text_lower = text.lower()

    if "acque veronesi" in text_lower or "acquedotto" in text_lower:
        return "Acque veronesi", "acqua"
    if "vivi" in text_lower and ("energia" in text_lower or "luce" in text_lower or "elettric" in text_lower):
        return "Vivi", "energia elettrica"
    if "reshma" in text_lower or "jebun" in text_lower or "jeebun" in text_lower:
        return "Reshma", "pulizie"
    if "enel" in text_lower:
        return "Enel", "energia elettrica"
    if "eni" in text_lower and "gas" in text_lower:
        return "Eni Gas", "gas"
    if "vodafone" in text_lower or "tim" in text_lower or "wind" in text_lower:
        return "Telefonia", "wifi"

    # Fallback: usa il nome del file
    basename = os.path.basename(filepath).lower()
    if "acqua" in basename or "idric" in basename:
        return "Acqua", "acqua"
    if "luce" in basename or "energia" in basename or "enel" in basename:
        return "Energia", "energia elettrica"
    if "pulizie" in basename or "reshma" in basename:
        return "Pulizie", "pulizie"

    return "Sconosciuto", "ordinarie"


def _detect_property(text: str) -> int:
    """
    Rileva la proprietà dall'indirizzo nel testo PDF.
    VIA FONTE 5 → Caldiero 5, VIA FONTE 7 → Caldiero 7
    """
    # Cerca pattern indirizzo specifico
    text_upper = text.upper()

    # Pattern specifici per le due proprietà
    if re.search(r"FONTE\s+5\b|VIA\s+FONTE,?\s*5\b", text_upper):
        return 5
    if re.search(r"FONTE\s+7\b|VIA\s+FONTE,?\s*7\b", text_upper):
        return 7

    # Cerca "CALDIERO 5" o "CALDIERO 7" nel testo
    if re.search(r"CALDIERO\s+5\b", text_upper):
        return 5
    if re.search(r"CALDIERO\s+7\b", text_upper):
        return 7

    return 0  # proprietà non rilevata


def _covers_both_properties(text: str, supplier: str) -> bool:
    """Verifica se la fattura copre entrambe le proprietà."""
    text_upper = text.upper()
    has_5 = bool(re.search(r"FONTE\s+5\b|CALDIERO\s+5\b", text_upper))
    has_7 = bool(re.search(r"FONTE\s+7\b|CALDIERO\s+7\b", text_upper))

    # Reshma fattura sempre entrambe le proprietà insieme
    if "reshma" in supplier.lower() or "jeebun" in supplier.lower():
        return True

    return has_5 and has_7


def parse_pdf_invoice(filepath: str) -> List[Cost]:
    """
    Estrae dati da una fattura PDF.
    Restituisce lista di Cost (normalmente 1, ma 2 per fatture che coprono entrambe le proprietà).
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("Installa pdfplumber: pip install pdfplumber")

    try:
        with pdfplumber.open(filepath) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        raise ValueError(f"Errore lettura PDF {filepath}: {e}")

    if not text.strip():
        raise ValueError(f"PDF vuoto o non leggibile: {filepath}")

    supplier, category = _detect_supplier(text, filepath)
    amount = _extract_amount(text, supplier)
    invoice_date = _extract_date(text)
    invoice_num = _extract_invoice_number(text)

    if amount == 0.0:
        raise ValueError(f"Importo non trovato nel PDF: {filepath}")

    if _covers_both_properties(text, supplier):
        # Split 50/50
        half = round(amount / 2, 2)
        return [
            Cost(
                property_num=5,
                date=invoice_date or date.today(),
                amount=-half,
                category=category,
                supplier=supplier,
                invoice_num=invoice_num,
                invoice_date=invoice_date,
                source_file=filepath,
            ),
            Cost(
                property_num=7,
                date=invoice_date or date.today(),
                amount=-half,
                category=category,
                supplier=supplier,
                invoice_num=invoice_num,
                invoice_date=invoice_date,
                source_file=filepath,
            ),
        ]
    else:
        property_num = _detect_property(text)
        return [
            Cost(
                property_num=property_num,
                date=invoice_date or date.today(),
                amount=-amount,
                category=category,
                supplier=supplier,
                invoice_num=invoice_num,
                invoice_date=invoice_date,
                source_file=filepath,
            )
        ]
