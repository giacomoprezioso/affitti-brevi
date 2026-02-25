"""
Parser per fatture PDF (bollette e fatture servizi).

Fornitori supportati (riconosciuti automaticamente dal testo):
  - Acque Veronesi  → acqua
  - Enel            → energia elettrica / gas (in base al contenuto)
  - Gritti Energia  → energia elettrica / gas
  - Vivi / ViviGas  → energia elettrica / gas
  - Vodafone        → wifi
  - Reshma Jeebun   → pulizie (copre entrambe le proprietà → 2 Cost 50/50)
  - 1K HOME s.r.l.s → fee (se GESTIONE LOCAZIONE TURISTICA) o rimborso bollette
  - Spurgo/idraulico → manutenzione straordinaria (rilevato da parole chiave)
  - Saccomani/Entratel → imposta di registro locazione
  - Generico        → fallback su parole chiave / nome file

Il fornitore viene rilevato dal testo del PDF, non dal nome del file.
La proprietà viene rilevata dall'indirizzo di fornitura nel testo.

NOTE SUI FORMATI:
  - Enel:   "Totale da pagare 108,17 €"  (numero poi €, NO due punti)
  - Gritti: "Totale da pagare 121,00"     (a fine riga sommario)
  - Vivi:   "TOTALE\nDA PAGARE:\n137,45 €" (multiriga, con :/€)
  - Vodafone: "Importo totale\n19,90 €"  o "Totale da pagare\n19,90 €"
  - Acque Veronesi: "Totale Bolletta € 144,54"  (€ poi numero)
  - Reshma: "Netto a pagare 317,50 €"
  - 1KHome: "€ 244,00" nella sezione SCADENZE (oppure importo finale)
  - Gritti: numero in colonna a destra di "Totale da pagare"
  - Spurgo: "Totale E 236,68" oppure "Totale € 236,68"
"""

import re
import os
from datetime import datetime, date
from typing import List, Optional

import sys
sys.path.insert(0, os.path.dirname(os.path.dirname(os.path.abspath(__file__))))

from core.models import Cost


# ─── Rilevamento fornitore ───────────────────────────────────────────────────

def _detect_supplier(text: str, filepath: str) -> tuple[str, str]:
    """
    Riconosce il fornitore dal testo PDF.
    Restituisce (supplier_name, category).
    """
    tl = text.lower()

    # 1K HOME deve venire PRIMA di altri check perché il testo contiene
    # riferimenti a "bollette acqua/gas" che potrebbero confondere.
    if "1k home" in tl or "1khome" in tl or "gestionekhome" in tl:
        # Distingui fee da rimborso bollette
        if "rimborso spese anticipate" in tl or "rimborso bollette" in tl:
            return "1KHome", "rimborso bollette"
        return "1KHome", "fee"

    if "reshma" in tl or "jeebun" in tl:
        return "Reshma", "pulizie"

    # Spurgo / idraulico PRIMA di altri check per evitare falsi positivi
    if any(k in tl for k in ["autospurgo", "spurgo", "tubazioni", "disotturazione", "idraulic"]):
        return "Idraulico", "manutenzione straordinaria"

    # Gritti Energia (viene prima di Enel/Vivi)
    if "gritti" in tl:
        # Distingui energia da gas dal titolo esplicito o dalle chiavi specifiche
        if "bolletta per la fornitura di gas" in tl or "codice pdr" in tl or "codice punto di riconsegna" in tl:
            return "Gritti", "gas"
        if "bolletta per la fornitura di energia elettrica" in tl or "codice pod" in tl or "codice punto di prelievo" in tl:
            return "Gritti", "energia elettrica"
        # Fallback: numero bolletta inizia con G=gas, E=energia
        m = re.search(r"NUMERO\s+BOLLETTA\s*\n?(\d{4}[EG])", text, re.IGNORECASE)
        if m:
            return "Gritti", "gas" if "G" in m.group(1).upper() else "energia elettrica"
        return "Gritti", "energia elettrica"

    # Vivi / ViviGas (viene PRIMA di Enel perché le bollette Vivi citano "E-DISTRIBUZIONE" e "Enel")
    if "vivigas" in tl or "vivi energia" in tl or "vivienergia" in tl or "viviweb" in tl:
        # Distingui gas da elettricità dal titolo esplicito della bolletta sintetica
        if "bolletta sintetica - gas" in tl:
            return "Vivi", "gas"
        if "bolletta sintetica - energia elettrica" in tl or "vivigas" not in tl:
            # vivigas nel nome azienda indica gas; altrimenti è luce
            if "vivigas s.p.a" in tl and "luce" not in tl:
                return "Vivi", "gas"
            return "Vivi", "energia elettrica"
        return "Vivi", "energia elettrica"

    # Enel (solo se intestatario/brand principale)
    # Il testo "enel energia" o "enel s.p.a" nel header indica fornitore Enel
    if re.search(r"^Enel\s+Energia", text, re.MULTILINE) or "enel energia s.p.a" in tl:
        # Distingui energia da gas: gas ha "Codice PDR", energia ha "Codice POD"
        if re.search(r"Codice\s+PDR\b", text, re.IGNORECASE) or "fornitura di gas naturale" in tl:
            return "Enel", "gas"
        if re.search(r"Codice\s+POD\b", text, re.IGNORECASE) or "energia elettrica" in tl:
            return "Enel", "energia elettrica"
        return "Enel", "energia elettrica"

    if "acque veronesi" in tl or "servizio idrico" in tl or "acquedotto" in tl:
        return "Acque Veronesi", "acqua"

    if "vodafone" in tl:
        return "Vodafone", "wifi"

    # Entratel / imposta di registro locazione
    if "entratel" in tl or "imposta di registro" in tl or "contratti di locazione" in tl:
        return "Agenzia Entrate", "imposta di registro"

    # Fallback su nome file
    basename = os.path.basename(filepath).lower()
    if "acqua" in basename or "idric" in basename:
        return "Acqua", "acqua"
    if "enel" in basename:
        if "gas" in basename:
            return "Enel", "gas"
        return "Enel", "energia elettrica"
    if "gritti" in basename:
        if "gas" in basename:
            return "Gritti", "gas"
        return "Gritti", "energia elettrica"
    if "vivi" in basename:
        if "gas" in basename:
            return "Vivi", "gas"
        return "Vivi", "energia elettrica"
    if "vodafone" in basename or "tim" in basename or "wind" in basename:
        return "Telefonia", "wifi"
    if "reshma" in basename or "pulizie" in basename:
        return "Pulizie", "pulizie"
    if "spurgo" in basename or "idraulico" in basename:
        return "Idraulico", "manutenzione straordinaria"

    return "Sconosciuto", "ordinarie"


# ─── Estrazione importo ─────────────────────────────────────────────────────

def _parse_amount(s: str) -> float:
    """Converte stringa tipo '1.066,22' o '108,17' in float."""
    s = s.strip().replace(".", "").replace(",", ".")
    return float(s)


def _extract_amount(text: str, supplier: str) -> float:
    """Estrae l'importo totale dalla fattura."""

    sl = supplier.lower()

    # ── Acque Veronesi ──
    # "Totale Bolletta € 144,54" oppure in prima riga dati "144,54"
    if "acque veronesi" in sl or "acqua" in sl:
        # Prima riga del PDF: "...#144,54#..."
        m = re.search(r"Totale\s+(?:Bolletta|fornitura)\s+[€\u20ac]?\s*([\d.]+,\d{2})", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))
        # Pattern nel sommario dati
        m = re.search(r"^[\d]+#[\d]+#([\d]+,\d{2})#", text, re.MULTILINE)
        if m:
            return _parse_amount(m.group(1))

    # ── Enel ──
    # "Totale da pagare 108,17 €" (NO due punti, € dopo il numero)
    if "enel" in sl:
        m = re.search(r"Totale\s+da\s+pagare\s+([\d.]+,\d{2})\s*[€\u20ac]", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))
        m = re.search(r"Totale\s+Bolletta\s+([\d.]+,\d{2})\s*[€\u20ac]", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))

    # ── Gritti ──
    # "Totale da pagare 121,00" (a fine riga)
    if "gritti" in sl:
        m = re.search(r"Totale\s+da\s+pagare\s+([\d.]+,\d{2})", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))
        m = re.search(r"Totale\s+Bolletta\s+([\d.]+,\d{2})", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))

    # ── Vivi ──
    # "TOTALE\nDA PAGARE:\n137,45 €" oppure "Totale bolletta 83,45 €"
    if "vivi" in sl:
        # Pattern in cima (sintetico)
        m = re.search(r"TOTALE\s+DA\s+PAGARE[:\s]+([\d.]+,\d{2})\s*[€\u20ac]", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))
        # Nel dettaglio bolletta (include canone RAI ecc.)
        m = re.search(r"Totale\s+da\s+pagare\s+([\d.]+,\d{2})\s*[€\u20ac]?", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))
        m = re.search(r"Totale\s+bolletta\s+([\d.]+,\d{2})\s*[€\u20ac]?", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))

    # ── Vodafone ──
    # "Importo totale\n19,90 €" oppure "Totale da pagare\n19,90 €"
    if "vodafone" in sl or "wifi" in sl or "telefon" in sl:
        m = re.search(r"(?:Importo\s+totale|Totale\s+da\s+pagare)\s*\n?\s*([\d.]+,\d{2})\s*[€\u20ac]?", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))
        # Anche come "19,90 €" standalone nella sezione conto
        m = re.search(r"Stato\s+dei\s+pagamenti.*?([\d.]+,\d{2})\s*[€\u20ac]", text, re.IGNORECASE | re.DOTALL)
        if m:
            return _parse_amount(m.group(1))

    # ── Reshma ──
    if "reshma" in sl or "pulizie" in sl:
        m = re.search(r"[Nn]etto\s+a\s+pagare\s+([\d.]+,\d{2})\s*[€\u20ac]?", text)
        if m:
            return _parse_amount(m.group(1))
        m = re.search(r"[Tt]otale\s+documento\s+([\d.]+,\d{2})\s*[€\u20ac]?", text)
        if m:
            return _parse_amount(m.group(1))
        m = re.search(r"[Ii]mporto\s+prodotti\s+o\s+servizi\s+([\d.]+,\d{2})\s*[€\u20ac]?", text)
        if m:
            return _parse_amount(m.group(1))

    # ── 1KHome ──
    # "SCADENZE\n04/09/2025: € 244,00" oppure importo con IVA al fondo
    if "1khome" in sl or "1k home" in sl:
        # Nella sezione SCADENZE: "DD/MM/YYYY: € 244,00"
        m = re.search(r"SCADENZE.*?(\d{2}/\d{2}/\d{4})[:\s]+[€\u20ac]?\s*([\d.]+,\d{2})", text, re.IGNORECASE | re.DOTALL)
        if m:
            return _parse_amount(m.group(2))
        # Totale con IVA
        m = re.search(r"[€\u20ac]\s*([\d.]+,\d{2})\s*$", text, re.MULTILINE)
        if m:
            return _parse_amount(m.group(1))

    # ── Spurgo / idraulico ──
    if "idraulico" in sl or "manutenzione" in sl:
        # "Totale E 236,68" oppure "Totale € 236,68"
        m = re.search(r"Totale\s+[E€\u20ac]\s*([\d.]+,\d{2})", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))
        # Riga dettaglio: "Imponibile IVA Importo scadenza Data scadenza\n194,00 0022 42,68 236,68 31/01/2026"
        # Prendo il secondo importo (totale con IVA) prima della data scadenza
        m = re.search(r"[\d,]+\s+\d{4}\s+[\d,]+\s+([\d.]+,\d{2})\s+\d{2}/\d{2}/\d{4}", text)
        if m:
            return _parse_amount(m.group(1))

    # ── Imposta di registro ──
    if "agenzia entrate" in sl or "registro" in sl:
        m = re.search(r"[Ii]mporto\s+addebitato.*?(?:Euro|EUR|€)\s*([\d.]+,\d{2})", text, re.DOTALL)
        if m:
            return _parse_amount(m.group(1))
        m = re.search(r"pari\s+a\s+Euro\s+([\d.]+,\d{2})", text, re.IGNORECASE)
        if m:
            return _parse_amount(m.group(1))

    # ── Fallback generico ──
    generic_patterns = [
        r"Totale\s+da\s+pagare\s*[:\s]+([\d.]+,\d{2})\s*[€\u20ac]?",
        r"TOTALE\s+DA\s+PAGARE\s*[:\s]+([\d.]+,\d{2})",
        r"Totale\s+[Bb]olletta\s*[:\s]+([\d.]+,\d{2})",
        r"Totale\s+[Ff]attura\s*[:\s]+([\d.]+,\d{2})",
        r"Totale\s+[Ff]ornitura\s+([\d.]+,\d{2})",
        r"[Nn]etto\s+a\s+pagare\s+([\d.]+,\d{2})",
        r"[Ii]mporto\s+[Tt]otale\s*[:\s]+([\d.]+,\d{2})",
        r"Totale\s+E\s+([\d.]+,\d{2})",           # Spurgo: "Totale E 236,68"
        r"[€\u20ac]\s*([\d.]+,\d{2})\b",
    ]
    for pattern in generic_patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            try:
                return _parse_amount(m.group(1))
            except ValueError:
                continue

    return 0.0


# ─── Estrazione data ─────────────────────────────────────────────────────────

def _extract_date(text: str, supplier: str = "") -> Optional[date]:
    """Estrae la data della fattura/bolletta."""

    sl = supplier.lower()

    # Pattern per Spurgo: "Fattura di Vendita 1.302/ 1 31/12/2025" → prende la data nella riga
    if "idraulico" in sl or "manutenzione" in sl:
        m = re.search(r"Fattura\s+di\s+Vendita\s+[\d./\s]+\s+(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
        if m:
            try:
                return datetime.strptime(m.group(1), "%d/%m/%Y").date()
            except ValueError:
                pass

    # Pattern per Gritti: "DATA\n10/03/2025"
    if "gritti" in sl:
        m = re.search(r"^DATA\s*\n(\d{2}/\d{2}/\d{4})", text, re.MULTILINE)
        if m:
            try:
                return datetime.strptime(m.group(1), "%d/%m/%Y").date()
            except ValueError:
                pass

    # Pattern per Vodafone: "emesso il\n06 aprile 2025"
    if "vodafone" in sl:
        m = re.search(r"emesso\s+il\s+(\d{1,2})\s+(\w+)\s+(\d{4})", text, re.IGNORECASE)
        if m:
            months = {
                "gennaio": 1, "febbraio": 2, "marzo": 3, "aprile": 4,
                "maggio": 5, "giugno": 6, "luglio": 7, "agosto": 8,
                "settembre": 9, "ottobre": 10, "novembre": 11, "dicembre": 12,
            }
            try:
                d = int(m.group(1))
                month = months.get(m.group(2).lower())
                y = int(m.group(3))
                if month:
                    return date(y, month, d)
            except ValueError:
                pass

    # Pattern per 1KHome: "FATTURA nr. 442/2025 del 04/09/2025"
    m = re.search(r"FATTURA\s+nr\.?\s+\S+\s+del\s+(\d{2}/\d{2}/\d{4})", text, re.IGNORECASE)
    if m:
        try:
            return datetime.strptime(m.group(1), "%d/%m/%Y").date()
        except ValueError:
            pass

    # Acque Veronesi: "del 23/04/2025" (prima occorrenza)
    # Reshma: "del 12/02/2025"
    priority_patterns = [
        r"[Dd]ata\s+di\s+emissione\s*[:\s]+(\d{2}/\d{2}/\d{4})",
        r"[Dd]ata\s+[Ff]attura\s*[:\s]+(\d{2}/\d{2}/\d{4})",
        r"[Dd]ata\s+emissione\s*[:\s]+(\d{2}/\d{2}/\d{4})",
        r"fattura.*?del\s+(\d{2}/\d{2}/\d{4})",
        r"Bolletta.*?del\s+(\d{2}/\d{2}/\d{4})",
        r"[Ff]attura\s+N(?:r)?\.?\s+\S+\s+del\s+(\d{2}/\d{2}/\d{4})",
        r"del\s+(\d{2}/\d{2}/\d{4})",
        # Gritti ha "DATA\n10/03/2025" già gestito sopra, ma anche inline
        r"DATA\s+(\d{2}/\d{2}/\d{4})",
        # Enel: "fattura elettronica n. XXXXX del 10/03/2025"
        r"(?:fattura elettronica|bolletta sintetica).*?del\s+(\d{2}/\d{2}/\d{4})",
        # Vivi: "Data di emissione: 01/09/2025"
        r"Data\s+di\s+emissione[:\s]+(\d{2}/\d{2}/\d{4})",
        # Spurgo: "31/12/2025" nel campo data documento
        r"Fattura\s+di\s+Vendita\s+[\d./]+\s+(\d{2}/\d{2}/\d{4})",
    ]
    for pattern in priority_patterns:
        m = re.search(pattern, text, re.IGNORECASE)
        if m:
            try:
                return datetime.strptime(m.group(1), "%d/%m/%Y").date()
            except ValueError:
                continue

    # Fallback: prendi la data più recente nel testo (escludo date lontane)
    all_dates = re.findall(r"\b(\d{2}/\d{2}/\d{4})\b", text)
    parsed = []
    for d in all_dates:
        try:
            dt = datetime.strptime(d, "%d/%m/%Y").date()
            if 2020 <= dt.year <= 2030:  # range ragionevole
                parsed.append(dt)
        except ValueError:
            continue
    if parsed:
        return max(parsed)

    return None


# ─── Estrazione numero fattura ───────────────────────────────────────────────

def _extract_invoice_number(text: str, supplier: str = "") -> str:
    """Estrae il numero fattura."""
    sl = supplier.lower()

    # 1KHome: "FATTURA nr. 442/2025"
    if "1khome" in sl or "1k home" in sl:
        m = re.search(r"FATTURA\s+nr\.?\s+(\S+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Reshma: "nr. FPR 33/25"
    if "reshma" in sl or "jeebun" in sl:
        m = re.search(r"nr\.?\s+(FPR\s+\S+|\d+/\d+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip().replace(" ", "")

    # Gritti: "NUMERO BOLLETTA\n2025E0178959"
    if "gritti" in sl:
        m = re.search(r"NUMERO\s+BOLLETTA\s*\n(\S+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Enel: "fattura elettronica n. 5220394653"
    if "enel" in sl:
        m = re.search(r"fattura\s+elettronica\s+n\.?\s+([\d]+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Vivi: numero fattura elettronica
    if "vivi" in sl:
        m = re.search(r"Numero\s+Fattura\s+Elettronica.*?n\.\s*\n?(\S+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()
        m = re.search(r"Fattura\s+del.*?(\d{15,20})", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Acque Veronesi: "Fattura N. 0001/2025/00462702"
    if "acque" in sl:
        m = re.search(r"Fattura\s+N\.?\s+([\d/]+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Vodafone: numero conto "TF68711466"
    if "vodafone" in sl:
        m = re.search(r"\b(TF\d{8,})\b", text)
        if m:
            return m.group(1)

    # Spurgo: "Fattura di Vendita 1.302/ 1 31/12/2025"  → "1.302/1"
    if "idraulico" in sl or "manutenzione" in sl:
        m = re.search(r"Fattura\s+di\s+Vendita\s+([\d.]+)\s*/\s*(\d+)\s+\d{2}/\d{2}/\d{4}", text, re.IGNORECASE)
        if m:
            return f"{m.group(1)}/{m.group(2)}"
        m = re.search(r"Fattura\s+di\s+Vendita\s+([\d./]+)", text, re.IGNORECASE)
        if m:
            return m.group(1).strip()

    # Generico
    generic = [
        r"FATTURA\s+NR\.?\s+(\S+)",
        r"[Ff]attura\s+[Nn]\.?\s+(\S+)",
        r"[Ff]attura\s+[Nn]r\.?\s+(\S+)",
        r"[Bb]olletta\s+[Nn]\.?\s+(\S+)",
        r"[Nn]umero\s+[Ff]attura\s*[:\s]+(\S+)",
        r"[Nn]\.?\s*(\d{4}/\d{4}/\d+)",
    ]
    for pattern in generic:
        m = re.search(pattern, text)
        if m:
            return m.group(1).strip()

    return ""


# ─── Rilevamento proprietà ───────────────────────────────────────────────────

def _detect_property(text: str) -> int:
    """
    Rileva la proprietà dall'indirizzo nel testo PDF.
    VIA FONTE 5 → Caldiero 5, VIA FONTE 7 → Caldiero 7
    Restituisce 0 se non rilevato (richiede intervento utente).

    Usa sezioni specifiche dove disponibile per evitare falsi positivi
    (es. Acque Veronesi ha "SACCOMANI CHIARA VIA FONTE 7" come intestatario/corrispondenza
    anche nelle bollette di VIA FONTE 5).
    """
    tu = text.upper()

    # ── Cerca nella sezione "Indirizzo di fornitura" (Acque Veronesi, Gritti, Enel) ──
    # Estraiamo il blocco dopo "Indirizzo di fornitura:" o "INDIRIZZO PUNTO DI PRELIEVO"
    # o "INDIRIZZO PUNTO DI RICONSEGNA" per leggere solo l'indirizzo reale
    fornitura_patterns = [
        r"[Ii]ndirizzo\s+di\s+fornitura[:\s]+([^\n]+)",
        r"INDIRIZZO\s+PUNTO\s+DI\s+(?:PRELIEVO|RICONSEGNA)[:\s\(][^\)]*\)[^\n]*\n([^\n]+)",
        r"La\s+tua\s+fornitura\s+di\s+(?:energia|gas)[^\n]+\n([^\n]+)",  # Enel
        r"VIA\s+FONTE\s+(\d)\s*\n.*?CALDIERO",  # pattern diretto
    ]
    for pat in fornitura_patterns:
        m = re.search(pat, text, re.IGNORECASE)
        if m:
            addr = m.group(1).upper()
            if "FONTE 5" in addr or "FONTE,5" in addr:
                return 5
            if "FONTE 7" in addr or "FONTE,7" in addr:
                return 7

    # ── Cerca l'indirizzo dell'intestatario nell'header (Gritti, Vivi: standalone) ──
    # "VIA FONTE 7\n37042 CALDIERO VR"
    m = re.search(r"VIA\s+FONTE\s+(\d)\s*\n\s*3704[12]\s+CALDIERO", tu)
    if m:
        return int(m.group(1))

    # ── Vodafone: "Indirizzo: VIA FONTE 7, CALDIERO" ──
    m = re.search(r"Indirizzo[:\s]+VIA\s+FONTE\s+(\d)", text, re.IGNORECASE)
    if m:
        return int(m.group(1))

    # ── Spurgo: "VIA FONTE 7\n37042 CALDIERO" nel destinatario ──
    m = re.search(r"VIA\s+FONTE\s+(\d)\b", tu)
    if m:
        prop = int(m.group(1))
        if prop in (5, 7):
            return prop

    return 0  # Non rilevato → richiede intervento utente


def _covers_both_properties(text: str, supplier: str) -> bool:
    """Verifica se la fattura copre entrambe le proprietà."""

    # Reshma fattura sempre entrambe le proprietà
    if "reshma" in supplier.lower() or "jeebun" in supplier.lower():
        return True

    # Per Acque Veronesi/Gritti/Enel/Vivi: una bolletta = una fornitura (un indirizzo)
    # Non facciamo split per bollette utenze singole
    sl = supplier.lower()
    if any(k in sl for k in ["acque", "enel", "gritti", "vivi", "vodafone", "idraulico"]):
        return False

    # 1KHome con VIA FONTE 5 e VIA FONTE 7 nella descrizione delle voci
    if "1khome" in sl or "1k home" in sl:
        # Le fatture "rimborso bollette" NON vengono splittate
        # (sono intestate a Massimo Prezioso, non alla gestione turistica)
        if "RIMBORSO SPESE" in text.upper() or "RIMBORSO BOLLETTE" in text.upper():
            return False

        tu = text.upper()

        # Cerca VIA FONTE 5 e VIA FONTE 7 entrambi nel testo
        has_5 = bool(re.search(r"VIA\s+FONTE\s*5\b", tu))
        has_7 = bool(re.search(r"VIA\s+FONTE\s*7\b", tu))
        if has_5 and has_7:
            return True

        # Fallback: se ci sono 2 voci separate di "GESTIONE LOCAZIONE TURISTICA"
        # ciascuna con un proprio importo (pdfplumber a volte non estrae "VIA FONTE 7")
        # Pattern: "GESTIONE LOCAZIONE TURISTICA\n€ 200,00\n..." ripetuto DUE VOLTE
        # con importi separati (la fattura singola ha UNA voce con descrizione lunga)
        importo_matches = re.findall(
            r"GESTIONE\s+LOCAZIONE\s+TURISTICA[^\n]*\n[^\n]*?\d+[,\.]\d{2}", tu
        )
        if len(importo_matches) >= 2:
            return True

        # Seconda euristica: l'importo nella sezione SCADENZE è il doppio di quello
        # nella sezione RIEPILOGO IVA per la stessa aliquota
        # (fattura singola: 200€ imponibile; fattura doppia: 400€ imponibile)
        imponibile_match = re.search(r"IMPONIBILE\s+[€\u20ac]\s*([\d.,]+)", tu)
        scadenza_match = re.search(r"SCADENZE.*?[€\u20ac]\s*([\d.,]+)", tu, re.DOTALL)
        if imponibile_match and scadenza_match:
            try:
                imponibile = float(imponibile_match.group(1).replace(".", "").replace(",", "."))
                scadenza = float(scadenza_match.group(1).replace(".", "").replace(",", "."))
                # Se l'imponibile è >= 400 (due voci da 200 ciascuna), è fattura doppia
                if imponibile >= 350:
                    return True
            except ValueError:
                pass

        return False

    return False


# ─── Parsing principale ──────────────────────────────────────────────────────

def parse_pdf_invoice(filepath: str) -> List[Cost]:
    """
    Estrae dati da una fattura PDF.

    Restituisce lista di Cost:
    - Normalmente 1 elemento
    - 2 elementi per fatture che coprono entrambe le proprietà (Reshma, 1KHome 5+7)
    - Raise ValueError se il PDF è vuoto/scansionato o l'importo non è trovato

    Se la proprietà è 0 (ambigua), la Cost viene restituita con property_num=0
    e il chiamante (app.py) deve chiedere all'utente.
    """
    try:
        import pdfplumber
    except ImportError:
        raise ImportError("Installa pdfplumber: pip install pdfplumber")

    try:
        with pdfplumber.open(filepath) as pdf:
            text = "\n".join(page.extract_text() or "" for page in pdf.pages)
    except Exception as e:
        raise ValueError(f"Errore lettura PDF: {e}")

    if not text.strip():
        fname = os.path.basename(filepath)
        raise ValueError(
            f"PDF non leggibile (probabilmente scansionato): {fname}. "
            "Inserisci i dati manualmente."
        )

    supplier, category = _detect_supplier(text, filepath)
    amount = _extract_amount(text, supplier)
    invoice_date = _extract_date(text, supplier)
    invoice_num = _extract_invoice_number(text, supplier)

    if amount == 0.0:
        fname = os.path.basename(filepath)
        raise ValueError(
            f"Importo non trovato in '{fname}' (fornitore: {supplier}). "
            "Verifica il PDF o inserisci i dati manualmente."
        )

    cost_date = invoice_date or date.today()

    # ── Fattura per entrambe le proprietà → split 50/50 ──
    if _covers_both_properties(text, supplier):
        half = round(amount / 2, 2)
        return [
            Cost(
                property_num=5,
                date=cost_date,
                amount=-half,
                category=category,
                supplier=supplier,
                invoice_num=invoice_num,
                invoice_date=invoice_date,
                source_file=os.path.basename(filepath),
            ),
            Cost(
                property_num=7,
                date=cost_date,
                amount=-half,
                category=category,
                supplier=supplier,
                invoice_num=invoice_num,
                invoice_date=invoice_date,
                source_file=os.path.basename(filepath),
            ),
        ]

    # ── Fattura per singola proprietà ──
    property_num = _detect_property(text)

    # Fallback dal nome file se property_num è 0
    if property_num == 0:
        fname_lower = os.path.basename(filepath).lower()
        if "caldiero 5" in fname_lower or "caldiero5" in fname_lower:
            property_num = 5
        elif "caldiero 7" in fname_lower or "caldiero7" in fname_lower:
            property_num = 7

    return [
        Cost(
            property_num=property_num,  # 0 = ambiguo, richiede intervento utente
            date=cost_date,
            amount=-amount,
            category=category,
            supplier=supplier,
            invoice_num=invoice_num,
            invoice_date=invoice_date,
            source_file=os.path.basename(filepath),
        )
    ]
