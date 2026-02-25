"""
Modelli dati: Booking (prenotazione) e Cost (costo/fattura).
"""

from dataclasses import dataclass, field
from datetime import date
from typing import Optional


@dataclass
class Booking:
    """Una prenotazione da Airbnb, Booking o affitto diretto."""
    platform: str           # "airbnb" | "booking" | "diretto"
    property_num: int       # 5 = Caldiero 5, 7 = Caldiero 7
    guest_name: str
    check_in: date
    check_out: date
    nights: int
    gross_amount: float     # importo lordo (quello che paga l'ospite)
    commission: float       # commissione piattaforma (negativo o 0)
    payment_charge: float   # costo transazione pagamento (negativo o 0)
    vat: float              # IVA su servizi piattaforma (negativo o 0)
    net_amount: float       # importo netto all'host
    withholding_tax: float  # ritenuta fiscale (negativo o 0)
    confirm_code: str       # codice conferma / numero prenotazione
    source_file: str        # file di origine (traceability)


@dataclass
class Cost:
    """Una spesa: bolletta, fattura pulizie, fee gestione, ecc."""
    property_num: int       # 5 | 7 | 0 (entrambe le proprietà)
    date: date
    amount: float           # sempre negativo
    category: str           # "acqua" | "energia elettrica" | "pulizie" | ...
    supplier: str           # "Acque veronesi" | "Vivi" | "Reshma" | ...
    invoice_num: str = ""
    invoice_date: Optional[date] = None
    source_file: str = ""
