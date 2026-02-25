"""
Configurazione centralizzata - modifica qui i percorsi e i mapping.
"""

# Percorso del file Excel principale
EXCEL_PATH = r"C:\Users\giaco\claude code\affitti brevi app\affitti brevi 2026.xlsx"

# Nome del foglio principale
SHEET_ELENCO = "elenco"

# Mapping nome annuncio Airbnb → numero proprietà (substring case-insensitive)
AIRBNB_LISTING_MAP = {
    "family retreat": 5,
    "tranquillità nel verde": 7,
    "tranquillit": 7,  # fallback per encoding broken
}

# Mapping nome struttura Booking → numero proprietà (substring case-insensitive)
BOOKING_PROPERTY_MAP = {
    "family retreat": 5,
    "tranquillit": 7,
}

# Mapping colonne Excel 1-indexed (A=1, B=2, ...)
COL_MAP = {
    "caldiero":       1,   # A - numero proprietà (5 o 7)
    "dal":            2,   # B - data inizio
    "al":             3,   # C - data fine
    "mese":           4,   # D - mese (formula =MONTH(B{row}))
    "tax":            5,   # E - T-A-X
    "importo":        6,   # F - importo netto
    "tipo":           7,   # G - tipo record
    "causale":        8,   # H - causale
    "ente":           9,   # I - piattaforma/fornitore
    "nominativo":    10,   # J - nome ospite/intestatario
    "documento":     11,   # K - tipo documento
    "nr":            12,   # L - codice prenotazione / numero fattura
    "data":          13,   # M - data operazione
    "periodo":       14,   # N - periodo riferimento
    "intestata_a":   15,   # O - intestata a
    "giorni":        16,   # P - numero notti/giorni
    "inviato_1k":    17,   # Q - inviato a 1K home
    "ritenuta":      18,   # R - ritenuta d'acconto
    "incassato":     19,   # S - importo incassato netto
    "lordo":         20,   # T - importo lordo
    "commission":    21,   # U - commissione piattaforma
    "payment_charge": 22,  # V - costo pagamento
    "vat":           23,   # W - IVA su servizi piattaforma
    # colonna 24 (X) vuota
    "euro_gg":       25,   # Y - costo al giorno
}
