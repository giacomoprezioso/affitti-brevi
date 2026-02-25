# Setup: Affitti Brevi su Streamlit Cloud + Google Sheets

## 1. Crea il Google Sheet

1. Vai su [sheets.google.com](https://sheets.google.com) e crea un nuovo foglio
2. Rinominalo "Affitti Brevi 2026"
3. Crea questi fogli (tab in basso):
   - `elenco`
4. Nella riga 1 del foglio `elenco`, incolla queste intestazioni:
   ```
   caldiero	dal	al	mese	tax	importo	tipo	causale	ente	nominativo	documento	nr	data	periodo	intestata_a	giorni	inviato_1k	ritenuta	incassato	lordo	commission	payment_charge	vat	euro_gg	piattaforma_raw	source_file
   ```
5. Copia l'ID del foglio dalla URL:
   `https://docs.google.com/spreadsheets/d/`**`QUESTO_È_L_ID`**`/edit`

## 2. Crea Service Account Google

1. Vai su [console.cloud.google.com](https://console.cloud.google.com)
2. Crea nuovo progetto (o usa uno esistente) → chiama "affitti-brevi"
3. Attiva **Google Sheets API**:
   - Menu → API e servizi → Libreria → cerca "Google Sheets API" → Attiva
4. Crea Service Account:
   - Menu → API e servizi → Credenziali → Crea credenziali → Account di servizio
   - Nome: `affitti-brevi`
   - Salta i passaggi 2 e 3
5. Clicca sull'account creato → tab **Chiavi** → Aggiungi chiave → JSON
   - Scarica il file JSON (es. `affitti-brevi-abc123.json`)
6. **Condividi il Google Sheet** con l'email del service account:
   - Apri il file JSON → copia il campo `client_email` (es. `affitti-brevi@...iam.gserviceaccount.com`)
   - Nel Google Sheet → Condividi → incolla l'email → Editor

## 3. Configura secrets localmente (per test)

Crea `.streamlit/secrets.toml` (NON committare su Git!):

```toml
[google_sheets]
spreadsheet_id = "IL_TUO_ID_SHEET"

[gcp_service_account]
type = "service_account"
project_id = "..."
private_key_id = "..."
private_key = "-----BEGIN RSA PRIVATE KEY-----\n...\n-----END RSA PRIVATE KEY-----\n"
client_email = "affitti-brevi@...iam.gserviceaccount.com"
client_id = "..."
auth_uri = "https://accounts.google.com/o/oauth2/auth"
token_uri = "https://oauth2.googleapis.com/token"
auth_provider_x509_cert_url = "https://www.googleapis.com/oauth2/v1/certs"
client_x509_cert_url = "..."
universe_domain = "googleapis.com"
```

Tutti i valori li trovi nel file JSON scaricato al passo 2.5.

## 4. Deploy su GitHub + Streamlit Cloud

```bash
cd "C:\Users\giaco\claude code\affitti-brevi"
git init
git add .
git commit -m "Initial commit"
git remote add origin https://github.com/TUO_USERNAME/affitti-brevi.git
git push -u origin main
```

Poi:
1. Vai su [share.streamlit.io](https://share.streamlit.io)
2. Connetti il tuo GitHub → seleziona repo `affitti-brevi`
3. File principale: `app.py`
4. **Secrets**: vai su Settings → Secrets → incolla il contenuto di `secrets.toml`
5. Deploy!

## 5. Tuo padre

Manda il link dell'app (es. `https://affitti-brevi.streamlit.app`).
Nessun install, funziona da browser su qualsiasi dispositivo.
