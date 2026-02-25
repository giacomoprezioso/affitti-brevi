"""
Affitti Brevi - Gestione automatica Caldiero 5 & 7
Web app su Streamlit Community Cloud con Google Sheets come storage.
"""

import streamlit as st
import pandas as pd
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parsers.airbnb import parse_airbnb_csv
from parsers.booking import parse_booking_xlsx
from parsers.pdf_invoice import parse_pdf_invoice
from core.sheets import save_to_sheets, load_elenco_from_sheets

st.set_page_config(
    page_title="Affitti Brevi - Caldiero",
    page_icon="🏠",
    layout="wide",
)

st.title("🏠 Affitti Brevi - Caldiero 5 & 7")


# ── Verifica connessione Google Sheets ──────────────────────────────────────
def check_sheets_connection() -> bool:
    try:
        _ = st.secrets["gcp_service_account"]
        _ = st.secrets["google_sheets"]["spreadsheet_id"]
        return True
    except Exception:
        return False


with st.sidebar:
    st.header("Stato connessione")
    if check_sheets_connection():
        st.success("✓ Google Sheets connesso")
    else:
        st.error("✗ Credenziali mancanti")
        st.caption("Configura `.streamlit/secrets.toml`")

    st.divider()
    st.caption("**Come esportare i file:**")
    with st.expander("Airbnb CSV"):
        st.write("Account → Transazioni → Esporta CSV")
    with st.expander("Booking XLSX"):
        st.write("Extranet → Finance → Pagamenti → Esporta")
    with st.expander("PDF Fatture"):
        st.write("Carica direttamente le fatture PDF")


# ── Tabs ────────────────────────────────────────────────────────────────────
tab_import, tab_report, tab_costs = st.tabs(["📥 Importa", "📊 Prenotazioni", "💶 Costi"])


# ============================================================
# TAB 1: IMPORTA
# ============================================================
with tab_import:
    st.header("Importa nuovi dati")
    st.write("Carica uno o più file — il sistema riconosce il tipo automaticamente.")

    uploaded_files = st.file_uploader(
        "Trascina qui i file o clicca per selezionare",
        accept_multiple_files=True,
        type=["csv", "xlsx", "xls", "pdf"],
        help="Airbnb → .csv  |  Booking → .xlsx  |  Fatture → .pdf",
    )

    if uploaded_files:
        all_bookings = []
        all_costs = []
        errors = []
        parse_results = []

        with st.spinner("Lettura file in corso..."):
            for uploaded_file in uploaded_files:
                fname = uploaded_file.name.lower()
                suffix = os.path.splitext(fname)[1]

                with tempfile.NamedTemporaryFile(delete=False, suffix=suffix) as tmp:
                    tmp.write(uploaded_file.getbuffer())
                    tmp_path = tmp.name

                try:
                    if suffix == ".csv":
                        bookings = parse_airbnb_csv(tmp_path)
                        all_bookings.extend(bookings)
                        parse_results.append({
                            "File": uploaded_file.name,
                            "Tipo": "Airbnb CSV",
                            "Trovati": f"{len(bookings)} prenotazioni",
                        })

                    elif suffix in (".xlsx", ".xls"):
                        bookings = parse_booking_xlsx(tmp_path)
                        all_bookings.extend(bookings)
                        parse_results.append({
                            "File": uploaded_file.name,
                            "Tipo": "Booking XLSX",
                            "Trovati": f"{len(bookings)} prenotazioni",
                        })

                    elif suffix == ".pdf":
                        costs = parse_pdf_invoice(tmp_path)
                        all_costs.extend(costs)
                        desc = " | ".join(
                            f"Caldiero {c.property_num}: €{abs(c.amount):.2f}" for c in costs
                        )
                        parse_results.append({
                            "File": uploaded_file.name,
                            "Tipo": "Fattura PDF",
                            "Trovati": desc,
                        })

                except Exception as e:
                    errors.append(f"**{uploaded_file.name}**: {e}")
                finally:
                    try:
                        os.unlink(tmp_path)
                    except Exception:
                        pass

        # Riepilogo parsing
        if parse_results:
            st.subheader("File riconosciuti")
            st.dataframe(pd.DataFrame(parse_results), use_container_width=True, hide_index=True)

        if errors:
            for e in errors:
                st.error(e)

        # Anteprima prenotazioni
        if all_bookings:
            st.subheader(f"Prenotazioni ({len(all_bookings)})")
            unknown = [b for b in all_bookings if b.property_num == 0]
            if unknown:
                st.warning(f"⚠️ {len(unknown)} prenotazioni con proprietà non riconosciuta — controlla `config.py`.")

            preview = []
            for b in all_bookings:
                preview.append({
                    "Piattaforma": b.platform.capitalize(),
                    "Proprietà": f"Caldiero {b.property_num}" if b.property_num else "❓",
                    "Ospite": b.guest_name,
                    "Check-in": b.check_in.strftime("%d/%m/%Y") if b.check_in else "—",
                    "Check-out": b.check_out.strftime("%d/%m/%Y") if b.check_out else "—",
                    "Notti": b.nights,
                    "Lordo €": f"{b.gross_amount:.2f}",
                    "Netto €": f"{b.net_amount:.2f}",
                    "Ritenuta €": f"{b.withholding_tax:.2f}" if b.withholding_tax else "—",
                    "Codice": b.confirm_code,
                })
            st.dataframe(pd.DataFrame(preview), use_container_width=True, hide_index=True)

        # Anteprima costi
        if all_costs:
            st.subheader(f"Costi ({len(all_costs)})")
            costs_preview = []
            for c in all_costs:
                costs_preview.append({
                    "Proprietà": f"Caldiero {c.property_num}" if c.property_num else "Entrambe",
                    "Data": c.date.strftime("%d/%m/%Y") if c.date else "—",
                    "Fornitore": c.supplier,
                    "Categoria": c.category,
                    "Importo €": f"{c.amount:.2f}",
                    "N. Fattura": c.invoice_num or "—",
                })
            st.dataframe(pd.DataFrame(costs_preview), use_container_width=True, hide_index=True)

        # Conferma importazione
        if all_bookings or all_costs:
            st.divider()
            col1, col2 = st.columns([1, 4])
            with col1:
                dry_run = st.checkbox("Dry run", value=False, help="Anteprima senza salvare")
            with col2:
                if st.button("✅ Salva su Google Sheets", type="primary"):
                    with st.spinner("Salvataggio in corso..."):
                        try:
                            added_b, added_c, skipped = save_to_sheets(
                                all_bookings, all_costs, dry_run=dry_run
                            )
                            if dry_run:
                                st.info(
                                    f"**Dry run:** verrebbero salvate **{added_b}** prenotazioni "
                                    f"e **{added_c}** costi. Saltati (già presenti): {len(skipped)}."
                                )
                            else:
                                st.success(
                                    f"✓ Salvato: **{added_b}** prenotazioni, **{added_c}** costi. "
                                    f"Saltati (già presenti): {len(skipped)}."
                                )
                                if skipped:
                                    with st.expander("Codici saltati"):
                                        st.write(", ".join(str(s) for s in skipped))
                                st.cache_data.clear()
                        except Exception as e:
                            st.error(f"Errore: {e}")


# ============================================================
# TAB 2: REPORT PRENOTAZIONI
# ============================================================
with tab_report:
    st.header("Prenotazioni")

    if not check_sheets_connection():
        st.warning("Connessione Google Sheets non configurata.")
    else:
        try:
            with st.spinner("Caricamento dati..."):
                df = load_elenco_from_sheets()

            if df.empty:
                st.info("Nessun dato nel foglio. Importa i file dal tab Importa.")
            else:
                # Filtra solo righe prenotazioni (importo positivo o tipo incasso)
                df_b = df[
                    df["tipo"].astype(str).str.lower().str.contains("incasso", na=False) |
                    df["causale"].astype(str).str.lower().str.contains("da clienti", na=False)
                ].copy()

                # Tipizzazione
                for col in ["importo", "lordo", "incassato", "ritenuta", "commission", "giorni"]:
                    if col in df_b.columns:
                        df_b[col] = pd.to_numeric(df_b[col], errors="coerce").fillna(0)

                if "dal" in df_b.columns:
                    df_b["check_in"] = pd.to_datetime(df_b["dal"], errors="coerce")
                    df_b["anno_mese"] = df_b["check_in"].dt.strftime("%Y-%m").fillna("N/D")
                else:
                    df_b["anno_mese"] = "N/D"

                if "caldiero" in df_b.columns:
                    df_b["proprieta"] = df_b["caldiero"].apply(
                        lambda x: f"Caldiero {int(float(x))}" if str(x) not in ("", "nan", "None") else "N/D"
                    )

                if df_b.empty:
                    st.info("Nessuna prenotazione trovata.")
                else:
                    # Filtri
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        anni = sorted(df_b["anno_mese"].str[:4].unique().tolist(), reverse=True)
                        anni = ["Tutti"] + [a for a in anni if a != "N/D"]
                        sel_anno = st.selectbox("Anno", anni)
                    with col2:
                        props = ["Tutte"] + sorted(df_b["proprieta"].unique().tolist())
                        sel_prop = st.selectbox("Proprietà", props)
                    with col3:
                        col_ente = "ente" if "ente" in df_b.columns else "piattaforma_raw"
                        platforms = ["Tutte"] + sorted(df_b[col_ente].astype(str).unique().tolist())
                        sel_platform = st.selectbox("Piattaforma", platforms)

                    df_f = df_b.copy()
                    if sel_anno != "Tutti":
                        df_f = df_f[df_f["anno_mese"].str.startswith(sel_anno)]
                    if sel_prop != "Tutte":
                        df_f = df_f[df_f["proprieta"] == sel_prop]
                    if sel_platform != "Tutte":
                        df_f = df_f[df_f[col_ente].astype(str) == sel_platform]

                    # KPI
                    st.subheader("KPI")
                    k1, k2, k3, k4 = st.columns(4)
                    k1.metric("Prenotazioni", len(df_f))
                    k2.metric("Lordo totale €", f"{df_f['lordo'].sum():.2f}" if "lordo" in df_f else "—")
                    k3.metric("Netto totale €", f"{df_f['incassato'].sum():.2f}" if "incassato" in df_f else "—")
                    k4.metric("Notti totali", int(df_f["giorni"].sum()) if "giorni" in df_f else "—")

                    st.divider()

                    # Pivot mese × proprietà
                    if "lordo" in df_f.columns and not df_f.empty:
                        st.subheader("Lordo per mese e proprietà (€)")
                        pivot = df_f.pivot_table(
                            values="lordo", index="anno_mese", columns="proprieta",
                            aggfunc="sum", fill_value=0, margins=True, margins_name="TOTALE"
                        )
                        st.dataframe(pivot.round(2), use_container_width=True)

                    st.divider()

                    # Per piattaforma
                    if col_ente in df_f.columns:
                        st.subheader("Per piattaforma")
                        by_plat = df_f.groupby(col_ente).agg(
                            prenotazioni=(col_ente, "count"),
                            lordo_totale=("lordo", "sum"),
                            netto_totale=("incassato", "sum"),
                            notti=("giorni", "sum"),
                        ).reset_index().round(2)
                        st.dataframe(by_plat, use_container_width=True, hide_index=True)

                    st.divider()

                    # Elenco completo
                    st.subheader("Elenco prenotazioni")
                    cols_show = [c for c in [
                        "anno_mese", "proprieta", col_ente, "nominativo",
                        "dal", "al", "giorni", "lordo", "commission", "incassato", "nr"
                    ] if c in df_f.columns]
                    st.dataframe(
                        df_f[cols_show].sort_values("anno_mese", ascending=False),
                        use_container_width=True, hide_index=True
                    )

                    # Export CSV
                    csv = df_f.to_csv(index=False).encode("utf-8")
                    st.download_button(
                        "⬇️ Scarica CSV", csv,
                        file_name=f"prenotazioni_{sel_anno}.csv",
                        mime="text/csv"
                    )

        except Exception as e:
            st.error(f"Errore caricamento: {e}")
            st.exception(e)


# ============================================================
# TAB 3: COSTI
# ============================================================
with tab_costs:
    st.header("Costi e Spese")

    if not check_sheets_connection():
        st.warning("Connessione Google Sheets non configurata.")
    else:
        try:
            with st.spinner("Caricamento dati..."):
                df = load_elenco_from_sheets()

            if df.empty:
                st.info("Nessun dato.")
            else:
                df_c = df[
                    pd.to_numeric(df["importo"], errors="coerce").fillna(0) < 0
                ].copy()

                for col in ["importo"]:
                    df_c[col] = pd.to_numeric(df_c[col], errors="coerce").fillna(0)

                if df_c.empty:
                    st.info("Nessun costo trovato.")
                else:
                    # Filtri
                    col1, col2 = st.columns(2)
                    with col1:
                        if "caldiero" in df_c.columns:
                            df_c["proprieta"] = df_c["caldiero"].apply(
                                lambda x: f"Caldiero {int(float(x))}" if str(x) not in ("", "nan") else "N/D"
                            )
                            props = ["Tutte"] + sorted(df_c["proprieta"].unique().tolist())
                            sel_prop_c = st.selectbox("Proprietà", props, key="costs_prop")
                    with col2:
                        if "causale" in df_c.columns:
                            cats = ["Tutte"] + sorted(df_c["causale"].astype(str).unique().tolist())
                            sel_cat = st.selectbox("Categoria", cats)

                    df_fc = df_c.copy()
                    if "proprieta" in df_fc.columns and sel_prop_c != "Tutte":
                        df_fc = df_fc[df_fc["proprieta"] == sel_prop_c]
                    if "causale" in df_fc.columns and sel_cat != "Tutte":
                        df_fc = df_fc[df_fc["causale"].astype(str) == sel_cat]

                    # KPI
                    k1, k2 = st.columns(2)
                    k1.metric("Costi totali €", f"{df_fc['importo'].sum():.2f}")
                    k2.metric("Voci", len(df_fc))

                    st.divider()

                    # Riepilogo per categoria
                    if "causale" in df_fc.columns and "ente" in df_fc.columns:
                        st.subheader("Per categoria")
                        summary = df_fc.groupby(["causale", "ente"])["importo"].sum().reset_index()
                        summary = summary.sort_values("importo").round(2)
                        st.dataframe(summary, use_container_width=True, hide_index=True)

                        # Grafico
                        chart_data = df_fc.groupby("causale")["importo"].sum().abs()
                        st.bar_chart(chart_data)

                    st.divider()

                    # Elenco
                    st.subheader("Elenco costi")
                    cols_c = [c for c in ["proprieta", "dal", "causale", "ente", "importo", "documento"] if c in df_fc.columns]
                    st.dataframe(df_fc[cols_c].sort_values("dal", ascending=False) if "dal" in df_fc.columns else df_fc[cols_c],
                                 use_container_width=True, hide_index=True)

        except Exception as e:
            st.error(f"Errore caricamento costi: {e}")
            st.exception(e)
