"""
Affitti Brevi - Gestione automatica Caldiero 5 & 7
Web app su Streamlit Community Cloud con Google Sheets come storage.
"""

import streamlit as st
import pandas as pd
import io
import os
import sys
import tempfile

sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

from parsers.airbnb import parse_airbnb_csv
from parsers.booking import parse_booking_xlsx
from parsers.pdf_invoice import parse_pdf_invoice
from core.sheets import save_to_sheets, load_elenco_from_sheets
from core.models import Cost

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


# ── Export helper ────────────────────────────────────────────────────────────
def df_to_excel_bytes(df: pd.DataFrame) -> bytes:
    """Converte DataFrame in bytes XLSX per il download."""
    buf = io.BytesIO()
    with pd.ExcelWriter(buf, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name="Dati")
    return buf.getvalue()


# ── Tabs ─────────────────────────────────────────────────────────────────────
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
        # Costi con proprietà ambigua (=0) da risolvere con l'utente
        ambiguous_costs: list[Cost] = []

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
                        # Separa costi ambigui (property_num=0) da quelli chiari
                        clear_costs = [c for c in costs if c.property_num != 0]
                        amb_costs = [c for c in costs if c.property_num == 0]
                        all_costs.extend(clear_costs)
                        ambiguous_costs.extend(amb_costs)

                        desc_parts = []
                        for c in clear_costs:
                            desc_parts.append(f"Caldiero {c.property_num}: €{abs(c.amount):.2f}")
                        for c in amb_costs:
                            desc_parts.append(f"⚠️ Proprietà sconosciuta: €{abs(c.amount):.2f}")
                        parse_results.append({
                            "File": uploaded_file.name,
                            "Tipo": "Fattura PDF",
                            "Trovati": " | ".join(desc_parts) if desc_parts else "—",
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

        # ── Risoluzione proprietà ambigua ──────────────────────────────────
        if ambiguous_costs:
            st.subheader("⚠️ Proprietà non riconosciuta — assegna manualmente")
            st.info(
                "Per alcune fatture non è stato possibile rilevare automaticamente "
                "a quale proprietà appartengono. Seleziona l'appartamento corretto."
            )
            for idx, cost in enumerate(ambiguous_costs):
                col1, col2 = st.columns([3, 1])
                with col1:
                    st.markdown(
                        f"**{cost.supplier}** — {cost.category} — "
                        f"€{abs(cost.amount):.2f} — {cost.date} — "
                        f"Fattura: {cost.invoice_num or '—'} — "
                        f"File: {cost.source_file}"
                    )
                with col2:
                    prop_choice = st.selectbox(
                        "Appartamento",
                        options=[5, 7],
                        key=f"ambig_prop_{idx}",
                        format_func=lambda x: f"Caldiero {x}",
                    )
                # Aggiorna la proprietà
                ambiguous_costs[idx] = Cost(
                    property_num=prop_choice,
                    date=cost.date,
                    amount=cost.amount,
                    category=cost.category,
                    supplier=cost.supplier,
                    invoice_num=cost.invoice_num,
                    invoice_date=cost.invoice_date,
                    source_file=cost.source_file,
                )
            # Aggiungi i costi risolti alla lista
            all_costs.extend(ambiguous_costs)

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
                    "Proprietà": f"Caldiero {c.property_num}" if c.property_num else "❓",
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
                # Filtra solo righe prenotazioni
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
                    df_b["mese_label"] = df_b["check_in"].dt.strftime("%b %Y").fillna("N/D")
                else:
                    df_b["anno_mese"] = "N/D"
                    df_b["mese_label"] = "N/D"

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

                    # ── Pivot Lordo per mese × proprietà ──
                    if "lordo" in df_f.columns and not df_f.empty:
                        st.subheader("Lordo per mese e proprietà (€)")
                        pivot_lordo = df_f.pivot_table(
                            values="lordo", index="anno_mese", columns="proprieta",
                            aggfunc="sum", fill_value=0, margins=True, margins_name="TOTALE"
                        )
                        st.dataframe(pivot_lordo.round(2), use_container_width=True)

                    st.divider()

                    # ── Pivot Netto per mese × proprietà ──
                    if "incassato" in df_f.columns and not df_f.empty:
                        st.subheader("Netto per mese e proprietà (€)")
                        pivot_netto = df_f.pivot_table(
                            values="incassato", index="anno_mese", columns="proprieta",
                            aggfunc="sum", fill_value=0, margins=True, margins_name="TOTALE"
                        )
                        st.dataframe(pivot_netto.round(2), use_container_width=True)

                    st.divider()

                    # ── Per piattaforma ──
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

                    # ── Elenco completo ──
                    st.subheader("Elenco prenotazioni")
                    cols_show = [c for c in [
                        "anno_mese", "proprieta", col_ente, "nominativo",
                        "dal", "al", "giorni", "lordo", "commission", "incassato", "nr"
                    ] if c in df_f.columns]
                    df_show = df_f[cols_show].sort_values("anno_mese", ascending=False)
                    st.dataframe(df_show, use_container_width=True, hide_index=True)

                    # ── Export ──
                    st.divider()
                    st.subheader("Esporta")
                    col_csv, col_xlsx = st.columns(2)
                    with col_csv:
                        csv = df_show.to_csv(index=False).encode("utf-8")
                        st.download_button(
                            "⬇️ Scarica CSV",
                            csv,
                            file_name=f"prenotazioni_{sel_anno}.csv",
                            mime="text/csv",
                        )
                    with col_xlsx:
                        xlsx = df_to_excel_bytes(df_show)
                        st.download_button(
                            "⬇️ Scarica Excel",
                            xlsx,
                            file_name=f"prenotazioni_{sel_anno}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
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

                # Calcola colonne utili
                if "dal" in df_c.columns:
                    df_c["data_dt"] = pd.to_datetime(df_c["dal"], errors="coerce")
                    df_c["anno_mese"] = df_c["data_dt"].dt.strftime("%Y-%m").fillna("N/D")
                    df_c["mese_label"] = df_c["data_dt"].dt.strftime("%b %Y").fillna("N/D")
                else:
                    df_c["anno_mese"] = "N/D"
                    df_c["mese_label"] = "N/D"

                if "caldiero" in df_c.columns:
                    df_c["proprieta"] = df_c["caldiero"].apply(
                        lambda x: f"Caldiero {int(float(x))}" if str(x) not in ("", "nan") else "N/D"
                    )

                if df_c.empty:
                    st.info("Nessun costo trovato.")
                else:
                    # Filtri
                    col1, col2, col3 = st.columns(3)
                    with col1:
                        anni_c = sorted(df_c["anno_mese"].str[:4].unique().tolist(), reverse=True)
                        anni_c = ["Tutti"] + [a for a in anni_c if a != "N/D"]
                        sel_anno_c = st.selectbox("Anno", anni_c, key="costs_anno")
                    with col2:
                        if "proprieta" in df_c.columns:
                            props_c = ["Tutte"] + sorted(df_c["proprieta"].unique().tolist())
                            sel_prop_c = st.selectbox("Proprietà", props_c, key="costs_prop")
                        else:
                            sel_prop_c = "Tutte"
                    with col3:
                        if "causale" in df_c.columns:
                            cats = ["Tutte"] + sorted(df_c["causale"].astype(str).unique().tolist())
                            sel_cat = st.selectbox("Categoria", cats)
                        else:
                            sel_cat = "Tutte"

                    df_fc = df_c.copy()
                    if sel_anno_c != "Tutti":
                        df_fc = df_fc[df_fc["anno_mese"].str.startswith(sel_anno_c)]
                    if "proprieta" in df_fc.columns and sel_prop_c != "Tutte":
                        df_fc = df_fc[df_fc["proprieta"] == sel_prop_c]
                    if "causale" in df_fc.columns and sel_cat != "Tutte":
                        df_fc = df_fc[df_fc["causale"].astype(str) == sel_cat]

                    # KPI
                    k1, k2, k3 = st.columns(3)
                    k1.metric("Costi totali €", f"{df_fc['importo'].sum():.2f}")
                    k2.metric("Voci", len(df_fc))
                    if "proprieta" in df_fc.columns:
                        n_props = df_fc["proprieta"].nunique()
                        k3.metric("Proprietà coinvolte", n_props)

                    st.divider()

                    # ── Pivot mese × proprietà ──
                    if "proprieta" in df_fc.columns and "anno_mese" in df_fc.columns and not df_fc.empty:
                        st.subheader("Spese per mese e proprietà (€)")
                        pivot_c = df_fc.pivot_table(
                            values="importo", index="anno_mese", columns="proprieta",
                            aggfunc="sum", fill_value=0, margins=True, margins_name="TOTALE"
                        )
                        st.dataframe(pivot_c.round(2), use_container_width=True)

                    st.divider()

                    # ── Per categoria ──
                    if "causale" in df_fc.columns and "ente" in df_fc.columns:
                        st.subheader("Per categoria e fornitore")
                        summary = df_fc.groupby(["causale", "ente"])["importo"].sum().reset_index()
                        summary = summary.sort_values("importo").round(2)
                        st.dataframe(summary, use_container_width=True, hide_index=True)

                        st.subheader("Ripartizione per categoria")
                        chart_data = df_fc.groupby("causale")["importo"].sum().abs()
                        st.bar_chart(chart_data)

                    st.divider()

                    # ── Elenco ──
                    st.subheader("Elenco costi")
                    cols_c = [c for c in [
                        "proprieta", "anno_mese", "dal", "causale", "ente", "importo", "documento"
                    ] if c in df_fc.columns]
                    df_c_show = (
                        df_fc[cols_c].sort_values("anno_mese", ascending=False)
                        if "anno_mese" in df_fc.columns
                        else df_fc[cols_c]
                    )
                    st.dataframe(df_c_show, use_container_width=True, hide_index=True)

                    # ── Export ──
                    st.divider()
                    st.subheader("Esporta")
                    col_csv_c, col_xlsx_c = st.columns(2)
                    with col_csv_c:
                        csv_c = df_c_show.to_csv(index=False).encode("utf-8")
                        st.download_button(
                            "⬇️ Scarica CSV",
                            csv_c,
                            file_name=f"costi_{sel_anno_c}.csv",
                            mime="text/csv",
                        )
                    with col_xlsx_c:
                        xlsx_c = df_to_excel_bytes(df_c_show)
                        st.download_button(
                            "⬇️ Scarica Excel",
                            xlsx_c,
                            file_name=f"costi_{sel_anno_c}.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        )

        except Exception as e:
            st.error(f"Errore caricamento costi: {e}")
            st.exception(e)
