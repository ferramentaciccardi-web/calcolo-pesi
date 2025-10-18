import streamlit as st
import pandas as pd
from io import StringIO
from openpyxl import load_workbook

st.set_page_config(page_title="Calcolatore Peso Prodotti", layout="wide")

st.title("⚖️ Calcolatore Peso Prodotti - Foglio 2")

# --- FUNZIONI ---
def read_products_from_excel(file):
    wb = load_workbook(file, data_only=True, read_only=True)
    if "Foglio 2" not in wb.sheetnames:
        st.error(f"'Foglio 2' non trovato. Fogli disponibili: {wb.sheetnames}")
        return pd.DataFrame()
    ws = wb["Foglio 2"]

    rows = []
    for row in ws.iter_rows(min_row=3, values_only=True):
        code = row[0]
        name = row[1]
        flag = row[2]
        unit_weight = None
        try:
            unit_weight = row[8]  # colonna I
        except IndexError:
            unit_weight = None

        if code is None and name is None and unit_weight is None:
            continue

        rows.append({
            "Codice": str(code).strip() if code else "",
            "Nome": str(name).strip() if name else "",
            "Flag": str(flag).strip() if flag else "",
            "Peso_unitario_kg": float(unit_weight) if unit_weight else None
        })
    return pd.DataFrame(rows)

def calcola_peso(unit_weight, pezzi, flag):
    try:
        peso = float(unit_weight) * float(pezzi)
        if "√" in str(flag):
            peso /= 2
        return peso
    except Exception:
        return 0.0


# --- UPLOAD FILE EXCEL ---
uploaded_file = st.file_uploader("📁 Carica file Excel (Foglio 2 richiesto)", type=["xlsx", "xlsm"])

if uploaded_file:
    prodotti_df = read_products_from_excel(uploaded_file)

    if not prodotti_df.empty:
        st.success(f"✅ File caricato correttamente ({len(prodotti_df)} prodotti trovati)")

        # --- SELEZIONE PRODOTTO ---
        col1, col2 = st.columns([1, 3])
        with col1:
            lista_prodotti = [f"{row.Codice} - {row.Nome}" for _, row in prodotti_df.iterrows()]
            selected = st.selectbox("Seleziona prodotto", [""] + lista_prodotti)
        with col2:
            st.info("Oppure inserisci i dati manualmente 👇")

        if selected:
            idx = lista_prodotti.index(selected)
            prodotto = prodotti_df.iloc[idx]
            codice = prodotto["Codice"]
            nome = prodotto["Nome"]
            flag = prodotto["Flag"]
            peso_unit = prodotto["Peso_unitario_kg"]
        else:
            codice = st.text_input("Codice prodotto")
            nome = st.text_input("Nome prodotto")
            flag = st.text_input("Flag (√ per dimezzare)")
            peso_unit = st.number_input("Peso unitario (kg)", min_value=0.0, step=0.01)

        pezzi = st.number_input("Numero pezzi", min_value=0.0, step=1.0)

        # --- CALCOLO E AGGIUNTA AL CARRELLO ---
        if "carrello" not in st.session_state:
            st.session_state.carrello = []

        if st.button("🛒 Aggiungi al riepilogo"):
            peso_tot = calcola_peso(peso_unit, pezzi, flag)
            st.session_state.carrello.append({
                "Codice": codice,
                "Nome": nome,
                "Flag": flag,
                "Peso_unitario_kg": peso_unit,
                "Pezzi": pezzi,
                "Kg": peso_tot
            })
            st.success("Prodotto aggiunto!")

        # --- MOSTRA RIEPILOGO ---
        if st.session_state.carrello:
            st.subheader("📋 Riepilogo prodotti")

            df_cart = pd.DataFrame(st.session_state.carrello)
            df_cart.index += 1
            st.dataframe(df_cart, use_container_width=True)

            totale = df_cart["Kg"].sum()
            st.metric("Totale peso (kg)", f"{totale:.4f}")

            # --- SCARICA CSV ---
            csv_buffer = StringIO()
            df_cart.to_csv(csv_buffer, index=False)
            st.download_button(
                label="💾 Scarica riepilogo CSV",
                data=csv_buffer.getvalue(),
                file_name="riepilogo_pesi.csv",
                mime="text/csv"
            )

            # --- PULISCI ---
            if st.button("🗑️ Svuota riepilogo"):
                st.session_state.carrello = []
                st.rerun()
    else:
        st.warning("Il foglio 'Foglio 2' non contiene prodotti validi.")
else:
    st.info("Carica un file Excel per iniziare.")
