import streamlit as st
import pandas as pd
from openpyxl import Workbook
from openpyxl.styles import PatternFill, Font
import io
import re
import requests

st.set_page_config(page_title="Scyavuru Lead Manager", page_icon="🍯", layout="wide")

COLORS = ["E6F2FF", "FFF2E6", "E6FFE6", "FFE6E6", "F2E6FF", "FFFFE6", "E6FFFF", "FFE6FF", "F0F0F0", "E6F9FF"]
APOLLO_API_KEY = "gNGHsp8mMIyM2Pam5tvxVg"  # Chiave di Elisa

st.title("🍯 Scyavuru Lead Manager Pro")
st.markdown("Strumento aziendale per l'estrazione autonoma e la pulizia dei database GDO.")

# --- UTILITY SCRAPING APOLLO ---
def get_domain_from_company(company_name: str) -> str:
    if not company_name or company_name == "Da verificare":
        return None
    key = company_name.lower().strip()
    clean = re.sub(r'\b(s\.?p\.?a\.?|s\.?r\.?l\.?|group|italia|gmbh|ltd|inc)\b', '', key, flags=re.IGNORECASE)
    clean = re.sub(r'[^a-z0-9]', '', clean.strip())
    if clean and len(clean) > 3:
        return f"{clean}.com"
    return None

def run_search(ruolo: str, azienda: str, location: str, max_profili: int, apollo_api_key: str):
    if not apollo_api_key:
        raise ValueError("Inserisci la tua API Key di Apollo per procedere.")

    url = "https://api.apollo.io/v1/mixed_people/search"
    payload = {
        "q_keywords": ruolo,
        "person_locations": [location],
        "per_page": min(max_profili, 100)
    }
    
    if azienda:
        domain = get_domain_from_company(azienda)
        if domain:
            payload["q_organization_domains"] = domain

    headers = {
        "Cache-Control": "no-cache", 
        "Content-Type": "application/json",
        "X-Api-Key": apollo_api_key
    }
    
    response = requests.post(url, headers=headers, json=payload, timeout=20)
    response.raise_for_status()
    data = response.json()
    
    people = data.get("people", [])
    results = []

    for person in people[:max_profili]:
        nome = f"{person.get('first_name', '')} {person.get('last_name', '')}".strip()
        qualifica = person.get("title", "N/D")
        org = person.get("organization", {})
        co_name = org.get("name", "Da verificare") if org else "Da verificare"
        linkedin_url = person.get("linkedin_url", "")
        email = person.get("email", "Non trovata")
        
        results.append({
            "Categoria (Ricerca)": ruolo,
            "Nome": nome.strip(),
            "Qualifica": qualifica,
            "Azienda": co_name,
            "Email": email,
            "Link Profilo": linkedin_url
        })
        
    return results

# --- TABS ---
tab1, tab2 = st.tabs(["🔍 Estrazione Lead (Scraping)", "🧹 Pulizia Database (Deep Cleaning)"])

# ----------------------------------------------------
# TAB 1: ESTRAZIONE (SCRAPING)
# ----------------------------------------------------
with tab1:
    st.header("Estrazione Autonoma da LinkedIn (via Apollo)")
    st.markdown("Usa questo strumento per cercare nuovi contatti. Inserisci la qualifica, il paese e l'API Key.")
    
    col_a, col_b = st.columns(2)
    with col_a:
        ruolo_input = st.text_input("Qualifica da cercare (es. Buyer Food, Category Manager)", value="Buyer Food")
        azienda_input = st.text_input("Azienda Specifica (Opzionale, es. Esselunga)", value="")
        location_input = st.text_input("Nazione / Città (es. Italy, Milan)", value="Italy")
    
    with col_b:
        max_profili_input = st.number_input("Numero Massimo di Profili", min_value=1, max_value=100, value=20)
        
    if st.button("🛰️ Avvia Scraping"):
        with st.spinner('Scraping in corso... Ricerca profili e aggiramento blocchi...'):
            try:
                risultati = run_search(ruolo_input, azienda_input, location_input, max_profili_input, APOLLO_API_KEY)
                if risultati:
                    df_scraping = pd.DataFrame(risultati)
                    st.success(f"Trovati {len(risultati)} profili!")
                    st.dataframe(df_scraping)
                    
                    # Export Excel Base
                    output_scraping = io.BytesIO()
                    with pd.ExcelWriter(output_scraping, engine="openpyxl") as writer:
                        df_scraping.to_excel(writer, index=False, sheet_name="Estrazione")
                    output_scraping.seek(0)
                    
                    st.download_button(
                        label="📥 Scarica Excel Grezzo",
                        data=output_scraping,
                        file_name=f"Estrazione_{ruolo_input.replace(' ', '_')}.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="secondary"
                    )
                else:
                    st.info("Nessun profilo trovato con questi criteri.")
            except Exception as e:
                st.error(f"Errore durante lo scraping: {e}")

# ----------------------------------------------------
# TAB 2: PULIZIA (DEEP CLEANING)
# ----------------------------------------------------
with tab2:
    st.header("Pulizia e Formattazione File")
    st.markdown("Carica qui i file grezzi per rimuovere i fuori target, eliminare i duplicati e creare l'Excel Colorato.")
    
    uploaded_file = st.file_uploader("1️⃣ Carica l'estrazione grezza (Excel)", type=["xlsx"])

    st.markdown("### 2️⃣ Impostazioni di Filtraggio")
    col1, col2 = st.columns(2)

    with col1:
        whitelist_input = st.text_area(
            "Parole Chiave OBBLIGATORIE (Whitelist)", 
            value="supermercato, ipermercato, discount, grocery, gdo, conad, coop, esselunga, carrefour, lidl, eurospin, md, penny, crai, selex, pam, despar, aldi, tigros, vege, food, alimentare, dolci, confectionery, biscott, fmcg, horeca, gastronomia, retail, cash & carry",
            height=150
        )

    with col2:
        blacklist_input = st.text_area(
            "Settori da CANCELLARE (Blacklist)", 
            value="immobiliare, real estate, bank, banca, assicurazion, insurance, fashion, abbigliamento, software, it, tech, technology, hotel, tourism, turismo, automotive, auto, telecom, medical, farma, pharma, hr, recruiting, legal, avvocato, construction, costruzioni, edilizia, elettrotecnica, furniture, arredamento, energia, energy, renewable, logistica, packaging, plastic, metal, steel, electronic, brico, fai da te, design, architett, media, marketing, formazion, scuola, clinic, hospital",
            height=150
        )

    if uploaded_file is not None:
        if st.button("🚀 Avvia Purificazione e Formattazione", type="primary"):
            with st.spinner('Analisi in corso...'):
                try:
                    xls = pd.ExcelFile(uploaded_file)
                    df_list = []
                    for sheet in xls.sheet_names:
                        df = pd.read_excel(xls, sheet_name=sheet)
                        if 'Categoria (Ricerca)' not in df.columns:
                            df['Categoria (Ricerca)'] = sheet
                        df_list.append(df)
                        
                    df_tot = pd.concat(df_list, ignore_index=True)
                    original_len = len(df_tot)
                    
                    whitelist = [x.strip().lower() for x in whitelist_input.split(',')]
                    blacklist = [x.strip().lower() for x in blacklist_input.split(',')]
                    
                    wl_regex = re.compile('|'.join([rf'\b{w}\b' for w in whitelist if w]), re.IGNORECASE)
                    bl_regex = re.compile('|'.join([rf'\b{w}\b' for w in blacklist if w]), re.IGNORECASE)
                    
                    valid_rows = []
                    for idx, row in df_tot.iterrows():
                        azienda = str(row.get('Azienda', '')).lower()
                        qualifica = str(row.get('Qualifica', '')).lower()
                        text_to_check = f"{azienda} {qualifica}"
                        
                        if bl_regex.search(text_to_check):
                            continue
                        if wl_regex.search(text_to_check):
                            valid_rows.append(row)
                            
                    df_clean = pd.DataFrame(valid_rows)
                    
                    cols_to_drop = [c for c in df_clean.columns if 'email' in c.lower()]
                    df_clean.drop(columns=cols_to_drop, inplace=True, errors='ignore')
                    
                    df_clean['Nome_str'] = df_clean.get('Nome', '').astype(str).str.strip().str.title()
                    df_clean['Cognome_str'] = df_clean.get('Cognome', '').astype(str).str.strip().str.title()
                    if 'Link Profilo' in df_clean.columns:
                        df_clean = df_clean.drop_duplicates(subset=['Link Profilo'])
                    df_clean = df_clean.drop_duplicates(subset=['Nome_str', 'Cognome_str'])
                    df_clean.drop(columns=['Nome_str', 'Cognome_str'], inplace=True, errors='ignore')
                    
                    final_len = len(df_clean)
                    st.success(f"✅ Operazione Completata! Contatti originali: {original_len} | Contatti perfetti: {final_len} | Eliminati: {original_len - final_len}")
                    
                    wb = Workbook()
                    ws = wb.active
                    ws.title = "Database GDO"
                    
                    cols = list(df_clean.columns)
                    for c_idx, col_name in enumerate(cols, 1):
                        ws.cell(row=1, column=c_idx, value=col_name).font = Font(bold=True)
                        
                    categorie = df_clean['Categoria (Ricerca)'].unique() if 'Categoria (Ricerca)' in df_clean.columns else []
                    cat_color_map = {cat: PatternFill(start_color=COLORS[i % len(COLORS)], end_color=COLORS[i % len(COLORS)], fill_type="solid") for i, cat in enumerate(categorie)}
                    
                    for r_idx, row_list in enumerate(df_clean.values.tolist(), 2):
                        cat_value = row_list[cols.index('Categoria (Ricerca)')] if 'Categoria (Ricerca)' in cols else None
                        for c_idx, value in enumerate(row_list, 1):
                            cell = ws.cell(row=r_idx, column=c_idx, value=value)
                            if cat_value and cat_value in cat_color_map:
                                cell.fill = cat_color_map[cat_value]
                                
                    output = io.BytesIO()
                    wb.save(output)
                    output.seek(0)
                    
                    st.download_button(
                        label="📥 SCARICA EXCEL PULITO E COLORATO",
                        data=output,
                        file_name="Scyavuru_Database_Pulito.xlsx",
                        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                        type="primary"
                    )
                    
                except Exception as e:
                    st.error(f"Errore durante l'elaborazione: {e}")

