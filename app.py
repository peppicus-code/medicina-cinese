import os
import re
import pandas as pd
import streamlit as st

# ---------------------------------------------------------
# CONFIGURAZIONE PAGINA
# ---------------------------------------------------------
st.set_page_config(
    page_title="Medicina Cinese - Ricettario Formule",
    layout="wide",
    initial_sidebar_state="expanded"
)

# ---------------------------------------------------------
# STILE CSS PERSONALIZZATO (Compattazione layout anti-scroll)
# ---------------------------------------------------------
st.markdown("""
    <style>
    .block-container {
        padding-top: 1.5rem !important;
        padding-bottom: 0rem !important;
        padding-left: 3rem !important;
        padding-right: 3rem !important;
    }
    html, body, [data-testid="stAppViewContainer"] {
        font-size: 17px !important;
    }
    h1 { font-size: 2.2rem !important; color: #1E3A8A !important; margin-bottom: 0.5rem !important; }
    h2 { font-size: 1.7rem !important; color: #0D9488 !important; }
    h3 { font-size: 1.3rem !important; color: #4338CA !important; margin-top: 0.5rem !important; margin-bottom: 0.5rem !important; }
    p, li, label { font-size: 1.1rem !important; line-height: 1.5 !important; }
    .stMetric { background-color: #F0FDF4; padding: 10px; border-radius: 8px; border: 1px solid #BBF7D0; }
    
    .custom-box {
        background-color: #F8FAFC;
        padding: 10px;
        border-radius: 6px;
        border: 1px solid #E2E8F0;
    }
    </style>
""", unsafe_allow_html=True)

# ---------------------------------------------------------
# 1. CREAZIONE E CARICAMENTO DATABASE EXCEL
# ---------------------------------------------------------
DB_FILE = "database_formule.xlsx"

def inizializza_database():
    if not os.path.exists(DB_FILE):
        dati = {
            "Nome Pinyin": [
                "Si Jun Zi Tang", "Qing Wei San", "Xiao Yao San", "Liu Wei Di Huang Wan", 
                "Gui Pi Tang", "Yin Qiao San", "Du Huo Ji Sheng Tang", "Ba Zhen Tang", 
                "Er Chen Tang", "Suan Zao Ren Tang"
            ],
            "Nome Italiano": [
                "Decotto dei Quattro Gentiluomini", "Polvere per Pulire lo Stomaco", "Polvere del Libero Vagabondaggio", "Pillola di Sei Erbe con Rehmannia",
                "Decotto per Nutrire la Milza", "Polvere di Caprifoglio e Forsizia", "Decotto di Du Huo e Vischio", "Decotto delle Otto Preziosità",
                "Decotto dei Due Ingredienti Vecchi", "Decotto di Semi di Giaggiolo"
            ],
            "Categoria": [
                "Tonici del Qi", "Chiarificatori del Calore", "Regolatori del Qi / Armonizzanti", "Tonici dello Yin",
                "Tonici di Cuore e Milza / Sangue", "Disperdenti il Calore-Vento Esterno", "Espellenti Vento-Umidità-Freddo", "Tonici di Qi e Sangue",
                "Trasformatori del Flemma-Umidità", "Nutrienti il Cuore e Calmanti lo Spirito"
            ],
            "Azione Energetica": [
                "Tonifica il Qi di Milza e Stomaco debole", "Purifica il Calore dello Stomaco e nutre lo Yin", "Muove il Qi di Fegato, dissolve la stasi e tonifica la Milza", "Nutre lo Yin di Fegato e Reni",
                "Tonifica il Qi e il Sangue, nutre il Cuore", "Dissipa il Vento-Calore, purifica le tossine", "Espelle Vento-Umidità, tonifica Reni e Fegato, allevia il dolore", "Tonifica fortemente il Qi e il Sangue",
                "Asciuga l'Umidità e trasforma il Flemma", "Nutre il Sangue del Fegato, calma lo Shen"
            ],
            "Sintomi": [
                "Inappetenza, feci molli, diarrea cronica, stanchezza cronica, voce debole, gonfiore addominale, digestione lenta, colon irritabile, astenia, milza debole, pancia gonfia", 
                "Gastrite acuta, bruciore di stomaco, mal di denti, gengive gonfie, alitosi, reflusso acido, gengivite, ulcerazioni della bocca, acidità di stomaco", 
                "Crampi addominali, spasmi intestinali, irritabilità, sindrome premestruale, mestruazioni dolorose, ciclo irregolare, ansia, sbalzi d'umore, tensione al seno, fegato, colon, emicrania da stress, mal di testa da tensione, crampi muscolari", 
                "Diabete, ipertensione, pressione alta, menopausa, vampate di calore, occhi secchi, vista debole, fegato, sudorazione notturna, dolore lombare, acufeni, ronzio orecchie, bocca secca, calore vuoto",
                "Insonnia iniziale, difficoltà ad addormentarsi, palpitazioni, memoria debole, ansia, pallore, sonno leggero, stanchezza mentale, carenza di sangue di cuore", 
                "Febbre, mal di gola, tosse secca, cefalea iniziale, mal di testa da influenza, raffreddore, laringite, calore vento esterno, placche gola", 
                "Sciatalgia, torcicollo, rigidità nucale, dolori articolari da freddo, artrosi, lombalgia cronica, reumatismi, vento umidità freddo, mal di schiena, cervicale", 
                "Grave spossatezza, vertigini, ciclo scarso, convalescenza, anemia, pallore cutaneo, debolezza dopo il parto, deficit qi e sangue, stanchezza estrema",
                "Gastrite cronica, nausea, vomito, tosse con escrezioni abbondanti e bianche, vertigini, flemma umidità, catarro, senso di pienezza al petto", 
                "Insonnia profonda, risvegli notturni frequenti, sonno disturbato da incubi, bocca secca, agitazione, palpitazioni serali, deficit yin fegato"
            ],
            "Ingredienti": [
                "Ren Shen (9), Bai Zhu (9), Fu Ling (9), Zhi Gan Cao (6)",
                "Huang Lian (6), Shi Gao (15), Sheng Di Huang (9), Mu Dan Pi (9), Dang Gui (9)",
                "Chai Hu (9), Dang Gui (9), Bai Shao (9), Bai Zhu (9), Fu Ling (9), Zhi Gan Cao (6)",
                "Shu Di Huang (24), Shan Zhu Yu (12), Shan Yao (12), Fu Ling (9), Mu Dan Pi (9), Ze Xie (9)",
                "Dang Shen (12), Huang Qi (12), Bai Zhu (9), Fu Ling (9), Suan Zao Ren (12), Long Yan Rou (12), Dang Gui (9), Zhi Gan Cao (6)",
                "Jin Yin Hua (15), Lian Qiao (15), Jie Geng (12), Niu Bang Zi (12), Bo He (6)",
                "Du Huo (9), Qin Jiao (9), Fang Feng (9), Xi Xin (3), Sang Ji Sheng (15), Du Zhong (12), Niuxi (9), Dang Gui (9), Bai Shao (9), Shu Di Huang (9), Chuan Xiong (6), Ren Shen (6), Fu Ling (9), Gan Cao (6)",
                "Ren Shen (9), Bai Zhu (9), Fu Ling (9), Zhi Gan Cao (6), Shu Di Huang (12), Bai Shao (9), Dang Gui (9), Chuan Xiong (6)",
                "Ban Xia (12), Chen Pi (12), Fu Ling (9), Zhi Gan Cao (4)",
                "Suan Zao Ren (15), Fu Ling (12), Zhi Mu (9), Chuan Xiong (6), Gan Cao (3)"
            ],
            "Prezzi Erbe": [
                "Ren Shen:0.25, Bai Zhu:0.08, Fu Ling:0.06, Zhi Gan Cao:0.05",
                "Huang Lian:0.18, Shi Gao:0.04, Sheng Di Huang:0.09, Mu Dan Pi:0.12, Dang Gui:0.15",
                "Chai Hu:0.12, Dang Gui:0.15, Bai Shao:0.08, Bai Zhu:0.08, Fu Ling:0.06, Zhi Gan Cao:0.05",
                "Shu Di Huang:0.10, Shan Zhu Yu:0.14, Shan Yao:0.07, Fu Ling:0.06, Mu Dan Pi:0.12, Ze Xie:0.05",
                "Dang Shen:0.14, Huang Qi:0.10, Bai Zhu:0.08, Fu Ling:0.06, Suan Zao Ren:0.20, Long Yan Rou:0.11, Dang Gui:0.15, Zhi Gan Cao:0.05",
                "Jin Yin Hua:0.16, Lian Qiao:0.14, Jie Geng:0.07, Niu Bang Zi:0.06, Bo He:0.09",
                "Du Huo:0.08, Qin Jiao:0.10, Fang Feng:0.11, Xi Xin:0.22, Sang Ji Sheng:0.07, Du Zhong:0.09, Niuxi:0.08, Dang Gui:0.15, Bai Shao:0.08, Shu Di Huang:0.10, Chuan Xiong:0.09, Ren Shen:0.25, Fu Ling:0.06, Gan Cao:0.05",
                "Ren Shen:0.25, Bai Zhu:0.08, Fu Ling:0.06, Zhi Gan Cao:0.05, Shu Di Huang:0.10, Bai Shao:0.08, Dang Gui:0.15, Chuan Xiong:0.09",
                "Ban Xia:0.15, Chen Pi:0.08, Fu Ling:0.06, Zhi Gan Cao:0.05",
                "Suan Zao Ren:0.20, Fu Ling:0.12, Zhi Mu:0.09, Chuan Xiong:0.06, Gan Cao:0.05"
            ],
            "Durata Base Giorni": [7, 5, 7, 14, 10, 3, 14, 10, 7, 7],
            "Fornitori Italia": [
                "Erboristerie specializzate in fitoterapia cinese; Farmacie galeniche autorizzate.",
                "Erboristerie cliniche; Negozi biologici con reparto erbe.",
                "Erboristerie di quartiere; Store online di fitoterapia.",
                "Farmacie con laboratorio galenico; Distributori autorizzati di MTC.",
                "Erboristerie professionali; Negozi specializzati in alimentazione orientale.",
                "Erboristerie standard; Negozi di tisane ed erbe sfuse.",
                "Farmacie specializzate MTC; Centri di distribuzione fitoterapica.",
                "Erboristerie fornite; Store integratori naturali.",
                "Erboristerie tradizionali; Negozi biologici dedicati.",
                "Erboristerie specializzate in rimedi per il sonno; Farmacie galeniche."
            ],
            "Preparazione": [
                "1. In 500ml di acqua fredda per 15 minuti.\n2. Portare a ebollizione e cuocere a fuoco lento per 40 minuti fino a dimezzare il liquido.\n3. Filtrare e bere tiepido.",
                "1. Mettere in ammollo Shi Gao per 10 minuti separatamente.\n2. Far bollire Shi Gao per 10 informazioni, poi aggiungere le altre erbe e cuocere per 20 minuti.\n3. Filtrare.",
                "1. Mettere in ammollo le erbe in 400ml di acqua fredda.\n2. Portare a ebollizione a fuoco moderato per 25 minuti.\n3. Consumare caldo lontano dai pasti.",
                "1. Bollire le radici dure in 500ml d'acqua per 35 minuti.\n2. Aggiungere le erbe rimanenti e proseguire per 10 minuti.\n3. Filtrare.",
                "1. Immergere in 600ml di acqua fredda.\n2. Cuocere a fuoco lento per 45 minuti per estrarre le sostanze tonificanti.\n3. Filtrare.",
                "1. Portare l'acqua a ebollizione.\n2. Versare le erbe e far bollire per massimo 10-15 minuti (cottura breve per oli essenziali).\n3. Filtrare.",
                "1. Bollire le cortecce e radici in 700ml d'acqua a fuoco basso per 50 minuti.\n2. Filtrare il decotto ottenuto e consumarlo in due dosi.",
                "1. Cuocere in 500ml d'acqua a fuoco medio-basso per 40 informazioni.\n2. Filtrare il liquido residuo e berlo tiepido.",
                "1. Cuocere la Ban Xia trattata e il Chen Pi insieme per 30 minuti in 400ml d'acqua.\n2. Filtrare accuratamente prima del consumo.",
                "1. Preparare il decotto cuocendo a fuoco controllato per 30 minuti.\n2. Assumere l'estratto liquido 1 ora prima di coricarsi."
            ]
        }
        df = pd.DataFrame(dati)
        df.to_excel(DB_FILE, index=False)

if os.path.exists(DB_FILE):
    os.remove(DB_FILE)
inizializza_database()
df_db = pd.read_excel(DB_FILE)

# ---------------------------------------------------------
# 2. DIZIONARIO ENCICLOPEDICO ESTESO CON NUOVI DISTURBI
# ---------------------------------------------------------
ENCICLOPEDIA_MTC = {
    "occhi": {
        "sinonimi": ["occhi secchi", "vista debole", "visione offuscata", "occhi arrossati"],
        "diagnosi": "In MTC, **gli occhi sono l'apertura del Fegato**. La secchezza indica deficit di Yin di Fegato e Reni."
    },
    "fegato": {
        "sinonimi": ["stasi fegato", "rabbia", "fegato grasso", "ipocondrio"],
        "diagnosi": "Il **Fegato** muove il Qi. Se bloccato da stress, causa spasmi e contratture."
    },
    "colon": {
        "sinonimi": ["colon irritabile", "intestino", "feci molli", "stipsi", "colite", "pancia gonfia"],
        "diagnosi": "Riflette una **Disarmonia tra Fegato e Milza**. Lo stress blocca il transito intestinale causando spasmi."
    },
    "diarrea": {
        "sinonimi": ["feci liquide", "diarrea cronica", "scariche", "feci acquose"],
        "diagnosi": "Evidenzia un **Deficit di Qi di Milza** che non trattiene i liquidi nell'addome."
    },
    "diabete": {
        "sinonimi": ["sindrome xiao ke", "bocca secca", "sete intensa", "glicemia"],
        "diagnosi": "Inquadrato come **Xiao Ke**. È un calore interno che consuma i liquidi per deficit di Yin."
    },
    "gastrite": {
        "sinonimi": ["bruciore di stomaco", "reflusso acido", "nausea", "iperacidità", "acidità"],
        "diagnosi": "Rappresenta il **Fuoco di Stomaco**. L'energia gastrica risale (Qi ribelle) invece di scendere."
    },
    "menopausa": {
        "sinonimi": ["vampate", "sudorazione notturna", "vampate di calore"],
        "diagnosi": "Dovuta al declino naturale dello **Yin dei Reni**, che causa risalite improvvise di calore."
    },
    "sciatalgia": {
        "sinonimi": ["dolori articolari", "lombare", "freddo", "lombalgia", "mal di schiena", "cervicale", "torcicollo"],
        "diagnosi": "Ostruzione da **Vento-Freddo-Umidità** (Sindrome Bi) che blocca la circolazione nei meridiani."
    },
    "ipertensione": {
        "sinonimi": ["pressione alta", "pressione", "acufeni", "ronzio orecchie"],
        "diagnosi": "Espressione della risalita dello **Yang del Fegato** che non viene ancorato a causa di un deficit di Yin."
    },
    "insonnia": {
        "sinonimi": ["ansia", "agitazione", "incubi", "non dormo", "risvegli", "difficoltà ad addormentarsi"],
        "diagnosi": "Lo **Shen** (Spirito del Cuore) non è ancorato di notte per deficit di Sangue o Calore Vuoto."
    },
    "stanchezza": {
        "sinonimi": ["spossatezza", "stanchezza cronica", "qi debole", "astenia", "stanchezza estrema"],
        "diagnosi": "Indica un **Deficit del Qi di Milza**, l'organo che estrae l'energia vitale dai cibi."
    },
    "mestruazioni": {
        "sinonimi": ["mestruazioni dolorose", "ciclo irregolare", "sindrome premestruale", "ciclo scarso", "ciclo"],
        "diagnosi": "I dolori e le irregolarità legati al **Ciclo Mestruale** derivano da una **Stasi di Sangue e Qi di Fegato**. Il Fegato immagazzina il Sangue; se l'energia si blocca, il flusso diventa doloroso, scuro o irregolare. La formula muove il sangue e decongestiona il bacino."
    },
    "mal di testa": {
        "sinonimi": ["emicrania da stress", "mal di testa da tensione", "cefalea", "cefalea iniziale"],
        "diagnosi": "Il **Mal di testa** (Cefalea) può derivare da una risalita dello Yang di Fegato (tipo pulsante e laterale) o da un attacco di Vento Esterno (tipo tensivo o frontale). La formula purifica il calore in alto e sblocca i meridiani cranici."
    },
    "crampi muscolari": {
        "sinonimi": ["crampi muscolari", "spasmi", "contrazioni", "rigidità nucale"],
        "diagnosi": "I **Crampi muscolari** indicano che il **Sangue del Fegato non nutre i tendini** ('Il Fegato governa i tendini'). Quando i tessuti rimangono asciutti e privi di nutrimento ematico, si contraggono dolorosamente. La formula nutre il sangue per ammorbidire i muscoli."
    }
}

def trova_diagnosi_universale(testo_ricerca):
    testo_clean = testo_ricerca.lower().strip()
    if not testo_clean:
        return "Nessuna chiave inserita."
    for chiave, info in ENCICLOPEDIA_MTC.items():
        if testo_clean in chiave or any(testo_clean in s for s in info["sinonimi"]):
            return info["diagnosi"]
    return "La formula selezionata agisce ripristinando l'equilibrio energetico degli organi coinvolti."

def espandi_ricerca(chiave_ricerca):
    chiave_clean = chiave_ricerca.lower().strip()
    termini_ricerca = [chiave_clean]
    for chiave, info in ENCICLOPEDIA_MTC.items():
        if chiave_clean in chiave or any(chiave_clean in s for s in info["sinonimi"]):
            termini_ricerca.append(chiave)
            termini_ricerca.extend(info["sinonimi"])
    return list(set(termini_ricerca))

# KEY ROTATION PATTERN
if "reset_counter" not in st.session_state:
    st.session_state.reset_counter = 0

# ---------------------------------------------------------
# INTERFACCIA UTENTE - SIDEBAR
# ---------------------------------------------------------
st.sidebar.header("🔍 Ricerca e Parametri")

if st.sidebar.button("🔄 Azzera Tutto"):
    st.session_state.reset_counter += 1

ricerca_input = st.sidebar.text_input(
    "Cerca sintomo o patologia (es: mestruazioni, mal di testa, crampi muscolari, colon):", 
    key=f"ricerca_{st.session_state.reset_counter}"
)

# Filtro formule dinamico
formule_filtrate = df_db
if ricerca_input:
    termini_espanditi = espandi_ricerca(ricerca_input)
    pattern = "|".join(termini_espanditi)
    formule_filtrate = df_db[
        df_db['Nome Pinyin'].str.contains(pattern, case=False, na=False) |
        df_db['Nome Italiano'].str.contains(pattern, case=False, na=False) |
        df_db['Sintomi'].str.contains(pattern, case=False, na=False)
    ]

if ricerca_input and not formule_filtrate.empty:
    elenco_nomi = formule_filtrate['Nome Pinyin'].tolist()
else:
    elenco_nomi = ["-- Seleziona una formula o cerca una patologia --"] + df_db['Nome Pinyin'].tolist()

formula_selezionata = st.sidebar.selectbox(
    "Formula Selezionata:", 
    elenco_nomi,
    index=0,
    key=f"formula_{st.session_state.reset_counter}"
)

# NOVITÀ: Calcolo e aggiornamento automatico della durata corretta basata sul database
valore_giorni_default = 7
if formula_selezionata != "-- Seleziona una formula o cerca una patologia --":
    riga_temp = df_db[df_db['Nome Pinyin'] == formula_selezionata]
    if not riga_temp.empty:
        valore_giorni_default = int(riga_temp['Durata Base Giorni'].values[0])

giorni_trattamento = st.sidebar.slider(
    "Giorni di trattamento calcolati automaticamente:", 
    min_value=1, 
    max_value=14, 
    value=valore_giorni_default,
    key=f"slider_{formula_selezionata}_{st.session_state.reset_counter}"
)

# ---------------------------------------------------------
# LAYOUT PRINCIPALE - FISSO ANTI-SCROLL
# ---------------------------------------------------------
st.title("🌿 Studio Medico MTC - Ricettario Digitale")

if formula_selezionata == "-- Seleziona una formula o cerca una patologia --":
    st.info("👋 Benvenuto. Inserisci una qualsiasi patologia occidentale o un sintomo nella barra laterale a sinistra per attivare in tempo reale il ricettario clinico.")
else:
    riga_formula = df_db[df_db['Nome Pinyin'] == formula_selezionata].iloc[0]

    col1, col2 = st.columns([1.1, 0.9], gap="large")

    # --- COLONNA 1: SCHEDA FORMULA E DIAGNOSTICA ---
    with col1:
        st.subheader(f"📋 {riga_formula['Nome Pinyin']} ({riga_formula['Nome Italiano']})")
        
        with st.container(height=170):
            st.markdown(f"**Categoria:** {riga_formula['Categoria']}")
            st.markdown(f"**Azione:** {riga_formula['Azione Energetica']}")
            st.markdown(f"**Sintomi:** {riga_formula['Sintomi']}")
            st.markdown(f"**Preparazione:** {riga_formula['Preparazione']}")
        
        st.markdown("### ☯️ Traduzione Diagnostica Orientale")
        with st.container(height=140):
            testo_diagnostica = trova_diagnosi_universale(ricerca_input)
            st.markdown(f"<div class='custom-box'>{testo_diagnostica}</div>", unsafe_allow_html=True)

        st.markdown("### 🏪 Reperibilità dei Componenti")
        with st.container(height=140):
            st.markdown(f"**Acquisto:** {riga_formula['Fornitori Italia']}")
            st.markdown("""
            * 🍃 **ERBORISTERIE**: Reperibili in estratto secco o radici sfuse presso erboristerie fisiche fornite o store online specializzati in fitoterapia.
            * 🛒 **SUPERMERCATI**: Alcune radici o estratti base sono presenti nei supermercati biologici o nel reparto integratori alimentari naturali.
            * 🏬 **NEGOZI SPECIALIZZATI**: Disponibili come radici sfuse tradizionali intere o preparati grezzi presso i supermercati asiatici ed empori orientali.
            """)

    # --- COLONNA 2: DOSAGGI, METRICHE E GRAFICO COMPATTO ---
    with col2:
        st.subheader("⚖️ Dosaggi e Calcolo Costi")
        
        ingredienti_raw = str(riga_formula['Ingredienti'])
        prezzi_raw = str(riga_formula['Prezzi Erbe'])
        
        prezzi_dict = {}
        for p_entry in prezzi_raw.split(','):
            if ':' in p_entry:
                k, v = p_entry.split(':')
                try:
                    prezzi_dict[k.strip()] = float(v.strip())
                except ValueError:
                    pass

        nomi_erbe = []
        grammi_totali = []
        grammi_base_lista = []
        costi_erbe = []
        lista_ingr = ingredienti_raw.split(',')
        
        for ingr in lista_ingr:
            try:
                match = re.search(r'([a-zA-Z\s]+)\s*\((\d+)\)', ingr)
                if match:
                    erba_nome = match.group(1).strip()
                    grammi_base = float(match.group(2).strip())
                    g_tot = grammi_base * giorni_trattamento
                    prezzo_singolo_g = prezzi_dict.get(erba_nome, 0.05)
                    costo_parziale = g_tot * prezzo_singolo_g
                    
                    nomi_erbe.append(erba_nome)
                    grammi_base_lista.append(grammi_base)
                    grammi_totali.append(g_tot)
                    costi_erbe.append(round(costo_parziale, 2))
            except Exception:
                continue

        df_costi_calcolati = pd.DataFrame({
            "Erba (Pinyin)": nomi_erbe,
            "Grammi Totali": grammi_totali,
            "Costo (€)": costi_erbe
        })
        
        st.dataframe(df_costi_calcolati, use_container_width=True, hide_index=True, height=150)
        
        costo_totale_ricetta = sum(costi_erbe)
        st.metric(label=f"Spesa Totale Stimata ({giorni_trattamento} Giorni)", value=f"€ {costo_totale_ricetta:.2f}")
        
        if nomi_erbe:
            df_chart = pd.DataFrame({
                "Erba": nomi_erbe,
                "Spesa (€)": costi_erbe
            }).set_index("Erba")
            st.bar_chart(df_chart, y="Spesa (€)", color="#0D9488", height=150)

    # --- FINALIZZAZIONE E STAMPA RICETTA INTERA ---
    st.markdown("<hr style='margin-top:0.2rem; margin-bottom:0.2rem;'>", unsafe_allow_html=True)
    
    col_paz, col_btn = st.columns([1.1, 0.9], gap="medium")
    with col_paz:
        paziente_nome = st.text_input(
            "Nome Paziente (Opzionale):", 
            key=f"paziente_{st.session_state.reset_counter}",
            label_visibility="collapsed",
            placeholder="Inserisci Nome e Cognome Paziente..."
        )
    with col_btn:
        testo_stampa = "==================================================\n"
        testo_stampa += "              RICETTA MEDICINA TRADIZIONALE CINESE\n"
        testo_stampa += "==================================================\n"
        testo_stampa += f"Paziente: {paziente_nome if paziente_nome else 'N.D.'}\n"
        testo_stampa += f"Formula: {riga_formula['Nome Pinyin']} ({riga_formula['Nome Italiano']})\n"
        testo_stampa += f"Categoria: {riga_formula['Categoria']}\n"
        testo_stampa += f"Azione Energetica: {riga_formula['Azione Energetica']}\n"
        testo_stampa += f"Fattore Moltiplicatore Applicato: x{giorni_trattamento}\n\n"
        testo_stampa += "--------------------------------------------------\n"
        testo_stampa += "DOSAGGI DEGLI INGREDIENTI CALCOLATI:\n"
        for i in range(len(nomi_erbe)):
            testo_stampa += f"- {nomi_erbe[i]}: {grammi_totali[i]:.1f} g (Base giornaliera: {grammi_base_lista[i]:.1f}g)\n"
        testo_stampa += "--------------------------------------------------\n\n"
        testo_stampa += "MODALITÀ DI PREPARAZIONE (DECOTTO):\n"
        testo_stampa += f"{riga_formula['Preparazione']}\n\n"
        testo_stampa += "REPERIBILITÀ IN ITALIA:\n"
        testo_stampa += "• **ERBORISTERIE**: Reperibili in estratto secco o radici sfuse presso erboristerie fisiche fornite o store online specializzati in fitoterapia.\n\n"
        testo_stampa += "• **SUPERMERCATI**: Alcune radici o estratti base sono presenti nei supermercati biologici o nel reparto integratori alimentari naturali.\n\n"
        testo_stampa += "• **NEGOZI SPECIALIZZATI**: Disponibili come radici sfuse tradizionali intere o preparati grezzi presso i supermercati asiatici ed empori orientali.\n\n"
        testo_stampa += "==================================================\n"
        testo_stampa += "Documento stampabile generato dall'applicazione web.\n"
        testo_stampa += "==================================================\n"

        st.download_button(
            label="💾 Scarica Ricetta Pronto Stampa (.txt)",
            data=testo_stampa,
            file_name=f"ricetta_{formula_selezionata.replace(' ', '_')}.txt",
            mime="text/plain",
            use_container_width=True
        )
