import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import numpy as np
from datetime import datetime, date
from supabase import create_client, Client

# ============================================================
# CONFIGURATION DE LA PAGE
# ============================================================
st.set_page_config(
    page_title="SKAB — Dashboard CI",
    page_icon="🛡️",
    layout="wide",
    initial_sidebar_state="expanded"
)

st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 28px; font-weight: 700; }
    [data-testid="stMetricLabel"] { font-size: 13px; color: #888; }
    .stAlert { margin-top: 10px; }
    div[data-testid="stHorizontalBlock"] > div { gap: 1rem; }
    .section-title {
        font-size: 16px;
        font-weight: 700;
        color: #1a1a2e;
        border-left: 4px solid #e63946;
        padding-left: 10px;
        margin-bottom: 12px;
    }
    .filter-badge {
        display: inline-block;
        background-color: #e63946;
        color: white;
        padding: 6px 12px;
        border-radius: 20px;
        font-size: 12px;
        margin: 4px 4px 4px 0;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# CONNEXION SUPABASE
# ============================================================
@st.cache_resource
def init_supabase() -> Client:
    """Initialise le client Supabase à partir des secrets Streamlit Cloud."""
    url  = st.secrets["SUPABASE_URL"]
    key  = st.secrets["SUPABASE_KEY"]
    return create_client(url, key)

try:
    supabase = init_supabase()
    SUPABASE_OK = True
except Exception:
    SUPABASE_OK = False

# ============================================================
# MOTEUR DE LECTURE DES FICHIERS EXCEL
# ============================================================
def get_available_sheets(file):
    """Retourne la liste des feuilles disponibles dans le classeur."""
    try:
        xls = pd.ExcelFile(file)
        return xls.sheet_names
    except Exception:
        return []


def find_sheet_containing_keyword(file, keywords):
    """Cherche la première feuille contenant l'un des mots-clés dans ses colonnes."""
    sheet_names = get_available_sheets(file)
    for sheet in sheet_names:
        try:
            df_test = pd.read_excel(file, sheet_name=sheet, header=None, nrows=20)
            df_str = df_test.astype(str).values.flatten()
            if any(keyword in cell for cell in df_str for keyword in keywords):
                return sheet
        except Exception:
            continue
    return None


def load_and_clean(file, keywords, sheet=None):
    """
    Charge et nettoie les données depuis une feuille Excel.
    Si `sheet` est None, cherche automatiquement la feuille contenant les keywords.
    """
    try:
        # Si aucune feuille n'est spécifiée, la chercher automatiquement
        if sheet is None:
            sheet = find_sheet_containing_keyword(file, keywords)
            if sheet is None:
                return pd.DataFrame()
        
        # Essayer de lire la feuille
        file.seek(0)
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
        
        if df_raw.empty:
            return pd.DataFrame()

        # Chercher la ligne d'en-tête
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = [str(val).strip() for val in row.values]
            if any(keyword in cell for cell in row_str for keyword in keywords):
                header_idx = idx
                break

        if header_idx is None:
            for idx, row in df_raw.iterrows():
                row_str = [str(val).strip() for val in row.values]
                if any("ID" in s or "N°" in s or "Code" in s for s in row_str):
                    header_idx = idx
                    break

        if header_idx is None:
            header_idx = 0

        file.seek(0)
        df = pd.read_excel(file, sheet_name=sheet, skiprows=header_idx)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(subset=[df.columns[0]]) if not df.empty else df
        df = df.dropna(how='all')

        if not df.empty:
            df = df[~df[df.columns[0]].astype(str).str.contains(
                "Une anomalie|Un plan|Saisissez|Une ligne", na=False
            )]
            df['Fichier Source'] = file.name

        return df
    except Exception as e:
        return pd.DataFrame()


def process_consolidation(files):
    """Consolide les données de tous les fichiers uploadés."""
    all_data = {"MISSIONS": [], "POINTS": [], "ANOMALIES": [], "PLANS": []}
    
    for f in files:
        # Missions
        df_missions = load_and_clean(f, ["N° Mission", "Mission"])
        if not df_missions.empty:
            all_data["MISSIONS"].append(df_missions)
        
        # Points de contrôle
        df_points = load_and_clean(f, ["ID Point", "Point de Contrôle"])
        if not df_points.empty:
            all_data["POINTS"].append(df_points)
        
        # Anomalies
        df_anomalies = load_and_clean(f, ["ID Anomalie", "Anomalie"])
        if not df_anomalies.empty:
            all_data["ANOMALIES"].append(df_anomalies)
        
        # Plans d'action
        df_plans = load_and_clean(f, ["ID Plan", "Plan"])
        if not df_plans.empty:
            all_data["PLANS"].append(df_plans)

    return {k: pd.concat(v, ignore_index=True) if v else pd.DataFrame()
            for k, v in all_data.items()}


def get_safe_len(series, col_name):
    clean = series.dropna()
    max_c = int(clean.apply(lambda x: len(str(x))).max()) if not clean.empty else 0
    return min(max(max_c, len(str(col_name))) + 3, 50)

# ============================================================
# PUSH VERS SUPABASE
# ============================================================
def push_to_supabase(data: dict, source_files):
    """Pousse les dataframes consolidés vers Supabase."""
    errors = []

    df_m = data["MISSIONS"]
    if not df_m.empty:
        col_num   = next((c for c in df_m.columns if 'Mission' in c or 'N°' in c), None)
        col_ag    = next((c for c in df_m.columns if 'Agence'  in c or 'Entité' in c or 'Site' in c), None)
        col_deb   = next((c for c in df_m.columns if 'Début'   in c or 'début'  in c or 'Start' in c), None)
        col_fin   = next((c for c in df_m.columns if 'Fin'     in c or 'fin'    in c or 'End'   in c), None)
        col_conf  = next((c for c in df_m.columns if 'conform' in c.lower()), None)
        col_stat  = next((c for c in df_m.columns if 'Statut'  in c or 'statut' in c), None)

        rows = []
        for _, r in df_m.iterrows():
            rows.append({
                "fichier_source":   str(r.get('Fichier Source', '')),
                "num_mission":      str(r[col_num])   if col_num  else None,
                "agence":           str(r[col_ag])    if col_ag   else None,
                "date_debut":       _safe_date(r[col_deb]) if col_deb else None,
                "date_fin":         _safe_date(r[col_fin]) if col_fin else None,
                "taux_conformite":  _safe_num(r[col_conf]) if col_conf else None,
                "statut":           str(r[col_stat])  if col_stat else None,
            })
        try:
            supabase.table("missions").insert(rows).execute()
        except Exception as e:
            errors.append(f"Missions : {e}")

    df_a = data["ANOMALIES"]
    if not df_a.empty:
        col_id    = next((c for c in df_a.columns if 'ID Anomalie' in c), None)
        col_mis   = next((c for c in df_a.columns if 'Mission'     in c or 'N°' in c), None)
        col_ag    = next((c for c in df_a.columns if 'Agence'      in c or 'Site' in c or 'Entité' in c), None)
        col_pays  = next((c for c in df_a.columns if 'Pays'        in c), None)
        col_dom   = next((c for c in df_a.columns if 'Domaine'     in c or 'Type' in c), None)
        col_crit  = next((c for c in df_a.columns if 'critic'      in c.lower() or 'Critic' in c), None)
        col_desc  = next((c for c in df_a.columns if 'Description' in c or 'Libellé' in c), None)
        col_imp   = next((c for c in df_a.columns if 'Impact'      in c), None)
        col_stat  = next((c for c in df_a.columns if 'Statut'      in c or 'statut' in c), None)
        col_date  = next((c for c in df_a.columns if 'Date'        in c or 'date'   in c), None)

        rows = []
        for _, r in df_a.iterrows():
            rows.append({
                "fichier_source":   str(r.get('Fichier Source', '')),
                "id_anomalie":      str(r[col_id])   if col_id   else None,
                "num_mission":      str(r[col_mis])  if col_mis  else None,
                "agence":           str(r[col_ag])   if col_ag   else None,
                "pays":             str(r[col_pays]) if col_pays else None,
                "domaine":          str(r[col_dom])  if col_dom  else None,
                "criticite":        str(r[col_crit]) if col_crit else None,
                "description":      str(r[col_desc]) if col_desc else None,
                "impact_financier": _safe_num(r[col_imp])  if col_imp  else None,
                "statut":           str(r[col_stat]) if col_stat else None,
                "date_detection":   _safe_date(r[col_date]) if col_date else None,
            })
        try:
            supabase.table("anomalies").insert(rows).execute()
        except Exception as e:
            errors.append(f"Anomalies : {e}")

    df_p = data["POINTS"]
    if not df_p.empty:
        col_id   = next((c for c in df_p.columns if 'ID Point' in c), None)
        col_mis  = next((c for c in df_p.columns if 'Mission'  in c or 'N°' in c), None)
        col_ag   = next((c for c in df_p.columns if 'Agence'   in c or 'Site' in c), None)
        col_res  = next((c for c in df_p.columns if 'Résultat' in c or 'Result' in c), None)

        rows = []
        for _, r in df_p.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source', '')),
                "id_point":       str(r[col_id])  if col_id  else None,
                "num_mission":    str(r[col_mis]) if col_mis else None,
                "agence":         str(r[col_ag])  if col_ag  else None,
                "resultat":       str(r[col_res]) if col_res else None,
            })
        try:
            supabase.table("points_controle").insert(rows).execute()
        except Exception as e:
            errors.append(f"Points de contrôle : {e}")

    df_pl = data["PLANS"]
    if not df_pl.empty:
        col_id   = next((c for c in df_pl.columns if 'ID Plan'     in c), None)
        col_anom = next((c for c in df_pl.columns if 'Anomalie'    in c), None)
        col_ag   = next((c for c in df_pl.columns if 'Agence'      in c or 'Site' in c), None)
        col_resp = next((c for c in df_pl.columns if 'Responsable' in c), None)
        col_ech  = next((c for c in df_pl.columns if 'Échéance'    in c or 'echeance' in c.lower()), None)
        col_stat = next((c for c in df_pl.columns if 'Statut'      in c), None)

        rows = []
        for _, r in df_pl.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source', '')),
                "id_plan":        str(r[col_id])   if col_id   else None,
                "id_anomalie":    str(r[col_anom]) if col_anom else None,
                "agence":         str(r[col_ag])   if col_ag   else None,
                "responsable":    str(r[col_resp]) if col_resp else None,
                "date_echeance":  _safe_date(r[col_ech]) if col_ech else None,
                "statut":         str(r[col_stat]) if col_stat else None,
            })
        try:
            supabase.table("plans_action").insert(rows).execute()
        except Exception as e:
            errors.append(f"Plans d'action : {e}")

    return errors


def _safe_date(val):
    try:
        if pd.isna(val): return None
        if isinstance(val, (datetime, date)): return str(val)[:10]
        return str(pd.to_datetime(val))[:10]
    except Exception:
        return None


def _safe_num(val):
    try:
        v = pd.to_numeric(val, errors='coerce')
        return None if pd.isna(v) else float(v)
    except Exception:
        return None

# ============================================================
# CHARGEMENT DEPUIS SUPABASE
# ============================================================
@st.cache_data(ttl=120)
def load_from_supabase():
    """Charge toutes les tables depuis Supabase (cache 2 min)."""
    def fetch(table):
        try:
            res = supabase.table(table).select("*").execute()
            return pd.DataFrame(res.data) if res.data else pd.DataFrame()
        except Exception:
            return pd.DataFrame()

    return {
        "MISSIONS":  fetch("missions"),
        "POINTS":    fetch("points_controle"),
        "ANOMALIES": fetch("anomalies"),
        "PLANS":     fetch("plans_action"),
    }

# ============================================================
# INTERFACE — BARRE LATÉRALE (RESTRUCTURÉE)
# ============================================================
st.title("🛡️ Espace Chef de Département CI — Groupe SKAB")
st.subheader("Pilotage, Validation Métier et Consolidation des Missions 2026")

with st.sidebar:
    st.header("⚙️ Sources de données")

    source_mode = st.radio(
        "Mode de consultation :",
        ["📂 Fichiers Excel (Import local)", "☁️ Supabase (Base consolidée)"],
        index=1 if SUPABASE_OK else 0
    )

    st.divider()

    # --- ZONE D'UPLOAD ---
    st.markdown("📥 **Importation Terrain**")
    uploaded_files = st.file_uploader(
        "Déposez les fichiers des contrôleurs (.xlsx) :",
        type="xlsx",
        accept_multiple_files=True
    )

    if source_mode == "☁️ Supabase (Base consolidée)":
        if not SUPABASE_OK:
            st.error("❌ Supabase non configuré.")
        elif uploaded_files:
            st.info("📤 Fichiers détectés. Envoyer vers Supabase ?")
            if st.button("☁️ Envoyer vers Supabase", use_container_width=True, type="primary"):
                with st.spinner("Envoi en cours…"):
                    data_tmp = process_consolidation(uploaded_files)
                    errs = push_to_supabase(data_tmp, uploaded_files)
                if errs:
                    for e in errs: st.error(e)
                else:
                    st.success("✅ Données envoyées !")
                    load_from_supabase.clear()

    st.divider()

    # --- FILTRES (COLLAPSIBLE) ---
    with st.expander("🔎 **Filtres Avancés**", expanded=True):
        
        # Période
        periode_type = st.selectbox("Période", ["Toutes les données", "Mois", "Trimestre", "Année"])
        periode_val  = None

        if periode_type == "Mois":
            periode_val = st.selectbox("Choisir le mois :", [
                "Janvier","Février","Mars","Avril","Mai","Juin",
                "Juillet","Août","Septembre","Octobre","Novembre","Décembre"
            ], key="mois_filter")
        elif periode_type == "Trimestre":
            periode_val = st.selectbox("Choisir le trimestre :", 
                ["T1 (Jan-Mar)","T2 (Avr-Jun)","T3 (Jul-Sep)","T4 (Oct-Déc)"], key="trim_filter")
        elif periode_type == "Année":
            periode_val = st.selectbox("Choisir l'année :", 
                [str(y) for y in range(2023, 2028)], key="year_filter")

    st.divider()
    
    st.caption("Direction de l'Audit & Contrôle Interne")
    st.caption("© 2026 Groupe SKAB Nutrition")

# ============================================================
# CHARGEMENT DES DONNÉES
# ============================================================
if source_mode == "📂 Fichiers Excel (Import local)":
    if not uploaded_files:
        st.info("👋 Déposez les fichiers de contrôle des filiales pour initialiser le tableau de bord.")
        st.stop()
    data = process_consolidation(uploaded_files)
else:
    if not SUPABASE_OK:
        st.error("Connexion Supabase non disponible. Vérifiez vos secrets.")
        st.stop()

    with st.spinner("Chargement depuis Supabase…"):
        data = load_from_supabase()

    if uploaded_files:
        data_local = process_consolidation(uploaded_files)
        for key in data:
            parts = [df for df in [data[key], data_local[key]] if not df.empty]
            data[key] = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
        st.info(f"👁️ Prévisualisation : données Supabase + {len(uploaded_files)} fichier(s) uploadé(s).")

    if data["ANOMALIES"].empty and data["MISSIONS"].empty:
        st.info("📭 La base Supabase est vide. Déposez des fichiers et envoyez-les.")
        st.stop()

df_mis  = data["MISSIONS"]
df_pts  = data["POINTS"]
df_anom = data["ANOMALIES"]
df_plan = data["PLANS"]

# ============================================================
# NORMALISATION DES COLONNES
# ============================================================
COL_MAP = {
    "impact":    ["Impact","impact_financier","Impact Financier"],
    "criticite": ["Criticité","criticite","Criticite","Niveau de Criticité"],
    "domaine":   ["Domaine","domaine","Type"],
    "pays":      ["Pays","pays"],
    "agence":    ["Agence","agence","Site / Entité","Site","Entité"],
    "statut":    ["Statut","statut"],
    "date":      ["Date","date_detection","Date de Détection","created_at"],
    "conformite":["Taux de Conformité","taux_conformite","conform","Conformité"],
    "num_mis":   ["N° Mission","num_mission","Mission"],
    "id_anom":   ["ID Anomalie","id_anomalie"],
    "controleur":["Contrôleur","controleur","Agent","agent","Auditeur","auditeur"],
}

def find_col(df, key):
    for candidate in COL_MAP.get(key, []):
        match = next((c for c in df.columns if candidate.lower() in c.lower()), None)
        if match: return match
    return None

col_impact  = find_col(df_anom, "impact")
col_crit    = find_col(df_anom, "criticite")
col_domaine = find_col(df_anom, "domaine")
col_pays    = find_col(df_anom, "pays")
col_ag_a    = find_col(df_anom, "agence")
col_stat_a  = find_col(df_anom, "statut")
col_date_a  = find_col(df_anom, "date")
col_ctrl    = find_col(df_anom, "controleur")
col_tx_conf = find_col(df_mis,  "conformite")
col_ag_m    = find_col(df_mis,  "agence")

# ============================================================
# INITIALISER LES VARIABLES DE FILTRE DANS SESSION STATE
# ============================================================
if "agence_sel" not in st.session_state:
    st.session_state.agence_sel = None
if "ctrl_sel" not in st.session_state:
    st.session_state.ctrl_sel = []

# ============================================================
# RÉCUPÉRER LES VALEURS DISPONIBLES
# ============================================================
all_agences = []
all_controleurs = []

for df in [df_anom, df_mis]:
    if col_ag_a and col_ag_a in df.columns:
        all_agences += df[col_ag_a].dropna().astype(str).unique().tolist()

if col_ctrl and col_ctrl in df_anom.columns:
    all_controleurs = sorted(df_anom[col_ctrl].dropna().astype(str).unique().tolist())

all_agences = sorted(set(all_agences))

# ============================================================
# FILTRES MÉTIER DANS LA SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("---")
    st.header("📋 Filtres Métier")
    
    if all_agences:
        st.session_state.agence_sel = st.selectbox(
            "🏢 Agence / Entité :",
            ["🌍 Toutes les agences"] + all_agences,
            key="agence_filter",
            index=0
        )
    
    if all_controleurs:
        st.session_state.ctrl_sel = st.multiselect(
            "👤 Contrôleur(s) :",
            all_controleurs,
            key="ctrl_filter"
        )

agence_sel = st.session_state.agence_sel
ctrl_sel = st.session_state.ctrl_sel

# ============================================================
# FONCTIONS DE FILTRAGE
# ============================================================
def filter_by_agence(df, col):
    if agence_sel and agence_sel != "🌍 Toutes les agences" and col and col in df.columns:
        return df[df[col].astype(str) == agence_sel]
    return df

def filter_by_controleurs(df, col):
    if ctrl_sel and col and col in df.columns:
        return df[df[col].astype(str).isin(ctrl_sel)]
    return df

# ============================================================
# FILTRE TEMPOREL
# ============================================================
MOIS_MAP = {
    "Janvier":1,"Février":2,"Mars":3,"Avril":4,"Mai":5,"Juin":6,
    "Juillet":7,"Août":8,"Septembre":9,"Octobre":10,"Novembre":11,"Décembre":12
}
TRIM_MAP = {
    "T1 (Jan-Mar)":[1,2,3], "T2 (Avr-Jun)":[4,5,6],
    "T3 (Jul-Sep)":[7,8,9], "T4 (Oct-Déc)":[10,11,12]
}

def filter_by_period(df, col_date):
    if periode_type == "Toutes les données" or col_date is None or col_date not in df.columns:
        return df
    df = df.copy()
    df["_date_parsed"] = pd.to_datetime(df[col_date], errors="coerce")
    df = df.dropna(subset=["_date_parsed"])
    if periode_type == "Mois"       and periode_val:
        df = df[df["_date_parsed"].dt.month == MOIS_MAP[periode_val]]
    elif periode_type == "Trimestre" and periode_val:
        df = df[df["_date_parsed"].dt.month.isin(TRIM_MAP[periode_val])]
    elif periode_type == "Année"     and periode_val:
        df = df[df["_date_parsed"].dt.year == int(periode_val)]
    return df.drop(columns=["_date_parsed"])

# Appliquer les filtres
df_anom_f = filter_by_agence(filter_by_period(filter_by_controleurs(df_anom, col_ctrl), col_date_a), col_ag_a)
df_mis_f  = filter_by_agence(df_mis, col_ag_m)

# ============================================================
# BADGE DE FILTRES ACTIFS
# ============================================================
filtres_actifs = []
if agence_sel and agence_sel != "🌍 Toutes les agences": 
    filtres_actifs.append(f"🏢 {agence_sel}")
if ctrl_sel: 
    filtres_actifs.append(f"👤 {len(ctrl_sel)} contrôleur(s)")
if periode_type != "Toutes les données":              
    filtres_actifs.append(f"📅 {periode_type} : {periode_val}")

if filtres_actifs:
    st.markdown("**Filtres appliqués :** " + " | ".join(filtres_actifs))
    st.divider()

# ============================================================
# INTERFACE AVEC ONGLETS (TAB-BASED)
# ============================================================
tab1, tab2, tab3, tab4, tab5 = st.tabs([
    "📊 Vue d'Ensemble",
    "🔴 Anomalies EN COURS",
    "🏢 Par Agence",
    "📅 Tendances",
    "🛠️ Diagnostic"
])

# ============================================================
# TAB 1 : VUE D'ENSEMBLE
# ============================================================
with tab1:
    st.markdown("### 📊 Indicateurs Clés de Risque")
    k1, k2, k3, k4 = st.columns(4)

    with k1:
        impact_total = 0
        if col_impact and not df_anom_f.empty:
            impact_total = pd.to_numeric(df_anom_f[col_impact], errors='coerce').fillna(0).sum()
        st.metric("💰 Risque Financier", f"{impact_total:,.0f} FCFA")

    with k2:
        nb_critiques = 0
        if col_crit and not df_anom_f.empty:
            nb_critiques = df_anom_f[df_anom_f[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
        st.metric("🔴 Critiques", nb_critiques, 
                  delta="⚠️ Urgent" if nb_critiques > 0 else "✅ OK", delta_color="inverse")

    with k3:
        conformite_moyenne = 0
        if col_tx_conf and not df_mis_f.empty:
            raw_mean = pd.to_numeric(df_mis_f[col_tx_conf], errors='coerce').mean()
            if pd.notna(raw_mean):
                conformite_moyenne = raw_mean * 100 if raw_mean <= 1.0 else raw_mean
        st.metric("✅ Conformité", f"{conformite_moyenne:.1f}%" if conformite_moyenne > 0 else "N/A")

    with k4:
        nb_en_cours = 0
        if col_stat_a and not df_anom_f.empty:
            nb_en_cours = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains('EN COURS', na=False)].shape[0]
        st.metric("⏳ EN COURS", nb_en_cours, delta="À surveiller" if nb_en_cours > 0 else None)

    st.divider()

    # Graphiques principaux
    g1, g2 = st.columns(2)

    with g1:
        st.markdown('<div class="section-title">📊 Anomalies par Domaine</div>', unsafe_allow_html=True)
        if not df_anom_f.empty and col_domaine:
            df_dom = df_anom_f.dropna(subset=[col_domaine])
            color_map = {
                '🔴 Critique':'#e63946','Critique':'#e63946',
                '🟠 Majeur':'#f4a261','Majeur':'#f4a261',
                '🟡 Mineur':'#e9c46a','Mineur':'#e9c46a',
                '🟢 Faible':'#2a9d8f','Faible':'#2a9d8f'
            }
            fig = px.bar(df_dom, x=col_domaine, color=col_crit if col_crit else None,
                         barmode='group', color_discrete_map=color_map)
            fig.update_layout(height=300, margin=dict(l=0,r=0,t=20,b=0),
                              xaxis_title=None, yaxis_title="Nombre", plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Aucune donnée")

    with g2:
        st.markdown('<div class="section-title">🌍 Distribution par Pays</div>', unsafe_allow_html=True)
        if not df_anom_f.empty and col_pays:
            df_pays = df_anom_f.dropna(subset=[col_pays]).groupby(col_pays).size().reset_index(name="Nb")
            fig = px.pie(df_pays, values="Nb", names=col_pays, hole=0.4)
            fig.update_layout(height=300, margin=dict(l=0,r=0,t=20,b=0))
            st.plotly_chart(fig, use_container_width=True)
        else:
            st.info("Aucune donnée")

# ============================================================
# TAB 2 : ANOMALIES EN COURS
# ============================================================
with tab2:
    if not df_anom_f.empty and col_stat_a:
        df_en_cours = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains("EN COURS", na=False)].copy()

        if df_en_cours.empty:
            st.success("✅ Aucune anomalie EN COURS pour ce filtre.")
        else:
            c1, c2, c3 = st.columns(3)
            c1.metric("Total EN COURS", len(df_en_cours))
            impact_ec = pd.to_numeric(df_en_cours[col_impact], errors='coerce').fillna(0).sum() if col_impact else 0
            c2.metric("Impact EN COURS", f"{impact_ec:,.0f} FCFA")
            if col_crit:
                nb_crit = df_en_cours[df_en_cours[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
                c3.metric("Critiques", nb_crit)

            st.divider()

            # Graphe tendance
            if col_date_a and col_date_a in df_en_cours.columns:
                df_ec_tmp = df_en_cours.copy()
                df_ec_tmp["_mois"] = pd.to_datetime(df_ec_tmp[col_date_a], errors="coerce").dt.to_period("M").astype(str)
                df_trend = df_ec_tmp.dropna(subset=["_mois"]).groupby("_mois").size().reset_index(name="Nb")
                if not df_trend.empty:
                    fig = px.line(df_trend, x="_mois", y="Nb", markers=True, title="Tendance des EN COURS")
                    fig.update_layout(height=280, margin=dict(l=0,r=0,t=40,b=0), plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig, use_container_width=True)

            st.markdown("**📋 Liste détaillée :**")
            cols_show = [c for c in [
                find_col(df_en_cours, "id_anom"), col_ctrl, col_ag_a, col_crit, col_domaine, col_impact
            ] if c and c in df_en_cours.columns]
            
            st.dataframe(df_en_cours[cols_show] if cols_show else df_en_cours, 
                        hide_index=True, use_container_width=True, height=300)

            csv = df_en_cours.to_csv(index=False).encode("utf-8")
            st.download_button(
                "⬇️ Exporter EN COURS (.csv)",
                data=csv,
                file_name=f"SKAB_EN_COURS_{date.today()}.csv",
                mime="text/csv"
            )
    else:
        st.info("Aucune colonne 'Statut' détectée.")

# ============================================================
# TAB 3 : ANALYSE PAR AGENCE
# ============================================================
with tab3:
    agences_dispo = []
    if col_ag_a and not df_anom.empty:
        agences_dispo = sorted(df_anom[col_ag_a].dropna().astype(str).unique().tolist())

    if not agences_dispo:
        st.info("Aucune colonne Agence détectée.")
    else:
        ag_sel = st.selectbox("Sélectionner une agence :", agences_dispo, key="ag_detail_tab")
        df_ag = df_anom_f[df_anom_f[col_ag_a].astype(str) == ag_sel].copy() if col_ag_a else df_anom_f.copy()

        a1, a2, a3, a4 = st.columns(4)
        a1.metric("Total", len(df_ag))
        if col_crit:
            nb_crit = df_ag[df_ag[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
            a2.metric("Critiques", nb_crit)
        if col_stat_a:
            nb_ec = df_ag[df_ag[col_stat_a].astype(str).str.upper().str.contains('EN COURS', na=False)].shape[0]
            a3.metric("EN COURS", nb_ec)
        if col_impact:
            impact = pd.to_numeric(df_ag[col_impact], errors='coerce').fillna(0).sum()
            a4.metric("Impact", f"{impact:,.0f} FCFA")

        st.divider()

        # Graphe domaines
        if col_domaine and not df_ag.empty:
            df_dom_ag = df_ag.dropna(subset=[col_domaine]).groupby(col_domaine).size().reset_index(name="Nb")
            if not df_dom_ag.empty:
                fig = px.bar_polar(df_dom_ag, r="Nb", theta=col_domaine, color="Nb", 
                                   color_continuous_scale="Reds", title=f"Profil — {ag_sel}")
                fig.update_layout(height=350, margin=dict(l=0,r=0,t=50,b=0))
                st.plotly_chart(fig, use_container_width=True)

        st.markdown("**📋 Tableau détaillé :**")
        cols_ag = [c for c in [find_col(df_ag, "id_anom"), col_ctrl, col_crit, col_domaine, col_impact, col_stat_a] 
                   if c and c in df_ag.columns]
        st.dataframe(df_ag[cols_ag] if cols_ag else df_ag, hide_index=True, use_container_width=True, height=300)

# ============================================================
# TAB 4 : TENDANCES
# ============================================================
with tab4:
    if col_date_a and not df_anom_f.empty:
        df_per = df_anom_f.copy()
        df_per["_date"] = pd.to_datetime(df_per[col_date_a], errors="coerce")
        df_per = df_per.dropna(subset=["_date"])

        if df_per.empty:
            st.info("Aucune date valide.")
        else:
            df_per["Mois"] = df_per["_date"].dt.strftime("%Y-%m")
            p_type = st.radio("Granularité :", ["Mois", "Trimestre", "Année"], horizontal=True)

            if p_type == "Mois":
                df_grp = df_per.groupby("Mois").size().reset_index(name="Nb").sort_values("Mois")
                col_p = "Mois"
            else:
                df_per["Trim"] = "T" + df_per["_date"].dt.quarter.astype(str) + "-" + df_per["_date"].dt.year.astype(str)
                df_grp = df_per.groupby("Trim").size().reset_index(name="Nb")
                col_p = "Trim"

            p1, p2 = st.columns(2)
            with p1:
                fig_p1 = px.bar(df_grp, x=col_p, y="Nb", title="Nombre d'anomalies")
                fig_p1.update_layout(height=300, plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_p1, use_container_width=True)

            with p2:
                if col_impact:
                    df_impact = df_per.groupby(col_p if p_type == "Mois" else "Trim").apply(
                        lambda x: pd.to_numeric(x[col_impact], errors='coerce').sum()
                    ).reset_index(name="Impact")
                    df_impact.columns = [col_p, "Impact"] if p_type == "Mois" else ["Trim", "Impact"]
                    fig_p2 = px.line(df_impact, x=df_impact.columns[0], y="Impact", markers=True, 
                                    title="Impact financier")
                    fig_p2.update_layout(height=300, plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_p2, use_container_width=True)
    else:
        st.info("Aucune date disponible.")

# ============================================================
# TAB 5 : DIAGNOSTIC
# ============================================================
with tab5:
    st.markdown("### 🛠️ Diagnostic Qualité")

    alertes = []
    
    if not df_anom_f.empty and col_crit and col_impact:
        lignes_anom = df_anom_f[
            (df_anom_f[col_crit].astype(str).str.contains('Critique|🔴|Majeur|🟠', na=False)) &
            (pd.to_numeric(df_anom_f[col_impact], errors='coerce').fillna(0) == 0)
        ]
        for _, row in lignes_anom.iterrows():
            id_val = row.get('ID Anomalie', row.get('id_anomalie', 'N/A'))
            alertes.append(f"⚠️ Anomalie {id_val} sans impact financier")

    if alertes:
        for a in alertes[:10]:  # Limiter à 10
            st.warning(a)
    else:
        st.success("✅ Données validées")

    st.divider()
    st.markdown("**📌 Registre complet (filtré) :**")
    cols_diag = [c for c in [find_col(df_anom_f, "id_anom"), col_ctrl, col_ag_a, col_crit, col_domaine, col_impact, col_stat_a]
                 if c and c in df_anom_f.columns]
    if cols_diag:
        st.dataframe(df_anom_f[cols_diag].sort_values(by=col_impact or df_anom_f.columns[0], ascending=False) 
                    if col_impact and col_impact in df_anom_f.columns 
                    else df_anom_f[cols_diag], 
                    hide_index=True, use_container_width=True, height=400)
    else:
        st.dataframe(df_anom_f, hide_index=True, use_container_width=True, height=400)

# ============================================================
# FOOTER : EXPORT
# ============================================================
st.divider()
st.markdown("### 📤 Export Fichier Maître")

if st.button("🏗️ Générer le fichier consolidé", type="primary", use_container_width=True):
    output = io.BytesIO()
    with pd.ExcelWriter(output, engine='xlsxwriter') as writer:
        # Métadonnées
        df_meta = pd.DataFrame({
            "RAPPORT": ["Destinataire", "Généré par", "Date", "Filtres", "Statut"],
            "VALEURS": [
                "Management",
                "Dashboard SKAB",
                datetime.now().strftime("%d/%m/%Y %H:%M"),
                " | ".join(filtres_actifs) if filtres_actifs else "Aucun",
                "VALIDÉ"
            ]
        })
        df_meta.to_excel(writer, sheet_name="ACCUEIL", index=False)

        # Onglets données
        sheets = {
            "MISSIONS": data["MISSIONS"],
            "POINTS": data["POINTS"],
            "ANOMALIES": data["ANOMALIES"],
            "PLANS": data["PLANS"],
        }

        if col_stat_a and not df_anom.empty:
            df_ec = df_anom[df_anom[col_stat_a].astype(str).str.upper().str.contains("EN COURS", na=False)]
            if not df_ec.empty:
                sheets["ANOMALIES_EN_COURS"] = df_ec

        for name, df in sheets.items():
            if df is not None and not df.empty:
                df.to_excel(writer, sheet_name=name, index=False)
                ws = writer.sheets[name]
                for i, col in enumerate(df.columns):
                    ws.set_column(i, i, get_safe_len(df[col], col))

    st.success("✅ Fichier généré !")
    st.download_button(
        f"💾 SKAB_CONSOLIDÉ_{datetime.now().strftime('%Y%m%d')}.xlsx",
        data=output.getvalue(),
        file_name=f"SKAB_CONSOLIDÉ_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.caption("© 2026 Groupe SKAB — Direction Audit & Contrôle Interne")
