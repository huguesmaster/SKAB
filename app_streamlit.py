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
    layout="wide"
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
    </style>
""", unsafe_allow_html=True)

# ============================================================
# CONNEXION SUPABASE
# ============================================================
@st.cache_resource
def init_supabase() -> Client:
    """
    Initialise le client Supabase à partir des secrets Streamlit Cloud.
    Configurez dans Settings > Secrets :
        SUPABASE_URL = "https://xxxx.supabase.co"
        SUPABASE_KEY = "eyJhbGci..."
    """
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
def load_and_clean(file, sheet):
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
        if df_raw.empty:
            return pd.DataFrame()

        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = [str(val).strip() for val in row.values]
            if sheet == "MES_MISSIONS"     and any("N° Mission"  in s for s in row_str): header_idx = idx; break
            elif sheet == "POINTS_CONTROLE"  and any("ID Point"    in s for s in row_str): header_idx = idx; break
            elif sheet == "ANOMALIES"        and any("ID Anomalie" in s for s in row_str): header_idx = idx; break
            elif sheet == "PLANS_ACTION"     and any("ID Plan"     in s for s in row_str): header_idx = idx; break

        if header_idx is None:
            for idx, row in df_raw.iterrows():
                row_str = [str(val).strip() for val in row.values]
                if any("ID" in s or "N°" in s or "Code" in s for s in row_str):
                    header_idx = idx; break

        if header_idx is None:
            header_idx = 0

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
    except Exception:
        return pd.DataFrame()


def process_consolidation(files):
    all_data = {"MISSIONS": [], "POINTS": [], "ANOMALIES": [], "PLANS": []}
    for f in files:
        f.seek(0); all_data["MISSIONS"].append(load_and_clean(f, "MES_MISSIONS"))
        f.seek(0); all_data["POINTS"].append(load_and_clean(f, "POINTS_CONTROLE"))
        f.seek(0); all_data["ANOMALIES"].append(load_and_clean(f, "ANOMALIES"))
        f.seek(0); all_data["PLANS"].append(load_and_clean(f, "PLANS_ACTION"))

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
    """
    Pousse les dataframes consolidés vers Supabase.
    Adapte les noms de colonnes du template SKAB aux colonnes de la BDD.
    """
    errors = []

    # --- MISSIONS ---
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

    # --- ANOMALIES ---
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

    # --- POINTS DE CONTRÔLE ---
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

    # --- PLANS D'ACTION ---
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
# INTERFACE — BARRE LATÉRALE
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

    if source_mode == "📂 Fichiers Excel (Import local)":
        uploaded_files = st.file_uploader(
            "Déposez les fichiers des contrôleurs (.xlsx) :",
            type="xlsx",
            accept_multiple_files=True
        )
        if SUPABASE_OK and uploaded_files:
            st.divider()
            if st.button("☁️ Envoyer vers Supabase", use_container_width=True, type="primary"):
                with st.spinner("Envoi en cours…"):
                    data_tmp = process_consolidation(uploaded_files)
                    errs = push_to_supabase(data_tmp, uploaded_files)
                if errs:
                    for e in errs: st.error(e)
                else:
                    st.success("✅ Données envoyées vers Supabase !")
                    load_from_supabase.clear()
    else:
        uploaded_files = []
        if not SUPABASE_OK:
            st.error("❌ Supabase non configuré.\nAjoutez SUPABASE_URL et SUPABASE_KEY dans les secrets.")

    st.divider()

    # --- FILTRES GLOBAUX ---
    st.header("🔎 Filtres globaux")

    # Période
    periode_type = st.selectbox("Période", ["Toutes les données", "Mois", "Trimestre", "Année"])
    periode_val  = None

    if periode_type == "Mois":
        periode_val = st.selectbox("Choisir le mois :", [
            "Janvier","Février","Mars","Avril","Mai","Juin",
            "Juillet","Août","Septembre","Octobre","Novembre","Décembre"
        ])
    elif periode_type == "Trimestre":
        periode_val = st.selectbox("Choisir le trimestre :", ["T1 (Jan-Mar)","T2 (Avr-Jun)","T3 (Jul-Sep)","T4 (Oct-Déc)"])
    elif periode_type == "Année":
        periode_val = st.selectbox("Choisir l'année :", [str(y) for y in range(2023, 2028)])

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
    if data["ANOMALIES"].empty and data["MISSIONS"].empty:
        st.info("📭 La base Supabase est vide. Importez d'abord des fichiers via le mode Excel.")
        st.stop()

df_mis  = data["MISSIONS"]
df_pts  = data["POINTS"]
df_anom = data["ANOMALIES"]
df_plan = data["PLANS"]

# ============================================================
# NORMALISATION DES COLONNES (compatibilité Excel ↔ Supabase)
# ============================================================
# On mappe les noms de colonnes Supabase vers des alias communs
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
col_tx_conf = find_col(df_mis,  "conformite")
col_ag_m    = find_col(df_mis,  "agence")

# ============================================================
# FILTRE PAR AGENCE (dynamique selon les données disponibles)
# ============================================================
all_agences = []
for df, col in [(df_anom, col_ag_a), (df_mis, col_ag_m)]:
    if col and not df.empty:
        all_agences += df[col].dropna().astype(str).unique().tolist()
all_agences = sorted(set(all_agences))

agence_sel = None
if all_agences:
    with st.sidebar:
        agence_sel = st.selectbox(
            "🏢 Filtrer par Agence / Entité :",
            ["Toutes les agences"] + all_agences
        )

def filter_by_agence(df, col):
    if agence_sel and agence_sel != "Toutes les agences" and col and col in df.columns:
        return df[df[col].astype(str) == agence_sel]
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
df_anom_f = filter_by_agence(filter_by_period(df_anom, col_date_a), col_ag_a)
df_mis_f  = filter_by_agence(df_mis, col_ag_m)

# ============================================================
# BADGE DE FILTRES ACTIFS
# ============================================================
filtres_actifs = []
if agence_sel and agence_sel != "Toutes les agences": filtres_actifs.append(f"🏢 {agence_sel}")
if periode_type != "Toutes les données":              filtres_actifs.append(f"📅 {periode_type} : {periode_val}")

if filtres_actifs:
    st.info("**Filtres actifs :** " + " | ".join(filtres_actifs))

# ============================================================
# KPI PRINCIPAUX
# ============================================================
st.markdown("### 📊 Indicateurs de Risques")
k1, k2, k3, k4, k5 = st.columns(5)

with k1:
    impact_total = 0
    if col_impact and not df_anom_f.empty:
        impact_total = pd.to_numeric(df_anom_f[col_impact], errors='coerce').fillna(0).sum()
    st.metric("Risque Financier Cumulé", f"{impact_total:,.0f} FCFA")

with k2:
    nb_critiques = 0
    if col_crit and not df_anom_f.empty:
        nb_critiques = df_anom_f[df_anom_f[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
    st.metric("Anomalies Critiques", nb_critiques,
              delta="Action urgente" if nb_critiques > 0 else None, delta_color="inverse")

with k3:
    conformite_moyenne = 0
    if col_tx_conf and not df_mis_f.empty:
        raw_mean = pd.to_numeric(df_mis_f[col_tx_conf], errors='coerce').mean()
        if pd.notna(raw_mean):
            conformite_moyenne = raw_mean * 100 if raw_mean <= 1.0 else raw_mean
    st.metric("Taux de Conformité Moyen", f"{conformite_moyenne:.1f}%" if conformite_moyenne > 0 else "N/A")

with k4:
    # ── NOUVEAU KPI : Anomalies EN COURS ──
    nb_en_cours = 0
    if col_stat_a and not df_anom_f.empty:
        nb_en_cours = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains('EN COURS', na=False)].shape[0]
    st.metric("Anomalies EN COURS", nb_en_cours,
              delta="À surveiller" if nb_en_cours > 0 else None, delta_color="inverse")

with k5:
    nb_fichiers = len(uploaded_files) if source_mode == "📂 Fichiers Excel (Import local)" else "☁️ BDD"
    st.metric("Source", nb_fichiers)

st.divider()

# ============================================================
# GRAPHIQUES STANDARDS
# ============================================================
g1, g2 = st.columns(2)

with g1:
    st.markdown('<div class="section-title">🔍 Volume d\'Anomalies par Domaine</div>', unsafe_allow_html=True)
    if not df_anom_f.empty and col_domaine:
        df_graph1 = df_anom_f.dropna(subset=[col_domaine])
        color_opt = {
            '🔴 Critique':'#e63946','Critique':'#e63946',
            '🟠 Majeur':'#f4a261','Majeur':'#f4a261',
            '🟡 Mineur':'#e9c46a','Mineur':'#e9c46a',
            '🟢 Faible':'#2a9d8f','Faible':'#2a9d8f'
        }
        fig = px.bar(df_graph1, x=col_domaine, color=col_crit if col_crit else None,
                     barmode='group', color_discrete_map=color_opt)
        fig.update_layout(height=350, margin=dict(l=0,r=0,t=20,b=0),
                          xaxis_title=None, yaxis_title="Nombre d'écarts",
                          plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aucune anomalie détectée pour ce filtre.")

with g2:
    st.markdown('<div class="section-title">🌍 Alertes par Pays / Entité</div>', unsafe_allow_html=True)
    if not df_anom_f.empty and col_pays:
        df_pays_s = df_anom_f.dropna(subset=[col_pays]).groupby(col_pays).size().reset_index(name="Anomalies")
        fig = px.pie(df_pays_s, values="Anomalies", names=col_pays,
                     hole=.4, color_discrete_sequence=px.colors.qualitative.Safe)
        fig.update_layout(height=350, margin=dict(l=0,r=0,t=20,b=0))
        st.plotly_chart(fig, use_container_width=True)
    else:
        st.info("Aucune donnée géographique à cartographier.")

st.divider()

# ============================================================
# ██  NOUVEAU BLOC 1 : ANOMALIES "EN COURS"
# ============================================================
st.markdown("### 🔴 Suivi des Anomalies EN COURS")

if not df_anom_f.empty and col_stat_a:
    df_en_cours = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains("EN COURS", na=False)].copy()

    if df_en_cours.empty:
        st.success("✅ Aucune anomalie avec statut EN COURS pour ce filtre.")
    else:
        nb = len(df_en_cours)
        impact_ec = 0
        if col_impact:
            impact_ec = pd.to_numeric(df_en_cours[col_impact], errors='coerce').fillna(0).sum()

        c1, c2, c3 = st.columns(3)
        c1.metric("Total EN COURS", nb)
        c2.metric("Impact financier EN COURS", f"{impact_ec:,.0f} FCFA")
        if col_crit:
            nb_crit_ec = df_en_cours[df_en_cours[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
            c3.metric("Dont Critiques EN COURS", nb_crit_ec)

        # Graphe évolution des anomalies EN COURS dans le temps (si date disponible)
        if col_date_a and col_date_a in df_en_cours.columns:
            df_en_cours["_mois"] = pd.to_datetime(df_en_cours[col_date_a], errors="coerce").dt.to_period("M").astype(str)
            df_trend = df_en_cours.dropna(subset=["_mois"]).groupby("_mois").size().reset_index(name="Nb EN COURS")
            if not df_trend.empty:
                fig_trend = px.line(df_trend, x="_mois", y="Nb EN COURS",
                                    markers=True, title="Évolution mensuelle des anomalies EN COURS",
                                    color_discrete_sequence=["#e63946"])
                fig_trend.update_layout(height=280, margin=dict(l=0,r=0,t=40,b=0),
                                        xaxis_title="Mois", plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_trend, use_container_width=True)

        # Tableau détaillé
        st.markdown("**📋 Liste détaillée :**")
        cols_ec = [c for c in [
            find_col(df_en_cours, "id_anom"),
            col_ag_a, col_pays, col_crit, col_domaine,
            find_col(df_en_cours, "date"), col_impact,
            next((c for c in df_en_cours.columns if 'Description' in c or 'description' in c), None)
        ] if c and c in df_en_cours.columns]

        # Coloration conditionnelle selon criticité
        def color_crit(val):
            val = str(val)
            if   'Critique' in val or '🔴' in val: return 'background-color:#ffe0e0'
            elif 'Majeur'   in val or '🟠' in val: return 'background-color:#fff3e0'
            elif 'Mineur'   in val or '🟡' in val: return 'background-color:#fffde7'
            return ''

        df_show = df_en_cours[cols_ec] if cols_ec else df_en_cours
        if col_crit and col_crit in df_show.columns:
            st.dataframe(
                df_show.style.applymap(color_crit, subset=[col_crit]),
                hide_index=True, use_container_width=True, height=350
            )
        else:
            st.dataframe(df_show, hide_index=True, use_container_width=True, height=350)

        # Export CSV des EN COURS
        csv_ec = df_en_cours.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Exporter les anomalies EN COURS (.csv)",
            data=csv_ec,
            file_name=f"SKAB_anomalies_EN_COURS_{date.today()}.csv",
            mime="text/csv"
        )
else:
    st.info("Aucune colonne 'Statut' détectée dans les anomalies.")

st.divider()

# ============================================================
# ██  NOUVEAU BLOC 2 : CONSULTATION PAR AGENCE
# ============================================================
st.markdown("### 🏢 Analyse par Agence / Entité")

agences_dispo = []
if col_ag_a and not df_anom.empty:
    agences_dispo = sorted(df_anom[col_ag_a].dropna().astype(str).unique().tolist())

if not agences_dispo:
    st.info("Aucune colonne Agence/Site détectée dans les anomalies.")
else:
    ag_tab = st.selectbox("Sélectionner une agence pour l'analyse détaillée :", agences_dispo, key="ag_detail")
    df_ag = df_anom[df_anom[col_ag_a].astype(str) == ag_tab].copy()

    a1, a2, a3, a4 = st.columns(4)
    a1.metric("Anomalies totales", len(df_ag))

    nb_crit_ag = 0
    if col_crit:
        nb_crit_ag = df_ag[df_ag[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
    a2.metric("Dont Critiques", nb_crit_ag)

    nb_ec_ag = 0
    if col_stat_a:
        nb_ec_ag = df_ag[df_ag[col_stat_a].astype(str).str.upper().str.contains('EN COURS', na=False)].shape[0]
    a3.metric("EN COURS", nb_ec_ag)

    impact_ag = 0
    if col_impact:
        impact_ag = pd.to_numeric(df_ag[col_impact], errors='coerce').fillna(0).sum()
    a4.metric("Impact Financier", f"{impact_ag:,.0f} FCFA")

    # Graphe radar des domaines pour cette agence
    if col_domaine and not df_ag.empty:
        df_dom_ag = df_ag.dropna(subset=[col_domaine]).groupby(col_domaine).size().reset_index(name="Nb")
        if not df_dom_ag.empty:
            fig_ag = px.bar_polar(df_dom_ag, r="Nb", theta=col_domaine,
                                  color="Nb", color_continuous_scale="Reds",
                                  title=f"Profil d'anomalies — {ag_tab}")
            fig_ag.update_layout(height=380, margin=dict(l=0,r=0,t=50,b=0))
            st.plotly_chart(fig_ag, use_container_width=True)

    st.markdown(f"**📋 Toutes les anomalies de {ag_tab} :**")
    st.dataframe(df_ag, hide_index=True, use_container_width=True, height=300)

st.divider()

# ============================================================
# ██  NOUVEAU BLOC 3 : CONSULTATION PÉRIODIQUE
# ============================================================
st.markdown("### 📅 Analyse Périodique")

if col_date_a and not df_anom.empty:
    df_per = df_anom.copy()
    df_per["_date"] = pd.to_datetime(df_per[col_date_a], errors="coerce")
    df_per = df_per.dropna(subset=["_date"])

    if df_per.empty:
        st.info("Aucune date valide trouvée dans les anomalies.")
    else:
        df_per["Mois"]      = df_per["_date"].dt.strftime("%Y-%m")
        df_per["Trimestre"] = "T" + df_per["_date"].dt.quarter.astype(str) + " " + df_per["_date"].dt.year.astype(str)
        df_per["Année"]     = df_per["_date"].dt.year.astype(str)

        p_type = st.radio("Granularité :", ["Mois","Trimestre","Année"], horizontal=True, key="period_radio")

        df_grp = df_per.groupby(p_type).agg(
            Nb_Anomalies=(col_impact if col_impact else df_per.columns[0], "count"),
            **({ "Impact_Total": (col_impact, lambda x: pd.to_numeric(x, errors='coerce').sum()) } if col_impact else {})
        ).reset_index().sort_values(p_type)

        col_p1, col_p2 = st.columns(2)

        with col_p1:
            fig_p1 = px.bar(df_grp, x=p_type, y="Nb_Anomalies",
                            title=f"Nombre d'anomalies par {p_type.lower()}",
                            color_discrete_sequence=["#1d3557"])
            fig_p1.update_layout(height=300, margin=dict(l=0,r=0,t=40,b=0), plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_p1, use_container_width=True)

        with col_p2:
            if col_impact and "Impact_Total" in df_grp.columns:
                fig_p2 = px.line(df_grp, x=p_type, y="Impact_Total", markers=True,
                                 title=f"Impact financier cumulé par {p_type.lower()}",
                                 color_discrete_sequence=["#e63946"])
                fig_p2.update_layout(height=300, margin=dict(l=0,r=0,t=40,b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_p2, use_container_width=True)
            else:
                st.info("Colonne d'impact financier non détectée.")

        st.markdown(f"**📋 Synthèse par {p_type.lower()} :**")
        st.dataframe(df_grp, hide_index=True, use_container_width=True)

else:
    st.info("Aucune colonne de date détectée dans les anomalies pour l'analyse périodique.")

st.divider()

# ============================================================
# DIAGNOSTIC QUALITÉ
# ============================================================
st.markdown("### 🛠️ Diagnostic Qualité des Données")
q1, q2 = st.columns([1, 1])

with q1:
    st.markdown("**🚨 Incohérences détectées**")
    alertes_qualite = []
    col_num_mis_m = find_col(df_mis,  "num_mis")
    col_num_mis_p = find_col(df_pts,  "num_mis")

    if not df_anom_f.empty and col_crit and col_impact:
        lignes_anormales = df_anom_f[
            (df_anom_f[col_crit].astype(str).str.contains('Critique|🔴|Majeur|🟠', na=False)) &
            (pd.to_numeric(df_anom_f[col_impact], errors='coerce').fillna(0) == 0)
        ]
        for _, row in lignes_anormales.iterrows():
            id_val = row.get('ID Anomalie', row.get('id_anomalie', 'N/A'))
            alertes_qualite.append(
                f"⚠️ Anomalie **{id_val}** [{row[col_crit]}] avec impact financier nul."
            )

    if col_num_mis_m and not df_mis_f.empty:
        for m_id in df_mis_f[col_num_mis_m].dropna().unique():
            if "Une mission" in str(m_id) or str(m_id).startswith("N°"): continue
            if df_pts.empty or col_num_mis_p not in df_pts.columns or m_id not in df_pts[col_num_mis_p].values:
                alertes_qualite.append(f"❌ Mission **{m_id}** sans point de contrôle rattaché.")

    if alertes_qualite:
        for alerte in alertes_qualite: st.warning(alerte)
    else:
        st.success("✅ Diagnostic validé : Structure et intégrité OK.")

with q2:
    st.markdown("**📌 Registre des Anomalies (vue filtrée)**")
    if not df_anom_f.empty:
        cols_show = [c for c in [
            find_col(df_anom_f, "id_anom"), col_crit, col_pays, col_ag_a,
            next((c for c in df_anom_f.columns if 'Description' in c or 'description' in c), None),
            col_impact, col_stat_a
        ] if c and c in df_anom_f.columns]
        sort_col = col_impact if col_impact and col_impact in df_anom_f.columns else df_anom_f.columns[0]
        st.dataframe(
            df_anom_f.sort_values(by=sort_col, ascending=False)[cols_show],
            hide_index=True, use_container_width=True, height=300
        )
    else:
        st.info("Aucune anomalie à lister.")

st.divider()

# ============================================================
# EXPORTATION FICHIER MAÎTRE
# ============================================================
st.header("📤 Génération du Fichier Maître Consolidé")

if st.button("🏗️ Compiler et figer le Fichier Maître Consolidé", type="primary", use_container_width=True):
    output_buffer = io.BytesIO()
    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        df_meta = pd.DataFrame({
            "RAPPORT GLOBAL CI SKAB": [
                "Destinataire Principal","Généré par","Horodatage","Filtres appliqués","Niveau de validation"
            ],
            "MÉTADONNÉES": [
                "M. Élie DIGNOU (DAF)",
                "Chef de Département Contrôle Interne",
                datetime.now().strftime("%d/%m/%Y à %H:%M"),
                " | ".join(filtres_actifs) if filtres_actifs else "Aucun",
                "VÉRIFIÉ ET SCELLÉ"
            ]
        })
        df_meta.to_excel(writer, sheet_name="ACCUEIL_CONSO", index=False)
        writer.sheets["ACCUEIL_CONSO"].set_column('A:B', 40)

        onglets = {
            "CONSO_MISSIONS":       data["MISSIONS"],
            "CONSO_POINTS_CTRL":    data["POINTS"],
            "CONSO_ANOMALIES":      data["ANOMALIES"],
            "CONSO_PLANS_ACTION":   data["PLANS"],
        }

        # Onglet spécial EN COURS
        if col_stat_a and not df_anom.empty:
            df_ec_export = df_anom[df_anom[col_stat_a].astype(str).str.upper().str.contains("EN COURS", na=False)]
            if not df_ec_export.empty:
                onglets["ANOMALIES_EN_COURS"] = df_ec_export

        for sheet_name, dataframe in onglets.items():
            if dataframe is not None and not dataframe.empty:
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                for i, col in enumerate(dataframe.columns):
                    ws.set_column(i, i, get_safe_len(dataframe[col], col))

    st.success("🎉 Fichier Maître structuré avec succès — onglet ANOMALIES_EN_COURS inclus !")
    st.download_button(
        label=f"💾 Télécharger SKAB_MAITRE_CONSO_{datetime.now().strftime('%Y%m%d')}.xlsx",
        data=output_buffer.getvalue(),
        file_name=f"SKAB_MAITRE_CONSO_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.markdown("---")
st.caption("Direction Générale SKAB Nutrition — Application de Contrôle Interne — v2.0 Supabase")
