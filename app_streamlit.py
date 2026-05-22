import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import numpy as np
from datetime import datetime, date, timedelta
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
    .section-title {
        font-size: 18px;
        font-weight: 700;
        color: #1a1a2e;
        border-left: 4px solid #e63946;
        padding-left: 10px;
        margin-bottom: 12px;
        margin-top: 20px;
    }
    .kpi-container { margin-bottom: 20px; }
    .filter-badge {
        display: inline-block;
        background-color: #e63946;
        color: white;
        padding: 5px 10px;
        border-radius: 20px;
        font-size: 12px;
        margin-right: 5px;
        margin-bottom: 5px;
    }
    </style>
""", unsafe_allow_html=True)

# ============================================================
# CONNEXION SUPABASE
# ============================================================
@st.cache_resource
def init_supabase() -> Client:
    """Initialise le client Supabase."""
    url = st.secrets.get("SUPABASE_URL")
    key = st.secrets.get("SUPABASE_KEY")
    if not url or not key:
        return None
    return create_client(url, key)

supabase = init_supabase()
SUPABASE_OK = supabase is not None

# ============================================================
# UTILITAIRES DE DONNÉES
# ============================================================
def load_and_clean(file, sheet):
    """Charge et nettoie un fichier Excel."""
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
        if df_raw.empty:
            return pd.DataFrame()

        header_idx = None
        keywords = {
            "MES_MISSIONS": "N° Mission",
            "POINTS_CONTROLE": "ID Point",
            "ANOMALIES": "ID Anomalie",
            "PLANS_ACTION": "ID Plan"
        }

        for idx, row in df_raw.iterrows():
            row_str = [str(val).strip() for val in row.values]
            if keywords.get(sheet) and any(keywords[sheet] in s for s in row_str):
                header_idx = idx
                break

        if header_idx is None:
            for idx, row in df_raw.iterrows():
                row_str = [str(val).strip() for val in row.values]
                if any("ID" in s or "N°" in s or "Code" in s for s in row_str):
                    header_idx = idx
                    break

        header_idx = header_idx if header_idx is not None else 0

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
    """Consolide les données de plusieurs fichiers."""
    all_data = {
        "MISSIONS": [],
        "POINTS": [],
        "ANOMALIES": [],
        "PLANS": []
    }
    
    for f in files:
        f.seek(0)
        all_data["MISSIONS"].append(load_and_clean(f, "MES_MISSIONS"))
        f.seek(0)
        all_data["POINTS"].append(load_and_clean(f, "POINTS_CONTROLE"))
        f.seek(0)
        all_data["ANOMALIES"].append(load_and_clean(f, "ANOMALIES"))
        f.seek(0)
        all_data["PLANS"].append(load_and_clean(f, "PLANS_ACTION"))

    return {
        k: pd.concat(v, ignore_index=True) if v else pd.DataFrame()
        for k, v in all_data.items()
    }


def push_to_supabase(data: dict, source_files):
    """Pousse les données vers Supabase."""
    errors = []

    def _safe_date(val):
        try:
            if pd.isna(val):
                return None
            if isinstance(val, (datetime, date)):
                return str(val)[:10]
            return str(pd.to_datetime(val))[:10]
        except Exception:
            return None

    def _safe_num(val):
        try:
            v = pd.to_numeric(val, errors='coerce')
            return None if pd.isna(v) else float(v)
        except Exception:
            return None

    # --- MISSIONS ---
    df_m = data["MISSIONS"]
    if not df_m.empty:
        col_num = next((c for c in df_m.columns if 'Mission' in c or 'N°' in c), None)
        col_ag = next((c for c in df_m.columns if 'Agence' in c or 'Entité' in c or 'Site' in c), None)
        col_deb = next((c for c in df_m.columns if 'Début' in c or 'début' in c or 'Start' in c), None)
        col_fin = next((c for c in df_m.columns if 'Fin' in c or 'fin' in c or 'End' in c), None)
        col_conf = next((c for c in df_m.columns if 'conform' in c.lower()), None)
        col_stat = next((c for c in df_m.columns if 'Statut' in c or 'statut' in c), None)

        rows = []
        for _, r in df_m.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source', '')),
                "num_mission": str(r[col_num]) if col_num else None,
                "agence": str(r[col_ag]) if col_ag else None,
                "date_debut": _safe_date(r[col_deb]) if col_deb else None,
                "date_fin": _safe_date(r[col_fin]) if col_fin else None,
                "taux_conformite": _safe_num(r[col_conf]) if col_conf else None,
                "statut": str(r[col_stat]) if col_stat else None,
            })
        try:
            supabase.table("missions").insert(rows).execute()
        except Exception as e:
            errors.append(f"Missions : {e}")

    # --- ANOMALIES ---
    df_a = data["ANOMALIES"]
    if not df_a.empty:
        col_id = next((c for c in df_a.columns if 'ID Anomalie' in c), None)
        col_mis = next((c for c in df_a.columns if 'Mission' in c or 'N°' in c), None)
        col_ag = next((c for c in df_a.columns if 'Agence' in c or 'Site' in c or 'Entité' in c), None)
        col_pays = next((c for c in df_a.columns if 'Pays' in c), None)
        col_dom = next((c for c in df_a.columns if 'Domaine' in c or 'Type' in c), None)
        col_crit = next((c for c in df_a.columns if 'critic' in c.lower() or 'Critic' in c), None)
        col_desc = next((c for c in df_a.columns if 'Description' in c or 'Libellé' in c), None)
        col_imp = next((c for c in df_a.columns if 'Impact' in c), None)
        col_stat = next((c for c in df_a.columns if 'Statut' in c or 'statut' in c), None)
        col_date = next((c for c in df_a.columns if 'Date' in c or 'date' in c), None)
        col_ctrl = next((c for c in df_a.columns if 'Controleur' in c or 'Auditeur' in c or 'Responsable' in c), None)

        rows = []
        for _, r in df_a.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source', '')),
                "id_anomalie": str(r[col_id]) if col_id else None,
                "num_mission": str(r[col_mis]) if col_mis else None,
                "agence": str(r[col_ag]) if col_ag else None,
                "pays": str(r[col_pays]) if col_pays else None,
                "domaine": str(r[col_dom]) if col_dom else None,
                "criticite": str(r[col_crit]) if col_crit else None,
                "description": str(r[col_desc]) if col_desc else None,
                "impact_financier": _safe_num(r[col_imp]) if col_imp else None,
                "statut": str(r[col_stat]) if col_stat else None,
                "date_detection": _safe_date(r[col_date]) if col_date else None,
                "controleur": str(r[col_ctrl]) if col_ctrl else None,
            })
        try:
            supabase.table("anomalies").insert(rows).execute()
        except Exception as e:
            errors.append(f"Anomalies : {e}")

    # --- POINTS DE CONTRÔLE ---
    df_p = data["POINTS"]
    if not df_p.empty:
        col_id = next((c for c in df_p.columns if 'ID Point' in c), None)
        col_mis = next((c for c in df_p.columns if 'Mission' in c or 'N°' in c), None)
        col_ag = next((c for c in df_p.columns if 'Agence' in c or 'Site' in c), None)
        col_res = next((c for c in df_p.columns if 'Résultat' in c or 'Result' in c), None)

        rows = []
        for _, r in df_p.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source', '')),
                "id_point": str(r[col_id]) if col_id else None,
                "num_mission": str(r[col_mis]) if col_mis else None,
                "agence": str(r[col_ag]) if col_ag else None,
                "resultat": str(r[col_res]) if col_res else None,
            })
        try:
            supabase.table("points_controle").insert(rows).execute()
        except Exception as e:
            errors.append(f"Points de contrôle : {e}")

    # --- PLANS D'ACTION ---
    df_pl = data["PLANS"]
    if not df_pl.empty:
        col_id = next((c for c in df_pl.columns if 'ID Plan' in c), None)
        col_anom = next((c for c in df_pl.columns if 'Anomalie' in c), None)
        col_ag = next((c for c in df_pl.columns if 'Agence' in c or 'Site' in c), None)
        col_resp = next((c for c in df_pl.columns if 'Responsable' in c), None)
        col_ech = next((c for c in df_pl.columns if 'Échéance' in c or 'echeance' in c.lower()), None)
        col_stat = next((c for c in df_pl.columns if 'Statut' in c), None)

        rows = []
        for _, r in df_pl.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source', '')),
                "id_plan": str(r[col_id]) if col_id else None,
                "id_anomalie": str(r[col_anom]) if col_anom else None,
                "agence": str(r[col_ag]) if col_ag else None,
                "responsable": str(r[col_resp]) if col_resp else None,
                "date_echeance": _safe_date(r[col_ech]) if col_ech else None,
                "statut": str(r[col_stat]) if col_stat else None,
            })
        try:
            supabase.table("plans_action").insert(rows).execute()
        except Exception as e:
            errors.append(f"Plans d'action : {e}")

    return errors


@st.cache_data(ttl=120)
def load_from_supabase():
    """Charge les données depuis Supabase."""
    def fetch(table):
        try:
            res = supabase.table(table).select("*").execute()
            return pd.DataFrame(res.data) if res.data else pd.DataFrame()
        except Exception:
            return pd.DataFrame()

    return {
        "MISSIONS": fetch("missions"),
        "POINTS": fetch("points_controle"),
        "ANOMALIES": fetch("anomalies"),
        "PLANS": fetch("plans_action"),
    }


def find_col(df, key, mapping):
    """Trouve une colonne dans un DataFrame."""
    for candidate in mapping.get(key, []):
        match = next((c for c in df.columns if candidate.lower() in c.lower()), None)
        if match:
            return match
    return None


def get_safe_len(series, col_name):
    """Calcule la longueur sûre pour une colonne."""
    clean = series.dropna()
    max_c = int(clean.apply(lambda x: len(str(x))).max()) if not clean.empty else 0
    return min(max(max_c, len(str(col_name))) + 3, 50)

# ============================================================
# INTERFACE PRINCIPALE
# ============================================================
st.title("🛡️ Espace Chef de Département CI — Groupe SKAB")
st.subheader("Pilotage, Validation Métier et Consolidation des Missions 2026")

# ============================================================
# BARRE LATÉRALE — SOURCE DE DONNÉES
# ============================================================
with st.sidebar:
    st.header("⚙️ Configuration")

    source_mode = st.radio(
        "Mode de consultation :",
        ["📂 Fichiers Excel (Import local)", "☁️ Supabase (Base consolidée)"],
        index=1 if SUPABASE_OK else 0
    )

    st.divider()

    # Upload
    st.markdown("📥 **Importation Terrain**")
    uploaded_files = st.file_uploader(
        "Déposez les fichiers des contrôleurs (.xlsx) :",
        type="xlsx",
        accept_multiple_files=True
    )

    if source_mode == "☁️ Supabase (Base consolidée)":
        if not SUPABASE_OK:
            st.error("❌ Supabase non configuré.\nAjoutez SUPABASE_URL et SUPABASE_KEY dans les secrets.")
        elif uploaded_files:
            st.info("📤 Fichiers détectés. Cliquez pour les envoyer dans la base Supabase.")
            if st.button("☁️ Envoyer vers Supabase", use_container_width=True, type="primary"):
                with st.spinner("Envoi en cours…"):
                    data_tmp = process_consolidation(uploaded_files)
                    errs = push_to_supabase(data_tmp, uploaded_files)
                if errs:
                    for e in errs:
                        st.error(e)
                else:
                    st.success("✅ Données envoyées vers Supabase !")
                    load_from_supabase.clear()
        else:
            st.caption("💡 Déposez des fichiers ici pour les ajouter à la base Supabase.")

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
        st.info(f"👁️ Prévisualisation : données Supabase + {len(uploaded_files)} fichier(s) uploadé(s) (non encore sauvegardés).")

    if data["ANOMALIES"].empty and data["MISSIONS"].empty:
        st.info("📭 La base Supabase est vide. Déposez des fichiers et cliquez sur 'Envoyer vers Supabase'.")
        st.stop()

df_mis = data["MISSIONS"]
df_pts = data["POINTS"]
df_anom = data["ANOMALIES"]
df_plan = data["PLANS"]

# ============================================================
# MAPPING DES COLONNES
# ============================================================
COL_MAP = {
    "impact": ["Impact", "impact_financier", "Impact Financier"],
    "criticite": ["Criticité", "criticite", "Criticite", "Niveau de Criticité"],
    "domaine": ["Domaine", "domaine", "Type"],
    "pays": ["Pays", "pays"],
    "agence": ["Agence", "agence", "Site / Entité", "Site", "Entité"],
    "statut": ["Statut", "statut"],
    "date": ["Date", "date_detection", "Date de Détection", "created_at"],
    "conformite": ["Taux de Conformité", "taux_conformite", "conform", "Conformité"],
    "num_mis": ["N° Mission", "num_mission", "Mission"],
    "id_anom": ["ID Anomalie", "id_anomalie"],
    "controleur": ["Controleur", "controleur", "Auditeur", "auditeur", "Responsable"],
    "desc": ["Description", "description", "Libellé", "libellé"],
}

col_impact = find_col(df_anom, "impact", COL_MAP)
col_crit = find_col(df_anom, "criticite", COL_MAP)
col_domaine = find_col(df_anom, "domaine", COL_MAP)
col_pays = find_col(df_anom, "pays", COL_MAP)
col_ag_a = find_col(df_anom, "agence", COL_MAP)
col_stat_a = find_col(df_anom, "statut", COL_MAP)
col_date_a = find_col(df_anom, "date", COL_MAP)
col_ctrl_a = find_col(df_anom, "controleur", COL_MAP)
col_desc_a = find_col(df_anom, "desc", COL_MAP)
col_tx_conf = find_col(df_mis, "conformite", COL_MAP)
col_ag_m = find_col(df_mis, "agence", COL_MAP)

# ============================================================
# SECTION FILTRES AVANCÉS
# ============================================================
st.markdown("### 🔍 Filtres Avancés")

with st.expander("📋 Afficher/Masquer les filtres", expanded=True):
    
    # Grille de filtres
    f_col1, f_col2, f_col3, f_col4 = st.columns(4)

    # Filtre Agence
    all_agences = []
    for df, col in [(df_anom, col_ag_a), (df_mis, col_ag_m)]:
        if col and not df.empty:
            all_agences += df[col].dropna().astype(str).unique().tolist()
    all_agences = sorted(set(all_agences))

    with f_col1:
        selected_agences = st.multiselect(
            "🏢 Agences / Entités",
            all_agences,
            default=None,
            key="filter_agences"
        )

    # Filtre Pays
    all_pays = []
    if col_pays and not df_anom.empty:
        all_pays = sorted(df_anom[col_pays].dropna().astype(str).unique().tolist())

    with f_col2:
        selected_pays = st.multiselect(
            "🌍 Pays",
            all_pays,
            default=None,
            key="filter_pays"
        )

    # Filtre Contrôleur
    all_controleurs = []
    if col_ctrl_a and not df_anom.empty:
        all_controleurs = sorted(df_anom[col_ctrl_a].dropna().astype(str).unique().tolist())

    with f_col3:
        selected_controleurs = st.multiselect(
            "👤 Contrôleurs",
            all_controleurs,
            default=None,
            key="filter_controleurs"
        )

    # Filtre Domaine
    all_domaines = []
    if col_domaine and not df_anom.empty:
        all_domaines = sorted(df_anom[col_domaine].dropna().astype(str).unique().tolist())

    with f_col4:
        selected_domaines = st.multiselect(
            "📁 Domaines",
            all_domaines,
            default=None,
            key="filter_domaines"
        )

    # Deuxième ligne de filtres
    f_col5, f_col6, f_col7, f_col8 = st.columns(4)

    # Filtre Criticité
    all_criticites = []
    if col_crit and not df_anom.empty:
        all_criticites = sorted(df_anom[col_crit].dropna().astype(str).unique().tolist())

    with f_col5:
        selected_criticites = st.multiselect(
            "🚨 Criticités",
            all_criticites,
            default=None,
            key="filter_criticites"
        )

    # Filtre Statut
    all_statuts = []
    if col_stat_a and not df_anom.empty:
        all_statuts = sorted(df_anom[col_stat_a].dropna().astype(str).unique().tolist())

    with f_col6:
        selected_statuts = st.multiselect(
            "📌 Statuts",
            all_statuts,
            default=None,
            key="filter_statuts"
        )

    # Plage de dates
    with f_col7:
        if col_date_a and not df_anom.empty:
            df_anom_tmp = df_anom.copy()
            df_anom_tmp[col_date_a] = pd.to_datetime(df_anom_tmp[col_date_a], errors="coerce")
            min_date = df_anom_tmp[col_date_a].min()
            max_date = df_anom_tmp[col_date_a].max()

            if pd.notna(min_date) and pd.notna(max_date):
                date_range = st.date_input(
                    "📅 Plage de dates",
                    value=(min_date.date(), max_date.date()),
                    min_value=min_date.date(),
                    max_value=max_date.date(),
                    key="filter_dates"
                )
                if len(date_range) == 2:
                    date_start, date_end = date_range
                else:
                    date_start, date_end = None, None
            else:
                date_start, date_end = None, None
        else:
            date_start, date_end = None, None

    # Filtre Impact Financier
    with f_col8:
        if col_impact and not df_anom.empty:
            df_anom_imp = df_anom.copy()
            df_anom_imp[col_impact] = pd.to_numeric(df_anom_imp[col_impact], errors="coerce")
            min_impact = df_anom_imp[col_impact].min()
            max_impact = df_anom_imp[col_impact].max()

            if pd.notna(min_impact) and pd.notna(max_impact):
                impact_range = st.slider(
                    "💰 Impact Financier (FCFA)",
                    min_value=int(min_impact),
                    max_value=int(max_impact),
                    value=(int(min_impact), int(max_impact)),
                    step=100000,
                    key="filter_impact"
                )
            else:
                impact_range = (0, 0)
        else:
            impact_range = (0, 0)

# ============================================================
# APPLICATION DES FILTRES
# ============================================================
def apply_filters(df, filters):
    """Applique tous les filtres au DataFrame."""
    df = df.copy()

    if filters.get("agences") and col_ag_a and col_ag_a in df.columns:
        df = df[df[col_ag_a].astype(str).isin(filters["agences"])]

    if filters.get("pays") and col_pays and col_pays in df.columns:
        df = df[df[col_pays].astype(str).isin(filters["pays"])]

    if filters.get("controleurs") and col_ctrl_a and col_ctrl_a in df.columns:
        df = df[df[col_ctrl_a].astype(str).isin(filters["controleurs"])]

    if filters.get("domaines") and col_domaine and col_domaine in df.columns:
        df = df[df[col_domaine].astype(str).isin(filters["domaines"])]

    if filters.get("criticites") and col_crit and col_crit in df.columns:
        df = df[df[col_crit].astype(str).isin(filters["criticites"])]

    if filters.get("statuts") and col_stat_a and col_stat_a in df.columns:
        df = df[df[col_stat_a].astype(str).isin(filters["statuts"])]

    if filters.get("dates") and col_date_a and col_date_a in df.columns:
        df[col_date_a] = pd.to_datetime(df[col_date_a], errors="coerce")
        date_start, date_end = filters["dates"]
        df = df[(df[col_date_a] >= pd.Timestamp(date_start)) & (df[col_date_a] <= pd.Timestamp(date_end))]

    if filters.get("impact") and col_impact and col_impact in df.columns:
        df[col_impact] = pd.to_numeric(df[col_impact], errors="coerce")
        impact_min, impact_max = filters["impact"]
        df = df[(df[col_impact] >= impact_min) & (df[col_impact] <= impact_max)]

    return df


filters = {
    "agences": selected_agences if selected_agences else None,
    "pays": selected_pays if selected_pays else None,
    "controleurs": selected_controleurs if selected_controleurs else None,
    "domaines": selected_domaines if selected_domaines else None,
    "criticites": selected_criticites if selected_criticites else None,
    "statuts": selected_statuts if selected_statuts else None,
    "dates": (date_start, date_end) if date_start and date_end else None,
    "impact": impact_range if impact_range != (0, 0) else None,
}

df_anom_f = apply_filters(df_anom, filters)
df_mis_f = apply_filters(df_mis, filters)
df_pts_f = apply_filters(df_pts, filters)
df_plan_f = apply_filters(df_plan, filters)

# Afficher les filtres actifs
active_filters = []
if selected_agences:
    active_filters.append(f"🏢 {len(selected_agences)} agence(s)")
if selected_pays:
    active_filters.append(f"🌍 {len(selected_pays)} pays")
if selected_controleurs:
    active_filters.append(f"👤 {len(selected_controleurs)} contrôleur(s)")
if selected_domaines:
    active_filters.append(f"📁 {len(selected_domaines)} domaine(s)")
if selected_criticites:
    active_filters.append(f"🚨 {len(selected_criticites)} criticité(s)")
if selected_statuts:
    active_filters.append(f"📌 {len(selected_statuts)} statut(s)")
if date_start and date_end:
    active_filters.append(f"📅 {date_start} → {date_end}")
if impact_range and impact_range != (0, 0):
    active_filters.append(f"💰 {impact_range[0]:,} - {impact_range[1]:,} FCFA")

if active_filters:
    filter_html = " ".join([f'<span class="filter-badge">{f}</span>' for f in active_filters])
    st.markdown(f"<div>{filter_html}</div>", unsafe_allow_html=True)
    st.caption(f"Résultat : {len(df_anom_f)} anomalie(s) | {len(df_mis_f)} mission(s)")
    st.divider()

# ============================================================
# KPI PRINCIPAUX
# ============================================================
st.markdown("<div class='section-title'>📊 Indicateurs de Risques</div>", unsafe_allow_html=True)

kpi_col1, kpi_col2, kpi_col3, kpi_col4, kpi_col5 = st.columns(5)

with kpi_col1:
    impact_total = 0
    if col_impact and not df_anom_f.empty:
        impact_total = pd.to_numeric(df_anom_f[col_impact], errors='coerce').fillna(0).sum()
    st.metric("💰 Risque Financier", f"{impact_total:,.0f} FCFA")

with kpi_col2:
    nb_critiques = 0
    if col_crit and not df_anom_f.empty:
        nb_critiques = df_anom_f[df_anom_f[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
    st.metric("🚨 Critiques", nb_critiques)

with kpi_col3:
    nb_en_cours = 0
    if col_stat_a and not df_anom_f.empty:
        nb_en_cours = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains('EN COURS', na=False)].shape[0]
    st.metric("⏳ En Cours", nb_en_cours)

with kpi_col4:
    conformite_moyenne = 0
    if col_tx_conf and not df_mis_f.empty:
        raw_mean = pd.to_numeric(df_mis_f[col_tx_conf], errors='coerce').mean()
        if pd.notna(raw_mean):
            conformite_moyenne = raw_mean * 100 if raw_mean <= 1.0 else raw_mean
    st.metric("✅ Conformité", f"{conformite_moyenne:.1f}%" if conformite_moyenne > 0 else "N/A")

with kpi_col5:
    nb_total = len(df_anom_f)
    st.metric("📌 Anomalies", nb_total)

st.divider()

# ============================================================
# GRAPHIQUES - LIGNE 1
# ============================================================
st.markdown("<div class='section-title'>📈 Analyses Visuelles</div>", unsafe_allow_html=True)

graph_col1, graph_col2 = st.columns(2)

with graph_col1:
    st.markdown("**Anomalies par Domaine**")
    if not df_anom_f.empty and col_domaine:
        df_graph1 = df_anom_f.dropna(subset=[col_domaine])
        fig1 = px.bar(
            df_graph1,
            x=col_domaine,
            color=col_crit if col_crit else None,
            title="",
            color_discrete_map={
                '🔴 Critique': '#e63946', 'Critique': '#e63946',
                '🟠 Majeur': '#f4a261', 'Majeur': '#f4a261',
                '🟡 Mineur': '#e9c46a', 'Mineur': '#e9c46a',
                '🟢 Faible': '#2a9d8f', 'Faible': '#2a9d8f'
            }
        )
        fig1.update_layout(height=350, margin=dict(l=0, r=0, t=0, b=0), showlegend=False,
                          plot_bgcolor='rgba(0,0,0,0)', paper_bgcolor='rgba(0,0,0,0)')
        st.plotly_chart(fig1, use_container_width=True)
    else:
        st.info("Aucune donnée disponible pour ce graphique")

with graph_col2:
    st.markdown("**Anomalies par Pays**")
    if not df_anom_f.empty and col_pays:
        df_pays = df_anom_f.dropna(subset=[col_pays]).groupby(col_pays).size().reset_index(name="Anomalies")
        fig2 = px.pie(
            df_pays,
            values="Anomalies",
            names=col_pays,
            hole=.4,
            color_discrete_sequence=px.colors.qualitative.Safe,
            title=""
        )
        fig2.update_layout(height=350, margin=dict(l=0, r=0, t=0, b=0))
        st.plotly_chart(fig2, use_container_width=True)
    else:
        st.info("Aucune donnée disponible pour ce graphique")

st.divider()

# ============================================================
# SECTION DÉTAILLÉE - ANOMALIES EN COURS
# ============================================================
st.markdown("<div class='section-title'>🔴 Anomalies EN COURS</div>", unsafe_allow_html=True)

if not df_anom_f.empty and col_stat_a:
    df_en_cours = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains("EN COURS", na=False)].copy()

    if df_en_cours.empty:
        st.success("✅ Aucune anomalie EN COURS")
    else:
        ec_col1, ec_col2, ec_col3 = st.columns(3)

        with ec_col1:
            st.metric("Total EN COURS", len(df_en_cours))

        with ec_col2:
            impact_ec = 0
            if col_impact:
                impact_ec = pd.to_numeric(df_en_cours[col_impact], errors='coerce').fillna(0).sum()
            st.metric("Impact EN COURS", f"{impact_ec:,.0f} FCFA")

        with ec_col3:
            nb_crit_ec = 0
            if col_crit:
                nb_crit_ec = df_en_cours[df_en_cours[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
            st.metric("Critiques EN COURS", nb_crit_ec)

        # Tableau EN COURS
        st.markdown("**Liste des anomalies EN COURS :**")
        cols_ec = [c for c in [
            find_col(df_en_cours, "id_anom", COL_MAP),
            col_ag_a,
            col_pays,
            col_ctrl_a,
            col_domaine,
            col_crit,
            col_date_a,
            col_impact,
        ] if c and c in df_en_cours.columns]

        df_show_ec = df_en_cours[cols_ec].copy() if cols_ec else df_en_cours.copy()
        st.dataframe(df_show_ec, hide_index=True, use_container_width=True, height=300)

        # Export EN COURS
        csv_ec = df_en_cours.to_csv(index=False).encode("utf-8")
        st.download_button(
            "⬇️ Exporter anomalies EN COURS (.csv)",
            data=csv_ec,
            file_name=f"SKAB_EN_COURS_{date.today()}.csv",
            mime="text/csv",
            use_container_width=True
        )
else:
    st.info("Aucune colonne Statut détectée")

st.divider()

# ============================================================
# SECTION DÉTAILLÉE - PAR AGENCE
# ============================================================
st.markdown("<div class='section-title'>🏢 Analyse par Agence</div>", unsafe_allow_html=True)

agences_dispo = []
if col_ag_a and not df_anom.empty:
    agences_dispo = sorted(df_anom[col_ag_a].dropna().astype(str).unique().tolist())

if not agences_dispo:
    st.info("Aucune agence détectée")
else:
    ag_tab = st.selectbox("Sélectionner une agence :", agences_dispo, key="ag_detail")
    df_ag = df_anom[df_anom[col_ag_a].astype(str) == ag_tab].copy()

    ag_col1, ag_col2, ag_col3, ag_col4 = st.columns(4)

    with ag_col1:
        st.metric("Total Anomalies", len(df_ag))

    with ag_col2:
        nb_crit_ag = 0
        if col_crit:
            nb_crit_ag = df_ag[df_ag[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0]
        st.metric("Critiques", nb_crit_ag)

    with ag_col3:
        nb_ec_ag = 0
        if col_stat_a:
            nb_ec_ag = df_ag[df_ag[col_stat_a].astype(str).str.upper().str.contains('EN COURS', na=False)].shape[0]
        st.metric("EN COURS", nb_ec_ag)

    with ag_col4:
        impact_ag = 0
        if col_impact:
            impact_ag = pd.to_numeric(df_ag[col_impact], errors='coerce').fillna(0).sum()
        st.metric("Impact Financier", f"{impact_ag:,.0f} FCFA")

    # Graphe par domaine pour l'agence
    if col_domaine and not df_ag.empty:
        df_dom_ag = df_ag.dropna(subset=[col_domaine]).groupby(col_domaine).size().reset_index(name="Nb")
        if not df_dom_ag.empty:
            fig_ag = px.bar_polar(
                df_dom_ag,
                r="Nb",
                theta=col_domaine,
                color="Nb",
                color_continuous_scale="Reds",
                title=f"Profil d'anomalies — {ag_tab}",
                height=400
            )
            fig_ag.update_layout(margin=dict(l=0, r=0, t=40, b=0))
            st.plotly_chart(fig_ag, use_container_width=True)

    # Tableau agence
    st.markdown(f"**Toutes les anomalies de {ag_tab} :**")
    cols_ag = [c for c in [
        find_col(df_ag, "id_anom", COL_MAP),
        col_crit,
        col_domaine,
        col_pays,
        col_ctrl_a,
        col_stat_a,
        col_date_a,
        col_impact,
    ] if c and c in df_ag.columns]

    df_show_ag = df_ag[cols_ag].copy() if cols_ag else df_ag.copy()
    st.dataframe(df_show_ag, hide_index=True, use_container_width=True, height=300)

st.divider()

# ============================================================
# SECTION DÉTAILLÉE - ANALYSE TEMPORELLE
# ============================================================
st.markdown("<div class='section-title'>📅 Analyse Temporelle</div>", unsafe_allow_html=True)

if col_date_a and not df_anom_f.empty:
    df_per = df_anom_f.copy()
    df_per[col_date_a] = pd.to_datetime(df_per[col_date_a], errors="coerce")
    df_per = df_per.dropna(subset=[col_date_a])

    if df_per.empty:
        st.info("Aucune date valide")
    else:
        df_per["Mois"] = df_per[col_date_a].dt.strftime("%Y-%m")
        df_per["Trimestre"] = "T" + df_per[col_date_a].dt.quarter.astype(str) + " " + df_per[col_date_a].dt.year.astype(str)
        df_per["Année"] = df_per[col_date_a].dt.year.astype(str)

        p_type = st.radio("Granularité :", ["Mois", "Trimestre", "Année"], horizontal=True, key="period_radio")

        df_grp = df_per.groupby(p_type).agg({
            col_impact if col_impact else df_per.columns[0]: ["count", lambda x: pd.to_numeric(x, errors='coerce').sum() if col_impact else 0]
        }).reset_index().sort_values(p_type)

        graph_per_col1, graph_per_col2 = st.columns(2)

        with graph_per_col1:
            df_grp_count = df_per.groupby(p_type).size().reset_index(name="Anomalies")
            fig_p1 = px.bar(
                df_grp_count,
                x=p_type,
                y="Anomalies",
                color_discrete_sequence=["#1d3557"],
                title=f"Anomalies par {p_type.lower()}"
            )
            fig_p1.update_layout(height=300, margin=dict(l=0, r=0, t=40, b=0), plot_bgcolor='rgba(0,0,0,0)')
            st.plotly_chart(fig_p1, use_container_width=True)

        with graph_per_col2:
            if col_impact:
                df_grp_impact = df_per.groupby(p_type)[col_impact].apply(lambda x: pd.to_numeric(x, errors='coerce').sum()).reset_index(name="Impact")
                fig_p2 = px.line(
                    df_grp_impact,
                    x=p_type,
                    y="Impact",
                    markers=True,
                    color_discrete_sequence=["#e63946"],
                    title=f"Impact financier par {p_type.lower()}"
                )
                fig_p2.update_layout(height=300, margin=dict(l=0, r=0, t=40, b=0), plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_p2, use_container_width=True)
            else:
                st.info("Colonne impact non détectée")
else:
    st.info("Aucune date détectée")

st.divider()

# ============================================================
# DIAGNOSTIC QUALITÉ
# ============================================================
st.markdown("<div class='section-title'>🛠️ Diagnostic Qualité</div>", unsafe_allow_html=True)

qual_col1, qual_col2 = st.columns(2)

with qual_col1:
    st.markdown("**🚨 Incohérences**")
    alertes = []

    if not df_anom_f.empty and col_crit and col_impact:
        lignes_anom = df_anom_f[
            (df_anom_f[col_crit].astype(str).str.contains('Critique|🔴|Majeur|🟠', na=False)) &
            (pd.to_numeric(df_anom_f[col_impact], errors='coerce').fillna(0) == 0)
        ]
        for _, row in lignes_anom.iterrows():
            id_val = row.get('id_anomalie', row.get('ID Anomalie', 'N/A'))
            alertes.append(f"⚠️ Anomalie {id_val} : Impact nul")

    if alertes:
        for alerte in alertes:
            st.warning(alerte)
    else:
        st.success("✅ Aucune incohérence détectée")

with qual_col2:
    st.markdown("**📊 Résumé des données**")
    st.info(f"""
    - **Anomalies** : {len(df_anom_f)} (filtrées) / {len(df_anom)} (total)
    - **Missions** : {len(df_mis_f)} (filtrées) / {len(df_mis)} (total)
    - **Points de contrôle** : {len(df_pts_f)} (filtrés) / {len(df_pts)} (total)
    - **Plans d'action** : {len(df_plan_f)} (filtrés) / {len(df_plan)} (total)
    """)

st.divider()

# ============================================================
# EXPORT FICHIER MAÎTRE
# ============================================================
st.markdown("<div class='section-title'>📤 Export Fichier Maître</div>", unsafe_allow_html=True)

if st.button("🏗️ Générer le fichier consolidé", type="primary", use_container_width=True):
    output_buffer = io.BytesIO()

    with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
        # Onglet accueil
        df_meta = pd.DataFrame({
            "RAPPORT SKAB 2026": [
                "Chef de Département",
                "Généré le",
                "Filtres appliqués",
                "Validation"
            ],
            "DÉTAILS": [
                "CI - Contrôle Interne",
                datetime.now().strftime("%d/%m/%Y à %H:%M"),
                " | ".join(active_filters) if active_filters else "Aucun",
                "SCELLÉ"
            ]
        })
        df_meta.to_excel(writer, sheet_name="ACCUEIL", index=False)
        writer.sheets["ACCUEIL"].set_column('A:B', 35)

        # Onglets de données
        sheets = {
            "MISSIONS": df_mis,
            "POINTS_CONTROLE": df_pts,
            "ANOMALIES": df_anom,
            "PLANS_ACTION": df_plan,
            "ANOMALIES_EN_COURS": df_anom[df_anom[col_stat_a].astype(str).str.upper().str.contains("EN COURS", na=False)] if col_stat_a else pd.DataFrame(),
        }

        for sheet_name, dataframe in sheets.items():
            if dataframe is not None and not dataframe.empty:
                dataframe.to_excel(writer, sheet_name=sheet_name, index=False)
                ws = writer.sheets[sheet_name]
                for i, col in enumerate(dataframe.columns):
                    ws.set_column(i, i, get_safe_len(dataframe[col], col))

    st.success("✅ Fichier généré avec succès !")
    st.download_button(
        label=f"💾 Télécharger SKAB_MAITRE_{datetime.now().strftime('%Y%m%d')}.xlsx",
        data=output_buffer.getvalue(),
        file_name=f"SKAB_MAITRE_{datetime.now().strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )

st.divider()

st.markdown("""
---
<div style="text-align: center; color: #888; font-size: 12px;">
<p>© 2026 Groupe SKAB Nutrition — Direction Audit & Contrôle Interne</p>
<p>Dashboard v3.0 — Interface modulaire avec filtres avancés</p>
</div>
""", unsafe_allow_html=True)
