import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import io
import numpy as np
from datetime import datetime, date, timedelta
from supabase import create_client, Client

# ============================================================
# CONFIGURATION PAGE
# ============================================================
st.set_page_config(page_title="SKAB — Dashboard CI", page_icon="🛡️", layout="wide")

st.markdown("""
<style>
[data-testid="stMetricValue"]    { font-size: 24px; font-weight: 700; }
[data-testid="stMetricLabel"]    { font-size: 11px; color: #6b7280; text-transform: uppercase; letter-spacing:.05em; }
[data-testid="metric-container"] { background:#f8fafc; border:1px solid #e2e8f0; border-radius:10px; padding:14px 18px; }
button[data-baseweb="tab"]       { font-size:13px; font-weight:600; }
.sec { font-size:14px; font-weight:700; color:#1e293b; border-left:4px solid #e63946;
       padding-left:10px; margin:10px 0 14px 0; }
.badge { display:inline-block; background:#fef3c7; color:#92400e; border:1px solid #fcd34d;
         border-radius:20px; padding:2px 11px; font-size:11px; font-weight:600; margin:2px; }
.kpi-alert { background:#fff1f2; border:1px solid #fecdd3; border-radius:8px; padding:10px 16px;
             font-size:13px; color:#be123c; font-weight:600; margin:4px 0; }
.kpi-ok    { background:#f0fdf4; border:1px solid #bbf7d0; border-radius:8px; padding:10px 16px;
             font-size:13px; color:#15803d; font-weight:600; margin:4px 0; }
section[data-testid="stSidebar"] > div { padding-top:.8rem; }
</style>
""", unsafe_allow_html=True)

# ============================================================
# SUPABASE
# ============================================================
@st.cache_resource
def init_supabase() -> Client:
    return create_client(st.secrets["SUPABASE_URL"], st.secrets["SUPABASE_KEY"])

try:
    supabase = init_supabase()
    SUPABASE_OK = True
except Exception:
    SUPABASE_OK = False

# ============================================================
# MOTEUR EXCEL
# ============================================================
def load_and_clean(file, sheet):
    try:
        df_raw = pd.read_excel(file, sheet_name=sheet, header=None)
        if df_raw.empty: return pd.DataFrame()
        targets = {"MES_MISSIONS":"N° Mission","POINTS_CONTROLE":"ID Point",
                   "ANOMALIES":"ID Anomalie","PLANS_ACTION":"ID Plan"}
        header_idx = None
        for idx, row in df_raw.iterrows():
            row_str = [str(v).strip() for v in row.values]
            if sheet in targets and any(targets[sheet] in s for s in row_str):
                header_idx = idx; break
        if header_idx is None:
            for idx, row in df_raw.iterrows():
                row_str = [str(v).strip() for v in row.values]
                if any("ID" in s or "N°" in s or "Code" in s for s in row_str):
                    header_idx = idx; break
        if header_idx is None: header_idx = 0
        df = pd.read_excel(file, sheet_name=sheet, skiprows=header_idx)
        df.columns = [str(c).strip() for c in df.columns]
        df = df.dropna(subset=[df.columns[0]]) if not df.empty else df
        df = df.dropna(how='all')
        if not df.empty:
            df = df[~df[df.columns[0]].astype(str).str.contains(
                "Une anomalie|Un plan|Saisissez|Une ligne", na=False)]
            df['Fichier Source'] = file.name
            df['Contrôleur']     = file.name.replace('.xlsx','').replace('.xls','')
        return df
    except Exception:
        return pd.DataFrame()

def process_consolidation(files):
    all_data = {"MISSIONS":[],"POINTS":[],"ANOMALIES":[],"PLANS":[]}
    for f in files:
        f.seek(0); all_data["MISSIONS"].append(load_and_clean(f,"MES_MISSIONS"))
        f.seek(0); all_data["POINTS"].append(load_and_clean(f,"POINTS_CONTROLE"))
        f.seek(0); all_data["ANOMALIES"].append(load_and_clean(f,"ANOMALIES"))
        f.seek(0); all_data["PLANS"].append(load_and_clean(f,"PLANS_ACTION"))
    return {k: pd.concat(v,ignore_index=True) if v else pd.DataFrame() for k,v in all_data.items()}

def get_safe_len(series, col_name):
    clean = series.dropna()
    max_c = int(clean.apply(lambda x: len(str(x))).max()) if not clean.empty else 0
    return min(max(max_c, len(str(col_name)))+3, 50)

# ============================================================
# HELPERS
# ============================================================
def _safe_date(val):
    try:
        if pd.isna(val): return None
        if isinstance(val,(datetime,date)): return str(val)[:10]
        return str(pd.to_datetime(val))[:10]
    except: return None

def _safe_num(val):
    try:
        v = pd.to_numeric(val, errors='coerce')
        return None if pd.isna(v) else float(v)
    except: return None

COLOR_CRIT = {
    'Critique':'#e63946','🔴 Critique':'#e63946',
    'Majeur':'#f4a261',  '🟠 Majeur':'#f4a261',
    'Mineur':'#e9c46a',  '🟡 Mineur':'#e9c46a',
    'Faible':'#2a9d8f',  '🟢 Faible':'#2a9d8f',
}

def style_crit(val):
    v = str(val)
    if 'Critique' in v or '🔴' in v: return 'background-color:#ffe4e6;font-weight:bold'
    elif 'Majeur' in v or '🟠' in v:  return 'background-color:#fff3e0'
    elif 'Mineur' in v or '🟡' in v:  return 'background-color:#fefce8'
    return ''

def apply_style(df, col):
    if col and col in df.columns:
        try:    return df.style.map(style_crit, subset=[col])
        except: return df.style.applymap(style_crit, subset=[col])
    return df

# ============================================================
# PUSH SUPABASE
# ============================================================
def push_to_supabase(data, source_files):
    errors = []
    def fc(df, *keys):
        for k in keys:
            m = next((c for c in df.columns if k.lower() in c.lower()), None)
            if m: return m
        return None

    # MISSIONS
    df_m = data["MISSIONS"]
    if not df_m.empty:
        rows = []
        for _, r in df_m.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source','')),
                "controleur":     str(r[fc(df_m,'Contrôleur','controleur')]) if fc(df_m,'Contrôleur','controleur') else None,
                "num_mission":    str(r[fc(df_m,'Mission','N°')]) if fc(df_m,'Mission','N°') else None,
                "agence":         str(r[fc(df_m,'Agence','Entité','Site')]) if fc(df_m,'Agence','Entité','Site') else None,
                "date_debut":     _safe_date(r[fc(df_m,'Début')]) if fc(df_m,'Début') else None,
                "date_fin":       _safe_date(r[fc(df_m,'Fin')]) if fc(df_m,'Fin') else None,
                "taux_conformite":_safe_num(r[fc(df_m,'conform')]) if fc(df_m,'conform') else None,
                "statut":         str(r[fc(df_m,'Statut','statut')]) if fc(df_m,'Statut','statut') else None,
            })
        try: supabase.table("missions").insert(rows).execute()
        except Exception as e: errors.append(f"Missions : {e}")

    # ANOMALIES
    df_a = data["ANOMALIES"]
    if not df_a.empty:
        rows = []
        for _, r in df_a.iterrows():
            rows.append({
                "fichier_source":   str(r.get('Fichier Source','')),
                "controleur":       str(r[fc(df_a,'Contrôleur')]) if fc(df_a,'Contrôleur') else None,
                "id_anomalie":      str(r[fc(df_a,'ID Anomalie')]) if fc(df_a,'ID Anomalie') else None,
                "num_mission":      str(r[fc(df_a,'Mission','N°')]) if fc(df_a,'Mission','N°') else None,
                "agence":           str(r[fc(df_a,'Agence','Site','Entité')]) if fc(df_a,'Agence','Site','Entité') else None,
                "pays":             str(r[fc(df_a,'Pays')]) if fc(df_a,'Pays') else None,
                "domaine":          str(r[fc(df_a,'Domaine','Type')]) if fc(df_a,'Domaine','Type') else None,
                "criticite":        str(r[fc(df_a,'critic')]) if fc(df_a,'critic') else None,
                "description":      str(r[fc(df_a,'Description','Libellé')]) if fc(df_a,'Description','Libellé') else None,
                "impact_financier": _safe_num(r[fc(df_a,'Impact')]) if fc(df_a,'Impact') else None,
                "statut":           str(r[fc(df_a,'Statut','statut')]) if fc(df_a,'Statut','statut') else None,
                "date_detection":   _safe_date(r[fc(df_a,'Date','date')]) if fc(df_a,'Date','date') else None,
            })
        try: supabase.table("anomalies").insert(rows).execute()
        except Exception as e: errors.append(f"Anomalies : {e}")

    # POINTS
    df_p = data["POINTS"]
    if not df_p.empty:
        rows = []
        for _, r in df_p.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source','')),
                "controleur":     str(r[fc(df_p,'Contrôleur')]) if fc(df_p,'Contrôleur') else None,
                "id_point":       str(r[fc(df_p,'ID Point')]) if fc(df_p,'ID Point') else None,
                "num_mission":    str(r[fc(df_p,'Mission','N°')]) if fc(df_p,'Mission','N°') else None,
                "agence":         str(r[fc(df_p,'Agence','Site')]) if fc(df_p,'Agence','Site') else None,
                "resultat":       str(r[fc(df_p,'Résultat','Result')]) if fc(df_p,'Résultat','Result') else None,
            })
        try: supabase.table("points_controle").insert(rows).execute()
        except Exception as e: errors.append(f"Points : {e}")

    # PLANS
    df_pl = data["PLANS"]
    if not df_pl.empty:
        rows = []
        for _, r in df_pl.iterrows():
            rows.append({
                "fichier_source": str(r.get('Fichier Source','')),
                "controleur":     str(r[fc(df_pl,'Contrôleur')]) if fc(df_pl,'Contrôleur') else None,
                "id_plan":        str(r[fc(df_pl,'ID Plan')]) if fc(df_pl,'ID Plan') else None,
                "id_anomalie":    str(r[fc(df_pl,'Anomalie')]) if fc(df_pl,'Anomalie') else None,
                "agence":         str(r[fc(df_pl,'Agence','Site')]) if fc(df_pl,'Agence','Site') else None,
                "responsable":    str(r[fc(df_pl,'Responsable')]) if fc(df_pl,'Responsable') else None,
                "date_echeance":  _safe_date(r[fc(df_pl,'Échéance','echeance')]) if fc(df_pl,'Échéance','echeance') else None,
                "statut":         str(r[fc(df_pl,'Statut')]) if fc(df_pl,'Statut') else None,
            })
        try: supabase.table("plans_action").insert(rows).execute()
        except Exception as e: errors.append(f"Plans : {e}")

    return errors

# ============================================================
# LECTURE SUPABASE
# ============================================================
@st.cache_data(ttl=120)
def load_from_supabase():
    def fetch(t):
        try:
            res = supabase.table(t).select("*").execute()
            return pd.DataFrame(res.data) if res.data else pd.DataFrame()
        except: return pd.DataFrame()
    return {"MISSIONS":fetch("missions"),"POINTS":fetch("points_controle"),
            "ANOMALIES":fetch("anomalies"),"PLANS":fetch("plans_action")}

# ============================================================
# SIDEBAR
# ============================================================
with st.sidebar:
    st.markdown("## 🛡️ SKAB — CI")
    st.markdown("---")
    st.markdown("#### ⚙️ Source")
    source_mode = st.radio("", ["📂 Excel (local)","☁️ Supabase (BDD)"],
                           index=1 if SUPABASE_OK else 0, label_visibility="collapsed")
    st.markdown("---")
    st.markdown("#### 📥 Import fichiers terrain")
    uploaded_files = st.file_uploader("Fichiers contrôleurs (.xlsx)", type="xlsx",
                                      accept_multiple_files=True, label_visibility="collapsed")
    if source_mode == "☁️ Supabase (BDD)":
        if not SUPABASE_OK:
            st.error("Supabase non configuré.")
        elif uploaded_files:
            if st.button("☁️ Envoyer vers Supabase", use_container_width=True, type="primary"):
                with st.spinner("Envoi…"):
                    errs = push_to_supabase(process_consolidation(uploaded_files), uploaded_files)
                if errs:
                    for e in errs: st.error(e)
                else:
                    st.success("✅ Envoyé !")
                    load_from_supabase.clear()
        else:
            st.caption("Déposez des fichiers pour alimenter la BDD.")
    st.markdown("---")
    st.markdown("#### 🔎 Filtres globaux")
    ph_ctrl   = st.empty()
    ph_agence = st.empty()
    st.markdown("**📅 Période**")
    MOIS_MAP = {"Janvier":1,"Février":2,"Mars":3,"Avril":4,"Mai":5,"Juin":6,
                "Juillet":7,"Août":8,"Septembre":9,"Octobre":10,"Novembre":11,"Décembre":12}
    TRIM_MAP = {"T1 (Jan-Mar)":[1,2,3],"T2 (Avr-Jun)":[4,5,6],
                "T3 (Jul-Sep)":[7,8,9],"T4 (Oct-Déc)":[10,11,12]}
    periode_type = st.selectbox("", ["Toutes","Mois","Trimestre","Année"], label_visibility="collapsed")
    periode_val  = None
    if periode_type == "Mois":      periode_val = st.selectbox("Mois",      list(MOIS_MAP.keys()))
    elif periode_type == "Trimestre": periode_val = st.selectbox("Trimestre", list(TRIM_MAP.keys()))
    elif periode_type == "Année":   periode_val = st.selectbox("Année",     [str(y) for y in range(2023,2028)])
    st.markdown("---")
    st.caption("Direction Audit & CI — © 2026 SKAB")

# ============================================================
# CHARGEMENT DONNÉES
# ============================================================
if source_mode == "📂 Excel (local)":
    if not uploaded_files:
        st.title("🛡️ Dashboard Contrôle Interne — SKAB")
        st.info("👋 Déposez les fichiers des contrôleurs dans la barre latérale pour démarrer.")
        st.stop()
    data = process_consolidation(uploaded_files)
else:
    if not SUPABASE_OK: st.error("Supabase indisponible."); st.stop()
    with st.spinner("Chargement BDD…"): data = load_from_supabase()
    if uploaded_files:
        dl = process_consolidation(uploaded_files)
        for k in data:
            parts = [d for d in [data[k], dl[k]] if not d.empty]
            data[k] = pd.concat(parts, ignore_index=True) if parts else pd.DataFrame()
        st.info(f"👁️ Prévisualisation : BDD + {len(uploaded_files)} fichier(s) local(aux) non sauvegardés")
    if data["ANOMALIES"].empty and data["MISSIONS"].empty:
        st.title("🛡️ Dashboard Contrôle Interne — SKAB")
        st.info("📭 Base vide — déposez des fichiers et cliquez 'Envoyer vers Supabase'.")
        st.stop()

df_mis  = data["MISSIONS"]
df_pts  = data["POINTS"]
df_anom = data["ANOMALIES"]
df_plan = data["PLANS"]

# ============================================================
# NORMALISATION COLONNES
# ============================================================
COL_MAP = {
    "impact":     ["Impact","impact_financier","Impact Financier"],
    "criticite":  ["Criticité","criticite","Criticite","Niveau de Criticité"],
    "domaine":    ["Domaine","domaine","Type"],
    "pays":       ["Pays","pays"],
    "agence":     ["Agence","agence","Site / Entité","Site","Entité"],
    "controleur": ["Contrôleur","controleur","Controleur","Fichier Source","fichier_source"],
    "statut":     ["Statut","statut"],
    "date":       ["Date","date_detection","Date de Détection","created_at"],
    "conformite": ["Taux de Conformité","taux_conformite","conform","Conformité"],
    "num_mis":    ["N° Mission","num_mission","Mission"],
    "id_anom":    ["ID Anomalie","id_anomalie"],
    "id_plan":    ["ID Plan","id_plan"],
    "description":["Description","description","Libellé"],
    "responsable":["Responsable","responsable"],
    "echeance":   ["Échéance","echeance","date_echeance"],
    "statut_plan":["Statut","statut"],
    "date_deb":   ["Début","date_debut"],
    "date_fin":   ["Fin","date_fin"],
}

def find_col(df, key):
    for c in COL_MAP.get(key, []):
        m = next((x for x in df.columns if c.lower() in x.lower()), None)
        if m: return m
    return None

col_impact   = find_col(df_anom, "impact")
col_crit     = find_col(df_anom, "criticite")
col_domaine  = find_col(df_anom, "domaine")
col_pays     = find_col(df_anom, "pays")
col_ag_a     = find_col(df_anom, "agence")
col_ctrl_a   = find_col(df_anom, "controleur")
col_stat_a   = find_col(df_anom, "statut")
col_date_a   = find_col(df_anom, "date")
col_desc_a   = find_col(df_anom, "description")
col_id_a     = find_col(df_anom, "id_anom")
col_tx_conf  = find_col(df_mis,  "conformite")
col_ag_m     = find_col(df_mis,  "agence")
col_ctrl_m   = find_col(df_mis,  "controleur")
col_stat_m   = find_col(df_mis,  "statut")
col_deb_m    = find_col(df_mis,  "date_deb")
col_fin_m    = find_col(df_mis,  "date_fin")
col_num_m    = find_col(df_mis,  "num_mis")
# PLANS
col_id_pl    = find_col(df_plan, "id_plan")
col_anom_pl  = find_col(df_plan, "id_anom")
col_ag_pl    = find_col(df_plan, "agence")
col_ctrl_pl  = find_col(df_plan, "controleur")
col_resp_pl  = find_col(df_plan, "responsable")
col_ech_pl   = find_col(df_plan, "echeance")
col_stat_pl  = find_col(df_plan, "statut_plan")

# ============================================================
# FILTRES DYNAMIQUES
# ============================================================
all_ctrl = sorted(set(
    (df_anom[col_ctrl_a].dropna().astype(str).tolist() if col_ctrl_a and not df_anom.empty else []) +
    (df_mis[col_ctrl_m].dropna().astype(str).tolist()  if col_ctrl_m and not df_mis.empty  else [])
))
all_agences = sorted(set(
    (df_anom[col_ag_a].dropna().astype(str).tolist() if col_ag_a and not df_anom.empty else []) +
    (df_mis[col_ag_m].dropna().astype(str).tolist()  if col_ag_m and not df_mis.empty  else [])
))

ctrl_sel   = ph_ctrl.selectbox("👤 Contrôleur",    ["Tous"] + all_ctrl)
agence_sel = ph_agence.selectbox("🏢 Agence",       ["Toutes"] + all_agences)

# ============================================================
# FONCTION FILTRE UNIVERSELLE
# ============================================================
def filt(df, col_ag=None, col_ctrl=None, col_date=None):
    if agence_sel != "Toutes" and col_ag and col_ag in df.columns:
        df = df[df[col_ag].astype(str) == agence_sel]
    if ctrl_sel != "Tous" and col_ctrl and col_ctrl in df.columns:
        df = df[df[col_ctrl].astype(str) == ctrl_sel]
    if col_date and col_date in df.columns and periode_type != "Toutes":
        df = df.copy()
        df["_d"] = pd.to_datetime(df[col_date], errors="coerce")
        df = df.dropna(subset=["_d"])
        if   periode_type == "Mois"       and periode_val: df = df[df["_d"].dt.month == MOIS_MAP[periode_val]]
        elif periode_type == "Trimestre"  and periode_val: df = df[df["_d"].dt.month.isin(TRIM_MAP[periode_val])]
        elif periode_type == "Année"      and periode_val: df = df[df["_d"].dt.year == int(periode_val)]
        df = df.drop(columns=["_d"])
    return df

df_anom_f = filt(df_anom, col_ag_a, col_ctrl_a, col_date_a)
df_mis_f  = filt(df_mis,  col_ag_m,  col_ctrl_m,  col_deb_m)
df_plan_f = filt(df_plan, col_ag_pl, col_ctrl_pl, col_ech_pl)

# ============================================================
# EN-TÊTE
# ============================================================
st.title("🛡️ Dashboard Contrôle Interne — Groupe SKAB")

badges = []
if ctrl_sel   != "Tous":   badges.append(f"👤 {ctrl_sel}")
if agence_sel != "Toutes": badges.append(f"🏢 {agence_sel}")
if periode_type != "Toutes": badges.append(f"📅 {periode_type} : {periode_val}")
if badges:
    st.markdown("**Filtres actifs :** " + "".join(f'<span class="badge">{b}</span>' for b in badges),
                unsafe_allow_html=True)

# ============================================================
# KPI ROW
# ============================================================
impact_total = pd.to_numeric(df_anom_f[col_impact], errors='coerce').fillna(0).sum() if col_impact and not df_anom_f.empty else 0
nb_crit      = df_anom_f[df_anom_f[col_crit].astype(str).str.contains('Critique|🔴',na=False)].shape[0] if col_crit and not df_anom_f.empty else 0
nb_ec        = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains('EN COURS',na=False)].shape[0] if col_stat_a and not df_anom_f.empty else 0
raw_conf     = pd.to_numeric(df_mis_f[col_tx_conf], errors='coerce').mean() if col_tx_conf and not df_mis_f.empty else None
conformite   = (raw_conf*100 if raw_conf and raw_conf<=1.0 else raw_conf) if raw_conf and not pd.isna(raw_conf) else 0

# Plans en retard
nb_retard = 0
if col_ech_pl and not df_plan_f.empty:
    df_plan_f["_ech"] = pd.to_datetime(df_plan_f[col_ech_pl], errors="coerce")
    if col_stat_pl:
        nb_retard = df_plan_f[
            (df_plan_f["_ech"] < pd.Timestamp.today()) &
            (~df_plan_f[col_stat_pl].astype(str).str.upper().str.contains("CLÔTUR|CLOTUR|TERMINÉ|DONE", na=False))
        ].shape[0]
    df_plan_f = df_plan_f.drop(columns=["_ech"])

k1,k2,k3,k4,k5,k6,k7 = st.columns(7)
k1.metric("💰 Risque Financier",  f"{impact_total:,.0f} F")
k2.metric("🔴 Critiques",         nb_crit,  delta="Urgent"  if nb_crit>0  else None, delta_color="inverse")
k3.metric("⏳ EN COURS",          nb_ec,    delta="À suivre" if nb_ec>0   else None, delta_color="inverse")
k4.metric("⚠️ Plans en retard",   nb_retard,delta="Retard"  if nb_retard>0 else None, delta_color="inverse")
k5.metric("✅ Conformité Moy.",   f"{conformite:.1f}%" if conformite>0 else "N/A")
k6.metric("📋 Missions",          len(df_mis_f))
k7.metric("📁 Total Anomalies",   len(df_anom_f))

st.markdown("---")

# ============================================================
# ONGLETS
# ============================================================
tab1,tab2,tab3,tab4,tab5,tab6,tab7 = st.tabs([
    "📊 Vue d'ensemble",
    "🔴 Anomalies EN COURS",
    "📋 Plans d'action",
    "🏢 Par Agence",
    "👤 Par Contrôleur",
    "📅 Périodique",
    "📄 Rapport & Export",
])

# ══════════════════════════════════════════════════════
# TAB 1 — VUE D'ENSEMBLE
# ══════════════════════════════════════════════════════
with tab1:
    c1, c2 = st.columns(2)

    with c1:
        st.markdown('<div class="sec">📊 Anomalies par Domaine & Criticité</div>', unsafe_allow_html=True)
        if not df_anom_f.empty and col_domaine:
            fig = px.bar(df_anom_f.dropna(subset=[col_domaine]), x=col_domaine,
                         color=col_crit if col_crit else None, barmode='group',
                         color_discrete_map=COLOR_CRIT)
            fig.update_layout(height=320, margin=dict(l=0,r=0,t=5,b=0),
                              plot_bgcolor='rgba(0,0,0,0)', legend_title_text="Criticité",
                              xaxis_title=None, yaxis_title="Nb")
            st.plotly_chart(fig, use_container_width=True)
        else: st.info("Aucune donnée de domaine.")

    with c2:
        st.markdown('<div class="sec">🌍 Répartition géographique</div>', unsafe_allow_html=True)
        if not df_anom_f.empty and col_pays:
            dp = df_anom_f.dropna(subset=[col_pays]).groupby(col_pays).size().reset_index(name="Nb")
            fig2 = px.pie(dp, values="Nb", names=col_pays, hole=.45,
                          color_discrete_sequence=px.colors.qualitative.Safe)
            fig2.update_traces(textposition='inside', textinfo='percent+label')
            fig2.update_layout(height=320, margin=dict(l=0,r=0,t=5,b=20), showlegend=False)
            st.plotly_chart(fig2, use_container_width=True)
        else: st.info("Aucune donnée géographique.")

    # Jauge de conformité
    if conformite > 0:
        st.markdown('<div class="sec">🎯 Jauge de Conformité Globale</div>', unsafe_allow_html=True)
        fig_gauge = go.Figure(go.Indicator(
            mode="gauge+number+delta",
            value=conformite,
            delta={"reference": 80, "suffix":"%"},
            number={"suffix":"%"},
            gauge={
                "axis":{"range":[0,100]},
                "bar":{"color":"#1d3557"},
                "steps":[
                    {"range":[0,50],"color":"#fee2e2"},
                    {"range":[50,75],"color":"#fef9c3"},
                    {"range":[75,100],"color":"#dcfce7"},
                ],
                "threshold":{"line":{"color":"#e63946","width":3},"thickness":0.75,"value":80}
            },
            title={"text":"Taux moyen de conformité"}
        ))
        fig_gauge.update_layout(height=260, margin=dict(l=20,r=20,t=30,b=10))
        st.plotly_chart(fig_gauge, use_container_width=True)

    # Tableau complet
    st.markdown('<div class="sec">📌 Registre des anomalies (vue filtrée)</div>', unsafe_allow_html=True)
    if not df_anom_f.empty:
        cols_r = [c for c in [col_id_a,col_ctrl_a,col_ag_a,col_crit,col_domaine,col_desc_a,col_impact,col_stat_a] if c and c in df_anom_f.columns]
        sort_c = col_impact if col_impact and col_impact in df_anom_f.columns else df_anom_f.columns[0]
        st.dataframe(apply_style(df_anom_f.sort_values(by=sort_c,ascending=False)[cols_r], col_crit),
                     hide_index=True, use_container_width=True, height=300)
    else: st.info("Aucune anomalie.")

    # Diagnostic qualité
    with st.expander("🛠️ Diagnostic Qualité"):
        alerts=[]
        col_np = find_col(df_pts,"num_mis")
        if not df_anom_f.empty and col_crit and col_impact:
            ko = df_anom_f[
                (df_anom_f[col_crit].astype(str).str.contains('Critique|🔴|Majeur|🟠',na=False)) &
                (pd.to_numeric(df_anom_f[col_impact],errors='coerce').fillna(0)==0)]
            for _,r in ko.iterrows():
                v = r.get('ID Anomalie',r.get('id_anomalie','N/A'))
                alerts.append(f"⚠️ Anomalie **{v}** [{r[col_crit]}] — impact financier nul.")
        if col_num_m and not df_mis_f.empty:
            for m in df_mis_f[col_num_m].dropna().unique():
                if "Une mission" in str(m) or str(m).startswith("N°"): continue
                if df_pts.empty or col_np not in df_pts.columns or m not in df_pts[col_np].values:
                    alerts.append(f"❌ Mission **{m}** sans point de contrôle rattaché.")
        for a in alerts: st.warning(a)
        if not alerts: st.success("✅ Intégrité des données : OK")

# ══════════════════════════════════════════════════════
# TAB 2 — ANOMALIES EN COURS
# ══════════════════════════════════════════════════════
with tab2:
    if df_anom_f.empty or not col_stat_a:
        st.info("Colonne 'Statut' non détectée.")
    else:
        df_ec = df_anom_f[df_anom_f[col_stat_a].astype(str).str.upper().str.contains("EN COURS",na=False)].copy()
        if df_ec.empty:
            st.success("✅ Aucune anomalie EN COURS pour ce filtre.")
        else:
            imp_ec   = pd.to_numeric(df_ec[col_impact],errors='coerce').fillna(0).sum() if col_impact else 0
            crit_ec  = df_ec[df_ec[col_crit].astype(str).str.contains('Critique|🔴',na=False)].shape[0] if col_crit else 0
            mj_ec    = df_ec[df_ec[col_crit].astype(str).str.contains('Majeur|🟠',na=False)].shape[0] if col_crit else 0

            m1,m2,m3,m4 = st.columns(4)
            m1.metric("Total EN COURS",    len(df_ec))
            m2.metric("Impact Financier",  f"{imp_ec:,.0f} F")
            m3.metric("Dont Critiques",    crit_ec)
            m4.metric("Dont Majeures",     mj_ec)
            st.markdown("")

            c1, c2 = st.columns(2)
            with c1:
                # Évolution temporelle
                if col_date_a and col_date_a in df_ec.columns:
                    df_ec["_m"] = pd.to_datetime(df_ec[col_date_a],errors="coerce").dt.to_period("M").astype(str)
                    dt = df_ec.dropna(subset=["_m"]).groupby("_m").size().reset_index(name="Nb EN COURS")
                    if not dt.empty:
                        st.markdown('<div class="sec">📈 Évolution mensuelle</div>',unsafe_allow_html=True)
                        fig_t = px.area(dt, x="_m", y="Nb EN COURS", markers=True,
                                        color_discrete_sequence=["#e63946"])
                        fig_t.update_layout(height=260, margin=dict(l=0,r=0,t=5,b=0),
                                            xaxis_title=None, plot_bgcolor='rgba(0,0,0,0)')
                        st.plotly_chart(fig_t, use_container_width=True)
            with c2:
                # Répartition par domaine
                if col_domaine and col_domaine in df_ec.columns:
                    st.markdown('<div class="sec">🗂️ Par Domaine</div>',unsafe_allow_html=True)
                    dd = df_ec.dropna(subset=[col_domaine]).groupby(col_domaine).size().reset_index(name="Nb")
                    fig_d = px.pie(dd, values="Nb", names=col_domaine, hole=.4,
                                   color_discrete_sequence=px.colors.qualitative.Pastel)
                    fig_d.update_layout(height=260, margin=dict(l=0,r=0,t=5,b=0), showlegend=True)
                    st.plotly_chart(fig_d, use_container_width=True)

            st.markdown('<div class="sec">📋 Liste détaillée</div>',unsafe_allow_html=True)
            cols_ec = [c for c in [col_id_a,col_ctrl_a,col_ag_a,col_pays,col_crit,
                                   col_domaine,col_date_a,col_impact,col_desc_a] if c and c in df_ec.columns]
            st.dataframe(apply_style(df_ec[cols_ec] if cols_ec else df_ec, col_crit),
                         hide_index=True, use_container_width=True, height=380)

            st.download_button("⬇️ Exporter EN COURS (.csv)",
                               data=df_ec.to_csv(index=False).encode("utf-8"),
                               file_name=f"SKAB_EN_COURS_{date.today()}.csv", mime="text/csv")

# ══════════════════════════════════════════════════════
# TAB 3 — PLANS D'ACTION
# ══════════════════════════════════════════════════════
with tab3:
    if df_plan_f.empty:
        st.info("Aucun plan d'action dans les données.")
    else:
        # KPI plans
        nb_plans  = len(df_plan_f)
        nb_clos   = df_plan_f[col_stat_pl].astype(str).str.upper().str.contains("CLÔTUR|CLOTUR|TERMINÉ|DONE",na=False).sum() if col_stat_pl else 0
        nb_ouvert = nb_plans - nb_clos

        # Plans en retard (échéance dépassée et non clôturés)
        df_plan_f2 = df_plan_f.copy()
        if col_ech_pl:
            df_plan_f2["_ech2"] = pd.to_datetime(df_plan_f2[col_ech_pl], errors="coerce")
            mask_retard = df_plan_f2["_ech2"] < pd.Timestamp.today()
            if col_stat_pl:
                mask_ouvert = ~df_plan_f2[col_stat_pl].astype(str).str.upper().str.contains("CLÔTUR|CLOTUR|TERMINÉ|DONE",na=False)
                nb_r = (mask_retard & mask_ouvert).sum()
            else:
                nb_r = mask_retard.sum()
        else:
            nb_r = 0

        p1,p2,p3,p4 = st.columns(4)
        p1.metric("Total Plans",       nb_plans)
        p2.metric("Plans Ouverts",     nb_ouvert)
        p3.metric("Plans Clôturés",    int(nb_clos))
        p4.metric("⚠️ En Retard",      nb_r, delta="Retard" if nb_r>0 else None, delta_color="inverse")
        st.markdown("")

        c1, c2 = st.columns(2)
        with c1:
            if col_stat_pl:
                st.markdown('<div class="sec">📊 Statut des Plans</div>',unsafe_allow_html=True)
                ds = df_plan_f[col_stat_pl].value_counts().reset_index()
                ds.columns = ["Statut","Nb"]
                fig_s = px.bar(ds, x="Statut", y="Nb", color="Statut",
                               color_discrete_sequence=px.colors.qualitative.Set2)
                fig_s.update_layout(height=280,margin=dict(l=0,r=0,t=5,b=0),
                                    plot_bgcolor='rgba(0,0,0,0)',showlegend=False)
                st.plotly_chart(fig_s, use_container_width=True)
        with c2:
            if col_ech_pl:
                st.markdown('<div class="sec">📅 Échéances à venir (30 j)</div>',unsafe_allow_html=True)
                df_ech = df_plan_f.copy()
                df_ech["_e"] = pd.to_datetime(df_ech[col_ech_pl], errors="coerce")
                horizon = pd.Timestamp.today() + timedelta(days=30)
                df_soon = df_ech[(df_ech["_e"] >= pd.Timestamp.today()) & (df_ech["_e"] <= horizon)]
                if df_soon.empty:
                    st.info("Aucune échéance dans les 30 prochains jours.")
                else:
                    cols_s = [c for c in [col_id_pl,col_anom_pl,col_ag_pl,col_resp_pl,col_ech_pl,col_stat_pl]
                              if c and c in df_soon.columns]
                    st.dataframe(df_soon[cols_s] if cols_s else df_soon,
                                 hide_index=True, use_container_width=True, height=260)

        # Tableau complet plans
        st.markdown('<div class="sec">📋 Tous les Plans d\'action</div>',unsafe_allow_html=True)
        cols_pl = [c for c in [col_id_pl,col_ctrl_pl,col_ag_pl,col_anom_pl,
                               col_resp_pl,col_ech_pl,col_stat_pl] if c and c in df_plan_f.columns]
        st.dataframe(df_plan_f[cols_pl] if cols_pl else df_plan_f,
                     hide_index=True, use_container_width=True, height=320)

        # Alerte plans en retard
        if nb_r > 0 and col_ech_pl:
            st.markdown('<div class="sec">🚨 Plans en Retard</div>',unsafe_allow_html=True)
            df_ret = df_plan_f2[mask_retard & mask_ouvert].drop(columns=["_ech2"], errors="ignore") if col_stat_pl else df_plan_f2[mask_retard].drop(columns=["_ech2"],errors="ignore")
            st.dataframe(df_ret[cols_pl] if cols_pl else df_ret,
                         hide_index=True, use_container_width=True, height=200)

        st.download_button("⬇️ Exporter Plans d'action (.csv)",
                           data=df_plan_f.to_csv(index=False).encode("utf-8"),
                           file_name=f"SKAB_PLANS_{date.today()}.csv", mime="text/csv")

# ══════════════════════════════════════════════════════
# TAB 4 — PAR AGENCE
# ══════════════════════════════════════════════════════
with tab4:
    agences_dispo = sorted(df_anom[col_ag_a].dropna().astype(str).unique()) if col_ag_a and not df_anom.empty else []
    if not agences_dispo:
        st.info("Aucune colonne Agence détectée.")
    else:
        ag = st.selectbox("Sélectionner une agence :", agences_dispo, key="ag_tab")
        df_ag = df_anom[df_anom[col_ag_a].astype(str)==ag].copy()

        a1,a2,a3,a4 = st.columns(4)
        imp_ag = pd.to_numeric(df_ag[col_impact],errors='coerce').fillna(0).sum() if col_impact else 0
        crit_ag= df_ag[df_ag[col_crit].astype(str).str.contains('Critique|🔴',na=False)].shape[0] if col_crit else 0
        ec_ag  = df_ag[df_ag[col_stat_a].astype(str).str.upper().str.contains('EN COURS',na=False)].shape[0] if col_stat_a else 0
        a1.metric("Anomalies",       len(df_ag))
        a2.metric("Critiques",       crit_ag)
        a3.metric("EN COURS",        ec_ag)
        a4.metric("Impact Financier",f"{imp_ag:,.0f} F")

        c1,c2 = st.columns(2)
        with c1:
            if col_domaine and not df_ag.empty:
                st.markdown('<div class="sec">🔍 Profil risque par domaine</div>',unsafe_allow_html=True)
                dd = df_ag.dropna(subset=[col_domaine]).groupby(col_domaine).size().reset_index(name="Nb")
                if not dd.empty:
                    fig_r = px.bar_polar(dd,r="Nb",theta=col_domaine,color="Nb",
                                         color_continuous_scale="Reds",title=f"Profil — {ag}")
                    fig_r.update_layout(height=320,margin=dict(l=0,r=0,t=40,b=0))
                    st.plotly_chart(fig_r,use_container_width=True)
        with c2:
            if col_crit and not df_ag.empty:
                st.markdown('<div class="sec">📊 Criticité</div>',unsafe_allow_html=True)
                dc = df_ag.dropna(subset=[col_crit])[col_crit].value_counts().reset_index()
                dc.columns=["Criticité","Nb"]
                fig_c = px.bar(dc,x="Criticité",y="Nb",color="Criticité",
                               color_discrete_map=COLOR_CRIT)
                fig_c.update_layout(height=320,margin=dict(l=0,r=0,t=5,b=0),
                                    plot_bgcolor='rgba(0,0,0,0)',showlegend=False)
                st.plotly_chart(fig_c,use_container_width=True)

        st.markdown(f'<div class="sec">📋 Anomalies de {ag}</div>',unsafe_allow_html=True)
        cols_ag=[c for c in [col_id_a,col_ctrl_a,col_crit,col_domaine,col_date_a,col_impact,col_stat_a,col_desc_a] if c and c in df_ag.columns]
        st.dataframe(apply_style(df_ag[cols_ag] if cols_ag else df_ag, col_crit),
                     hide_index=True,use_container_width=True,height=300)

        st.download_button(f"⬇️ Exporter {ag} (.csv)",
                           data=df_ag.to_csv(index=False).encode("utf-8"),
                           file_name=f"SKAB_{ag}_{date.today()}.csv",mime="text/csv")

# ══════════════════════════════════════════════════════
# TAB 5 — PAR CONTRÔLEUR
# ══════════════════════════════════════════════════════
with tab5:
    if not col_ctrl_a or df_anom.empty:
        st.info("Colonne 'Contrôleur' non détectée — elle est déduite automatiquement du nom du fichier Excel.")
    else:
        ctrl_list = sorted(df_anom[col_ctrl_a].dropna().astype(str).unique())
        if not ctrl_list:
            st.info("Aucun contrôleur identifié.")
        else:
            # Tableau comparatif
            st.markdown('<div class="sec">📊 Comparaison entre contrôleurs</div>',unsafe_allow_html=True)
            agg = {"Nb_Anomalies":(col_ctrl_a,"count")}
            if col_impact: agg["Impact_Total"]=(col_impact,lambda x: pd.to_numeric(x,errors='coerce').sum())
            if col_crit:   agg["Nb_Critiques"]=(col_crit,lambda x: x.astype(str).str.contains('Critique|🔴',na=False).sum())
            df_comp = df_anom.groupby(col_ctrl_a).agg(**agg).reset_index().rename(columns={col_ctrl_a:"Contrôleur"})
            if col_stat_a:
                ec = df_anom[df_anom[col_stat_a].astype(str).str.upper().str.contains("EN COURS",na=False)].groupby(col_ctrl_a).size().reset_index(name="EN COURS").rename(columns={col_ctrl_a:"Contrôleur"})
                df_comp = df_comp.merge(ec,on="Contrôleur",how="left").fillna(0)

            c1,c2,c3 = st.columns(3)
            with c1:
                fig1=px.bar(df_comp,x="Contrôleur",y="Nb_Anomalies",color="Nb_Anomalies",
                             color_continuous_scale="Blues",title="Anomalies totales")
                fig1.update_layout(height=280,margin=dict(l=0,r=0,t=35,b=0),
                                   plot_bgcolor='rgba(0,0,0,0)',showlegend=False)
                st.plotly_chart(fig1,use_container_width=True)
            with c2:
                if "Impact_Total" in df_comp.columns:
                    fig2=px.bar(df_comp,x="Contrôleur",y="Impact_Total",color="Impact_Total",
                                 color_continuous_scale="Oranges",title="Impact financier")
                    fig2.update_layout(height=280,margin=dict(l=0,r=0,t=35,b=0),
                                       plot_bgcolor='rgba(0,0,0,0)',showlegend=False)
                    st.plotly_chart(fig2,use_container_width=True)
            with c3:
                if "Nb_Critiques" in df_comp.columns:
                    fig3=px.bar(df_comp,x="Contrôleur",y="Nb_Critiques",color="Nb_Critiques",
                                 color_continuous_scale="Reds",title="Anomalies critiques")
                    fig3.update_layout(height=280,margin=dict(l=0,r=0,t=35,b=0),
                                       plot_bgcolor='rgba(0,0,0,0)',showlegend=False)
                    st.plotly_chart(fig3,use_container_width=True)

            st.dataframe(df_comp,hide_index=True,use_container_width=True)
            st.markdown("---")

            # Drill-down
            st.markdown('<div class="sec">🔎 Fiche détaillée d\'un contrôleur</div>',unsafe_allow_html=True)
            ctrl_d = st.selectbox("Choisir :",ctrl_list,key="ctrl_dd")
            df_ctrl = df_anom[df_anom[col_ctrl_a].astype(str)==ctrl_d].copy()

            d1,d2,d3,d4 = st.columns(4)
            imp_c  = pd.to_numeric(df_ctrl[col_impact],errors='coerce').fillna(0).sum() if col_impact else 0
            crit_c = df_ctrl[df_ctrl[col_crit].astype(str).str.contains('Critique|🔴',na=False)].shape[0] if col_crit else 0
            ec_c   = df_ctrl[df_ctrl[col_stat_a].astype(str).str.upper().str.contains('EN COURS',na=False)].shape[0] if col_stat_a else 0
            d1.metric("Anomalies",       len(df_ctrl))
            d2.metric("Critiques",       crit_c)
            d3.metric("EN COURS",        ec_c)
            d4.metric("Impact Financier",f"{imp_c:,.0f} F")

            # Missions de ce contrôleur
            if not df_mis.empty and col_ctrl_m:
                df_mis_c = df_mis[df_mis[col_ctrl_m].astype(str)==ctrl_d]
                if not df_mis_c.empty:
                    with st.expander(f"📋 Missions de {ctrl_d} ({len(df_mis_c)})"):
                        st.dataframe(df_mis_c,hide_index=True,use_container_width=True)

            cols_c=[c for c in [col_id_a,col_ag_a,col_pays,col_crit,col_domaine,col_date_a,col_impact,col_stat_a,col_desc_a] if c and c in df_ctrl.columns]
            st.dataframe(apply_style(df_ctrl[cols_c] if cols_c else df_ctrl, col_crit),
                         hide_index=True,use_container_width=True,height=320)
            st.download_button(f"⬇️ Exporter {ctrl_d}",
                               data=df_ctrl.to_csv(index=False).encode("utf-8"),
                               file_name=f"SKAB_{ctrl_d}_{date.today()}.csv",mime="text/csv")

# ══════════════════════════════════════════════════════
# TAB 6 — ANALYSE PÉRIODIQUE
# ══════════════════════════════════════════════════════
with tab6:
    if not col_date_a or df_anom.empty:
        st.info("Aucune colonne de date détectée.")
    else:
        df_per = df_anom.copy()
        df_per["_d"] = pd.to_datetime(df_per[col_date_a],errors="coerce")
        df_per = df_per.dropna(subset=["_d"])
        if df_per.empty:
            st.info("Aucune date valide.")
        else:
            df_per["Mois"]      = df_per["_d"].dt.strftime("%Y-%m")
            df_per["Trimestre"] = "T"+df_per["_d"].dt.quarter.astype(str)+" "+df_per["_d"].dt.year.astype(str)
            df_per["Année"]     = df_per["_d"].dt.year.astype(str)

            pt = st.radio("Granularité :",["Mois","Trimestre","Année"],horizontal=True,key="pt")

            agg2={"Nb_Anomalies":(col_impact if col_impact else df_per.columns[0],"count")}
            if col_impact: agg2["Impact_Total"]=(col_impact,lambda x:pd.to_numeric(x,errors='coerce').sum())
            df_grp = df_per.groupby(pt).agg(**agg2).reset_index().sort_values(pt)

            g1,g2 = st.columns(2)
            with g1:
                st.markdown('<div class="sec">📊 Nb anomalies</div>',unsafe_allow_html=True)
                fig_a=px.bar(df_grp,x=pt,y="Nb_Anomalies",color_discrete_sequence=["#1d3557"])
                fig_a.update_layout(height=280,margin=dict(l=0,r=0,t=5,b=0),
                                    xaxis_title=None,plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_a,use_container_width=True)
            with g2:
                if "Impact_Total" in df_grp.columns:
                    st.markdown('<div class="sec">💰 Impact financier cumulé</div>',unsafe_allow_html=True)
                    fig_b=px.line(df_grp,x=pt,y="Impact_Total",markers=True,
                                   color_discrete_sequence=["#e63946"])
                    fig_b.update_layout(height=280,margin=dict(l=0,r=0,t=5,b=0),
                                        xaxis_title=None,plot_bgcolor='rgba(0,0,0,0)')
                    st.plotly_chart(fig_b,use_container_width=True)

            if col_crit:
                st.markdown('<div class="sec">🎨 Ventilation par Criticité</div>',unsafe_allow_html=True)
                dc2=df_per.dropna(subset=[col_crit]).groupby([pt,col_crit]).size().reset_index(name="Nb")
                fig_c=px.bar(dc2,x=pt,y="Nb",color=col_crit,barmode="stack",
                              color_discrete_map=COLOR_CRIT)
                fig_c.update_layout(height=280,margin=dict(l=0,r=0,t=5,b=0),
                                    xaxis_title=None,plot_bgcolor='rgba(0,0,0,0)')
                st.plotly_chart(fig_c,use_container_width=True)

            # Heatmap agence x periode
            if col_ag_a and col_ag_a in df_per.columns:
                st.markdown('<div class="sec">🗺️ Heatmap Agence × Période</div>',unsafe_allow_html=True)
                hm = df_per.groupby([col_ag_a,pt]).size().unstack(fill_value=0)
                fig_h=px.imshow(hm,color_continuous_scale="Reds",aspect="auto",
                                 text_auto=True,title="Nombre d'anomalies")
                fig_h.update_layout(height=350,margin=dict(l=0,r=0,t=40,b=0))
                st.plotly_chart(fig_h,use_container_width=True)

            st.markdown('<div class="sec">📋 Synthèse chiffrée</div>',unsafe_allow_html=True)
            st.dataframe(df_grp,hide_index=True,use_container_width=True)

# ══════════════════════════════════════════════════════
# TAB 7 — RAPPORT & EXPORT
# ══════════════════════════════════════════════════════
with tab7:
    st.markdown('<div class="sec">📄 Rapport de synthèse exécutif</div>',unsafe_allow_html=True)

    # Synthèse narrative automatique
    now_str = datetime.now().strftime("%d/%m/%Y à %H:%M")
    filtre_str = " | ".join(badges) if badges else "Aucun filtre appliqué"

    top_dom = "N/A"
    if col_domaine and not df_anom_f.empty:
        s = df_anom_f[col_domaine].value_counts()
        if not s.empty: top_dom = s.index[0]

    top_ag = "N/A"
    if col_ag_a and not df_anom_f.empty:
        s2 = df_anom_f[col_ag_a].value_counts()
        if not s2.empty: top_ag = s2.index[0]

    pct_ec = f"{nb_ec/len(df_anom_f)*100:.1f}%" if len(df_anom_f)>0 else "N/A"

    st.markdown(f"""
| Indicateur | Valeur |
|---|---|
| 📅 Date de génération | {now_str} |
| 🔎 Filtres appliqués | {filtre_str} |
| 📋 Missions analysées | **{len(df_mis_f)}** |
| ⚠️ Anomalies totales | **{len(df_anom_f)}** |
| 🔴 Anomalies critiques | **{nb_crit}** |
| ⏳ Anomalies EN COURS | **{nb_ec}** ({pct_ec} du total) |
| ⚠️ Plans en retard | **{nb_retard}** |
| 💰 Risque financier cumulé | **{impact_total:,.0f} FCFA** |
| ✅ Taux de conformité moyen | **{conformite:.1f}%** |
| 🏆 Domaine le plus exposé | **{top_dom}** |
| 🏢 Agence la plus exposée | **{top_ag}** |
""")

    # Alertes prioritaires
    st.markdown('<div class="sec">🚨 Points d\'attention prioritaires</div>',unsafe_allow_html=True)
    if nb_crit > 0:
        st.markdown(f'<div class="kpi-alert">🔴 {nb_crit} anomalie(s) CRITIQUE(S) nécessitent une action immédiate</div>',unsafe_allow_html=True)
    if nb_retard > 0:
        st.markdown(f'<div class="kpi-alert">⚠️ {nb_retard} plan(s) d\'action en RETARD — relancer les responsables</div>',unsafe_allow_html=True)
    if conformite > 0 and conformite < 70:
        st.markdown(f'<div class="kpi-alert">📉 Taux de conformité ({conformite:.1f}%) en dessous du seuil acceptable (70%)</div>',unsafe_allow_html=True)
    if nb_crit == 0 and nb_retard == 0 and conformite >= 70:
        st.markdown('<div class="kpi-ok">✅ Aucune alerte critique — situation sous contrôle</div>',unsafe_allow_html=True)

    st.markdown("---")
    st.markdown('<div class="sec">📤 Générer le Fichier Maître Consolidé</div>',unsafe_allow_html=True)

    col_exp1, col_exp2 = st.columns(2)
    with col_exp1:
        inclure_ec   = st.checkbox("Inclure onglet Anomalies EN COURS",  value=True)
        inclure_ret  = st.checkbox("Inclure onglet Plans en Retard",     value=True)
    with col_exp2:
        inclure_ctrl = st.checkbox("Inclure onglet Synthèse Contrôleurs", value=True)
        inclure_synth= st.checkbox("Inclure onglet Synthèse Périodique",  value=True)

    if st.button("🏗️ Compiler et télécharger le Fichier Maître", type="primary", use_container_width=True):
        buf = io.BytesIO()
        with pd.ExcelWriter(buf, engine='xlsxwriter') as writer:
            wb = writer.book

            # Styles xlsxwriter
            fmt_title  = wb.add_format({'bold':True,'font_size':14,'bg_color':'#1e293b','font_color':'#ffffff','border':1})
            fmt_header = wb.add_format({'bold':True,'bg_color':'#e63946','font_color':'#ffffff','border':1,'align':'center'})
            fmt_cell   = wb.add_format({'border':1,'align':'left'})
            fmt_num    = wb.add_format({'border':1,'num_format':'#,##0'})

            # ACCUEIL
            pd.DataFrame({
                "RAPPORT GLOBAL CI SKAB":["Destinataire","Généré par","Horodatage","Filtres","Validation"],
                "MÉTADONNÉES":["M. Élie DIGNOU (DAF)",
                               "Chef de Département Contrôle Interne",
                               now_str, filtre_str, "VÉRIFIÉ ET SCELLÉ"]
            }).to_excel(writer, sheet_name="ACCUEIL", index=False)
            writer.sheets["ACCUEIL"].set_column('A:B',45)

            def write_sheet(name, df):
                if df is not None and not df.empty:
                    df.to_excel(writer, sheet_name=name, index=False)
                    ws = writer.sheets[name]
                    for i,c in enumerate(df.columns):
                        ws.set_column(i,i,get_safe_len(df[c],c))

            write_sheet("MISSIONS",        data["MISSIONS"])
            write_sheet("POINTS_CONTROLE", data["POINTS"])
            write_sheet("ANOMALIES",       data["ANOMALIES"])
            write_sheet("PLANS_ACTION",    data["PLANS"])

            if inclure_ec and col_stat_a and not df_anom.empty:
                df_ec_x = df_anom[df_anom[col_stat_a].astype(str).str.upper().str.contains("EN COURS",na=False)]
                write_sheet("ANOMALIES_EN_COURS", df_ec_x)

            if inclure_ret and nb_retard > 0 and col_ech_pl:
                df_r2 = df_plan.copy()
                df_r2["_e"]=pd.to_datetime(df_r2[col_ech_pl],errors="coerce")
                mask_r=df_r2["_e"]<pd.Timestamp.today()
                if col_stat_pl:
                    mask_o=~df_r2[col_stat_pl].astype(str).str.upper().str.contains("CLÔTUR|CLOTUR|TERMINÉ|DONE",na=False)
                    df_ret_x=df_r2[mask_r&mask_o].drop(columns=["_e"],errors="ignore")
                else:
                    df_ret_x=df_r2[mask_r].drop(columns=["_e"],errors="ignore")
                write_sheet("PLANS_EN_RETARD", df_ret_x)

            if inclure_ctrl and col_ctrl_a and not df_anom.empty:
                agg3={"Nb_Anomalies":(col_ctrl_a,"count")}
                if col_impact: agg3["Impact_Total"]=(col_impact,lambda x:pd.to_numeric(x,errors='coerce').sum())
                df_synth_c=df_anom.groupby(col_ctrl_a).agg(**agg3).reset_index()
                write_sheet("SYNTHESE_CONTROLEURS", df_synth_c)

            if inclure_synth and col_date_a and not df_anom.empty:
                dp2=df_anom.copy()
                dp2["_d"]=pd.to_datetime(dp2[col_date_a],errors="coerce")
                dp2=dp2.dropna(subset=["_d"])
                if not dp2.empty:
                    dp2["Mois"]=dp2["_d"].dt.strftime("%Y-%m")
                    agg4={"Nb_Anomalies":(col_impact if col_impact else dp2.columns[0],"count")}
                    if col_impact: agg4["Impact_Total"]=(col_impact,lambda x:pd.to_numeric(x,errors='coerce').sum())
                    write_sheet("SYNTHESE_PERIODIQUE", dp2.groupby("Mois").agg(**agg4).reset_index())

        st.success("🎉 Fichier Maître compilé avec succès !")
        st.download_button(
            label=f"💾 Télécharger SKAB_MAITRE_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            data=buf.getvalue(),
            file_name=f"SKAB_MAITRE_{datetime.now().strftime('%Y%m%d_%H%M')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            use_container_width=True
        )

st.markdown("---")
st.caption("Direction Générale SKAB Nutrition — Contrôle Interne v4.0 | Supabase Edition")
