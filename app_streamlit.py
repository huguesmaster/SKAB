import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from datetime import datetime
from sqlalchemy import create_engine, text, inspect

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="SKAB - Console Supervision Contrôle Interne",
    page_icon="🛡️",
    layout="wide"
)

st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 26px; }
    .stAlert { margin-top: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 20px; }
    .stTabs [data-baseweb="tab"] { height: 45px; font-weight: bold; font-size: 15px; }
    </style>
""", unsafe_allow_html=True)


# --- 1. CONNEXION À LA BASE DE DONNÉES ---
try:
    conn = st.connection("supabase", type="sql")
    creds = st.secrets["connections"]["supabase"]
    db_url = f"postgresql://{creds['username']}:{creds['password']}@{creds['host']}:{creds['port']}/{creds['database']}"
    engine = create_engine(db_url)
except Exception as e:
    st.error("🔒 Erreur de connexion : Les identifiants Supabase sont introuvables ou erronés dans les Secrets.")
    st.stop()


# --- 2. FONCTIONS DE NETTOYAGE ET CHARGEMENT ---
def clean_column_name(col):
    """Nettoie proprement les en-têtes pour éviter les erreurs de syntaxe PostgreSQL"""
    s = str(col).strip().lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")
    s = re.sub(r"[/\-()°’'%]", " ", s)
    s = re.sub(r"\s+", "_", s)
    return s.strip("_")

def load_table(table_name):
    """Charge une table SQL sous forme de DataFrame, retourne un DF vide si la table n'existe pas"""
    try:
        return pd.DataFrame(conn.query(f"SELECT * FROM {table_name};", ttl="2s"))
    except Exception:
        return pd.DataFrame()


# --- 3. ARCHITECTURE DE L'APPLICATION ---
st.title("🛡️ SKAB NUTRITION — Console de Supervision du Contrôle Interne")
st.caption("Espace d'administration et d'analyse de données - Réservé au Chef de Département")

tabs = st.tabs([
    "📥 Injection des Livrables", 
    "📊 Tableaux de Bord", 
    "🔍 Requêteur SQL & Base de Données"
])


# ==============================================================================
# ONGLET 1 : INJECTION CENTRALISÉE AVEC CRÉATION AUTOMATIQUE DES TABLES
# ==============================================================================
with tabs[0]:
    st.header("🗂️ Centralisation et Structuration des rapports terrains")
    st.markdown("""
        Déposez ici le classeur Excel d'un contrôleur. Si les tables n'existent pas dans votre console **Supabase**, 
        **le système les créera automatiquement** avec les bonnes colonnes.
    """)
    
    src_file = st.file_uploader("Sélectionnez le fichier Excel à intégrer (.xlsx) :", type="xlsx")
    
    if src_file:
        target_sheets = {
            "ANOMALIES": "table_anomalies",
            "MES_MISSIONS": "table_missions",
            "POINTS_CONTROLE": "table_points_controle",
            "PLANS_ACTION": "table_plans_action"
        }
        
        try:
            excel_obj = pd.ExcelFile(src_file)
            available_sheets = excel_obj.sheet_names
            
            st.info(f"📁 Fichier détecté : `{src_file.name}` (Onglets présents : {', '.join(available_sheets)})")
            
            mode_import = st.radio(
                "Stratégie de stockage dans Supabase :",
                [
                    "Ajouter les données à la suite de l'historique existant (Crée la table si elle n'existe pas)",
                    "⚠️ Vider la base et réinitialiser toutes les tables à neuf (Écrase ou Recrée proprement)"
                ]
            )
            
            if st.button("🚀 Lancer la synchronisation globale des tables", type="primary", use_container_width=True):
                
                # Inspection des tables actuellement existantes sur Supabase
                inspector = inspect(engine)
                existing_tables = inspector.get_table_names()
                
                progress_bar = st.progress(0)
                success_count = 0
                
                for idx, (sheet_name, db_table) in enumerate(target_sheets.items()):
                    if sheet_name in available_sheets:
                        df_raw = pd.read_excel(src_file, sheet_name=sheet_name, header=None)
                        header_idx = 0
                        
                        for r_idx, row in df_raw.iterrows():
                            row_str = " ".join([str(v) for v in row.values])
                            if any(k in row_str for k in ["ID", "N°", "Date", "Statut"]):
                                header_idx = r_idx
                                break
                        
                        df_clean = pd.read_excel(src_file, sheet_name=sheet_name, skiprows=header_idx)
                        
                        if not df_clean.empty:
                            df_clean.columns = [clean_column_name(c) for c in df_clean.columns]
                            
                            first_col = df_clean.columns[0]
                            df_clean = df_clean.dropna(subset=[first_col])
                            df_clean = df_clean[~df_clean[first_col].astype(str).str.contains("une_anomalie|id_anomalie|exemple", na=False, case=False)]
                            
                            # Métadonnées d'audit
                            df_clean['meta_source_file'] = src_file.name
                            df_clean['meta_import_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            # --- GESTION DYNAMIQUE DE LA STRUCTURE SQL ---
                            if "Vider" in mode_import:
                                # Si l'utilisateur demande un reset complet, on force "replace" pour générer/écraser proprement
                                df_clean.to_sql(db_table, con=engine, if_exists="replace", index=False)
                                st.caption(f"⚡ Table `{db_table}` recréée à neuf et initialisée ({df_clean.shape[0]} lignes).")
                            else:
                                # Mode normal (Append)
                                if db_table in existing_tables:
                                    df_clean.to_sql(db_table, con=engine, if_exists="append", index=False)
                                    st.caption(f"✅ Table `{db_table}` mise à jour ({df_clean.shape[0]} lignes insérées).")
                                else:
                                    # Si la table n'existe pas du tout sur Supabase, on force la création automatique initiale via "replace"
                                    df_clean.to_sql(db_table, con=engine, if_exists="replace", index=False)
                                    st.caption(f"✨ Table `{db_table}` détectée absente : créée automatiquement avec succès ({df_clean.shape[0]} lignes).")
                            
                            success_count += 1
                    
                    progress_bar.progress((idx + 1) / len(target_sheets))
                
                if success_count > 0:
                    st.success(f"🎉 Opération validée. Les tables ont été matérialisées et alimentées sur Supabase !")
                    st.balloons()
                    
        except Exception as ex:
            st.error(f"❌ Erreur critique lors de l'injection : {ex}")


# ==============================================================================
# ONGLET 2 : TABLEAUX DE BORD
# ==============================================================================
with tabs[1]:
    st.header("📊 Consolidation Automatique du Groupe")
    
    df_anom = load_table("table_anomalies")
    df_miss = load_table("table_missions")
    
    if df_anom.empty:
        st.warning("💡 La base SQL ne contient actuellement aucune donnée. Veuillez injecter un classeur Excel dans le premier onglet pour l'initialiser.")
    else:
        df_anom.columns = [c.lower() for c in df_anom.columns]
        
        c_site = next((c for c in df_anom.columns if 'site' in c or 'entite' in c), df_anom.columns[2])
        c_date = next((c for c in df_anom.columns if 'date' in c), df_anom.columns[1])
        c_impact = next((c for c in df_anom.columns if 'impact' in c), None)
        c_crit = next((c for c in df_anom.columns if 'crit' in c), None)
        c_pays = next((c for c in df_anom.columns if 'pays' in c), None)
        c_statut = next((c for c in df_anom.columns if 'statut' in c), None)

        liste_pays = ["Toutes les filiales"] + list(df_anom[c_pays].dropna().unique()) if c_pays else ["Toutes les filiales"]
        pays_selectionne = st.selectbox("🌍 Périmètre d'analyse géographique :", liste_pays)
        
        if pays_selectionne != "Toutes les filiales" and c_pays:
            df_anom = df_anom[df_anom[c_pays] == pays_selectionne]

        st.markdown("### 📌 Indicateurs de Vulnérabilité Majeure")
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            impact_total = pd.to_numeric(df_anom[c_impact], errors='coerce').fillna(0).sum() if c_impact else 0
            st.metric("Risque Financier Global", f"{impact_total:,.0f} FCFA")
        with k2:
            nb_crit = df_anom[df_anom[c_crit].astype(str).str.contains('critique|🔴', na=False, case=False)].shape[0] if c_crit else 0
            st.metric("Alertes Critiques", nb_crit)
        with k3:
            nb_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS|OUVERT", na=False)].shape[0] if c_statut else 0
            st.metric("Anomalies non résolues", nb_encours)
        with k4:
            st.metric("Total Écarts en Base", df_anom.shape[0])

        st.divider()
        
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**🔍 Répartition des Risques par Criticité & Site**")
            if c_crit and c_site:
                fig = px.bar(df_anom, x=c_site, color=c_crit, barmode='stack')
                fig.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
                st.plotly_chart(fig, use_container_width=True)
        with g2:
            st.markdown("**📂 Origine des fichiers sources intégrés**")
            if 'meta_source_file' in df_anom.columns:
                df_src_summary = df_anom.groupby('meta_source_file').size().reset_index(name="Nombre d'anomalies")
                fig2 = px.pie(df_src_summary, values="Nombre d'anomalies", names='meta_source_file', hole=0.4)
                fig2.update_layout(height=300, margin=dict(l=10, r=10, t=10, b=10))
                st.plotly_chart(fig2, use_container_width=True)

        st.markdown("### 📋 Registre Général de Contrôle")
        st.dataframe(df_anom, hide_index=True, use_container_width=True)


# ==============================================================================
# ONGLET 3 : REQUÊTEUR SQL NATIVE
# ==============================================================================
with tabs[2]:
    st.header("🛢️ Console SQL & États Réels des Tables Supabase")
    
    st.subheader("🖋️ Saisir ou coller une requête SQL")
    ex_query = "SELECT table_name FROM information_schema.tables WHERE table_schema='public';"
    user_sql = st.text_area("Requête PostgreSQL :", value=ex_query, height=140)
    
    if st.button("⚡ Exécuter la requête SQL sur Supabase", type="primary"):
        if user_sql.strip():
            try:
                with engine.connect() as r_conn:
                    result_sql = r_conn.execute(text(user_sql))
                    if result_sql.returns_rows:
                        df_query_res = pd.DataFrame(result_sql.fetchall(), columns=result_sql.keys())
                        st.success(f"🎯 Requête exécutée. {df_query_res.shape[0]} lignes renvoyées.")
                        st.dataframe(df_query_res, use_container_width=True)
                    else:
                        st.success("✅ Requête exécutée avec succès (aucune ligne renvoyée).")
            except Exception as sql_err:
                st.error(f"❌ Erreur SQL : {sql_err}")
