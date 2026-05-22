import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from datetime import datetime
from sqlalchemy import create_engine, text

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


# --- 2. SÉCURITÉ RESSOURCE : CRÉATION DU SCHÉMA SI INEXISTANT (DDL EXPLICITE) ---
def init_database_schema(engine_pg):
    """Exécute des requêtes SQL natives pour garantir la création des tables dans Supabase"""
    queries = [
        """
        CREATE TABLE IF NOT EXISTS table_anomalies (
            id_anomalie TEXT, date_detection TEXT, site_entite TEXT, pays TEXT,
            type_domaine TEXT, niveau_criticite TEXT, description TEXT, cause_racine_identifiee TEXT,
            impact_estime_fcfa NUMERIC, responsable_traitement TEXT, statut TEXT, date_cloture TEXT,
            lien_plan_action TEXT, n_mission_rattachee TEXT, controleur TEXT,
            meta_source_file TEXT, meta_import_date TEXT
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS table_missions (
            n_mission TEXT, date_debut TEXT, date_fin TEXT, domaine TEXT, type_controle TEXT,
            pays_entite TEXT, site_agence TEXT, responsable_site TEXT, statut_mission TEXT,
            nb_points_oui NUMERIC, nb_points_non NUMERIC, nb_points_n_a NUMERIC,
            taux_conformite NUMERIC, nb_anomalies NUMERIC, commentaire_general TEXT,
            meta_source_file TEXT, meta_import_date TEXT
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS table_points_controle (
            id_point TEXT, n_mission TEXT, date TEXT, domaine TEXT, section TEXT, code_point TEXT,
            libelle_point_de_controle TEXT, resultat TEXT, observation_constat TEXT,
            piece_justificative TEXT, criticite_si_non TEXT, action_immediate TEXT, controleur TEXT,
            meta_source_file TEXT, meta_import_date TEXT
        );
        """,
        """
        CREATE TABLE IF NOT EXISTS table_plans_action (
            id_plan TEXT, id_anomalie_liee TEXT, action_corrective_a_mener TEXT, responsable TEXT,
            echeance TEXT, date_realisation TEXT, statut TEXT, avancement TEXT, commentaire_suivi TEXT,
            piece_de_preuve TEXT, controleur TEXT, meta_source_file TEXT, meta_import_date TEXT
        );
        """
    ]
    with engine_pg.connect() as schema_conn:
        transaction = schema_conn.begin()
        try:
            for q in queries:
                schema_conn.execute(text(q))
            transaction.commit()
        except Exception as schema_err:
            transaction.rollback()
            st.error(f"⚠️ Impossible d'initialiser le schéma SQL : {schema_err}")


def clean_column_name(col):
    """Nettoie proprement les en-têtes pour éviter les erreurs de syntaxe PostgreSQL"""
    s = str(col).strip().lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")
    s = re.sub(r"[/\-()°’'%]", " ", s)
    s = re.sub(r"\s+", "_", s)
    return s.strip("_")

def load_table(table_name):
    """Charge une table SQL sous forme de DataFrame, retourne un DF vide si absente"""
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
        Déposez ici le classeur Excel d'un contrôleur. Le système va forcer la création des tables manquantes 
        dans votre console **Supabase** avant de pousser les lignes de données.
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
                    "Ajouter les données à la suite de l'historique existant",
                    "⚠️ Vider la base et réinitialiser toutes les tables à neuf (Purge complète)"
                ]
            )
            
            if st.button("🚀 Lancer la synchronisation globale des tables", type="primary", use_container_width=True):
                
                # 🛡️ APPEL SÉCURITÉ CRITIQUE : Matérialise les tables si absentes de Supabase
                init_database_schema(engine)
                
                # Si l'utilisateur demande une réinitialisation complète
                if "Vider" in mode_import:
                    with engine.connect() as clear_conn:
                        trans = clear_conn.begin()
                        try:
                            for db_table in target_sheets.values():
                                clear_conn.execute(text(f"TRUNCATE TABLE {db_table};"))
                            trans.commit()
                            st.warning("🗑️ Base de données vidée (Contenu purgé). Injection des données neuves...")
                        except Exception as tr_ex:
                            trans.rollback()
                
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
                            
                            # Ajout des métadonnées
                            df_clean['meta_source_file'] = src_file.name
                            df_clean['meta_import_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            # Aligner les colonnes du DataFrame sur le schéma de la table SQL créée pour éviter tout écart
                            query_cols = conn.query(f"SELECT * FROM {db_table} LIMIT 0;")
                            db_cols = list(query_cols.columns)
                            
                            for c in db_cols:
                                if c not in df_clean.columns:
                                    df_clean[c] = None
                            df_clean = df_clean[db_cols]
                            
                            # Injection sécurisée par lot (Append forcé puisque la structure existe désormais)
                            df_clean.to_sql(db_table, con=engine, if_exists="append", index=False)
                            st.caption(f"✅ Données enregistrées dans `{db_table}` ({df_clean.shape[0]} lignes).")
                            success_count += 1
                    
                    progress_bar.progress((idx + 1) / len(target_sheets))
                
                if success_count > 0:
                    st.success(f"🎉 Opération validée. Les tables ont été créées et alimentées avec succès sur Supabase !")
                    st.balloons()
                    
        except Exception as ex:
            st.error(f"❌ Erreur critique lors de l'injection : {ex}")


# ==============================================================================
# ONGLET 2 : TABLEAUX DE BORD
# ==============================================================================
with tabs[1]:
    st.header("📊 Consolidation Automatique du Groupe")
    
    df_anom = load_table("table_anomalies")
    
    if df_anom.empty:
        st.warning("💡 La base SQL ne contient actuellement aucune donnée. Veuillez injecter un premier classeur Excel pour l'initialiser.")
    else:
        df_anom.columns = [c.lower() for c in df_anom.columns]
        
        c_site = next((c for c in df_anom.columns if 'site' in c or 'entite' in c), df_anom.columns[2])
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
                        st.success("✅ Requête exécutée avec succès.")
            except Exception as sql_err:
                st.error(f"❌ Erreur SQL : {sql_err}")
