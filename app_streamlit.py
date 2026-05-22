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


# --- 1. CONNEXION VIA L'URL DE LA BASE DE DONNÉES (DATABASE_URL) ---
try:
    # Récupération de l'URL complète depuis les Secrets de Streamlit
    # Rôle 'postgres' (Super-utilisateur) configuré par défaut via les paramètres Supabase
    DATABASE_URL = st.secrets["connections"]["supabase"]["DATABASE_URL"]
    
    # Création de l'engine SQL basé sur le driver psycopg2 natif
    engine = create_engine(DATABASE_URL)
except Exception as e:
    st.error("🔒 Impossible de charger l'URL de la base de données. Assurez-vous d'avoir configuré 'DATABASE_URL' dans vos secrets Streamlit.")
    st.stop()


# --- 2. SCHÉMA INITIAL : CRÉATION FORCÉE EN CAS DE BASE VIERGE ---
def force_init_supabase_schema(sql_engine):
    """Garantit la création des structures physiques si elles sont absentes de la console"""
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
    with sql_engine.connect() as admin_conn:
        trans = admin_conn.begin()
        try:
            for query in queries:
                admin_conn.execute(text(query))
            trans.commit()
        except Exception as err:
            trans.rollback()
            st.error(f"💥 Erreur lors de l'initialisation forcée des tables : {err}")


def clean_column_name(col):
    """Standardise les en-têtes de colonnes au format PostgreSQL (minuscules, sans caractères spéciaux)"""
    s = str(col).strip().lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")
    s = re.sub(r"[/\-()°’'%]", " ", s)
    s = re.sub(r"\s+", "_", s)
    return s.strip("_")

def load_table_from_supabase(table_name):
    """Charge de manière sécurisée une table sous forme de DataFrame pour les graphiques"""
    try:
        with engine.connect() as read_conn:
            return pd.read_sql_query(text(f"SELECT * FROM {table_name};"), read_conn)
    except Exception:
        return pd.DataFrame()


# --- 3. ARCHITECTURE DE L'APPLICATION ---
st.title("🛡️ SKAB NUTRITION — Console de Supervision du Contrôle Interne")
st.caption("Espace d'administration centralisé - Réservé au Chef de Département")

tabs = st.tabs([
    "📥 Injection des Livrables", 
    "📊 Tableaux de Bord", 
    "🔍 Requêteur SQL & Base de Données"
])


# ==============================================================================
# ONGLET 1 : INJECTION SÉCURISÉE DEPUIS EXCEL
# ==============================================================================
with tabs[0]:
    st.header("🗂️ Centralisation et Structuration des rapports terrains")
    st.markdown("""
        Déposez le classeur Excel d'un contrôleur. Le système va automatiquement se connecter à votre URI Supabase,
        vérifier les structures et pousser les nouvelles données.
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
                
                # 🚀 FORCE SYNC SCHEMA : On crée les tables si la base est vide
                force_init_supabase_schema(engine)
                
                # Optionnel : Vidage du contenu existant
                if "Vider" in mode_import:
                    with engine.connect() as clear_conn:
                        trans = clear_conn.begin()
                        try:
                            for db_table in target_sheets.values():
                                clear_conn.execute(text(f"TRUNCATE TABLE {db_table};"))
                            trans.commit()
                            st.warning("🗑️ Base de données vidée (Structures conservées). Remplissage en cours...")
                        except Exception:
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
                            
                            # Métadonnées d'administration
                            df_clean['meta_source_file'] = src_file.name
                            df_clean['meta_import_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            # Alignement parfait du DataFrame sur le schéma réel
                            with engine.connect() as col_conn:
                                query_cols = pd.read_sql_query(text(f"SELECT * FROM {db_table} LIMIT 0;"), col_conn)
                            
                            db_cols = list(query_cols.columns)
                            for c in db_cols:
                                if c not in df_clean.columns:
                                    df_clean[c] = None
                            df_clean = df_clean[db_cols]
                            
                            # Insertion via l'engine direct alimenté par DATABASE_URL
                            with engine.connect() as write_conn:
                                df_clean.to_sql(db_table, con=write_conn, if_exists="append", index=False)
                            
                            st.caption(f"✅ Synchro réussie pour `{db_table}` ({df_clean.shape[0]} lignes insérées).")
                            success_count += 1
                    
                    progress_bar.progress((idx + 1) / len(target_sheets))
                
                if success_count > 0:
                    st.success(f"🎉 Les données ont été packagées et injectées avec succès sur Supabase via l'URL administrative !")
                    st.balloons()
                    
        except Exception as ex:
            st.error(f"❌ Échec de l'intégration : {ex}")


# ==============================================================================
# ONGLET 2 : TABLEAUX DE BORD (CONSOLIDATION RAPIDE)
# ==============================================================================
with tabs[1]:
    st.header("📊 Consolidation Automatique du Groupe")
    
    df_anom = load_table_from_supabase("table_anomalies")
    
    if df_anom.empty:
        st.warning("💡 En attente d'un premier fichier d'importation pour générer les indicateurs Groupe.")
    else:
        df_anom.columns = [c.lower() for c in df_anom.columns]
        
        c_site = next((c for c in df_anom.columns if 'site' in c or 'entite' in c), df_anom.columns[2])
        c_impact = next((c for c in df_anom.columns if 'impact' in c), None)
        c_crit = next((c for c in df_anom.columns if 'crit' in c), None)
        c_pays = next((c for c in df_anom.columns if 'pays' in c), None)
        c_statut = next((c for c in df_anom.columns if 'statut' in c), None)

        liste_pays = ["Toutes les filiales"] + list(df_anom[c_pays].dropna().unique()) if c_pays else ["Toutes les filiales"]
        pays_selectionne = st.selectbox("🌍 Filtrer par filiale :", liste_pays)
        
        if pays_selectionne != "Toutes les filiales" and c_pays:
            df_anom = df_anom[df_anom[c_pays] == pays_selectionne]

        st.markdown("### 📌 Indicateurs Majeurs du Groupe")
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            impact_total = pd.to_numeric(df_anom[c_impact], errors='coerce').fillna(0).sum() if c_impact else 0
            st.metric("Risque Financier Global", f"{impact_total:,.0f} FCFA")
        with k2:
            nb_crit = df_anom[df_anom[c_crit].astype(str).str.contains('critique|🔴', na=False, case=False)].shape[0] if c_crit else 0
            st.metric("Alertes Critiques", nb_crit)
        with k3:
            nb_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS|OUVERT", na=False)].shape[0] if c_statut else 0
            st.metric("Anomalies en cours", nb_encours)
        with k4:
            st.metric("Écarts Totaux Répertoriés", df_anom.shape[0])

        st.divider()
        st.markdown("### 📋 Registre d'Audit Consolidé")
        st.dataframe(df_anom, hide_index=True, use_container_width=True)


# ==============================================================================
# ONGLET 3 : CONSOLE DE REQUÊTAGE DIRECTE POSTGRESQL
# ==============================================================================
with tabs[2]:
    st.header("🔍 Console SQL native (Accès direct Supabase)")
    st.markdown("Exécutez vos requêtes analytiques sur vos tables consolides.")
    
    st.subheader("🖋️ Éditeur de requêtes PostgreSQL")
    ex_query = "SELECT * FROM table_anomalies LIMIT 10;"
    user_sql = st.text_area("Saisir la requête SQL :", value=ex_query, height=120)
    
    if st.button("⚡ Exécuter la requête", type="primary"):
        if user_sql.strip():
            try:
                with engine.connect() as query_conn:
                    result_sql = query_conn.execute(text(user_sql))
                    if result_sql.returns_rows:
                        df_query_res = pd.DataFrame(result_sql.fetchall(), columns=result_sql.keys())
                        st.success(f"🎯 Requête exécutée. {df_query_res.shape[0]} lignes trouvées.")
                        st.dataframe(df_query_res, use_container_width=True)
                    else:
                        st.success("✅ Commande de modification exécutée avec succès sur le serveur.")
            except Exception as sql_err:
                st.error(f"❌ Erreur SQL renvoyée par Supabase : {sql_err}")
