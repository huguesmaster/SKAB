import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
import psycopg2  # Connecteur natif PostgreSQL
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


# --- 1. CONNEXION SÉCURISÉE VIA LA CHAÎNE DIRECTE DU POOLER ---
try:
    # Récupération de la chaîne complète stockée de manière étanche dans vos secrets Streamlit
    DATABASE_URL = st.secrets["connections"]["supabase"]["DATABASE_URL"]
    
    # Création de l'engine avec le paramètre d'optimisation du pooler
    engine = create_engine(DATABASE_URL, pool_pre_ping=True)
    
except Exception as e:
    st.error("🔒 Configuration manquante : Assurez-vous d'avoir défini la variable 'DATABASE_URL' dans vos Secrets Streamlit.")
    st.stop()


# --- 2. INITIALISATION FORCÉE DU SCHÉMA (ORDRES DDL ADMINISTRATEUR) ---
def force_init_supabase_schema(sql_engine):
    """Garantit la création physique des tables dans le schéma public de Supabase"""
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
            st.error(f"💥 Impossible d'initialiser les tables sur le serveur : {err}")


def clean_column_name(col):
    """Nettoie les en-têtes des fichiers Excel pour correspondre aux normes PostgreSQL"""
    s = str(col).strip().lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")
    s = re.sub(r"[/\-()°’'%]", " ", s)
    s = re.sub(r"\s+", "_", s)
    return s.strip("_")

def load_table_from_supabase(table_name):
    """Charge de manière sécurisée une table sous forme de DataFrame, évite le crash si vide"""
    try:
        with engine.connect() as read_conn:
            return pd.read_sql_query(text(f"SELECT * FROM {table_name};"), read_conn)
    except Exception:
        return pd.DataFrame()


# --- 3. ARCHITECTURE GENERALE DES ONGLETS ---
st.title("🛡️ SKAB NUTRITION — Console de Supervision du Contrôle Interne")
st.caption("Espace d'administration centralisé — Canal sécurisé Pooler Supabase")

tabs = st.tabs([
    "📥 Injection des Livrables", 
    "📊 Tableaux de Bord", 
    "🔍 Requêteur SQL & Base de Données"
])


# ==============================================================================
# ONGLET 1 : INJECTION SÉCURISÉE DES CLASSEURS TERRAINS
# ==============================================================================
with tabs[0]:
    st.header("🗂️ Centralisation et Structuration des rapports terrains")
    st.markdown("""
        Déposez ici le classeur Excel d'un contrôleur. Le système va forcer la création des tables manquantes 
        sur votre instance **Supabase** via le tunnel d'écriture sécurisé avant d'y transférer les lignes.
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
            
            st.info(f"📁 Fichier chargé en mémoire : `{src_file.name}` (Onglets détectés : {', '.join(available_sheets)})")
            
            mode_import = st.radio(
                "Stratégie de stockage dans la base de données :",
                [
                    "Ajouter les données à la suite de l'historique global",
                    "⚠️ Vider la base en ligne et réinitialiser toutes les tables à neuf (Purge complète)"
                ]
            )
            
            if st.button("🚀 Valider et injecter dans la table Supabase", type="primary", use_container_width=True):
                
                # 🔥 APPEL CRITIQUE : Crée les structures physiques si elles n'existent pas
                force_init_supabase_schema(engine)
                
                # Gestion de la réinitialisation complète si cochée
                if "Vider" in mode_import:
                    with engine.connect() as clear_conn:
                        trans = clear_conn.begin()
                        try:
                            for db_table in target_sheets.values():
                                clear_conn.execute(text(f"TRUNCATE TABLE {db_table};"))
                            trans.commit()
                            st.warning("🗑️ Base de données vidée avec succès. Début du transfert des lignes...")
                        except Exception:
                            trans.rollback()
                
                progress_bar = st.progress(0)
                success_count = 0
                
                for idx, (sheet_name, db_table) in enumerate(target_sheets.items()):
                    if sheet_name in available_sheets:
                        df_raw = pd.read_excel(src_file, sheet_name=sheet_name, header=None)
                        header_idx = 0
                        
                        # Détection dynamique de la ligne d'en-tête
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
                            
                            # Injection des métadonnées d'audit
                            df_clean['meta_source_file'] = src_file.name
                            df_clean['meta_import_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            # Alignement strict sur la structure réelle en base PostgreSQL
                            with engine.connect() as col_conn:
                                query_cols = pd.read_sql_query(text(f"SELECT * FROM {db_table} LIMIT 0;"), col_conn)
                            
                            db_cols = list(query_cols.columns)
                            for c in db_cols:
                                if c not in df_clean.columns:
                                    df_clean[c] = None
                            df_clean = df_clean[db_cols]
                            
                            # Écriture par lot sécurisée
                            with engine.connect() as write_conn:
                                df_clean.to_sql(db_table, con=write_conn, if_exists="append", index=False)
                            
                            st.caption(f"✅ Table `{db_table}` synchronisée ({df_clean.shape[0]} lignes enregistrées).")
                            success_count += 1
                    
                    progress_bar.progress((idx + 1) / len(target_sheets))
                
                if success_count > 0:
                    st.success(f"🎉 Succès ! Vos livrables ont été injectés avec succès sur votre instance Supabase.")
                    st.balloons()
                    
        except Exception as ex:
            st.error(f"❌ Erreur lors du traitement ou du transfert SQL : {ex}")


# ==============================================================================
# ONGLET 2 : RESTITUTION ET METRICS
# ==============================================================================
with tabs[1]:
    st.header("📊 Consolidation Automatique du Groupe")
    
    df_anom = load_table_from_supabase("table_anomalies")
    
    if df_anom.empty:
        st.warning("💡 La base de données en ligne est actuellement vide. Veuillez importer un premier fichier Excel.")
    else:
        df_anom.columns = [c.lower() for c in df_anom.columns]
        
        c_impact = next((c for c in df_anom.columns if 'impact' in c), None)
        c_crit = next((c for c in df_anom.columns if 'crit' in c), None)
        c_pays = next((c for c in df_anom.columns if 'pays' in c), None)
        c_statut = next((c for c in df_anom.columns if 'statut' in c), None)

        liste_pays = ["Toutes les filiales"] + list(df_anom[c_pays].dropna().unique()) if c_pays else ["Toutes les filiales"]
        pays_selectionne = st.selectbox("🌍 Filtrer par filiale géographique :", liste_pays)
        
        if pays_selectionne != "Toutes les filiales" and c_pays:
            df_anom = df_anom[df_anom[c_pays] == pays_selectionne]

        st.markdown("### 📌 Indicateurs Financiers et Opérationnels")
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            impact_total = pd.to_numeric(df_anom[c_impact], errors='coerce').fillna(0).sum() if c_impact else 0
            st.metric("Risque Financier Global", f"{impact_total:,.0f} FCFA")
        with k2:
            nb_crit = df_anom[df_anom[c_crit].astype(str).str.contains('critique|🔴', na=False, case=False)].shape[0] if c_crit else 0
            st.metric("Alertes Critiques", nb_crit)
        with k3:
            nb_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS|OUVERT", na=False)].shape[0] if c_statut else 0
            st.metric("Anomalies non clôturées", nb_encours)
        with k4:
            st.metric("Lignes d'écarts en base", df_anom.shape[0])

        st.divider()
        st.dataframe(df_anom, hide_index=True, use_container_width=True)


# ==============================================================================
# ONGLET 3 : REQUÊTEUR TECHNIQUE DE DIAGNOSTIC
# ==============================================================================
with tabs[2]:
    st.header("🔍 Console SQL native (Accès direct Supabase)")
    st.markdown("Utilisez cet espace pour auditer en temps réel l'état des tables de votre instance.")
    
    st.subheader("🖋️ Éditeur de requêtes PostgreSQL")
    ex_query = "SELECT table_name FROM information_schema.tables WHERE table_schema='public';"
    user_sql = st.text_area("Commande SQL :", value=ex_query, height=120)
    
    if st.button("⚡ Exécuter la commande", type="primary"):
        if user_sql.strip():
            try:
                with engine.connect() as query_conn:
                    result_sql = query_conn.execute(text(user_sql))
                    if result_sql.returns_rows:
                        df_query_res = pd.DataFrame(result_sql.fetchall(), columns=result_sql.keys())
                        st.success(f"🎯 Requête validée. {df_query_res.shape[0]} lignes trouvées.")
                        st.dataframe(df_query_res, use_container_width=True)
                    else:
                        st.success("✅ Script SQL exécuté avec succès sur Supabase.")
            except Exception as sql_err:
                st.error(f"❌ Erreur renvoyée par le serveur : {sql_err}")
