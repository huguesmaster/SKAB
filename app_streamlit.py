import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
import psycopg2  # Connecteur PostgreSQL requis sous le capot par SQLAlchemy
from datetime import datetime
from sqlalchemy import create_engine, text
import urllib.parse  # Module critique pour isoler et encoder le mot de passe

# --- CONFIGURATION STRICTE DE LA PAGE ---
st.set_page_config(
    page_title="SKAB - Console Supervision Contrôle Interne",
    page_icon="🛡️",
    layout="wide"
)

# Custom CSS pour optimiser la compacité de l'interface et la lisibilité des métriques
st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 24px; font-weight: bold; }
    .stAlert { margin-top: 8px; }
    .stTabs [data-baseweb="tab-list"] { gap: 15px; }
    .stTabs [data-baseweb="tab"] { height: 42px; font-weight: bold; font-size: 14px; }
    </style>
""", unsafe_allow_html=True)


# --- 1. ROUTAGE DYNAMIQUE & SÉCURISÉ DU CORRIDOR RÉSEAU (POOLER IPv4) ---
@st.cache_resource(show_spinner=False)
def get_database_engine():
    """Génère un moteur de connexion SQLAlchemy résistant aux coupures et aux erreurs d'encodage"""
    try:
        # Configuration immuable pointant vers votre projet Supabase 'fkqlhylqsyycaiuukewf'
        DB_HOST = "aws-0-eu-central-1.pooler.supabase.com"  # Hôte universel IPv4 (évite l'erreur Cannot assign requested address)
        DB_PORT = "6543"                                    # Port transactionnel du Pooler Supabase
        DB_NAME = "postgres"
        DB_USER = "postgres.fkqlhylqsyycaiuukewf"           # Identifiant requis par le système de routage Supabase
        
        # Extraction du mot de passe brut depuis les variables d'environnement sécurisées
        DB_PASSWORD = st.secrets["connections"]["supabase"]["password"]
        
        # Sécurisation : Traduction automatique des caractères conflictuels (ex: un '@' en début de chaîne devient '%40')
        encoded_password = urllib.parse.quote_plus(DB_PASSWORD)
        
        # Assemblage de l'URI de connexion
        DATABASE_URL = f"postgresql://{DB_USER}:{encoded_password}@{DB_HOST}:{DB_PORT}/{DB_NAME}?sslmode=require"
        
        # Création du moteur d'exécution avec vérification active de l'état du canal (pool_pre_ping)
        return create_engine(
            DATABASE_URL, 
            pool_pre_ping=True, 
            pool_recycle=1800
        )
    except KeyError:
        st.error("🔒 Configuration manquante : Le paramètre 'password' n'est pas défini dans les Secrets Streamlit.")
        st.stop()
    except Exception as e:
        st.error(f"❌ Échec de l'initialisation du connecteur de base de données : {e}")
        st.stop()

# Initialisation globale du moteur
engine = get_database_engine()


# --- 2. SÉCURITÉ DDL : CONTRÔLE ET INITIALISATION AUTOMATIQUE DU SCHÉMA ---
def force_init_supabase_schema(sql_engine):
    """Vérifie l'existence des tables et force leur création physique si elles n'existent pas"""
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
            st.error(f"💥 Erreur lors de la mise en conformité de la structure SQL : {err}")
            raise err


def clean_column_name(col):
    """Formate et nettoie les en-têtes Excel pour respecter les standards relationnels PostgreSQL"""
    s = str(col).strip().lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")
    s = re.sub(r"[/\-()°’'%]", " ", s)
    s = re.sub(r"\s+", "_", s)
    return s.strip("_")

def load_table_from_supabase(table_name):
    """Charge de façon étanche le contenu d'une table PostgreSQL en DataFrame, évite le crash si vide"""
    try:
        with engine.connect() as read_conn:
            return pd.read_sql_query(text(f"SELECT * FROM {table_name};"), read_conn)
    except Exception:
        # Retourne un DataFrame vide simulé en cas d'absence de la table ou de base vierge
        return pd.DataFrame()


# --- 3. ARCHITECTURE DE L'INTERFACE UTILISATEUR ---
st.title("🛡️ SKAB NUTRITION — Console de Supervision du Contrôle Interne")
st.caption("Système d'administration cloud sécurisé — Pipeline transactionnel validé")

tabs = st.tabs([
    "📥 Injection des Livrables Excel", 
    "📊 Tableaux de Bord Analytiques", 
    "🔍 Requêteur SQL & Console Diagnostic"
])


# ==============================================================================
# ONGLET 1 : PIPELINE D'INJECTION DES CLASSEURS DE TERRAIN
# ==============================================================================
with tabs[0]:
    st.header("🗂️ Centralisation et Structuration des rapports terrains")
    st.markdown("""
        Déposez le classeur Excel consolidé issu des missions de contrôle. L'application va valider 
        la structure de vos données, appliquer la structure SQL sur **Supabase** et pousser les nouvelles lignes.
    """)
    
    src_file = st.file_uploader("Sélectionnez le rapport Excel (.xlsx) :", type="xlsx")
    
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
                "Sélectionnez la stratégie d'alimentation de la base :",
                [
                    "Ajouter les lignes à la suite de l'historique de la base de données (Append)",
                    "⚠️ Vider complètement la base en ligne avant d'importer ce fichier (Purge et Remplacement)"
                ]
            )
            
            if st.button("🚀 Lancer la synchronisation vers Supabase", type="primary", use_container_width=True):
                
                # Validation / Création des tables physiques à la volée avant injection
                force_init_supabase_schema(engine)
                
                # Traitement de la purge complète si sélectionnée
                if "Vider" in mode_import:
                    with engine.connect() as clear_conn:
                        trans = clear_conn.begin()
                        try:
                            for db_table in target_sheets.values():
                                clear_conn.execute(text(f"TRUNCATE TABLE {db_table};"))
                            trans.commit()
                            st.warning("🗑️ Base de données en ligne réinitialisée à blanc.")
                        except Exception as e:
                            trans.rollback()
                            st.error(f"Impossible de purger les tables existantes : {e}")
                            st.stop()
                
                progress_bar = st.progress(0)
                success_count = 0
                
                # Itération et parsing des onglets cibles
                for idx, (sheet_name, db_table) in enumerate(target_sheets.items()):
                    if sheet_name in available_sheets:
                        df_raw = pd.read_excel(src_file, sheet_name=sheet_name, header=None)
                        header_idx = 0
                        
                        # Recherche dynamique de la ligne d'en-tête utile
                        for r_idx, row in df_raw.iterrows():
                            row_str = " ".join([str(v) for v in row.values])
                            if any(k in row_str for k in ["ID", "N°", "Date", "Statut", "Id"]):
                                header_idx = r_idx
                                break
                        
                        df_clean = pd.read_excel(src_file, sheet_name=sheet_name, skiprows=header_idx)
                        
                        if not df_clean.empty:
                            df_clean.columns = [clean_column_name(c) for c in df_clean.columns]
                            
                            # Élimination des lignes vides ou d'exemples
                            first_col = df_clean.columns[0]
                            df_clean = df_clean.dropna(subset=[first_col])
                            df_clean = df_clean[~df_clean[first_col].astype(str).str.contains("une_anomalie|id_anomalie|exemple", na=False, case=False)]
                            
                            # Injection des métadonnées de traçabilité
                            df_clean['meta_source_file'] = src_file.name
                            df_clean['meta_import_date'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
                            
                            # alignement structurel sur les colonnes de la base de données
                            with engine.connect() as col_conn:
                                query_cols = pd.read_sql_query(text(f"SELECT * FROM {db_table} LIMIT 0;"), col_conn)
                            
                            db_cols = list(query_cols.columns)
                            for c in db_cols:
                                if c not in df_clean.columns:
                                    df_clean[c] = None
                            df_clean = df_clean[db_cols]
                            
                            # Écriture finale par lot dans PostgreSQL
                            with engine.connect() as write_conn:
                                df_clean.to_sql(db_table, con=write_conn, if_exists="append", index=False)
                            
                            st.caption(f"💾 Écriture validée pour `{db_table}` ({df_clean.shape[0]} lignes transférées).")
                            success_count += 1
                    
                    progress_bar.progress((idx + 1) / len(target_sheets))
                
                if success_count > 0:
                    st.success(f"🎉 Traitement terminé ! Vos données ont été nettoyées, converties et poussées sur Supabase.")
                    st.balloons()
                else:
                    st.error("❌ Aucun onglet correspondant aux formats attendus (ANOMALIES, MES_MISSIONS...) n'a pu être traité.")
                    
        except Exception as ex:
            st.error(f"❌ Erreur lors du traitement ou du transfert SQL : {ex}")


# ==============================================================================
# ONGLET 2 : RESTITUTION ET KPI GRAPHIQUES DU GROUPE
# ==============================================================================
with tabs[1]:
    st.header("📊 Tableau de Bord de Restitution Opérationnelle")
    
    df_anom = load_table_from_supabase("table_anomalies")
    
    if df_anom.empty:
        st.warning("💡 Aucune donnée disponible en ligne. Injectez un premier livrable pour activer le tableau de bord.")
    else:
        # Standardisation temporaire pour l'analyse des colonnes
        df_anom.columns = [c.lower() for c in df_anom.columns]
        
        c_impact = next((c for c in df_anom.columns if 'impact' in c), None)
        c_crit = next((c for c in df_anom.columns if 'crit' in c), None)
        c_pays = next((c for c in df_anom.columns if 'pays' in c), None)
        c_statut = next((c for c in df_anom.columns if 'statut' in c), None)

        # Filtre géographique intelligent
        liste_pays = ["Toutes les filiales"] + list(df_anom[c_pays].dropna().unique()) if c_pays else ["Toutes les filiales"]
        pays_selectionne = st.selectbox("🌍 Filtrer par filiale géographique :", liste_pays)
        
        if pays_selectionne != "Toutes les filiales" and c_pays:
            df_anom = df_anom[df_anom[c_pays] == pays_selectionne]

        st.markdown("### 📌 Indicateurs Clés de Risque")
        k1, k2, k3, k4 = st.columns(4)
        with k1:
            impact_total = pd.to_numeric(df_anom[c_impact], errors='coerce').fillna(0).sum() if c_impact else 0
            st.metric("Risque Financier Cumulé", f"{impact_total:,.0f} FCFA")
        with k2:
            nb_crit = df_anom[df_anom[c_crit].astype(str).str.contains('critique|🔴', na=False, case=False)].shape[0] if c_crit else 0
            st.metric("Alertes Critiques", nb_crit)
        with k3:
            nb_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS|OUVERT", na=False)].shape[0] if c_statut else 0
            st.metric("Anomalies en suspens", nb_encours)
        with k4:
            st.metric("Total Écarts Enregistrés", df_anom.shape[0])

        st.divider()
        st.subheader("📋 Registre Consolidated des Écarts")
        st.dataframe(df_anom, hide_index=True, use_container_width=True)


# ==============================================================================
# ONGLET 3 : ACCÈS DIRECT ET SÉCURISÉ POUR AUDIT SQL
# ==============================================================================
with tabs[2]:
    st.header("🔍 Console SQL native")
    st.markdown("Exécutez des ordres en lecture seule pour diagnostiquer l'état exact de votre serveur distant.")
    
    default_query = "SELECT table_name, row_security FROM information_schema.tables WHERE table_schema='public';"
    user_sql = st.text_area("Saisissez votre requête PostgreSQL :", value=default_query, height=100)
    
    if st.button("⚡ Exécuter l'analyse", type="primary"):
        if user_sql.strip():
            try:
                with engine.connect() as query_conn:
                    result_sql = query_conn.execute(text(user_sql))
                    if result_sql.returns_rows:
                        df_query_res = pd.DataFrame(result_sql.fetchall(), columns=result_sql.keys())
                        st.success(f"🎯 Requête réussie. {df_query_res.shape[0]} lignes renvoyées.")
                        st.dataframe(df_query_res, use_container_width=True)
                    else:
                        st.success("✅ Ordre SQL exécuté avec succès (aucun jeu de résultat retourné).")
            except Exception as sql_err:
                st.error(f"❌ Retour d'erreur du moteur Supabase : {sql_err}")
