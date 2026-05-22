import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime
from sqlalchemy import create_engine

# --- CONFIGURATION SÉCURISÉE DE LA PAGE ---
st.set_page_config(
    page_title="SKAB - Système Intégré de Contrôle Interne (Supabase Sécurisé)",
    page_icon="🛡️",
    layout="wide"
)

# Style CSS pour masquer les éléments de debug et moderniser l'interface
st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 28px; }
    .stAlert { margin-top: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: bold; font-size: 16px; }
    /* Masquage strict des messages d'erreur techniques SQLAlchemy bruts */
    .stException { display: none; }
    </style>
""", unsafe_allow_html=True)


# --- 1. SÉCURISATION ABSOLUE DES IDENTIFIANTS DE LA BASE DE DONNÉES ---
# AUCUNE INFORMATION EN CLAIR (Host, Password, Port, etc.) n'est présente dans ce script.
# Tout est extrait dynamiquement depuis l'espace chiffré "Secrets" de Streamlit Cloud.
try:
    # Connexion native Streamlit pour les lectures de tables avec cache automatique
    conn = st.connection("supabase", type="sql")
    
    # Extraction sécurisée des secrets pour le moteur d'écriture SQLAlchemy (df.to_sql)
    creds = st.secrets["connections"]["supabase"]
    db_url = f"postgresql://{creds['username']}:{creds['password']}@{creds['host']}:{creds['port']}/{creds['database']}"
    engine = create_engine(db_url)
except Exception as e:
    st.error("🔒 Erreur de sécurité ou de connexion : Les identifiants Supabase sont manquants ou incorrects dans vos Streamlit Secrets.")
    st.info("Veuillez configurer la section [connections.supabase] dans les paramètres Secrets de votre Dashboard Streamlit Cloud.")
    st.stop()


# --- 2. FONCTIONS DE CHARGEMENT ET DE NETTOYAGE SÉCURISÉES ---
def clean_column_name(col):
    """Nettoie drastiquement les en-têtes Excel pour les rendre 100% compatibles avec PostgreSQL"""
    return (str(col).strip()
            .replace(" ", "_")
            .replace("/", "_")
            .replace("-", "_")
            .replace("(", "")
            .replace(")", "")
            .replace("é", "e")
            .replace("è", "e")
            .replace("ê", "e")
            .replace("à", "a")
            .replace("ç", "c")
            .replace("°", "no")
            .replace("’", "_")
            .replace("'", "_")
            .lower())

def load_table_from_supabase(table_name):
    """Charge une table depuis Supabase sous forme de DataFrame Pandas avec gestion d'absence"""
    try:
        query = f"SELECT * FROM {table_name};"
        df = conn.query(query, ttl="30s") # Cache court de 30 secondes pour le temps réel
        return pd.DataFrame(df)
    except Exception:
        # Si la table n'existe pas encore (première exécution), retourne un DataFrame vide
        return pd.DataFrame()

def get_safe_len(series, col_name):
    """Calcule intelligemment la largeur des colonnes pour l'export Excel scellé du DAF"""
    clean = series.dropna()
    if not clean.empty:
        max_c = int(clean.apply(lambda x: len(str(x))).max())
    else:
        max_c = 0
    return min(max(max_c, len(str(col_name))) + 3, 50)


# --- 3. ARCHITECTURE DU WORKFLOW DU GROUPE SKAB ---
tab_chef, tab_terrain = st.tabs([
    "📊 ESPACE CHEF DE DÉPARTEMENT (Supervision, Filtres & Conso)", 
    "📥 ESPACE CONTRÔLEURS TERRAINS (Saisie Directe & Création Automatique)"
])


# ==============================================================================
# ONGLETS 1 : ESPACE CHEF DE DÉPARTEMENT (CONSULTATION MACRO)
# ==============================================================================
with tab_chef:
    st.title("🛡️ Direction du Contrôle Interne — Groupe SKAB")
    st.subheader("Suivi de l'intégrité et tableau de bord en temps réel")
    
    # Lecture en direct depuis Supabase
    df_anom = load_table_from_supabase("anomalies")
    
    if df_anom.empty:
        st.warning("💡 La table 'anomalies' n'a pas encore été initialisée dans Supabase ou est totalement vide. Allez sur l'onglet 'Espace Contrôleurs' pour injecter votre premier fichier Excel et la créer automatiquement.")
    else:
        # Standardisation des en-têtes récupérés de la base pour éviter les conflits de casse
        df_anom.columns = [str(c).lower().strip() for c in df_anom.columns]
        
        # Détection dynamique des colonnes clés nettoyées
        c_site = next((c for c in df_anom.columns if 'site' in c or 'entite' in c or 'agence' in c), df_anom.columns[2])
        c_date = next((c for c in df_anom.columns if 'date' in c), df_anom.columns[1])
        c_impact = next((c for c in df_anom.columns if 'impact' in c), None)
        c_crit = next((c for c in df_anom.columns if 'crit' in c), None)
        c_domaine = next((c for c in df_anom.columns if 'domaine' in c or 'type' in c), None)
        c_pays = next((c for c in df_anom.columns if 'pays' in c), None)
        c_statut = next((c for c in df_anom.columns if 'statut' in c), None)

        # --- FILTRES DE CONSULTATION DANS LA BARRE LATÉRALE ---
        st.sidebar.header("🎛️ Filtres de Supervision")
        
        # Filtre 1 : Sélection de l'Agence / Site
        liste_agences = ["Toutes les agences"] + list(df_anom[c_site].dropna().unique())
        agence_choisie = st.sidebar.selectbox("🏢 Filtrer par Agence / Site :", liste_agences)
        if agence_choisie != "Toutes les agences":
            df_anom = df_anom[df_anom[c_site] == agence_choisie]
            
        # Filtre 2 : Sélection Périodique (Mois, Trimestre, Année)
        df_anom['date_format_systeme'] = pd.to_datetime(df_anom[c_date], errors='coerce')
        df_anom = df_anom.dropna(subset=['date_format_systeme'])
        
        type_periode = st.sidebar.radio("Fréquence temporelle :", ["Toutes les dates", "Mensuelle", "Trimestrielle", "Annuelle"])
        
        if type_periode == "Mensuelle":
            df_anom['mois_annee'] = df_anom['date_format_systeme'].dt.to_period('M').astype(str)
            choix_mois = st.sidebar.selectbox("Sélectionner le Mois :", sorted(df_anom['mois_annee'].unique(), reverse=True))
            df_anom = df_anom[df_anom['mois_annee'] == choix_mois]
            
        elif type_periode == "Trimestrielle":
            df_anom['trim_annee'] = df_anom['date_format_systeme'].dt.to_period('Q').astype(str)
            choix_trim = st.sidebar.selectbox("Sélectionner le Trimestre :", sorted(df_anom['trim_annee'].unique(), reverse=True))
            df_anom = df_anom[df_anom['trim_annee'] == choix_trim]
            
        elif type_periode == "Annuelle":
            df_anom['annee_systeme'] = df_anom['date_format_systeme'].dt.year
            choix_annee = st.sidebar.selectbox("Sélectionner l'Année :", sorted(df_anom['annee_systeme'].unique(), reverse=True))
            df_anom = df_anom[df_anom['annee_systeme'] == choix_annee]

        # --- AFFICHAGE DES MESURES FINANCIÈRES ET ALERTES (KPI) ---
        st.markdown("### 📊 Indicateurs de Performance Métier (Données SQL)")
        k1, k2, k3, k4 = st.columns(4)
        
        with k1:
            impact_total = pd.to_numeric(df_anom[c_impact], errors='coerce').fillna(0).sum() if c_impact else 0
            st.metric("Risque Financier Cumulé", f"{impact_total:,.0f} FCFA")
        with k2:
            nb_critiques = df_anom[df_anom[c_crit].astype(str).str.contains('critique|🔴', na=False, case=False)].shape[0] if c_crit else 0
            st.metric("Anomalies Critiques", nb_critiques, delta="Attention Requise" if nb_critiques > 0 else None, delta_color="inverse")
        with k3:
            # Traitement ultra-précis demandé du statut "EN COURS"
            nb_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS", na=False)].shape[0] if c_statut else 0
            st.metric("Missions / Alertes [EN COURS]", nb_encours)
        with k4:
            st.metric("Total des Écarts en Base", df_anom.shape[0])

        st.divider()

        # --- REPRÉSENTATIONS GRAPHIQUES INTERACTIVES ---
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**🔍 Répartition des Risques par Domaine d'Activité**")
            if c_domaine:
                fig = px.bar(df_anom, x=c_domaine, color=c_crit if c_crit else None, barmode='group',
                             color_discrete_map={'🔴 critique': '#FF4B4B', '🔴 Critique': '#FF4B4B', '🟠 majeur': '#FFA500', '🟠 Majeur': '#FFA500'})
                fig.update_layout(height=320, margin=dict(l=0, r=0, t=10, b=0), xaxis_title=None, yaxis_title="Volume")
                st.plotly_chart(fig, use_container_width=True)
        with g2:
            st.markdown("**🌍 Provenance des Alertes par Filiale**")
            if c_pays:
                df_p = df_anom.groupby(c_pays).size().reset_index(name="Volume")
                fig2 = px.pie(df_p, values="Volume", names=c_pays, hole=.4, color_discrete_sequence=px.colors.qualitative.Pastel)
                fig2.update_layout(height=320, margin=dict(l=0, r=0, t=10, b=0))
                st.plotly_chart(fig2, use_container_width=True)

        # --- ZONE COMMANDEUR : VISUALISATION STRICTE DES ALERTES "EN COURS" ---
        st.divider()
        st.markdown("### ⏳ Registre Spécifique des Incidents au Statut ''EN COURS''")
        if c_statut:
            df_uniquement_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS", na=False)]
            if not df_uniquement_encours.empty:
                st.dataframe(df_uniquement_encours, hide_index=True, use_container_width=True)
            else:
                st.success("✅ Intégrité Parfaite : Aucune anomalie n'est actuellement enregistrée avec le statut 'EN COURS'.")

        # Table globale filtrée pour consultation
        st.markdown("### 📋 Données complètes de la période sélectionnée")
        st.dataframe(df_anom, hide_index=True, use_container_width=True)

        # --- CLÔTURE DU COMPILATEUR POUR LE DAF ---
        st.divider()
        st.header("📤 Extraction et Scellé du Rapport Maître")
        st.write("Générez le Fichier Excel Unique de Livraison destiné au Directeur Financier (DAF).")
        
        if st.button("🏗️ Compiler et figer les données pour M. Élie DIGNOU (DAF)", type="primary", use_container_width=True):
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                # Métadonnées de garde
                df_meta = pd.DataFrame({
                    "SYSTÈME SÉCURISÉ DU CONTRÔLE INTERNE": ["Destinataire Officiel", "Émetteur", "Horodatage d'extraction", "Périmètre Filtré", "Statut de Livraison"],
                    "MÉTADONNÉES SKAB NUTRITION": ["M. Élie DIGNOU (DAF)", "Chef de Département Contrôle Interne", datetime.now().strftime("%d/%m/%Y à %H:%M"), agence_choisie, "SCELLÉ ET SÉCURISÉ"]
                })
                df_meta.to_excel(writer, sheet_name="ACCUEIL_CONSO", index=False)
                writer.sheets["ACCUEIL_CONSO"].set_column('A:B', 38)
                
                # Table de consolidation
                df_anom.to_excel(writer, sheet_name="CONSO_ANOMALIES", index=False)
                ws = writer.sheets["CONSO_ANOMALIES"]
                for i, col in enumerate(df_anom.columns):
                    ws.set_column(i, i, get_safe_len(df_anom[col], col))
                    
            st.success("🎉 Le livrable maître a été figé sur la base de vos filtres dynamiques.")
            st.download_button(
                label="💾 Télécharger le Fichier Excel Consolidé (.xlsx)",
                data=output_buffer.getvalue(),
                file_name=f"SKAB_RAPPORT_MAITRE_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# ==============================================================================
# ONGLETS 2 : PORTAIL SÉCURISÉ DES CONTRÔLEURS TERRAINS (INJECTION DIRECTE)
# ==============================================================================
with tab_terrain:
    st.title("📥 Portail d'Injection Automatique pour Agents de Terrain")
    st.subheader("Zéro mail, zéro ressaisie — Génération automatique de la structure SQL")

    # --- MÉTHODE 1 : LECTURE ET CRÉATION DE LA TABLE VIA FICHIER EXCEL TEMPLATE ---
    st.markdown("#### 📁 Méthode 1 : Charger un fichier Excel individuel (ex: Jean-Pierre MVA)")
    file_uploaded = st.file_uploader("Glissez-déposez votre fichier de travail (.xlsx) :", type="xlsx")
    
    if file_uploaded:
        try:
            # 1. Analyse brute pour isoler la ligne d'en-tête réelle (Ligne 5 du modèle SKAB)
            df_raw_t = pd.read_excel(file_uploaded, sheet_name="ANOMALIES", header=None)
            
            header_idx = 0
            for idx, row in df_raw_t.iterrows():
                if any("ID Anomalie" in str(s) for s in row.values):
                    header_idx = idx
                    break
            
            # 2. Chargement propre à partir de la ligne détectée
            df_to_inject = pd.read_excel(file_uploaded, sheet_name="ANOMALIES", skiprows=header_idx)
            
            if df_to_inject.empty:
                st.error("⚠️ Le fichier chargé ne contient aucune ligne de données dans l'onglet 'ANOMALIES'.")
            else:
                # 3. 🛡️ SÉCURISATION ET CRÉATION AUTOMATIQUE DES COLONNES POUR POSTGRESQL
                # Cette ligne nettoie à la volée tous les noms de colonnes (minuscules, pas d'espaces, pas de caractères spéciaux)
                df_to_inject.columns = [clean_column_name(c) for c in df_to_inject.columns]
                
                # Nettoyage des lignes de consignes utilisateur du template Excel
                df_to_inject = df_to_inject.dropna(subset=[df_to_inject.columns[0]])
                df_to_inject = df_to_inject[~df_to_inject[df_to_inject.columns[0]].astype(str).str.contains("une_anomalie|id_anomalie", na=False, case=False)]
                
                # Intégration des marqueurs de traçabilité système
                df_to_inject['fichier_source'] = file_uploaded.name
                df_to_inject['date_saisie_base'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                st.write(f"📝 **Aperçu des données prêtes à structurer la base Supabase ({df_to_inject.shape[0]} lignes détectées) :**")
                st.dataframe(df_to_inject, hide_index=True)

                # --- SÉLECTION DU MODE D'INJECTION POUR GÉRER LA CRÉATION AUTOMATIQUE ---
                st.markdown("##### ⚙️ Paramétrage de l'écriture en Base SQL :")
                mode_écriture = st.radio(
                    "Définir le comportement de l'injection :",
                    [
                        "Initialiser / Recréer la table (Utile si la table n'existe pas du tout dans Supabase)",
                        "Ajouter à la suite (À utiliser pour ajouter vos données sans effacer le travail des collègues)"
                    ]
                )
                
                if_exists_param = "replace" if "Initialiser" in mode_écriture else "append"

                if st.button("🚀 Exécuter df.to_sql() et synchroniser avec Supabase", type="primary"):
                    # COMMANDE MAÎTRE DE CRÉATION ET INJECTION AUTOMATIQUE
                    df_to_inject.to_sql("anomalies", con=engine, if_exists=if_exists_param, index=False)
                    
                    st.success(f"🔥 Opération réussie ! Les colonnes ont été créées/alignées automatiquement et les {df_to_inject.shape[0]} lignes de données sont sécurisées dans Supabase.")
                    st.balloons()
                    
        except Exception as err:
            st.error("❌ Erreur lors du traitement de la structure de données.")
            st.info("Vérifiez que votre fichier possède bien l'onglet nommé 'ANOMALIES' avec la colonne 'ID Anomalie'.")

    st.divider()

    # --- MÉTHODE 2 : FORMULAIRE DE SAISIE RAPIDE DIRECT EN LIGNE ---
    st.markdown("#### 📝 Méthode 2 : Formulaire de Saisie Directe à la volée (Sans fichier)")
    with st.form("form_saisie_directe"):
        c1, c2, c3 = st.columns(3)
        with c1:
            f_id = st.text_input("ID Anomalie *", value="ANOM-2026-")
            f_pays = st.selectbox("Filiale / Pays", ["Cameroun", "Tchad", "Gabon", "RCA", "Congo"])
            f_site = st.text_input("Site / Agence / Entité *", placeholder="Ex: Moundou Agence")
        with c2:
            f_date = st.date_input("Date du constat")
            f_domaine = st.selectbox("Type / Domaine", ["AGENCES", "FERMES", "TRESORERIE", "COMPTABILITE", "ACHATS_STOCK", "PATRIMOINE"])
            f_crit = st.selectbox("Niveau de criticité", ["🟢 Faible", "🟡 Mineur", "🟠 Majeur", "🔴 Critique"])
        with c3:
            f_impact = st.number_input("Impact Financier Estimé (FCFA)", min_value=0, step=50000)
            f_resp = st.text_input("Responsable de traitement")
            f_statut = st.selectbox("Statut initial", ["Ouvert", "En cours", "Résolu"])
            
        f_desc = st.text_area("Description détaillée et factuelle des écarts *")
        f_cause = st.text_area("Cause racine identifiée")
        
        btn_submit = st.form_submit_button("💾 Enregistrer directement dans Supabase")
        
        if btn_submit:
            if len(f_id) <= 10 or not f_site or not f_desc:
                st.error("⚠️ Veuillez remplir tous les champs obligatoires marqués d'un astérisque (*).")
            else:
                # Création du DataFrame avec des noms de colonnes déjà nettoyés et standardisés pour correspondre au format SQL
                dict_form = {
                    "id_anomalie": [f_id], 
                    "date_detection": [str(f_date)], 
                    "site___entite": [f_site],
                    "pays": [f_pays], 
                    "type___domaine": [f_domaine], 
                    "niveau_criticite": [f_crit],
                    "description": [f_desc], 
                    "cause_racine_identifiee": [f_cause], 
                    "impact_estime_fcfa": [f_impact],
                    "responsable_traitement": [f_resp], 
                    "statut": [f_statut], 
                    "fichier_source": ["Formulaire Streamlit Saisie Directe"],
                    "date_saisie_base": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
                }
                df_form = pd.DataFrame(dict_form)
                
                try:
                    # Injection directe à la suite (append)
                    df_form.to_sql("anomalies", con=engine, if_exists="append", index=False)
                    st.success(f"🔥 Enregistrement validé ! L'anomalie **{f_id}** a été poussée dans votre base Supabase.")
                except Exception as form_err:
                    st.error("⚠️ Impossible d'insérer via le formulaire. Assurez-vous d'avoir initialisé la table au moins une fois avec la Méthode 1 (Import Excel).")
