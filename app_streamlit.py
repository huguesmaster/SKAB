import streamlit as st
import pandas as pd
import plotly.express as px
import io
from datetime import datetime
from sqlalchemy import create_engine

# --- CONFIGURATION DE LA PAGE ---
st.set_page_config(
    page_title="SKAB - Système Intégré de Contrôle Interne (Supabase)",
    page_icon="🛡️",
    layout="wide"
)

# Optimisation visuelle CSS
st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 28px; }
    .stAlert { margin-top: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: bold; font-size: 16px; }
    </style>
""", unsafe_allow_html=True)

# --- 1. CONNEXION SÉCURISÉE À SUPABASE ---
try:
    # Connexion native Streamlit pour les requêtes de lecture (bénéficie du cache)
    conn = st.connection("supabase", type="sql")
    
    # Récupération des secrets pour reconstruire l'engine SQLAlchemy nécessaire pour le to_sql() en écriture
    creds = st.secrets["connections"]["supabase"]
    db_url = f"postgresql://{creds['username']}:{creds['password']}@{creds['host']}:{creds['port']}/{creds['database']}"
    engine = create_engine(db_url)
except Exception as e:
    st.error(f"❌ Erreur de configuration ou de connexion à Supabase : {e}")
    st.stop()

# --- 2. FONCTIONS OUTILS ---
def load_table(table_name):
    """Charge une table Supabase sous forme de DataFrame Pandas (avec cache d'une minute)"""
    try:
        query = f"SELECT * FROM {table_name};"
        df = conn.query(query, ttl="1m")
        return pd.DataFrame(df)
    except Exception:
        # Si la table n'existe pas encore dans Supabase, retourne un DataFrame vide
        return pd.DataFrame()

def get_safe_len(series, col_name):
    """Ajuste dynamiquement la largeur des colonnes lors de l'export Excel"""
    clean = series.dropna()
    if not clean.empty:
        max_c = int(clean.apply(lambda x: len(str(x))).max())
    else:
        max_c = 0
    return min(max(max_c, len(str(col_name))) + 3, 50)


# --- 3. WORKFLOW EN ONGLETS ---
tab_chef, tab_terrain = st.tabs([
    "📊 ESPACE CHEF DE DÉPARTEMENT (Supervision & Conso)", 
    "📥 ESPACE CONTRÔLEURS TERRAINS (Saisie & Injection Directe)"
])


# ==============================================================================
# ONGLET 1 : INTERFACE CHEF DE DÉPARTEMENT (CONSULTATION EN TEMPS RÉEL)
# ==============================================================================
with tab_chef:
    st.title("🛡️ Espace Chef de Département — Groupe SKAB")
    st.subheader("Pilotage macro et consolidation automatique depuis Supabase")
    
    # Chargement en temps réel des tables depuis Supabase
    df_anom = load_table("anomalies")
    df_mis = load_table("missions")
    df_pts = load_table("points_controle")
    
    if df_anom.empty:
        st.info("💡 Aucune anomalie n'a été trouvée dans Supabase. Les contrôleurs peuvent utiliser l'onglet de saisie pour alimenter la base.")
    else:
        # Nettoyage systématique des en-têtes
        df_anom.columns = [str(c).strip() for c in df_anom.columns]
        
        # --- FILTRES LATÉRAUX DYNAMIQUES ---
        st.sidebar.header("🎛️ Paramètres de Consultation")
        
        # 1. Consultation par Agence / Site
        col_site = 'Site / Entité' if 'Site / Entité' in df_anom.columns else ('Site / Agence' if 'Site / Agence' in df_anom.columns else df_anom.columns[2])
        liste_agences = ["Toutes les agences"] + list(df_anom[col_site].dropna().unique())
        agence_choisie = st.sidebar.selectbox("🏢 Sélectionner l'Agence :", liste_agences)
        
        if agence_choisie != "Toutes les agences":
            df_anom = df_anom[df_anom[col_site] == agence_choisie]
            
        # 2. Consultation Périodique (Mois, Trimestre, Année)
        col_date = 'Date détection' if 'Date détection' in df_anom.columns else 'Date'
        if col_date in df_anom.columns:
            df_anom['Date_Format'] = pd.to_datetime(df_anom[col_date], errors='coerce')
            df_anom = df_anom.dropna(subset=['Date_Format'])
            
            st.sidebar.subheader("📅 Période")
            type_periode = st.sidebar.radio("Fréquence :", ["Toutes", "Mensuelle", "Trimestrielle", "Annuelle"])
            
            if type_periode == "Mensuelle":
                df_anom['Mois_Annee'] = df_anom['Date_Format'].dt.to_period('M').astype(str)
                choix_mois = st.sidebar.selectbox("Choisir le Mois :", sorted(df_anom['Mois_Annee'].unique(), reverse=True))
                df_anom = df_anom[df_anom['Mois_Annee'] == choix_mois]
                
            elif type_periode == "Trimestrielle":
                df_anom['Trim_Annee'] = df_anom['Date_Format'].dt.to_period('Q').astype(str)
                choix_trim = st.sidebar.selectbox("Choisir le Trimestre :", sorted(df_anom['Trim_Annee'].unique(), reverse=True))
                df_anom = df_anom[df_anom['Trim_Annee'] == choix_trim]
                
            elif type_periode == "Annuelle":
                df_anom['Annee'] = df_anom['Date_Format'].dt.year
                choix_annee = st.sidebar.selectbox("Choisir l'Année :", sorted(df_anom['Annee'].unique(), reverse=True))
                df_anom = df_anom[df_anom['Annee'] == choix_annee]

        # Mappings des colonnes cibles
        col_impact = next((c for c in df_anom.columns if 'Impact' in c), None)
        col_crit = next((c for c in df_anom.columns if 'critic' in c or 'Critic' in c), None)
        col_domaine = next((c for c in df_anom.columns if 'Domaine' in c or 'Type' in c), None)
        col_pays = next((c for c in df_anom.columns if 'Pays' in c), None)
        col_statut = next((c for c in df_anom.columns if 'Statut' in c), None)

        # --- AFFICHAGE DES KPI ---
        st.markdown("### 📊 Indicateurs clés issus de Supabase")
        k1, k2, k3, k4 = st.columns(4)
        
        with k1:
            impact_total = pd.to_numeric(df_anom[col_impact], errors='coerce').fillna(0).sum() if col_impact else 0
            st.metric("Risque Financier Global", f"{impact_total:,.0f} FCFA")
        with k2:
            nb_critiques = df_anom[df_anom[col_crit].astype(str).str.contains('Critique|🔴', na=False)].shape[0] if col_crit else 0
            st.metric("Anomalies Critiques", nb_critiques, delta="Alerte Rouge" if nb_critiques > 0 else None, delta_color="inverse")
        with k3:
            # Filtre strict demandé : Détection du statut "EN COURS"
            nb_encours = df_anom[df_anom[col_statut].astype(str).str.upper().str.contains("EN COURS", na=False)].shape[0] if col_statut else 0
            st.metric("Incidents [EN COURS]", nb_encours)
        with k4:
            st.metric("Lignes d'écarts extraites", df_anom.shape[0])

        st.divider()

        # --- GRAPHES PLOTLY ---
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**🔍 Volume par Domaine de contrôle**")
            if col_domaine:
                fig = px.bar(df_anom, x=col_domaine, color=col_crit if col_crit else None, barmode='group',
                             color_discrete_map={'🔴 Critique': '#FF4B4B', '🟠 Majeur': '#FFA500', '🟡 Mineur': '#FFD700', '🟢 Faible': '#2ECC71'})
                fig.update_layout(height=320, margin=dict(l=0, r=0, t=10, b=0), xaxis_title=None, yaxis_title="Quantité")
                st.plotly_chart(fig, use_container_width=True)
        with g2:
            st.markdown("**🌍 Origine Géographique des alertes**")
            if col_pays:
                df_p = df_anom.groupby(col_pays).size().reset_index(name="Total")
                fig2 = px.pie(df_p, values="Total", names=col_pays, hole=.4, color_discrete_sequence=px.colors.qualitative.Safe)
                fig2.update_layout(height=320, margin=dict(l=0, r=0, t=10, b=0))
                st.plotly_chart(fig2, use_container_width=True)

        # --- AFFICHAGE EXCLUSIF DES ANOMALIES "EN COURS" ---
        st.divider()
        st.markdown("### ⏳ Liste des Anomalies avec le Statut ''EN COURS''")
        if col_statut:
            df_uniquement_encours = df_anom[df_anom[col_statut].astype(str).str.upper().str.contains("EN COURS", na=False)]
            if not df_uniquement_encours.empty:
                cols_to_print = [c for c in ['ID Anomalie', col_crit, col_site, 'Description', col_impact, 'Responsable traitement'] if c in df_uniquement_encours.columns]
                st.dataframe(df_uniquement_encours[cols_to_print], hide_index=True, use_container_width=True)
            else:
                st.success("✅ Excellente nouvelle : Aucune anomalie n'est en attente au statut 'EN COURS'.")

        # --- REGISTRE DE CONSULTATION COMPLET ---
        st.markdown("### 📋 Table complète des anomalies (Données filtrées)")
        st.dataframe(df_anom, hide_index=True, use_container_width=True)

        # --- CLÔTURE DE SESSION & COMPILATION ---
        st.divider()
        st.header("📤 Finalisation du Fichier Maître")
        if st.button("🏗️ Exporter le Fichier Maître Scellé pour le DAF (Élie DIGNOU)", type="primary", use_container_width=True):
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                # Page d'accueil officielle
                df_meta = pd.DataFrame({
                    "SUIVI CONSOLIDÉ DU CONTRÔLE INTERNE": ["Bénéficiaire", "Auteur", "Généré le", "Périmètre d'agence", "Source"],
                    "LIVRABLE EXCLUSIF": ["M. Élie DIGNOU (DAF)", "Chef de Département CI", datetime.now().strftime("%d/%m/%Y à %H:%M"), agence_choisie, "Centralisation Supabase"]
                })
                df_meta.to_excel(writer, sheet_name="ACCUEIL_CONSO", index=False)
                writer.sheets["ACCUEIL_CONSO"].set_column('A:B', 38)
                
                # Onglet des anomalies filtrées
                df_anom.to_excel(writer, sheet_name="CONSO_ANOMALIES", index=False)
                ws = writer.sheets["CONSO_ANOMALIES"]
                for i, col in enumerate(df_anom.columns):
                    ws.set_column(i, i, get_safe_len(df_anom[col], col))
                    
            st.success("🎉 Compilation Excel effectuée avec succès depuis les données Supabase !")
            st.download_button(
                label="💾 Télécharger le livrable Excel Officiel",
                data=output_buffer.getvalue(),
                file_name=f"SKAB_MAITRE_SUPABASE_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# ==============================================================================
# ONGLET 2 : INTERFACE CONTRÔLEURS TERRAINS (INJECTION DIRECTE DANS SUPABASE)
# ==============================================================================
with tab_terrain:
    st.title("📥 Portail d'injection Terrain — Base de Données")
    st.subheader("Outil de transmission directe pour les contrôleurs (Zéro mail)")

    # MÉTHODE A : IMPORTATION DE LEUR TEMPLATE EXCEL HABITUEL
    st.markdown("#### 📁 Méthode 1 : Téléverser un fichier Excel Template SKAB")
    file_uploaded = st.file_uploader("Sélectionnez votre fichier Excel de contrôle :", type="xlsx")
    
    if file_uploaded:
        try:
            df_raw_t = pd.read_excel(file_uploaded, sheet_name="ANOMALIES", header=None)
            
            # Localisation automatique de la ligne 5 d'en-tête du template SKAB
            header_idx = 0
            for idx, row in df_raw_t.iterrows():
                if any("ID Anomalie" in str(s) for s in row.values):
                    header_idx = idx
                    break
            
            df_to_inject = pd.read_excel(file_uploaded, sheet_name="ANOMALIES", skiprows=header_idx)
            df_to_inject.columns = [str(c).strip() for c in df_to_inject.columns]
            df_to_inject = df_to_inject.dropna(subset=[df_to_inject.columns[0]]) # Nettoie les lignes vides
            
            # Suppression des lignes d'explications du template
            df_to_inject = df_to_inject[~df_to_inject[df_to_inject.columns[0]].astype(str).str.contains("Une anomalie", na=False)]
            
            # Marquage des métadonnées système avant transfert
            df_to_inject['Fichier Source'] = file_uploaded.name
            df_to_inject['Date Saisie Base'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

            st.write(f"📝 **Aperçu des {df_to_inject.shape[0]} lignes détectées prêtes pour Supabase :**")
            st.dataframe(df_to_inject, hide_index=True)

            if st.button("🚀 Valider et injecter dans la table Supabase", type="primary"):
                # COMMANDE CRITIQUE : Écriture directe dans la table PostgreSQL 'anomalies'
                df_to_inject.to_sql("anomalies", con=engine, if_exists="append", index=False)
                
                st.success(f"✅ Transmission réussie ! Les anomalies du fichier '{file_uploaded.name}' sont sauvegardées de façon permanente dans Supabase.")
                st.balloons()
        except Exception as err:
            st.error(f"❌ Erreur lors de l'analyse ou du transfert SQL : {err}")

    st.divider()

    # MÉTHODE B : FORMULAIRE WEB DE SAISIE MANUELLE RAPIDE
    st.markdown("#### 📝 Méthode 2 : Formulaire de saisie à la volée (Sans fichier)")
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
            f_impact = st.number_input("Impact Financier Estimé (FCFA)", min_value=0, step=100000)
            f_resp = st.text_input("Responsable de la remédiation")
            f_statut = st.selectbox("Statut initial", ["Ouvert", "En cours", "Résolu"])
            
        f_desc = st.text_area("Description détaillée et factuelle des écarts *")
        f_cause = st.text_area("Cause racine identifiée")
        
        btn_submit = st.form_submit_button("💾 Enregistrer immédiatement dans Supabase")
        
        if btn_submit:
            if len(f_id) <= 10 or not f_site or not f_desc:
                st.error("⚠️ Les champs obligatoires (*) doivent être correctement renseignés.")
            else:
                # Structuration de la ligne unique sous forme de DataFrame
                dict_form = {
                    "ID Anomalie": [f_id], "Date détection": [str(f_date)], "Site / Entité": [f_site],
                    "Pays": [f_pays], "Type / Domaine": [f_domaine], "Niveau criticité": [f_crit],
                    "Description": [f_desc], "Cause racine identifiée": [f_cause], "Impact estimé (FCFA)": [f_impact],
                    "Responsable traitement": [f_resp], "Statut": [f_statut], "Fichier Source": ["Formulaire Streamlit"],
                    "Date Saisie Base": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
                }
                df_form = pd.DataFrame(dict_form)
                
                # Injection de la ligne de formulaire dans Supabase
                df_form.to_sql("anomalies", con=engine, if_exists="append", index=False)
                st.success(f"🔥 Succès complet ! L'anomalie **{f_id}** est intégrée dans le Cloud Supabase.")
