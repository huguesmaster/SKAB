import streamlit as st
import pandas as pd
import plotly.express as px
import io
import re
from datetime import datetime
from sqlalchemy import create_engine, inspect

# --- CONFIGURATION SÉCURISÉE DE LA PAGE ---
st.set_page_config(
    page_title="SKAB - Système Intégré de Contrôle Interne (Supabase)",
    page_icon="🛡️",
    layout="wide"
)

st.markdown("""
    <style>
    [data-testid="stMetricValue"] { font-size: 28px; }
    .stAlert { margin-top: 10px; }
    .stTabs [data-baseweb="tab-list"] { gap: 24px; }
    .stTabs [data-baseweb="tab"] { height: 50px; white-space: pre-wrap; font-weight: bold; font-size: 16px; }
    .stException { display: none; }
    </style>
""", unsafe_allow_html=True)


# --- 1. SÉCURISATION DES IDENTIFIANTS VIA STREAMLIT SECRETS ---
try:
    conn = st.connection("supabase", type="sql")
    creds = st.secrets["connections"]["supabase"]
    db_url = f"postgresql://{creds['username']}:{creds['password']}@{creds['host']}:{creds['port']}/{creds['database']}"
    engine = create_engine(db_url)
except Exception as e:
    st.error("🔒 Erreur de configuration : Les identifiants Supabase sont manquants ou incorrects dans vos Streamlit Secrets.")
    st.stop()


# --- 2. FONCTIONS OUTILS ---
def clean_column_name(col):
    """Nettoie proprement les en-têtes pour PostgreSQL (pas d'accents, espaces ou caractères spéciaux doublés)"""
    s = str(col).strip().lower()
    s = s.replace("é", "e").replace("è", "e").replace("ê", "e").replace("à", "a").replace("ç", "c")
    s = re.sub(r"[/\-()°’']", " ", s)  # Remplace les séparateurs par des espaces
    s = re.sub(r"\s+", "_", s)         # Remplace les espaces multiples par un seul underscore
    return s.strip("_")

def load_table_from_supabase(table_name):
    """Charge une table depuis Supabase avec gestion de l'absence de table"""
    try:
        query = f"SELECT * FROM {table_name};"
        df = conn.query(query, ttl="10s")
        return pd.DataFrame(df)
    except Exception:
        return pd.DataFrame()

def force_match_supabase_columns(df_source, table_name, engine_pg):
    """🛡️ Aligne parfaitement le DataFrame sur la structure réelle de Supabase pour éviter le ValueError"""
    try:
        inspector = inspect(engine_pg)
        if table_name in inspector.get_table_names():
            # Récupère les colonnes réelles de la table Supabase
            db_columns = [col['name'] for col in inspector.get_columns(table_name)]
            
            # 1. Ajoute dans notre DataFrame les colonnes de la base qui manquent (remplies par du vide)
            for col in db_columns:
                if col not in df_source.columns:
                    df_source[col] = None
                    
            # 2. On ne garde STRICTEMENT que les colonnes qui existent dans la base de données
            df_source = df_source[db_columns]
    except Exception as e:
        pass
    return df_source

def get_safe_len(series, col_name):
    clean = series.dropna()
    if not clean.empty:
        max_c = int(clean.apply(lambda x: len(str(x))).max())
    else:
        max_c = 0
    return min(max(max_c, len(str(col_name))) + 3, 50)


# --- 3. ARCHITECTURE DES ONGLETS SKAB ---
tab_chef, tab_terrain = st.tabs([
    "📊 ESPACE CHEF DE DÉPARTEMENT (Supervision & Conso)", 
    "📥 ESPACE CONTRÔLEURS TERRAINS (Saisie Directe & Injection)"
])


# ==============================================================================
# ONGLETS 1 : INTERFACE DE SUPERVISION (CHEF DE DÉPARTEMENT)
# ==============================================================================
with tab_chef:
    st.title("🛡️ Direction du Contrôle Interne — Groupe SKAB")
    st.subheader("Suivi de l'intégrité et consolidation automatique")
    
    df_anom = load_table_from_supabase("anomalies")
    
    if df_anom.empty:
        st.warning("💡 La table 'anomalies' n'est pas encore initialisée ou est vide. Rendez-vous dans l'onglet 'Espace Contrôleurs' pour injecter un premier fichier Excel.")
    else:
        df_anom.columns = [str(c).lower().strip() for c in df_anom.columns]
        
        # Mappings dynamiques souples
        c_site = next((c for c in df_anom.columns if 'site' in c or 'entite' in c or 'agence' in c), df_anom.columns[2])
        c_date = next((c for c in df_anom.columns if 'date' in c), df_anom.columns[1])
        c_impact = next((c for c in df_anom.columns if 'impact' in c), None)
        c_crit = next((c for c in df_anom.columns if 'crit' in c), None)
        c_domaine = next((c for c in df_anom.columns if 'domaine' in c or 'type' in c), None)
        c_pays = next((c for c in df_anom.columns if 'pays' in c), None)
        c_statut = next((c for c in df_anom.columns if 'statut' in c), None)

        # Barre latérale des filtres
        st.sidebar.header("🎛️ Filtres de Supervision")
        
        liste_agences = ["Toutes les agences"] + list(df_anom[c_site].dropna().unique())
        agence_choisie = st.sidebar.selectbox("🏢 Filtrer par Agence / Site :", liste_agences)
        if agence_choisie != "Toutes les agences":
            df_anom = df_anom[df_anom[c_site] == agence_choisie]
            
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

        # KPIs
        st.markdown("### 📊 Indicateurs de Performance Métier (Données Supabase)")
        k1, k2, k3, k4 = st.columns(4)
        
        with k1:
            impact_total = pd.to_numeric(df_anom[c_impact], errors='coerce').fillna(0).sum() if c_impact else 0
            st.metric("Risque Financier Cumulé", f"{impact_total:,.0f} FCFA")
        with k2:
            nb_critiques = df_anom[df_anom[c_crit].astype(str).str.contains('critique|🔴', na=False, case=False)].shape[0] if c_crit else 0
            st.metric("Anomalies Critiques", nb_critiques, delta="Action Requise" if nb_critiques > 0 else None, delta_color="inverse")
        with k3:
            nb_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS", na=False)].shape[0] if c_statut else 0
            st.metric("Missions / Alertes [EN COURS]", nb_encours)
        with k4:
            st.metric("Total des Écarts en Base", df_anom.shape[0])

        st.divider()

        # Graphiques
        g1, g2 = st.columns(2)
        with g1:
            st.markdown("**🔍 Répartition des Risques par Domaine d'Activité**")
            if c_domaine:
                fig = px.bar(df_anom, x=c_domaine, color=c_crit if c_crit else None, barmode='group')
                fig.update_layout(height=320, margin=dict(l=0, r=0, t=10, b=0), xaxis_title=None, yaxis_title="Volume")
                st.plotly_chart(fig, use_container_width=True)
        with g2:
            st.markdown("**🌍 Provenance des Alertes par Filiale**")
            if c_pays:
                df_p = df_anom.groupby(c_pays).size().reset_index(name="Volume")
                fig2 = px.pie(df_p, values="Volume", names=c_pays, hole=.4)
                fig2.update_layout(height=320, margin=dict(l=0, r=0, t=10, b=0))
                st.plotly_chart(fig2, use_container_width=True)

        st.divider()
        st.markdown("### ⏳ Focus Exclusif sur les Incidents au Statut ''EN COURS''")
        if c_statut:
            df_uniquement_encours = df_anom[df_anom[c_statut].astype(str).str.upper().str.contains("EN COURS", na=False)]
            if not df_uniquement_encours.empty:
                st.dataframe(df_uniquement_encours, hide_index=True, use_container_width=True)
            else:
                st.success("✅ Intégrité Parfaite : Aucune anomalie n'est au statut 'EN COURS'.")

        st.markdown("### 📋 Registre Global des Données")
        st.dataframe(df_anom, hide_index=True, use_container_width=True)

        # Export DAF
        st.divider()
        st.header("📤 Finalisation et Rapport pour le DAF")
        if st.button("🏗️ Compiler le Fichier Maître Unique pour M. Élie DIGNOU (DAF)", type="primary", use_container_width=True):
            output_buffer = io.BytesIO()
            with pd.ExcelWriter(output_buffer, engine='xlsxwriter') as writer:
                df_meta = pd.DataFrame({
                    "SYSTÈME SÉCURISÉ DU CONTRÔLE INTERNE": ["Destinataire Officiel", "Émetteur", "Horodatage d'extraction", "Périmètre Filtré", "Statut de Livraison"],
                    "MÉTADONNÉES SKAB NUTRITION": ["M. Élie DIGNOU (DAF)", "Chef de Département Contrôle Interne", datetime.now().strftime("%d/%m/%Y à %H:%M"), agence_choisie, "SCELLÉ ET SÉCURISÉ"]
                })
                df_meta.to_excel(writer, sheet_name="ACCUEIL_CONSO", index=False)
                df_anom.to_excel(writer, sheet_name="CONSO_ANOMALIES", index=False)
                
            st.success("🎉 Le livrable maître a été figé sur la base de vos filtres dynamiques.")
            st.download_button(
                label="💾 Télécharger le Fichier Excel Consolidé (.xlsx)",
                data=output_buffer.getvalue(),
                file_name=f"SKAB_RAPPORT_MAITRE_{datetime.now().strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                use_container_width=True
            )


# ==============================================================================
# ONGLETS 2 : INTERFACE DE SAISIE ET TRANSFERT (CONTRÔLEURS TERRAINS)
# ==============================================================================
with tab_terrain:
    st.title("📥 Portail d'Injection Automatique pour Agents de Terrain")
    st.subheader("Zéro mail, zéro ressaisie — Synchronisation directe avec Supabase")

    # MÉTHODE 1 : INJECTION EXCEL
    st.markdown("#### 📁 Méthode 1 : Charger un fichier Excel individuel (ex: Jean-Pierre MVA)")
    file_uploaded = st.file_uploader("Glissez-déposez votre fichier de travail (.xlsx) :", type="xlsx")
    
    if file_uploaded:
        try:
            df_raw_t = pd.read_excel(file_uploaded, sheet_name="ANOMALIES", header=None)
            
            header_idx = 0
            for idx, row in df_raw_t.iterrows():
                if any("ID Anomalie" in str(s) for s in row.values):
                    header_idx = idx
                    break
            
            df_to_inject = pd.read_excel(file_uploaded, sheet_name="ANOMALIES", skiprows=header_idx)
            
            if df_to_inject.empty:
                st.error("⚠️ Le fichier chargé ne contient aucune ligne de données valides.")
            else:
                df_to_inject.columns = [clean_column_name(c) for c in df_to_inject.columns]
                df_to_inject = df_to_inject.dropna(subset=[df_to_inject.columns[0]])
                df_to_inject = df_to_inject[~df_to_inject[df_to_inject.columns[0]].astype(str).str.contains("une_anomalie|id_anomalie", na=False, case=False)]
                
                df_to_inject['fichier_source'] = file_uploaded.name
                df_to_inject['date_saisie_base'] = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

                st.markdown("##### ⚙️ Sélectionner le Mode d'écriture :")
                mode_ecriture = st.radio(
                    "Action à mener :",
                    [
                        "Ajouter à la suite (À utiliser au quotidien pour cumuler les rapports)",
                        "Initialiser / Recréer la table (⚠️ Écrase et recrée la table à neuf)"
                    ]
                )
                
                if_exists_param = "replace" if "Initialiser" in mode_ecriture else "append"

                # 🛡️ SÉCURITÉ DE CARTOGRAPHIE DES COLONNES SI MODE APPEND
                if if_exists_param == "append":
                    df_to_inject = force_match_supabase_columns(df_to_inject, "anomalies", engine)

                st.write(f"📝 **Aperçu final des données alignées pour Supabase ({df_to_inject.shape[0]} lignes) :**")
                st.dataframe(df_to_inject, hide_index=True)

                if st.button("🚀 Synchroniser le fichier avec Supabase", type="primary"):
                    df_to_inject.to_sql("anomalies", con=engine, if_exists=if_exists_param, index=False)
                    st.success(f"🔥 Opération réussie ! Les données Excel ont été poussées sans conflit structurel.")
                    st.balloons()
                    
        except Exception as err:
            st.error(f"❌ Une erreur est survenue lors de la synchronisation : {err}")

    st.divider()

    # MÉTHODE 2 : FORMULAIRE WEB DIRECT BRINDÉ
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
                raw_form_data = {
                    "ID Anomalie": [f_id], "Date détection": [str(f_date)], "Site / Entité": [f_site],
                    "Pays": [f_pays], "Type / Domaine": [f_domaine], "Niveau criticité": [f_crit],
                    "Description": [f_desc], "Cause racine identifiée": [f_cause], "Impact estimé (FCFA)": [f_impact],
                    "Responsable traitement": [f_resp], "Statut": [f_statut], "Fichier Source": ["Formulaire Web Direct"],
                    "Date Saisie Base": [datetime.now().strftime("%Y-%m-%d %H:%M:%S")]
                }
                
                df_form = pd.DataFrame(raw_form_data)
                df_form.columns = [clean_column_name(c) for c in df_form.columns]
                
                # 🛡️ Aligne aussi le formulaire sur la structure exacte de Supabase
                df_form = force_match_supabase_columns(df_form, "anomalies", engine)
                
                try:
                    df_form.to_sql("anomalies", con=engine, if_exists="append", index=False)
                    st.success(f"🔥 Enregistrement validé ! L'anomalie **{f_id}** est intégrée dans Supabase.")
                except Exception as ex:
                    st.error(f"⚠️ Échec de l'insertion : {ex}")
