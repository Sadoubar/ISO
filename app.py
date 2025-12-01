import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from plotly.subplots import make_subplots
import numpy as np
from datetime import datetime, timedelta
import warnings
import re

warnings.filterwarnings('ignore')

# =========================
# CONSTANTES RSE
# =========================
FACTEUR_CUMAC_TO_KWH = {
    'BAR-TH': 1 / 12.16,
    'BAR-EN': 1 / 17.29,
    'BAR-EQ': 1 / 11.12,
    'BAT-TH': 1 / 12.16,
    'AGRI-TH': 1 / 12.16,
    'BAT-EN': 1 / 17.29,
    'TRA': 1 / 0.9615,
    'IND': 1 / 8.11,
    'DEFAULT': 1 / 8.11
}

DUREE_VIE_EQUIPEMENT = {
    'BAR-TH': 17,
    'AGRI-TH': 17,
    'BAR-EN': 30,
    'BAR-EQ': 15,
    'BAT-TH': 17,
    'BAT-EN': 30,
    'TRA': 1,
    'IND': 10,
    'DEFAULT': 10
}

EMISSION_CO2_KWH = 0.057  # kg CO2 par kWh
CO2_PAR_ARBRE_AN = 25  # kg CO2 absorbé par arbre par an
CO2_PAR_VOITURE_AN = 2800  # kg CO2 par voiture par an
CONSO_MOYENNE_FOYER_KWH = 15312  # kWh/an (chauffage + élec)

# Configuration de la page
st.set_page_config(
    page_title="Dashboard CEE - Analyse Simplifiée",
    page_icon="⚡",
    layout="wide",
    initial_sidebar_state="expanded"
)

# Style CSS amélioré
st.markdown("""
    <style>
    .stMetric { 
        background-color: #ffffff; 
        padding: 10px; 
        border-radius: 8px; 
        border: 1px solid #f0f2f6;
        box-shadow: 0 1px 3px rgba(0,0,0,0.05); 
    }
    .alert-danger {
        background-color: #f8d7da;
        border: 1px solid #f5c6cb;
        padding: 10px;
        border-radius: 5px;
        color: #721c24;
    }
    .alert-warning {
        background-color: #fff3cd;
        border: 1px solid #ffeeba;
        padding: 10px;
        border-radius: 5px;
        color: #856404;
    }
    .alert-success {
        background-color: #d4edda;
        border: 1px solid #c3e6cb;
        padding: 10px;
        border-radius: 5px;
        color: #155724;
    }
    div[data-testid="stMetric"] {
        background-color: #ffffff;
        border: 1px solid #f0f2f6;
        padding: 10px;
        border-radius: 10px;
        box-shadow: 0 2px 4px rgba(0,0,0,0.05);
    }
    </style>
""", unsafe_allow_html=True)

# Titre principal
st.title("⚡ Dashboard CEE - Analyse & Performance")
st.markdown("---")


@st.cache_data
def load_and_process_data(uploaded_file):
    """Charge et traite les données CEE avec gestion complète des erreurs"""
    try:
        # Lire toutes les feuilles du fichier Excel
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        sheet_names = xls.sheet_names

        # Identifier la feuille principale (Data) et la feuille Synthese
        sheet_synthese = None
        sheet_data = None

        # Recherche intelligente des feuilles
        for name in sheet_names:
            if 'synthese' in name.lower() or 'synthèse' in name.lower():
                sheet_synthese = name
            elif 'data' in name.lower() or 'données' in name.lower() or 'export' in name.lower():
                sheet_data = name

        # Si pas trouvé explicitement, on prend la 1ère comme Data
        if sheet_data is None:
            sheet_data = sheet_names[0]

        # === TRAITEMENT DONNÉES PRINCIPALES (DATA) ===
        df = pd.read_excel(uploaded_file, sheet_name=sheet_data, engine='openpyxl')
        df.columns = df.columns.str.strip()

        # Conversion des dates
        date_columns = ['Date Validation', 'Date depot', 'Date de début', 'Date de fin',
                        'Date de la facture', 'Date Insertion']
        for col in date_columns:
            if col in df.columns:
                df[col] = pd.to_datetime(df[col], errors='coerce')

        # Calcul de l'année de dépôt
        if 'Date depot' in df.columns:
            df['Année_depot'] = df['Date depot'].dt.year
            df['Mois_depot'] = df['Date depot'].dt.month
            df['Date_depot_str'] = df['Date depot'].dt.strftime('%Y-%m')

        # Traitement colonne Erreur de saisi (si existe)
        if 'Erreur de saisi' in df.columns:
            df['Erreur de saisi'] = pd.to_numeric(df['Erreur de saisi'], errors='coerce').fillna(0)

        # === CRÉATION CLÉ DE JOINTURE (DEPOT CLEAN) ===
        # On essaie de trouver ou créer une colonne "Depot_Join" propre (ex: P5-14) sans suffixe (AK/GP)
        if 'Depot' in df.columns:
            df['Depot_Join'] = df['Depot'].astype(str).str.strip()
        elif 'N° DEPOT' in df.columns:
            # Extraction regex pour récupérer Pxx-xx et ignorer les suffixes AK/GP
            def clean_depot_number(val):
                if pd.isna(val): return ''
                s = str(val).strip()
                # Cherche un motif P + chiffres + tiret + chiffres (ex: P5-14)
                match = re.match(r'(P\d+-\d+)', s, re.IGNORECASE)
                if match:
                    return match.group(1).upper()
                return s  # Fallback si pas de match

            df['Depot_Join'] = df['N° DEPOT'].apply(clean_depot_number)
        else:
            # Fallback vide
            df['Depot_Join'] = ''

        # Gestion spéciale pour les équipements TRA et nettoyage des codes postaux
        if 'Code équipement' in df.columns and 'code postal' in df.columns and 'Ville' in df.columns:

            def extract_postal_code(postal_value):
                """Extrait le code postal d'une valeur mixte"""
                try:
                    if pd.isna(postal_value) or postal_value == '':
                        return ''

                    postal_str = str(postal_value).strip()

                    postal_match = re.search(r'\b(\d{5})\b', postal_str)
                    if postal_match:
                        return postal_match.group(1)

                    try:
                        return str(int(float(postal_str))).zfill(5)
                    except:
                        return ''
                except:
                    return ''

            df['is_TRA'] = df['Code équipement'].astype(str).str.startswith('TRA')
            df['ville_geo'] = df.apply(
                lambda row: str(row['Adresse des travaux']) if row['is_TRA']
                else str(row['Ville']), axis=1
            )

            df['code_postal_source'] = df.apply(
                lambda row: str(row['Ville']) if row['is_TRA']
                else str(row['code postal']), axis=1
            )

            df['code_postal_clean'] = df['code_postal_source'].apply(extract_postal_code)

            df['Département'] = df['code_postal_clean'].apply(
                lambda x: x[:2] if len(x) == 5 and x.isdigit() else ''
            )

            df['geo_valid'] = (df['Département'] != '') & (df['Département'].str.len() == 2) & (
                df['Département'].str.isnumeric())

        # Statut validation
        df['Statut'] = df['Date Validation'].apply(lambda x: 'Validé' if pd.notna(x) else 'En cours')

        # Calcul des délais
        if 'Date de début' in df.columns and 'Date de fin' in df.columns:
            df['Délai_travaux'] = (df['Date de fin'] - df['Date de début']).dt.days
            df['Délai_début_fin'] = df['Délai_travaux']  # Alias

        if 'Date depot' in df.columns and 'Date Validation' in df.columns:
            df['Délai_validation'] = (df['Date Validation'] - df['Date depot']).dt.days
            df['Délai_depot_validation'] = df['Délai_validation']  # Alias

        if 'Date de début' in df.columns and 'Date depot' in df.columns:
            df['Délai_devis_depot'] = (df['Date depot'] - df['Date de début']).dt.days

        if 'Date de fin' in df.columns and 'Date depot' in df.columns:
            df['Délai_fin_depot'] = (df['Date depot'] - df['Date de fin']).dt.days

        if 'Date de fin' in df.columns and 'Date Validation' in df.columns:
            df['Délai_fin_validation'] = (df['Date Validation'] - df['Date de fin']).dt.days

        if 'Date de début' in df.columns and 'Date Validation' in df.columns:
            df['Délai_début_validation'] = (df['Date Validation'] - df['Date de début']).dt.days

        # Gestion des volumes
        volume_columns = ['Total précarité', 'Total classique', 'Tableau Recapitulatif champ 23']
        for col in volume_columns:
            if col in df.columns:
                df[col] = pd.to_numeric(df[col], errors='coerce').fillna(0)

        if 'Total précarité' in df.columns:
            df['Total_précarité_MWh'] = df['Total précarité'] / 1000
        if 'Total classique' in df.columns:
            df['Total_classique_MWh'] = df['Total classique'] / 1000

        if 'Total précarité' in df.columns and 'Total classique' in df.columns:
            df['Volume_total'] = df['Total précarité'] + df['Total classique']
            df['Volume_total_MWh'] = df['Volume_total'] / 1000

        # === CALCULS RSE ===
        if 'Code équipement' in df.columns:
            code = df['Code équipement'].astype(str).str.strip().str.upper()
            df['CodeEquip_prefix'] = code.str.split('-').str[0]
            df['CodeEquip_sub'] = code.str.split('-').str[1].fillna('')

            df['FacteurKey'] = np.where(
                df['CodeEquip_sub'].isin(['TH', 'EN', 'EQ']),
                df['CodeEquip_prefix'] + '-' + df['CodeEquip_sub'],
                df['CodeEquip_prefix']
            )

            df['Facteur_Conversion'] = df['FacteurKey'].map(FACTEUR_CUMAC_TO_KWH).fillna(
                FACTEUR_CUMAC_TO_KWH['DEFAULT'])

            df['Duree_Vie'] = df['FacteurKey'].map(DUREE_VIE_EQUIPEMENT).fillna(
                DUREE_VIE_EQUIPEMENT['DEFAULT'])

            df['Secteur'] = df['CodeEquip_prefix'].map({
                'BAR': 'Résidentiel',
                'BAT': 'Tertiaire',
                'TRA': 'Transport',
                'AGRI': 'Agriculture',
                'IND': 'Industrie'
            }).fillna('Autre')

        # Nettoyage des types
        for col in df.columns:
            if df[col].dtype == 'object':
                df[col] = df[col].astype(str)
                df[col] = df[col].replace(['nan', 'None', 'NaT'], '')

        # === TRAITEMENT FEUILLE SYNTHÈSE ===
        df_synthese = None
        if sheet_synthese:
            df_synthese = pd.read_excel(uploaded_file, sheet_name=sheet_synthese, engine='openpyxl')
            df_synthese.columns = df_synthese.columns.str.strip()

            # Normalisation des noms de colonnes (DEPOT / DÉPÔT)
            rename_map = {}
            for col in df_synthese.columns:
                col_norm = col.strip().lower()
                if 'date' in col_norm and ('depot' in col_norm or 'dépôt' in col_norm):
                    rename_map[col] = 'Date Depot'
                elif (
                        'depot' in col_norm or 'dépôt' in col_norm) and 'date' not in col_norm and 'n°' not in col_norm and 'no' not in col_norm:
                    rename_map[col] = 'Depot'

            if rename_map:
                df_synthese.rename(columns=rename_map, inplace=True)

            # Sécurité : Si "Depot" n'a pas été trouvé
            if 'Depot' not in df_synthese.columns:
                candidates = [c for c in df_synthese.columns if 'depot' in c.lower() and 'n°' not in c.lower()]
                if candidates:
                    df_synthese.rename(columns={candidates[0]: 'Depot'}, inplace=True)

            # Renommage Type de rejet vers Type de Retrait
            for col in df_synthese.columns:
                if 'type' in col.lower() and ('rejet' in col.lower() or 'retrait' in col.lower()):
                    df_synthese.rename(columns={col: 'Type de Retrait'}, inplace=True)
                    break

            # Nettoyage et conversion des Volumes
            numeric_cols = ['Volume demande', 'Volume delivre', 'Volume retire',
                            'Nombre d\'opérations validées', 'Nombre demande operation']
            for col in numeric_cols:
                if col in df_synthese.columns:
                    df_synthese[col] = pd.to_numeric(
                        df_synthese[col].astype(str).str.replace(r'\s+', '', regex=True).str.replace(',', '.'),
                        errors='coerce'
                    ).fillna(0)

            # Nettoyage des Pourcentages
            pct_cols = ['% Delivrance', 'taux d\'acceptation']
            for col in pct_cols:
                if col in df_synthese.columns:
                    df_synthese[f'{col}_Num'] = pd.to_numeric(
                        df_synthese[col].astype(str).str.replace('%', '').str.replace(',', '.'),
                        errors='coerce'
                    ) / 100

            # Conversion Date Depot
            if 'Date Depot' in df_synthese.columns:
                df_synthese['Date Depot'] = pd.to_datetime(df_synthese['Date Depot'], errors='coerce', dayfirst=True)
                df_synthese['Jours_Instruction'] = (datetime.now() - df_synthese['Date Depot']).dt.days

            # Catégorisation Statut
            if 'DECISION DE DELIVRANCE' in df_synthese.columns:
                df_synthese['Statut_Synthese'] = df_synthese['DECISION DE DELIVRANCE'].apply(
                    lambda x: 'Validé' if (pd.notna(x) and str(x).strip() != '' and str(x) != '0') else 'En instruction'
                )

        return df, df_synthese

    except Exception as e:
        st.error(f"❌ Erreur lors du chargement des données : {str(e)}")
        return None, None


def calculate_rse_metrics(df, taux_efficacite=0.45):
    """Calcule les métriques RSE avec validation"""
    if 'Volume_total_MWh' not in df.columns:
        return {}

    # GWh cumac
    gwhc_total = df['Volume_total_MWh'].sum() / 1000

    # GWh réels annuels
    if 'Facteur_Conversion' in df.columns and 'Duree_Vie' in df.columns:
        df['kWh_reels_annuels'] = (df['Volume_total'] * df['Facteur_Conversion'] / df['Duree_Vie']) * taux_efficacite
        df['GWh_reels_annuels'] = df['kWh_reels_annuels'] / 1_000_000
        gwh_reels = df['GWh_reels_annuels'].sum()
    else:
        gwh_reels = gwhc_total * 0.082 * taux_efficacite

    # CO2 évité (tonnes/an)
    co2_evite = gwh_reels * 1_000_000 * EMISSION_CO2_KWH / 1000

    # Équivalences
    arbres_equivalent = co2_evite * 1000 / CO2_PAR_ARBRE_AN
    voitures_equivalent = co2_evite * 1000 / CO2_PAR_VOITURE_AN
    foyers_equivalent = gwh_reels * 1_000_000 / CONSO_MOYENNE_FOYER_KWH

    # Précarité
    if 'Total_précarité_MWh' in df.columns and 'Volume_total_MWh' in df.columns:
        volume_precarite = df['Total_précarité_MWh'].sum()
        volume_total = df['Volume_total_MWh'].sum()
        pct_precarite = (volume_precarite / volume_total * 100) if volume_total > 0 else 0
    else:
        pct_precarite = 0

    return {
        'gwhc_total': gwhc_total,
        'gwh_reels': gwh_reels,
        'co2_evite': co2_evite,
        'arbres_equivalent': arbres_equivalent,
        'voitures_equivalent': voitures_equivalent,
        'foyers_equivalent': foyers_equivalent,
        'pct_precarite': pct_precarite
    }


def create_kpi_cards(df, sla_days=60):
    """Crée les cartes KPI principales améliorées"""
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        total_volume_mwh = df['Volume_total_MWh'].sum() if 'Volume_total_MWh' in df.columns else 0
        st.metric(
            label="📊 Volume Total (MWh cumac)",
            value=f"{total_volume_mwh:,.1f}".replace(',', ' ')
        )

    with col2:
        total_dossiers = len(df)
        dossiers_valides = len(df[df['Statut'] == 'Validé'])
        taux_validation = (dossiers_valides / total_dossiers * 100) if total_dossiers > 0 else 0
        st.metric(
            label="✅ Taux de Validation",
            value=f"{taux_validation:.1f}%",
            delta=f"{dossiers_valides:,}/{total_dossiers:,}".replace(',', ' ')
        )

    with col3:
        if 'Délai_validation' in df.columns:
            delai_moyen = df['Délai_validation'].mean()
            st.metric(
                label="⏱️ Délai Validation Moyen",
                value=f"{delai_moyen:.0f} jours" if pd.notna(delai_moyen) else "N/A"
            )
        else:
            st.metric(label="⏱️ Délai Validation Moyen", value="N/A")

    with col4:
        montant_total = df[
            'Tableau Recapitulatif champ 23'].sum() if 'Tableau Recapitulatif champ 23' in df.columns else 0
        st.metric(
            label="💰 Montant Primes Total",
            value=f"{montant_total:,.0f} €".replace(',', ' ')
        )

    # Deuxième ligne de KPIs
    col5, col6, col7, col8 = st.columns(4)

    with col5:
        taux_conversion = (dossiers_valides / total_dossiers * 100) if total_dossiers > 0 else 0
        st.metric(
            label="🔄 Taux de Conversion",
            value=f"{taux_conversion:.1f}%",
            help="Dossiers validés / Total dossiers"
        )

    with col6:
        if total_dossiers > 0:
            volume_moyen = total_volume_mwh / total_dossiers
            st.metric(
                label="📈 Volume Moyen/Dossier",
                value=f"{volume_moyen:.2f} MWh"
            )
        else:
            st.metric(label="📈 Volume Moyen/Dossier", value="N/A")

    with col7:
        if 'Délai_validation' in df.columns:
            dossiers_en_cours = df[df['Statut'] == 'En cours']
            if 'Date depot' in df.columns and len(dossiers_en_cours) > 0:
                dossiers_en_cours_copy = dossiers_en_cours.copy()
                dossiers_en_cours_copy['Jours_depuis_depot'] = (
                        datetime.now() - dossiers_en_cours_copy['Date depot']).dt.days
                dossiers_bloques = len(dossiers_en_cours_copy[dossiers_en_cours_copy['Jours_depuis_depot'] > 90])
                pct_bloques = (dossiers_bloques / len(dossiers_en_cours) * 100) if len(dossiers_en_cours) > 0 else 0
                st.metric(
                    label="🚨 Dossiers Instruction (>90j)",
                    value=f"{dossiers_bloques}",
                    delta=f"{pct_bloques:.1f}% des en cours",
                    delta_color="inverse"
                )
            else:
                st.metric(label="🚨 Dossiers Instruction (>90j)", value="N/A")
        else:
            st.metric(label="🚨 Dossiers Instruction (>90j)", value="N/A")

    with col8:
        if 'Délai_validation' in df.columns:
            df_valides = df[df['Statut'] == 'Validé']
            if len(df_valides) > 0:
                dans_sla = len(df_valides[df_valides['Délai_validation'] <= sla_days])
                taux_sla = (dans_sla / len(df_valides) * 100)
                st.metric(
                    label=f"⚡ Taux SLA (<{sla_days}j)",
                    value=f"{taux_sla:.1f}%",
                    delta=f"{dans_sla}/{len(df_valides)}"
                )
            else:
                st.metric(label=f"⚡ Taux SLA (<{sla_days}j)", value="N/A")
        else:
            st.metric(label=f"⚡ Taux SLA (<{sla_days}j)", value="N/A")


def create_filters(df):
    """Crée la barre latérale avec les filtres"""
    st.sidebar.header("🎛️ Filtres")

    filters = {}

    # Filtre année de dépôt
    if 'Année_depot' in df.columns:
        annees_disponibles = sorted(df['Année_depot'].dropna().unique())
        if len(annees_disponibles) > 0:
            filters['annees'] = st.sidebar.multiselect(
                "Année de dépôt",
                options=annees_disponibles,
                default=annees_disponibles
            )

    # Filtre statut
    statuts_disponibles = df['Statut'].unique()
    filters['statuts'] = st.sidebar.multiselect(
        "Statut",
        options=statuts_disponibles,
        default=statuts_disponibles
    )

    # Filtre mandataire
    if 'Mandataire' in df.columns:
        mandataires_disponibles = df['Mandataire'].unique()
        filters['mandataires'] = st.sidebar.multiselect(
            "Mandataire",
            options=mandataires_disponibles,
            default=mandataires_disponibles
        )

    # Filtre type d'équipement
    if 'Code équipement' in df.columns:
        equipements_disponibles = sorted(df['Code équipement'].unique())
        filters['equipements'] = st.sidebar.multiselect(
            "Code équipement",
            options=equipements_disponibles,
            default=equipements_disponibles
        )

    # Filtre département
    if 'Département' in df.columns:
        departements_disponibles = sorted(df[df['Département'] != '']['Département'].unique())
        if len(departements_disponibles) > 0:
            filters['departements'] = st.sidebar.multiselect(
                "Département",
                options=departements_disponibles,
                default=departements_disponibles
            )

    return filters


def apply_filters(df, filters):
    """Applique les filtres au dataframe"""
    filtered_df = df.copy()

    if 'annees' in filters and len(filters['annees']) > 0:
        filtered_df = filtered_df[filtered_df['Année_depot'].isin(filters['annees'])]

    if 'statuts' in filters and len(filters['statuts']) > 0:
        filtered_df = filtered_df[filtered_df['Statut'].isin(filters['statuts'])]

    if 'mandataires' in filters and len(filters['mandataires']) > 0:
        filtered_df = filtered_df[filtered_df['Mandataire'].isin(filters['mandataires'])]

    if 'equipements' in filters and len(filters['equipements']) > 0:
        filtered_df = filtered_df[filtered_df['Code équipement'].isin(filters['equipements'])]

    if 'departements' in filters and len(filters['departements']) > 0:
        filtered_df = filtered_df[filtered_df['Département'].isin(filters['departements'])]

    return filtered_df


def create_volume_evolution_chart(df):
    """Crée l'analyse d'évolution des volumes et dépôts basée sur la Date de Dépôt"""
    st.header("📈 Évolution des Volumes et Dépôts")

    if 'Date depot' not in df.columns:
        st.warning("⚠️ La colonne 'Date depot' est requise pour cette analyse temporelle.")
        return

    if not pd.api.types.is_datetime64_any_dtype(df['Date depot']):
        try:
            df['Date depot'] = pd.to_datetime(df['Date depot'], errors='coerce')
        except Exception:
            st.warning("⚠️ La colonne 'Date depot' ne contient pas des dates valides.")
            return

    df_evol = df.dropna(subset=['Date depot']).copy()

    if len(df_evol) == 0:
        st.warning("Aucune donnée avec une date de dépôt valide.")
        return

    # Choix de la granularité
    granularite = st.selectbox(
        "Granularité temporelle",
        ["Mensuel", "Trimestriel", "Annuel"],
        key="granularite_volume"
    )

    # Préparation des données de groupement
    if granularite == "Mensuel":
        df_evol['Période_Tri'] = df_evol['Date depot'].dt.to_period('M')
        df_evol['Période'] = df_evol['Date depot'].dt.strftime('%Y-%m')
        title_suffix = "Mensuelle"

    elif granularite == "Trimestriel":
        df_evol['Période_Tri'] = df_evol['Date depot'].dt.to_period('Q')
        df_evol['Période'] = df_evol['Date depot'].dt.year.astype(str) + '-T' + df_evol['Date depot'].dt.quarter.astype(
            str)
        title_suffix = "Trimestrielle"

    else:  # Annuel
        df_evol['Période_Tri'] = df_evol['Date depot'].dt.to_period('Y')
        df_evol['Période'] = df_evol['Date depot'].dt.year.astype(str)
        title_suffix = "Annuelle"

    # Agrégation
    df_grouped = df_evol.groupby(['Période_Tri', 'Période']).agg({
        'Volume_total_MWh': 'sum',
        'N° DEPOT': 'count'
    }).reset_index()

    df_grouped.columns = ['Période_Tri', 'Période', 'Volume_MWh', 'Nb_Dossiers']
    df_grouped['Volume_GWh'] = df_grouped['Volume_MWh'] / 1000
    df_grouped = df_grouped.sort_values('Période_Tri')

    # Graphique double axe
    col1, col2 = st.columns(2)

    with col1:
        st.subheader(f"📊 Évolution {title_suffix} des Volumes")
        fig_volume = make_subplots(specs=[[{"secondary_y": True}]])

        fig_volume.add_trace(
            go.Bar(x=df_grouped['Période'], y=df_grouped['Volume_GWh'],
                   name='Volume (GWh)', marker_color='#457b9d'),
            secondary_y=False
        )

        fig_volume.add_trace(
            go.Scatter(x=df_grouped['Période'], y=df_grouped['Nb_Dossiers'],
                       name='Nb Dossiers', mode='lines+markers',
                       line=dict(color='#e76f51', width=3)),
            secondary_y=True
        )

        fig_volume.update_xaxes(title_text="Période")
        fig_volume.update_yaxes(title_text="Volume (GWh cumac)", secondary_y=False)
        fig_volume.update_yaxes(title_text="Nombre de dossiers", secondary_y=True)
        fig_volume.update_layout(height=400, hovermode='x unified')

        st.plotly_chart(fig_volume, width='stretch')

    with col2:
        st.subheader(f"📈 Taux de Croissance {title_suffix}")
        if len(df_grouped) > 1:
            df_grouped['Croissance_Volume_%'] = df_grouped['Volume_GWh'].pct_change() * 100

            fig_growth = go.Figure()

            croissance_values = df_grouped['Croissance_Volume_%'][1:]

            fig_growth.add_trace(go.Bar(
                x=df_grouped['Période'][1:],
                y=croissance_values,
                name='Croissance Volume',
                marker_color=['#2a9d8f' if x >= 0 else '#e76f51' for x in croissance_values],
                text=croissance_values.round(1),
                texttemplate='%{text:+.1f}%',
                textposition='outside',
                textfont=dict(size=12, color='black'),
                hovertemplate='<b>%{x}</b><br>Croissance: %{y:+.1f}%<extra></extra>'
            ))

            fig_growth.update_layout(
                xaxis_title="Période",
                yaxis_title="Taux de croissance du volume (%)",
                height=400,
                hovermode='x unified',
                showlegend=False,
                yaxis=dict(
                    gridcolor='lightgray',
                    zeroline=True,
                    zerolinewidth=2,
                    zerolinecolor='gray'
                )
            )
            fig_growth.add_hline(y=0, line_dash="dash", line_color="gray", line_width=1)
            st.plotly_chart(fig_growth, width='stretch')
        else:
            st.info("Pas assez de données pour calculer la croissance")

    # Nouvelles fiches CEE par an
    st.markdown("---")
    st.subheader("🆕 Nouvelles Fiches CEE par Année")

    if 'Code équipement' in df.columns and 'Année_depot' in df.columns:
        first_appearance = df.groupby('Code équipement').agg({
            'Année_depot': 'min',
            'Volume_total_MWh': 'sum'
        }).reset_index()
        first_appearance.columns = ['Code équipement', 'Première_Année', 'Volume_Total_GWh']
        first_appearance['Volume_Total_GWh'] = first_appearance['Volume_Total_GWh'] / 1000

        nouvelles_fiches = first_appearance.groupby('Première_Année').agg({
            'Code équipement': 'count',
            'Volume_Total_GWh': 'sum'
        }).reset_index()
        nouvelles_fiches.columns = ['Année', 'Nb_Nouvelles_Fiches', 'Volume_Total_GWh']

        col1, col2 = st.columns([2, 1])

        with col1:
            fig_new = make_subplots(specs=[[{"secondary_y": True}]])

            fig_new.add_trace(
                go.Bar(
                    x=nouvelles_fiches['Année'],
                    y=nouvelles_fiches['Nb_Nouvelles_Fiches'],
                    name='Nb nouvelles fiches',
                    marker_color='#5F27CD',
                    text=nouvelles_fiches['Nb_Nouvelles_Fiches'],
                    textposition='outside'
                ),
                secondary_y=False
            )

            fig_new.add_trace(
                go.Scatter(
                    x=nouvelles_fiches['Année'],
                    y=nouvelles_fiches['Volume_Total_GWh'],
                    name='Volume généré (GWh)',
                    mode='lines+markers',
                    line=dict(color='#10AC84', width=3),
                    marker=dict(size=8)
                ),
                secondary_y=True
            )

            fig_new.update_xaxes(title_text="Année")
            fig_new.update_yaxes(title_text="Nombre de nouvelles fiches", secondary_y=False)
            fig_new.update_yaxes(title_text="Volume généré (GWh cumac)", secondary_y=True)
            fig_new.update_layout(
                title="Nouvelles fiches CEE introduites par an et volumes générés",
                height=400,
                hovermode='x unified'
            )

            st.plotly_chart(fig_new, width='stretch')

        with col2:
            st.markdown("#### 📋 Détails par année")
            annee_selectionnee = st.selectbox(
                "Sélectionner une année",
                options=sorted(nouvelles_fiches['Année'].unique(), reverse=True)
            )

            nouvelles_cette_annee = first_appearance[
                first_appearance['Première_Année'] == annee_selectionnee
                ][['Code équipement', 'Volume_Total_GWh']].sort_values('Volume_Total_GWh', ascending=False)

            st.write(f"**{len(nouvelles_cette_annee)} nouvelle(s) fiche(s) en {int(annee_selectionnee)}:**")
            for _, row in nouvelles_cette_annee.iterrows():
                st.write(f"- {row['Code équipement']}: {row['Volume_Total_GWh']:.2f} GWh")

    # Répartition Précarité vs Classique
    st.markdown("---")
    col1, col2 = st.columns(2)

    with col1:
        st.subheader("🎯 Répartition Précarité vs Classique")
        if 'Total_précarité_MWh' in df.columns and 'Total_classique_MWh' in df.columns:
            volume_prec = df['Total_précarité_MWh'].sum() / 1000
            volume_class = df['Total_classique_MWh'].sum() / 1000

            fig_pie = go.Figure(data=[go.Pie(
                labels=['Précarité', 'Classique'],
                values=[volume_prec, volume_class],
                hole=0.4,
                marker_colors=['#e76f51', '#457b9d'],
                textinfo='label+percent+value',
                texttemplate='%{label}<br>%{percent}<br>%{value:.1f} GWh'
            )])
            fig_pie.update_layout(height=350)
            st.plotly_chart(fig_pie, width='stretch')

    with col2:
        st.subheader(f"📈 Évolution {title_suffix} Précarité/Classique (GWh cumac)")
        if 'Total_précarité_MWh' in df.columns and 'Total_classique_MWh' in df.columns:
            if granularite == "Mensuel":
                evol_prec = df.groupby('Date_depot_str').agg({
                    'Total_précarité_MWh': 'sum',
                    'Total_classique_MWh': 'sum'
                }).reset_index()
                evol_prec.columns = ['Période', 'Précarité', 'Classique']
            elif granularite == "Trimestriel":
                df_prec = df.dropna(subset=['Date depot']).copy()
                df_prec['Période_Tri'] = df_prec['Date depot'].dt.to_period('Q')
                df_prec['Période'] = df_prec['Date depot'].dt.year.astype(str) + '-T' + df_prec[
                    'Date depot'].dt.quarter.astype(str)

                evol_prec = df_prec.groupby(['Période_Tri', 'Période']).agg({
                    'Total_précarité_MWh': 'sum',
                    'Total_classique_MWh': 'sum'
                }).reset_index()
                evol_prec = evol_prec.sort_values('Période_Tri')
                evol_prec.columns = ['Période_Tri', 'Période', 'Précarité', 'Classique']

            else:
                df_prec = df.dropna(subset=['Date depot']).copy()
                df_prec['Période'] = df_prec['Date depot'].dt.year.astype(str)
                evol_prec = df_prec.groupby('Période').agg({
                    'Total_précarité_MWh': 'sum',
                    'Total_classique_MWh': 'sum'
                }).reset_index()
                evol_prec.columns = ['Période', 'Précarité', 'Classique']

            evol_prec['Précarité'] = evol_prec['Précarité'] / 1000
            evol_prec['Classique'] = evol_prec['Classique'] / 1000

            fig_evol_prec = go.Figure()
            fig_evol_prec.add_trace(go.Bar(
                x=evol_prec['Période'],
                y=evol_prec['Précarité'],
                name='Précarité',
                marker_color='#e76f51'
            ))
            fig_evol_prec.add_trace(go.Bar(
                x=evol_prec['Période'],
                y=evol_prec['Classique'],
                name='Classique',
                marker_color='#457b9d'
            ))
            fig_evol_prec.update_layout(
                barmode='stack',
                xaxis_title="Période",
                yaxis_title="Volume (GWh cumac)",
                height=350
            )
            st.plotly_chart(fig_evol_prec, width='stretch')


def create_geographic_analysis(df):
    """Analyse géographique"""
    st.header("🗺️ Analyse Géographique")

    if 'Département' not in df.columns or 'geo_valid' not in df.columns:
        st.warning("Les données géographiques ne sont pas disponibles")
        return

    df_geo = df[df['geo_valid'] == True].copy()

    if len(df_geo) == 0:
        st.warning("Aucune donnée géographique valide trouvée")
        return

    geo_agg = df_geo.groupby('Département').agg({
        'Volume_total_MWh': 'sum',
        'N° DEPOT': 'count'
    }).reset_index()
    geo_agg.columns = ['Département', 'Volume_GWh', 'Nb_Dossiers']
    geo_agg['Volume_GWh'] = geo_agg['Volume_GWh'] / 1000
    geo_agg = geo_agg.sort_values('Volume_GWh', ascending=False)

    geo_agg['Volume_Moyen'] = geo_agg['Volume_GWh'] / geo_agg['Nb_Dossiers']

    st.subheader("📊 Tableau Récapitulatif par Département")
    st.info(f"**{len(geo_agg)}** départements avec des opérations CEE")

    display_geo = geo_agg.copy()
    display_geo['Volume_GWh'] = display_geo['Volume_GWh'].apply(lambda x: f"{x:,.2f}".replace(',', ' '))
    display_geo['Nb_Dossiers'] = display_geo['Nb_Dossiers'].apply(lambda x: f"{x:,}".replace(',', ' '))
    display_geo['Volume_Moyen'] = display_geo['Volume_Moyen'].apply(lambda x: f"{x:,.2f}".replace(',', ' '))
    display_geo.columns = ['Département', 'Volume Total (GWh cumac)', 'Nb Dossiers', 'Volume Moyen (GWh/dossier)']

    st.dataframe(display_geo, width="stretch", height=600)

    # Statistiques récapitulatives
    st.markdown("---")
    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric("🏆 Département leader (volume)", geo_agg.iloc[0]['Département'])

    with col2:
        top_dept_volume = geo_agg.iloc[0]['Volume_GWh']
        st.metric("Volume du leader", f"{top_dept_volume:.2f} GWh")

    with col3:
        dept_plus_dossiers = geo_agg.sort_values('Nb_Dossiers', ascending=False).iloc[0]
        st.metric("🏆 Département leader (dossiers)", dept_plus_dossiers['Département'])

    with col4:
        nb_dossiers_leader = dept_plus_dossiers['Nb_Dossiers']
        st.metric("Nb dossiers du leader", f"{int(nb_dossiers_leader):,}".replace(',', ' '))


def create_equipment_analysis(df):
    """Analyse par équipement"""
    st.header("🔧 Analyse par Équipement")

    if 'Code équipement' not in df.columns:
        st.warning("Les données d'équipement ne sont pas disponibles")
        return

    equip_agg = df.groupby('Code équipement').agg({
        'Volume_total_MWh': 'sum',
        'N° DEPOT': 'count'
    }).reset_index()
    equip_agg.columns = ['Code équipement', 'Volume_GWh', 'Nb_Dossiers']
    equip_agg['Volume_GWh'] = equip_agg['Volume_GWh'] / 1000
    equip_agg = equip_agg.sort_values('Volume_GWh', ascending=False)

    col1, col2 = st.columns(2)

    with col1:
        st.subheader("📊 Top 15 Équipements - Volume")
        st.info("💡 **Les fiches CEE les plus performantes** en termes de volume de certificats générés")

        top_15 = equip_agg.head(15).sort_values('Volume_GWh')

        fig_equip = go.Figure(go.Bar(
            x=top_15['Volume_GWh'],
            y=top_15['Code équipement'],
            orientation='h',
            marker=dict(
                color=top_15['Volume_GWh'],
                colorscale='Greens',
                showscale=False,
                line=dict(color='rgb(0,100,0)', width=1)
            ),
            text=top_15['Volume_GWh'].round(2),
            textposition='outside',
            texttemplate='%{text:.2f} GWh',
            hovertemplate='<b>%{y}</b><br>Volume: %{x:.2f} GWh cumac<extra></extra>'
        ))

        fig_equip.update_layout(
            title="Top 15 équipements par volume (GWh cumac)",
            xaxis_title="Volume (GWh cumac)",
            yaxis_title="Code équipement",
            height=500,
            font=dict(size=11),
            plot_bgcolor='rgba(0,0,0,0)',
            paper_bgcolor='rgba(0,0,0,0)',
            margin=dict(l=150, r=50, t=50, b=50)
        )
        fig_equip.update_xaxes(showgrid=True, gridwidth=0.5, gridcolor='LightGray')

        st.plotly_chart(fig_equip, width='stretch')

    with col2:
        st.subheader("📋 Tableau détaillé")
        display_equip = equip_agg.head(20).copy()
        display_equip['Volume_GWh'] = display_equip['Volume_GWh'].apply(lambda x: f"{x:,.2f}".replace(',', ' '))
        display_equip['Nb_Dossiers'] = display_equip['Nb_Dossiers'].apply(lambda x: f"{x:,}".replace(',', ' '))
        display_equip.columns = ['Code équipement', 'Volume (GWh cumac)', 'Nb Dossiers']
        st.dataframe(display_equip, width="stretch", height=500)

    # Camembert de répartition
    st.markdown("---")
    st.subheader("🥧 Répartition des Volumes par Fiche CEE")

    col1, col2 = st.columns([2, 1])

    with col1:
        top_10_equip = equip_agg.head(10).copy()
        autres_volume = equip_agg.iloc[10:]['Volume_GWh'].sum()

        if autres_volume > 0:
            autres_row = pd.DataFrame([{
                'Code équipement': 'Autres',
                'Volume_GWh': autres_volume,
                'Nb_Dossiers': equip_agg.iloc[10:]['Nb_Dossiers'].sum()
            }])
            equip_pie_data = pd.concat([top_10_equip, autres_row], ignore_index=True)
        else:
            equip_pie_data = top_10_equip

        fig_pie = px.pie(
            equip_pie_data,
            values='Volume_GWh',
            names='Code équipement',
            title="Répartition des volumes par fiche CEE (Top 10 + Autres)",
            color_discrete_sequence=px.colors.qualitative.Set3
        )
        fig_pie.update_traces(
            textposition='inside',
            textinfo='label+percent',
            hovertemplate='<b>%{label}</b><br>Volume: %{value:.2f} GWh<br>Part: %{percent}<extra></extra>'
        )
        fig_pie.update_layout(height=450)
        st.plotly_chart(fig_pie, width='stretch')

    with col2:
        st.markdown("#### 📊 Statistiques")
        total_fiches = len(equip_agg)
        volume_total = equip_agg['Volume_GWh'].sum()
        top_3_volume = equip_agg.head(3)['Volume_GWh'].sum()
        pct_top_3 = (top_3_volume / volume_total * 100) if volume_total > 0 else 0

        st.metric("Nombre de fiches différentes", f"{total_fiches}")
        st.metric("Volume total", f"{volume_total:.2f} GWh")
        st.metric("Part du Top 3", f"{pct_top_3:.1f}%")

        st.markdown("#### 🏆 Podium")
        for i, row in equip_agg.head(3).iterrows():
            emoji = ["🥇", "🥈", "🥉"][list(equip_agg.head(3).index).index(i)]
            st.write(f"{emoji} **{row['Code équipement']}**: {row['Volume_GWh']:.2f} GWh")


def create_installer_performance(df):
    """Performance par installateur - VERSION AMÉLIORÉE"""
    st.header("🏢 Performance des Installateurs")

    if 'N° d\'identification du professionnel' not in df.columns:
        st.warning("Les données d'installateurs ne sont pas disponibles")
        return

    # Regroupement par SIREN
    siren_raison = df.groupby('N° d\'identification du professionnel')['Raison sociale du professionnel'].agg(
        lambda x: x.value_counts().index[0] if len(x) > 0 else ''
    ).reset_index()

    installer_agg = df.groupby('N° d\'identification du professionnel').agg({
        'Volume_total_MWh': 'sum',
        'N° DEPOT': 'count'
    }).reset_index()

    installer_agg = installer_agg.merge(siren_raison, on='N° d\'identification du professionnel')

    # FILTRES AVANCÉS
    st.subheader("🔍 Filtres")

    col_filter1, col_filter2, col_filter3 = st.columns(3)

    with col_filter1:
        min_dossiers = st.number_input(
            "Nombre minimum de dossiers",
            min_value=1,
            max_value=100,
            value=3,
            step=1,
            help="Filtrer les installateurs avec au moins X dossiers"
        )

    with col_filter2:
        min_volume = st.number_input(
            "Volume minimum (MWh)",
            min_value=0,
            max_value=100000,
            value=0,
            step=100,
            help="Filtrer les installateurs avec au moins X MWh"
        )

    with col_filter3:
        search_installer = st.text_input(
            "🔎 Rechercher un installateur",
            "",
            help="Recherche par nom (partielle)"
        )

    # Application des filtres
    installer_filtered = installer_agg[installer_agg['N° DEPOT'] >= min_dossiers].copy()
    if min_volume > 0:
        installer_filtered = installer_filtered[installer_filtered['Volume_total_MWh'] >= min_volume]
    if search_installer:
        installer_filtered = installer_filtered[
            installer_filtered['Raison sociale du professionnel'].str.contains(search_installer, case=False, na=False)
        ]

    installer_filtered = installer_filtered.sort_values('Volume_total_MWh', ascending=False)

    st.info(f"📊 **{len(installer_filtered)}** installateur(s) après filtrage (sur {len(installer_agg)} total)")

    st.markdown("---")

    # CAMEMBERT DES VOLUMES
    st.subheader("📊 Répartition des Volumes par Installateur")

    col_pie1, col_pie2 = st.columns([3, 1])

    with col_pie2:
        count_installers = len(installer_filtered)

        if count_installers > 5:
            max_val = min(50, count_installers)
            min_val = 5 if max_val > 5 else 1

            top_n_pie = st.slider(
                "Nombre d'installateurs à afficher",
                min_value=min_val,
                max_value=max_val,
                value=min(10, max_val),
                step=1 if max_val < 10 else 5,
                key="slider_pie"
            )
        else:
            top_n_pie = count_installers

    with col_pie1:
        if len(installer_filtered) > top_n_pie:
            top_installateurs = installer_filtered.head(top_n_pie).copy()
            autres_volume = installer_filtered.iloc[top_n_pie:]['Volume_total_MWh'].sum()
            autres_row = pd.DataFrame({
                'Raison sociale du professionnel': ['Autres'],
                'Volume_total_MWh': [autres_volume]
            })
            installer_volumes_plot = pd.concat(
                [top_installateurs[['Raison sociale du professionnel', 'Volume_total_MWh']],
                 autres_row], ignore_index=True)
        else:
            installer_volumes_plot = installer_filtered[['Raison sociale du professionnel', 'Volume_total_MWh']]

        if not installer_volumes_plot.empty:
            fig_camembert = px.pie(
                installer_volumes_plot,
                values='Volume_total_MWh',
                names='Raison sociale du professionnel',
                title=f'Répartition des Volumes (Top {top_n_pie} + Autres)',
                hole=0.4,
                color_discrete_sequence=px.colors.qualitative.Set3
            )
            fig_camembert.update_traces(textposition='inside', textinfo='percent+label')
            fig_camembert.update_layout(height=500)
            st.plotly_chart(fig_camembert, width='stretch')
        else:
            st.warning("Pas assez de données pour afficher le graphique.")

    st.markdown("---")

    # ANALYSE DES DÉLAIS PAR INSTALLATEUR
    st.subheader("⏱️ Analyse des Délais par Installateur")

    delais_columns = {
        'Délai Travaux (Fin - Début)': 'Délai_travaux',
        'Délai Dépôt (Dépôt - Fin)': 'Délai_fin_depot',
        'Délai Validation (Validation - Dépôt)': 'Délai_validation',
        'Délai Total (Validation - Fin)': 'Délai_fin_validation'
    }

    available_delais = {k: v for k, v in delais_columns.items() if v in df.columns}

    if not available_delais:
        st.warning("Les données de délais ne sont pas disponibles.")
        return

    # Calcul des moyennes de délais
    agg_delais = {'N° DEPOT': 'count'}
    for col in available_delais.values():
        agg_delais[col] = 'mean'

    df_delais = df.groupby('N° d\'identification du professionnel').agg(agg_delais).reset_index()
    df_delais = df_delais.merge(siren_raison, on='N° d\'identification du professionnel')

    df_delais = df_delais[df_delais['N° DEPOT'] >= min_dossiers]
    if search_installer:
        df_delais = df_delais[
            df_delais['Raison sociale du professionnel'].str.contains(search_installer, case=False, na=False)]

    if len(df_delais) == 0:
        st.warning("Aucun installateur pour l'analyse des délais avec les filtres actuels.")
        return

    choix_delai_label = st.selectbox("Choisir le type de délai à analyser", list(available_delais.keys()))
    choix_delai_col = available_delais[choix_delai_label]

    tri_ordre = st.radio("Trier par :", ["Délai Croissant (Plus rapide)", "Délai Décroissant (Plus lent)"],
                         horizontal=True)
    ascending_order = True if tri_ordre == "Délai Croissant (Plus rapide)" else False

    df_delais_sorted = df_delais.sort_values(choix_delai_col, ascending=ascending_order).head(15)

    fig_delais = px.bar(
        df_delais_sorted,
        x=choix_delai_col,
        y='Raison sociale du professionnel',
        orientation='h',
        title=f"Top 15 Installateurs - {choix_delai_label}",
        labels={choix_delai_col: "Délai moyen (jours)", 'Raison sociale du professionnel': 'Installateur'},
        text=choix_delai_col,
        color=choix_delai_col,
        color_continuous_scale='RdYlGn_r' if ascending_order else 'RdYlGn'
    )

    fig_delais.update_traces(texttemplate='%{text:.1f} j', textposition='outside')
    fig_delais.update_layout(height=500,
                             yaxis={'categoryorder': 'total ascending' if ascending_order else 'total descending'})
    st.plotly_chart(fig_delais, width='stretch')

    # Tableau récapitulatif
    with st.expander("📋 Voir le tableau détaillé des délais par installateur"):
        cols_table = ['Raison sociale du professionnel', 'N° DEPOT'] + list(available_delais.values())
        display_table = df_delais[cols_table].copy()

        rename_map = {v: k for k, v in available_delais.items()}
        rename_map['Raison sociale du professionnel'] = 'Installateur'
        rename_map['N° DEPOT'] = 'Nb Dossiers'
        display_table = display_table.rename(columns=rename_map)

        for col in available_delais.keys():
            if col in display_table.columns:
                display_table[col] = display_table[col].round(1)

        st.dataframe(display_table, width="stretch")


def create_controle_synthese_analysis(df_synthese, df_detail):
    """ONGLET CONTRÔLE AMÉLIORÉ - Calcul correct du taux de délivrance"""
    st.header("🛡️ Cockpit de Contrôle & Conformité")

    if df_synthese is None or len(df_synthese) == 0:
        st.warning("⚠️ Aucune donnée de synthèse disponible avec les filtres actuels.")
        return

    # FILTRAGE - On ne prend que les dépôts avec un volume délivré > 0
    # (c'est-à-dire ceux qui ont été traités et ont une décision)
    df_valid = df_synthese[
        (df_synthese['Statut_Synthese'] == 'Validé') &
        (df_synthese['Volume delivre'] > 0)
        ].copy()

    df_instruction = df_synthese[df_synthese['Statut_Synthese'] == 'En instruction'].copy()

    # KPIs GLOBAUX - Calcul basé uniquement sur les dépôts avec volume délivré
    st.subheader("📊 Performance de Délivrance & Retraits")

    # Calcul des volumes (en GWh) - UNIQUEMENT pour les dépôts avec volume délivré
    vol_demande = df_valid['Volume demande'].sum() / 1000000
    vol_delivre = df_valid['Volume delivre'].sum() / 1000000
    vol_retire = vol_demande - vol_delivre

    # Calcul des opérations
    ops_demande = df_valid['Nombre demande operation'].sum()
    ops_valide = df_valid['Nombre d\'opérations validées'].sum()
    ops_retire = ops_demande - ops_valide

    # Calcul des taux
    taux_retrait_vol = (vol_retire / vol_demande * 100) if vol_demande > 0 else 0
    taux_retrait_ops = (ops_retire / ops_demande * 100) if ops_demande > 0 else 0
    taux_delivrance = (vol_delivre / vol_demande * 100) if vol_demande > 0 else 0

    # Informations sur l'échantillon analysé
    total_depots_synthese = len(df_synthese)
    depots_avec_delivrance = len(df_valid)

    st.info(
        f"📊 **Analyse basée sur {depots_avec_delivrance} dépôts avec volume délivré** (sur {total_depots_synthese} dépôts affichés)")

    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    with kpi1:
        st.metric(
            "Volume Délivré",
            f"{vol_delivre:,.1f} GWh",
            delta=f"{depots_avec_delivrance} dépôts",
            help=f"Volume total délivré sur {depots_avec_delivrance} dépôts traités"
        )

    with kpi2:
        st.metric(
            "Taux Délivrance",
            f"{taux_delivrance:.1f}%",
            delta=f"-{100 - taux_delivrance:.1f}% vs Cible",
            delta_color="normal",
            help="Volume délivré / Volume demandé (uniquement dépôts avec délivrance)"
        )

    with kpi3:
        st.metric(
            "Taux Retrait (Volume)",
            f"{taux_retrait_vol:.2f}%",
            delta="Cible : 0%",
            delta_color="inverse",
            help=f"Volume retiré : {vol_retire:.2f} GWh"
        )

    with kpi4:
        st.metric(
            "Taux Retrait (Opérations)",
            f"{taux_retrait_ops:.2f}%",
            delta=f"{int(ops_retire)} ops retirées",
            delta_color="inverse",
            help=f"Sur {int(ops_demande)} opérations demandées"
        )

    st.markdown("---")

    # DÉTAILS DES CALCULS
    with st.expander("📊 Voir le détail des calculs de délivrance"):
        st.markdown("### Répartition des dépôts dans la feuille Synthèse")

        col_detail1, col_detail2, col_detail3 = st.columns(3)

        with col_detail1:
            st.metric(
                "Total dépôts Synthèse",
                f"{total_depots_synthese}",
                help="Nombre total de lignes dans la feuille Synthèse après filtrage"
            )

        with col_detail2:
            depots_en_instruction = len(df_instruction)
            st.metric(
                "En instruction",
                f"{depots_en_instruction}",
                help="Dépôts en cours de traitement (pas encore de décision)"
            )

        with col_detail3:
            st.metric(
                "Avec délivrance",
                f"{depots_avec_delivrance}",
                help="Dépôts validés avec un volume délivré > 0"
            )

        st.markdown("---")
        st.markdown("### Volumes analysés (GWh)")

        col_vol1, col_vol2, col_vol3 = st.columns(3)

        with col_vol1:
            st.metric(
                "Volume demandé",
                f"{vol_demande:.2f} GWh",
                help="Somme des volumes demandés pour les dépôts avec délivrance"
            )

        with col_vol2:
            st.metric(
                "Volume délivré",
                f"{vol_delivre:.2f} GWh",
                help="Somme des volumes effectivement délivrés"
            )

        with col_vol3:
            st.metric(
                "Volume retiré",
                f"{vol_retire:.2f} GWh",
                delta=f"{taux_retrait_vol:.2f}%",
                delta_color="inverse",
                help="Différence entre demandé et délivré"
            )

    st.markdown("---")

    # FOCUS ALERTE : Dépôts < 100%
    st.subheader("🚨 Focus Alertes : Dépôts avec Coupes (< 100%)")

    df_alert = df_valid[df_valid['Volume delivre'] < (df_valid['Volume demande'] - 1)].copy()

    if len(df_alert) > 0:
        df_alert['Perte (kWh)'] = df_alert['Volume demande'] - df_alert['Volume delivre']
        df_alert['Taux Réel'] = (df_alert['Volume delivre'] / df_alert['Volume demande'] * 100).round(1)

        st.warning(f"⚠️ **{len(df_alert)}** dépôt(s) ont subi des retraits partiels ou totaux.")

        col_alert1, col_alert2 = st.columns([2, 1])

        with col_alert1:
            cols_candidate = ['Depot', 'Date Depot', 'Volume demande', 'Volume delivre', 'Perte (kWh)', 'Taux Réel']
            cols_to_show = [c for c in cols_candidate if c in df_alert.columns]

            if 'Perte (kWh)' in cols_to_show:
                st.dataframe(
                    df_alert[cols_to_show].sort_values('Perte (kWh)', ascending=False),
                    width="stretch"
                )
            else:
                st.dataframe(df_alert[cols_to_show], width="stretch")

        with col_alert2:
            if 'Type de Retrait' in df_alert.columns:
                rejet_counts = df_alert['Type de Retrait'].value_counts()
                st.write("**Principaux motifs :**")
                st.table(rejet_counts)
            elif 'Type de rejet' in df_alert.columns:
                rejet_counts = df_alert['Type de rejet'].value_counts()
                st.write("**Principaux motifs :**")
                st.table(rejet_counts)
    else:
        st.success("✅ Aucun dépôt avec retrait détecté ! Performance 100%.")

    st.markdown("---")

    # GESTION DES RISQUES (EXPIRATION)
    st.subheader("⏳ Gestion des Risques : Dépôts In Extremis (Expiration)")
    st.info("Analyse des dossiers déposés dans les 3 derniers mois de validité (Date fin + 12 mois).")

    if df_detail is not None and 'Date de fin' in df_detail.columns and 'Date depot' in df_detail.columns:
        df_exp = df_detail.dropna(subset=['Date de fin', 'Date depot']).copy()
        df_exp['Date_Exp'] = df_exp['Date de fin'] + pd.DateOffset(months=12)
        df_exp['Zone_Risk'] = df_exp['Date_Exp'] - pd.DateOffset(months=3)

        mask_risk = (df_exp['Date depot'] >= df_exp['Zone_Risk']) & (df_exp['Date depot'] <= df_exp['Date_Exp'])
        df_risk = df_exp[mask_risk].copy()

        if len(df_risk) > 0:
            col_risk1, col_risk2 = st.columns([1, 2])

            with col_risk1:
                st.metric("Dossiers 'Limites'", len(df_risk), f"{(len(df_risk) / len(df_exp) * 100):.1f}% du total",
                          delta_color="inverse")
                st.caption("Ces dossiers augmentent le risque de retrait pour hors délai.")

            with col_risk2:
                if 'Raison sociale du professionnel' in df_risk.columns:
                    top_risk_installers = df_risk['Raison sociale du professionnel'].value_counts().reset_index()
                    top_risk_installers.columns = ['Installateur', 'Nb Dossiers Limites']
                    top_risk_installers = top_risk_installers.head(10)

                    fig_risk = px.bar(
                        top_risk_installers,
                        x='Nb Dossiers Limites',
                        y='Installateur',
                        orientation='h',
                        title="Top Installateurs : Dossiers déposés en fin de validité",
                        text='Nb Dossiers Limites',
                        color='Nb Dossiers Limites',
                        color_continuous_scale='Reds'
                    )
                    fig_risk.update_layout(yaxis={'categoryorder': 'total ascending'}, height=350)
                    st.plotly_chart(fig_risk, width='stretch')

            with st.expander("📋 Voir le détail des dossiers à risque d'expiration"):
                st.dataframe(
                    df_risk[['N° DEPOT', 'Date de fin', 'Date depot', 'Date_Exp', 'Raison sociale du professionnel']],
                    width="stretch")
        else:
            st.success("✅ Aucun dossier déposé en zone critique d'expiration.")
    else:
        st.warning("Données de détail manquantes pour l'analyse d'expiration.")


def create_iso_quality_analysis(df, df_synthese):
    """Version COMPACTE & ESTHÉTIQUE avec ÉVOLUTION TEMPORELLE"""
    st.header("✨ Pilotage Qualité ISO 9001")

    # LES 4 INDICATEURS CLÉS
    kpi1, kpi2, kpi3, kpi4 = st.columns(4)

    with kpi1:
        if 'Erreur de saisi' in df.columns:
            nb_erreurs = df['Erreur de saisi'].sum()
            total_lignes = len(df)
            taux_erreur = (nb_erreurs / total_lignes) * 100 if total_lignes > 0 else 0

            st.metric(
                label="📉 Taux Erreur CSV",
                value=f"{taux_erreur:.2f}%",
                delta="Cible: 0%",
                delta_color="inverse"
            )
        else:
            st.metric("Taux Erreur", "N/A")

    with kpi2:
        if df_synthese is not None and 'Volume demande' in df_synthese.columns:
            SEUIL = 50_000_000
            nb_sous = len(df_synthese[df_synthese['Volume demande'] < SEUIL])
            total = len(df_synthese)
            pct_ok = ((total - nb_sous) / total * 100) if total > 0 else 0

            st.metric(
                label="📦 Dépôts > 50 GWhc",
                value=f"{pct_ok:.1f}%",
                delta=f"-{nb_sous} dépôts non conformes",
                delta_color="normal"
            )
        else:
            st.metric("Conformité Vol.", "N/A")

    with kpi3:
        if 'Date depot' in df.columns and 'Date Insertion' in df.columns:
            df['Delai_Insertion'] = (df['Date depot'] - df['Date Insertion']).dt.days
            avg_delai = df['Delai_Insertion'].mean()
            hors_delai = len(df[df['Delai_Insertion'] > 14])

            st.metric(
                label="⏱️ Délai Insertion",
                value=f"{avg_delai:.1f} j",
                delta=f"{hors_delai} dossiers > 14j",
                delta_color="inverse"
            )
        else:
            st.metric("Délai Insertion", "N/A")

    with kpi4:
        if 'Date de fin' in df.columns and 'Date depot' in df.columns:
            df_exp = df.dropna(subset=['Date de fin', 'Date depot']).copy()
            df_exp['Date_Exp'] = df_exp['Date de fin'] + pd.DateOffset(months=12)
            df_exp['Zone_Risk'] = df_exp['Date_Exp'] - pd.DateOffset(months=3)

            nb_risk = len(
                df_exp[(df_exp['Date depot'] >= df_exp['Zone_Risk']) & (df_exp['Date depot'] <= df_exp['Date_Exp'])])
            total_dossiers = len(df_exp)
            taux_risk = (nb_risk / total_dossiers * 100) if total_dossiers > 0 else 0

            st.metric(
                label="⏳ Taux Risque Expiration",
                value=f"{taux_risk:.1f}%",
                delta=f"{nb_risk} dossiers critiques",
                delta_color="inverse"
            )
        else:
            st.metric("Risque Expiration", "N/A")

    st.markdown("---")

    # ÉVOLUTION TEMPORELLE
    st.subheader("📈 Évolution Temporelle")

    col_evol1, col_evol2 = st.columns(2)

    with col_evol1:
        if 'Delai_Insertion' in df.columns and 'Date depot' in df.columns:
            df_evol_ins = df.dropna(subset=['Delai_Insertion', 'Date depot']).copy()
            df_evol_ins['Mois'] = df_evol_ins['Date depot'].dt.to_period('M').astype(str)

            evol_ins = df_evol_ins.groupby('Mois')['Delai_Insertion'].mean().reset_index()

            fig_ins = px.line(
                evol_ins,
                x='Mois',
                y='Delai_Insertion',
                title="Évolution du Délai Moyen d'Insertion",
                markers=True
            )
            fig_ins.add_hline(y=14, line_dash="dash", line_color="red", annotation_text="Cible 14j")
            fig_ins.update_layout(height=300, xaxis_title=None, yaxis_title="Jours")
            st.plotly_chart(fig_ins, width='stretch')

    with col_evol2:
        if 'Date de fin' in df.columns and 'Date depot' in df.columns:
            df_exp = df.dropna(subset=['Date de fin', 'Date depot']).copy()
            df_exp['Date_Exp'] = df_exp['Date de fin'] + pd.DateOffset(months=12)
            df_exp['Zone_Risk'] = df_exp['Date_Exp'] - pd.DateOffset(months=3)
            df_exp['Is_Risk'] = (df_exp['Date depot'] >= df_exp['Zone_Risk']) & (
                    df_exp['Date depot'] <= df_exp['Date_Exp'])
            df_exp['Mois'] = df_exp['Date depot'].dt.to_period('M').astype(str)

            evol_risk = df_exp.groupby('Mois')['Is_Risk'].mean().reset_index()
            evol_risk['Taux_Risk'] = evol_risk['Is_Risk'] * 100

            fig_risk_evol = px.line(
                evol_risk,
                x='Mois',
                y='Taux_Risk',
                title="Évolution du Taux de Dossiers à Risque Expiration (%)",
                markers=True,
                color_discrete_sequence=['orange']
            )
            fig_risk_evol.update_layout(height=300, xaxis_title=None, yaxis_title="% Dossiers à Risque")
            st.plotly_chart(fig_risk_evol, width='stretch')

    st.markdown("---")

    # VISUALISATIONS GRAPHIQUES COMPACTES
    col_g1, col_g2 = st.columns([1, 1])

    with col_g1:
        st.caption("📊 **Distribution des délais d'insertion (Cible : 14j)**")
        if 'Delai_Insertion' in df.columns:
            df_clean_delai = df.dropna(subset=['Delai_Insertion'])
            fig_hist = px.histogram(
                df_clean_delai,
                x='Delai_Insertion',
                nbins=20,
                color_discrete_sequence=['#636efa']
            )
            fig_hist.add_vline(x=14, line_width=2, line_dash="dash", line_color="red", annotation_text="Cible")

            fig_hist.update_layout(
                height=250,
                margin=dict(l=0, r=0, t=0, b=0),
                yaxis_title=None,
                xaxis_title="Jours",
                plot_bgcolor='white',
                showlegend=False
            )
            st.plotly_chart(fig_hist, width='stretch')
        else:
            st.info("Données insuffisantes pour l'histogramme.")

    with col_g2:
        with st.expander("📋 Voir les anomalies (Erreurs & Volumes)", expanded=True):
            tab_d1, tab_d2 = st.tabs(["Erreurs CSV", "Volumes < 50GWh"])

            with tab_d1:
                if 'Erreur de saisi' in df.columns and df['Erreur de saisi'].sum() > 0:
                    st.dataframe(df[df['Erreur de saisi'] == 1], width="stretch", height=200)
                else:
                    st.success("Zéro erreur de saisie !")

            with tab_d2:
                if df_synthese is not None and 'Volume demande' in df_synthese.columns:
                    st.dataframe(df_synthese[df_synthese['Volume demande'] < 50_000_000], width="stretch", height=200)
                else:
                    st.info("Pas de données.")


def create_rse_analysis(df, taux_efficacite=0.45):
    """Analyse RSE condensée"""
    st.header("🌱 Analyse RSE - Impact Environnemental et Social")

    metrics = calculate_rse_metrics(df, taux_efficacite)

    # PARAMÈTRE TAUX EFFICACITÉ
    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Paramètres RSE")
    taux_efficacite_input = st.sidebar.slider(
        "Taux d'efficacité énergétique",
        min_value=0.30,
        max_value=0.60,
        value=0.45,
        step=0.05,
        help="Coefficient de réalisation des économies d'énergie réelles"
    )

    if taux_efficacite_input != taux_efficacite:
        metrics = calculate_rse_metrics(df, taux_efficacite_input)

    # KPIs ENVIRONNEMENTAUX
    st.subheader("🌍 Impact Environnemental")

    col1, col2, col3, col4 = st.columns(4)

    with col1:
        st.metric(
            "⚡ GWh cumac Total",
            f"{metrics['gwhc_total']:.2f} GWh"
        )

    with col2:
        st.metric(
            "🔋 GWh Réels/an",
            f"{metrics['gwh_reels']:.2f} GWh/an",
            help=f"Avec taux d'efficacité de {taux_efficacite_input * 100:.0f}%"
        )

    with col3:
        st.metric(
            "🌳 CO₂ Évité",
            f"{metrics['co2_evite']:,.0f} tonnes/an".replace(',', ' ')
        )

    with col4:
        st.metric(
            "🏠 Foyers Alimentés",
            f"{metrics['foyers_equivalent']:,.0f}".replace(',', ' ')
        )

    col5, col6 = st.columns(2)

    with col5:
        st.metric(
            "🌲 Arbres Équivalents",
            f"{metrics['arbres_equivalent']:,.0f} arbres".replace(',', ' '),
            help="Basé sur 25 kg CO₂/arbre/an"
        )

    with col6:
        st.metric(
            "🚗 Voitures Retirées",
            f"{metrics['voitures_equivalent']:,.0f} voitures".replace(',', ' '),
            help="Basé sur 2.8 tonnes CO₂/voiture/an"
        )

    st.markdown("---")

    # PERFORMANCE SOCIALE
    st.subheader("👥 Performance Sociale")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 💰 Répartition Précarité vs Classique")

        if 'Total_précarité_MWh' in df.columns and 'Total_classique_MWh' in df.columns:
            volume_prec = df['Total_précarité_MWh'].sum()
            volume_class = df['Total_classique_MWh'].sum()

            pct_prec = (volume_prec / (volume_prec + volume_class) * 100) if (volume_prec + volume_class) > 0 else 0

            fig_prec = go.Figure(data=[go.Pie(
                labels=['Précarité', 'Classique'],
                values=[volume_prec, volume_class],
                hole=0.5,
                marker_colors=['#e76f51', '#457b9d'],
                textinfo='label+percent',
                textfont_size=14
            )])

            fig_prec.update_layout(
                height=350,
                annotations=[dict(text=f'{pct_prec:.1f}%<br>Précarité',
                                  x=0.5, y=0.5, font_size=20, showarrow=False)]
            )

            st.plotly_chart(fig_prec, width='stretch')

            st.info(f"**{pct_prec:.1f}%** des volumes concernent des opérations en précarité énergétique")

    with col2:
        st.markdown("#### 🏗️ Répartition par Secteur")

        if 'Secteur' in df.columns:
            secteur_agg = df.groupby('Secteur')['Volume_total_MWh'].sum().reset_index()
            secteur_agg = secteur_agg.sort_values('Volume_total_MWh', ascending=False)

            fig_secteur = px.pie(
                secteur_agg,
                names='Secteur',
                values='Volume_total_MWh',
                title="",
                color_discrete_sequence=px.colors.qualitative.Set2
            )
            fig_secteur.update_traces(textinfo='label+percent', textfont_size=12)
            fig_secteur.update_layout(height=350, showlegend=True)

            st.plotly_chart(fig_secteur, width='stretch')

    st.markdown("---")

    # GRAPHIQUES CLÉS
    st.subheader("📊 Analyses Détaillées")

    col1, col2 = st.columns(2)

    with col1:
        st.markdown("#### 📈 Évolution Annuelle des Économies")

        if 'Année_depot' in df.columns and 'GWh_reels_annuels' in df.columns:
            evol_gwh = df.groupby('Année_depot')['GWh_reels_annuels'].sum().reset_index()

            fig_evol = px.area(
                evol_gwh,
                x='Année_depot',
                y='GWh_reels_annuels',
                title="",
                labels={'GWh_reels_annuels': 'GWh réels/an', 'Année_depot': 'Année'},
                color_discrete_sequence=['#2a9d8f']
            )
            fig_evol.update_layout(height=350)
            st.plotly_chart(fig_evol, width='stretch')

    with col2:
        st.markdown("#### 🏆 Top 5 Opérations par Impact CO₂")
        st.info("💡 Les opérations TRA (transport) sont exclues de ce classement")

        if 'Code équipement' in df.columns and 'GWh_reels_annuels' in df.columns:
            df_sans_tra = df[~df['Code équipement'].astype(str).str.startswith('TRA')].copy()

            if len(df_sans_tra) > 0:
                df_sans_tra['CO2_evite_tonnes_an'] = df_sans_tra[
                                                         'GWh_reels_annuels'] * 1_000_000 * EMISSION_CO2_KWH / 1000

                top_co2 = df_sans_tra.groupby('Code équipement')['CO2_evite_tonnes_an'].sum().reset_index()
                top_co2 = top_co2.nlargest(5, 'CO2_evite_tonnes_an').sort_values('CO2_evite_tonnes_an')

                fig_top_co2 = px.bar(
                    top_co2,
                    x='CO2_evite_tonnes_an',
                    y='Code équipement',
                    orientation='h',
                    title="",
                    labels={'CO2_evite_tonnes_an': 'CO₂ évité (tonnes/an)', 'Code équipement': ''},
                    color='CO2_evite_tonnes_an',
                    color_continuous_scale='Greens',
                    text=top_co2['CO2_evite_tonnes_an'].round(0)
                )
                fig_top_co2.update_traces(texttemplate='%{text:,.0f}', textposition='outside')
                fig_top_co2.update_layout(height=350, showlegend=False)
                st.plotly_chart(fig_top_co2, width='stretch')
            else:
                st.warning("Toutes les opérations sont de type TRA")
        else:
            st.warning("Données insuffisantes pour calculer l'impact CO₂")

    # RÉCAPITULATIF
    st.markdown("---")
    st.subheader("📋 Récapitulatif RSE")

    st.markdown(f"""
    ### 🎯 En résumé

    Grâce aux **{len(df):,}** dossiers CEE analysés :

    - 🌍 **{metrics['gwh_reels']:.2f} GWh** d'énergie économisée chaque année
    - 🌳 **{metrics['co2_evite']:,.0f} tonnes** de CO₂ évitées annuellement
    - 🏠 Équivalent à alimenter **{metrics['foyers_equivalent']:,.0f} foyers** pendant un an
    - 🌲 Impact comparable à **{metrics['arbres_equivalent']:,.0f} arbres** plantés
    - 🚗 Équivaut à retirer **{metrics['voitures_equivalent']:,.0f} voitures** de la circulation
    - 💰 **{metrics['pct_precarite']:.1f}%** des volumes concernent la **précarité énergétique**

    **Contribution significative à la transition énergétique et à la lutte contre le changement climatique ! 🌟**
    """.replace(',', ' '))


def main():
    """Interface principale"""
    # Upload de fichiers
    st.sidebar.header("📁 Chargement des données")

    uploaded_file = st.sidebar.file_uploader(
        "Fichier CEE principal (Excel)",
        type=['xlsx', 'xls'],
        help="Fichier contenant les données CEE (avec feuille Synthese si disponible)"
    )

    # Paramètre SLA
    st.sidebar.markdown("---")
    st.sidebar.subheader("⚙️ Paramètres")
    sla_days = st.sidebar.number_input(
        "Seuil SLA (jours)",
        min_value=30,
        max_value=120,
        value=60,
        step=10,
        help="Délai cible pour la validation des dossiers"
    )

    if uploaded_file is not None:
        with st.spinner('Chargement et traitement des données...'):
            df, df_synthese = load_and_process_data(uploaded_file)

        if df is not None:
            st.success(f"✅ Fichier chargé : {len(df)} dossiers trouvés.")
            if df_synthese is not None:
                st.info(f"✅ Feuille 'Synthese' détectée : {len(df_synthese)} dépôts.")

            # Création des filtres
            filters = create_filters(df)

            # Application des filtres sur le DataFrame principal (détail)
            filtered_df = apply_filters(df, filters)

            # === LOGIQUE DE FILTRAGE CASCADE POUR LA SYNTHÈSE ===
            # Utilisation de la clé de jointure 'Depot_Join' créée dans load_and_process_data
            # pour faire le lien entre les lignes de détails (N° DEPOT) et les lignes de synthèse (Depot)
            filtered_synthese = None
            if df_synthese is not None:
                if 'Depot_Join' in filtered_df.columns and 'Depot' in df_synthese.columns:
                    # On récupère tous les dépôts "racines" présents dans la vue filtrée
                    valid_depots = filtered_df['Depot_Join'].unique()

                    # On ne garde dans la synthèse que les lignes correspondant à ces dépôts racines
                    filtered_synthese = df_synthese[
                        df_synthese['Depot'].astype(str).isin(valid_depots)
                    ].copy()
                else:
                    # Fallback si colonnes manquantes
                    filtered_synthese = df_synthese.copy()
            # ====================================================

            if len(filtered_df) == 0:
                st.warning("⚠️ Aucune donnée ne correspond aux filtres sélectionnés.")
                return

            st.info(f"📊 {len(filtered_df)} dossiers affichés après filtrage")

            # KPIs principaux
            create_kpi_cards(filtered_df, sla_days)
            st.markdown("---")

            # Onglets pour organiser les analyses
            tab1, tab2, tab3, tab4, tab5, tab6 = st.tabs([
                "📈 Évolution Volumes",
                "🔧 Équipements",
                "🏢 Installateurs",
                "🛡️ Contrôle & Conformité",
                "🏆 Qualité ISO 9001",
                "🌱 RSE"
            ])

            with tab1:
                create_volume_evolution_chart(filtered_df)

            with tab2:
                create_equipment_analysis(filtered_df)

            with tab3:
                create_installer_performance(filtered_df)

            with tab4:
                # Utilisation de filtered_synthese au lieu de df_synthese
                create_controle_synthese_analysis(filtered_synthese, filtered_df)

            with tab5:
                # Utilisation de filtered_synthese pour les KPIs Qualité
                create_iso_quality_analysis(filtered_df, filtered_synthese)

            with tab6:
                create_rse_analysis(filtered_df)

            # Données brutes (optionnel)
            with st.expander("📋 Voir les données brutes (échantillon)"):
                st.dataframe(filtered_df.head(100), width="stretch")

    else:
        # Instructions d'utilisation
        st.info("""
        ### 🚀 Bienvenue sur le Dashboard CEE !

        **Instructions :**
        1. **Uploadez** votre fichier Excel CEE principal dans la barre latérale.
        2. **Configurez** les paramètres (SLA, filtres) dans le menu de gauche.
        3. **Explorez** les onglets d'analyse disponibles.

        **Fonctionnalités :**
        - ✨ **KPIs Avancés** : Taux de conversion, volume/dossier, dossiers bloqués, taux SLA
        - 📈 **Évolution optimisée** : Focus sur les volumes (GWh), nouvelles fiches avec volumes générés
        - 🔧 **Équipements enrichi** : Volumes en GWh + camembert de répartition + podium
        - 🏢 **Installateurs AMÉLIORÉ** : Camembert des volumes + Analyse des délais
        - 🛡️ **Contrôle & Conformité** : Filtrage croisé activé (Année, Mandataire impactent la synthèse) via clé racine (P5-14)
        - 🏆 **Qualité ISO 9001** : Taux d'erreur CSV, Volume min par dépôt, Délai d'insertion
        - 🌱 **RSE optimisé** : Impact environnemental

        **Format attendu :**
        - Fichier Excel contenant au moins une feuille de données.
        - Idéalement une feuille "Synthese" pour les KPIs spécifiques.
        """)


if __name__ == "__main__":
    main()