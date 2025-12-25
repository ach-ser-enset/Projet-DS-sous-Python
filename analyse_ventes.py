import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime, timedelta
import warnings
warnings.filterwarnings('ignore')

# Configuration des graphiques
plt.style.use('seaborn-v0_8-darkgrid')
sns.set_palette("husl")
plt.rcParams['figure.figsize'] = (15, 8)
plt.rcParams['font.size'] = 10

# ============================================================================
# ÉTAPE 1: CHARGEMENT ET NETTOYAGE DES DONNÉES
# ============================================================================

def load_and_clean_data(filepath='Sales.xlsx'):
    """
    Charge et nettoie les données de vente depuis un fichier Excel
    
    Args:
        filepath: Chemin vers le fichier Excel
        
    Returns:
        DataFrame pandas nettoyé
    """
    try:
        # Chargement des données
        df = pd.read_excel(filepath)
        print(f"✓ Données chargées: {df.shape[0]} lignes, {df.shape[1]} colonnes")
        
        # Affichage des premières lignes
        print("\nAperçu des données:")
        print(df.head())
        
        # Nettoyage des données
        print("\n--- Nettoyage des données ---")
        
        # 1. Suppression des doublons
        initial_rows = len(df)
        df = df.drop_duplicates()
        print(f"✓ Doublons supprimés: {initial_rows - len(df)}")
        
        # 2. Gestion des valeurs manquantes
        print(f"✓ Valeurs manquantes par colonne:")
        print(df.isnull().sum())
        
        # 3. Conversion des types de données
        if 'OrderDate' in df.columns:
            df['OrderDate'] = pd.to_datetime(df['OrderDate'], errors='coerce')
        if 'Ship Date' in df.columns:
            df['Ship Date'] = pd.to_datetime(df['Ship Date'], errors='coerce')
            
        # 4. Création de colonnes calculées si nécessaire
        if 'Sales' not in df.columns and 'Order Quantity' in df.columns and 'Unit Price' in df.columns:
            df['Sales'] = df['Order Quantity'] * df['Unit Price']
            
        if 'Profit' not in df.columns and 'Sales' in df.columns and 'Cost' in df.columns:
            df['Profit'] = df['Sales'] - df['Cost']
        
        print(f"\n✓ Données nettoyées: {df.shape[0]} lignes restantes")
        return df
        
    except FileNotFoundError:
        print(f"❌ Erreur: Le fichier '{filepath}' n'existe pas")
        # Création de données d'exemple pour la démonstration
        return create_sample_data()
    except Exception as e:
        print(f"❌ Erreur lors du chargement: {str(e)}")
        return create_sample_data()

def create_sample_data():
    """
    Crée un jeu de données d'exemple pour la démonstration
    """
    print("\n⚠️  Création de données d'exemple pour la démonstration...")
    
    np.random.seed(42)
    n_records = 1000
    
    # Génération des dates
    start_date = datetime(2017, 1, 1)
    end_date = datetime(2018, 12, 31)
    dates = pd.date_range(start=start_date, end=end_date, periods=n_records)
    
    # Données d'exemple
    df = pd.DataFrame({
        'OrderNumber': [f'SO-{str(i).zfill(7)}' for i in range(1, n_records + 1)],
        'OrderDate': dates,
        'Ship Date': dates + pd.Timedelta(days=7),
        'Customer Name': np.random.choice([f'Customer_{i}' for i in range(1, 51)], n_records),
        'Index': np.random.randint(1, 51, n_records),
        'Channel': np.random.choice(['Wholesale', 'Distributor', 'Export'], n_records, p=[0.5, 0.3, 0.2]),
        'Currency Code': np.random.choice(['USD', 'EUR', 'GBP', 'AUD', 'NZD'], n_records),
        'Product': np.random.choice([f'Product_{chr(65+i)}' for i in range(10)], n_records),
        'City': np.random.choice(['New York', 'London', 'Paris', 'Sydney', 'Tokyo', 
                                  'Berlin', 'Toronto', 'Auckland', 'Singapore', 'Dubai'], n_records),
        'Order Quantity': np.random.randint(1, 100, n_records),
        'Unit Price': np.random.uniform(10, 500, n_records),
        'Cost': np.random.uniform(5, 300, n_records),
    })
    
    # Calculs
    df['Sales'] = df['Order Quantity'] * df['Unit Price']
    df['Profit'] = df['Sales'] - (df['Order Quantity'] * df['Cost'])
    
    print(f"✓ Données d'exemple créées: {df.shape[0]} lignes")
    return df

# ============================================================================
# ÉTAPE 2: CRÉATION DE LA TABLE DE DATES
# ============================================================================

def create_date_table(df, date_column='OrderDate'):
    """
    Crée une table de dates complète pour l'analyse temporelle
    
    Args:
        df: DataFrame source
        date_column: Nom de la colonne de date
        
    Returns:
        DataFrame de dates avec tous les attributs temporels
    """
    print("\n--- Création de la table de dates ---")
    
    # Extraction des dates min et max
    min_date = df[date_column].min()
    max_date = df[date_column].max()
    
    # Création de la plage de dates complète
    date_range = pd.date_range(start=min_date, end=max_date, freq='D')
    
    # Construction de la table de dates
    date_table = pd.DataFrame({
        'Date': date_range,
        'Year': date_range.year,
        'Quarter': date_range.quarter,
        'Month': date_range.month,
        'MonthName': date_range.strftime('%B'),
        'MonthNameShort': date_range.strftime('%b'),
        'Week': date_range.isocalendar().week,
        'Day': date_range.day,
        'DayOfWeek': date_range.dayofweek,
        'DayName': date_range.strftime('%A'),
        'DayNameShort': date_range.strftime('%a'),
        'IsWeekend': (date_range.dayofweek >= 5).astype(int),
        'YearMonth': date_range.to_period('M').astype(str),
        'YearQuarter': date_range.to_period('Q').astype(str)
    })
    
    print(f"✓ Table de dates créée: {len(date_table)} jours de {min_date.date()} à {max_date.date()}")
    return date_table

# ============================================================================
# ÉTAPE 3: CALCUL DES MESURES ET INDICATEURS (KPI)
# ============================================================================

def calculate_kpis(df, date_table, current_year=2018, previous_year=2017):
    """
    Calcule tous les KPI requis pour l'analyse
    Conversion des formules DAX en Python/Pandas
    
    Args:
        df: DataFrame des ventes
        date_table: Table de dates
        current_year: Année courante pour l'analyse
        previous_year: Année précédente pour la comparaison
        
    Returns:
        Dictionary contenant tous les KPI
    """
    print("\n--- Calcul des indicateurs clés (KPI) ---")
    
    # Filtrage par année
    df_current = df[df['OrderDate'].dt.year == current_year]
    df_previous = df[df['OrderDate'].dt.year == previous_year]
    
    # Création du dictionnaire des KPI
    kpis = {}
    
    # 1. Total Sales
    kpis['Total Sales'] = df_current['Sales'].sum()
    
    # 2. Total Sales PY (Previous Year)
    kpis['Total Sales PY'] = df_previous['Sales'].sum()
    
    # 3. Total Sales/PY Var (Variation)
    kpis['Total Sales/PY Var'] = kpis['Total Sales'] - kpis['Total Sales PY']
    
    # 4. Total Sales/PY Var % (Variation en pourcentage)
    kpis['Total Sales/PY Var %'] = (kpis['Total Sales/PY Var'] / kpis['Total Sales PY'] * 100) if kpis['Total Sales PY'] != 0 else 0
    
    # 5. Total Order Quantity
    kpis['Total Order Quantity'] = df_current['Order Quantity'].sum()
    
    # 6. Total Profit
    kpis['Total Profit'] = df_current['Profit'].sum()
    
    # 7. Total Profit PY
    kpis['Total Profit PY'] = df_previous['Profit'].sum()
    
    # 8. Total Profit/PY Var
    kpis['Total Profit/PY Var'] = kpis['Total Profit'] - kpis['Total Profit PY']
    
    # 9. Total Profit/PY Var %
    kpis['Total Profit/PY Var %'] = (kpis['Total Profit/PY Var'] / kpis['Total Profit PY'] * 100) if kpis['Total Profit PY'] != 0 else 0
    
    # 10. Profit Margin %
    kpis['Profit Margin %'] = (kpis['Total Profit'] / kpis['Total Sales'] * 100) if kpis['Total Sales'] != 0 else 0
    
    # 11. Total Cost
    kpis['Total Cost'] = (df_current['Order Quantity'] * df_current['Cost']).sum()
    
    # 12. Total Order Quantity/PY
    kpis['Total Order Quantity/PY'] = df_previous['Order Quantity'].sum()
    
    # 13. Total Order Quantity/PY Var
    kpis['Total Order Quantity/PY Var'] = kpis['Total Order Quantity'] - kpis['Total Order Quantity/PY']
    
    # 14. Total Order Quantity/PY Var %
    kpis['Total Order Quantity/PY Var %'] = (kpis['Total Order Quantity/PY Var'] / kpis['Total Order Quantity/PY'] * 100) if kpis['Total Order Quantity/PY'] != 0 else 0
    
    # Affichage des KPI
    print("\n📊 Indicateurs clés de performance:")
    print("="*60)
    for key, value in kpis.items():
        if '%' in key:
            print(f"{key:35} : {value:>15.2f}%")
        else:
            print(f"{key:35} : {value:>15,.2f}")
    
    return kpis, df_current, df_previous

# ============================================================================
# ÉTAPE 4: ANALYSES COMPARATIVES
# ============================================================================

def analyze_sales_by_product(df_current, df_previous):
    """
    Analyse des ventes par produit avec comparaison année précédente
    """
    sales_by_product_current = df_current.groupby('Product')['Sales'].sum().sort_values(ascending=False)
    sales_by_product_previous = df_previous.groupby('Product')['Sales'].sum()
    
    comparison = pd.DataFrame({
        'Current Year': sales_by_product_current,
        'Previous Year': sales_by_product_previous
    }).fillna(0)
    
    comparison['Variation'] = comparison['Current Year'] - comparison['Previous Year']
    comparison['Variation %'] = (comparison['Variation'] / comparison['Previous Year'] * 100).replace([np.inf, -np.inf], 0)
    
    return comparison

def analyze_sales_by_month(df_current, df_previous):
    """
    Analyse des ventes par mois avec comparaison année précédente
    """
    df_current['Month'] = df_current['OrderDate'].dt.month
    df_previous['Month'] = df_previous['OrderDate'].dt.month
    
    sales_by_month_current = df_current.groupby('Month')['Sales'].sum()
    sales_by_month_previous = df_previous.groupby('Month')['Sales'].sum()
    
    comparison = pd.DataFrame({
        'Current Year': sales_by_month_current,
        'Previous Year': sales_by_month_previous
    }).fillna(0)
    
    comparison['Variation'] = comparison['Current Year'] - comparison['Previous Year']
    comparison['Variation %'] = (comparison['Variation'] / comparison['Previous Year'] * 100).replace([np.inf, -np.inf], 0)
    
    return comparison

def analyze_top_cities(df_current, top_n=5):
    """
    Analyse des top N villes par ventes
    """
    top_cities = df_current.groupby('City')['Sales'].sum().sort_values(ascending=False).head(top_n)
    return top_cities

def analyze_profit_by_channel(df_current, df_previous):
    """
    Analyse du profit par canal de vente avec comparaison
    """
    profit_by_channel_current = df_current.groupby('Channel')['Profit'].sum()
    profit_by_channel_previous = df_previous.groupby('Channel')['Profit'].sum()
    
    comparison = pd.DataFrame({
        'Current Year': profit_by_channel_current,
        'Previous Year': profit_by_channel_previous
    }).fillna(0)
    
    comparison['Variation'] = comparison['Current Year'] - comparison['Previous Year']
    comparison['Variation %'] = (comparison['Variation'] / comparison['Previous Year'] * 100).replace([np.inf, -np.inf], 0)
    
    return comparison

def analyze_customers(df_current, df_previous, top_n=5, bottom_n=5):
    """
    Analyse des top et last N clients
    """
    # Top clients
    top_customers_current = df_current.groupby('Customer Name')['Sales'].sum().sort_values(ascending=False).head(top_n)
    top_customers_previous = df_previous.groupby('Customer Name')['Sales'].sum()
    
    top_comparison = pd.DataFrame({
        'Current Year': top_customers_current,
        'Previous Year': [top_customers_previous.get(c, 0) for c in top_customers_current.index]
    })
    
    # Bottom clients
    bottom_customers_current = df_current.groupby('Customer Name')['Sales'].sum().sort_values(ascending=True).head(bottom_n)
    bottom_customers_previous = df_previous.groupby('Customer Name')['Sales'].sum()
    
    bottom_comparison = pd.DataFrame({
        'Current Year': bottom_customers_current,
        'Previous Year': [bottom_customers_previous.get(c, 0) for c in bottom_customers_current.index]
    })
    
    return top_comparison, bottom_comparison

# ============================================================================
# ÉTAPE 5: VISUALISATIONS
# ============================================================================

def create_kpi_cards(kpis):
    """
    Crée des cartes KPI visuelles
    """
    fig, axes = plt.subplots(2, 2, figsize=(16, 10))
    fig.suptitle('📊 Indicateurs Clés de Performance (KPI)', fontsize=20, fontweight='bold', y=1.02)
    
    # Card 1: Total Sales
    ax1 = axes[0, 0]
    ax1.text(0.5, 0.7, f"${kpis['Total Sales']:,.0f}", 
             ha='center', va='center', fontsize=36, fontweight='bold', color='#2E86AB')
    ax1.text(0.5, 0.4, 'Ventes Totales', ha='center', va='center', fontsize=16, color='gray')
    ax1.text(0.5, 0.2, f"vs Année Précédente: {kpis['Total Sales/PY Var %']:+.1f}%", 
             ha='center', va='center', fontsize=12, 
             color='green' if kpis['Total Sales/PY Var %'] > 0 else 'red')
    ax1.axis('off')
    ax1.set_facecolor('#f8f9fa')
    
    # Card 2: Total Profit
    ax2 = axes[0, 1]
    ax2.text(0.5, 0.7, f"${kpis['Total Profit']:,.0f}", 
             ha='center', va='center', fontsize=36, fontweight='bold', color='#06A77D')
    ax2.text(0.5, 0.4, 'Profit Total', ha='center', va='center', fontsize=16, color='gray')
    ax2.text(0.5, 0.2, f"vs Année Précédente: {kpis['Total Profit/PY Var %']:+.1f}%", 
             ha='center', va='center', fontsize=12,
             color='green' if kpis['Total Profit/PY Var %'] > 0 else 'red')
    ax2.axis('off')
    ax2.set_facecolor('#f8f9fa')
    
    # Card 3: Profit Margin
    ax3 = axes[1, 0]
    ax3.text(0.5, 0.7, f"{kpis['Profit Margin %']:.1f}%", 
             ha='center', va='center', fontsize=36, fontweight='bold', color='#F18F01')
    ax3.text(0.5, 0.4, 'Marge Bénéficiaire', ha='center', va='center', fontsize=16, color='gray')
    ax3.axis('off')
    ax3.set_facecolor('#f8f9fa')
    
    # Card 4: Total Orders
    ax4 = axes[1, 1]
    ax4.text(0.5, 0.7, f"{kpis['Total Order Quantity']:,.0f}", 
             ha='center', va='center', fontsize=36, fontweight='bold', color='#C73E1D')
    ax4.text(0.5, 0.4, 'Commandes Totales', ha='center', va='center', fontsize=16, color='gray')
    ax4.text(0.5, 0.2, f"vs Année Précédente: {kpis['Total Order Quantity/PY Var %']:+.1f}%", 
             ha='center', va='center', fontsize=12,
             color='green' if kpis['Total Order Quantity/PY Var %'] > 0 else 'red')
    ax4.axis('off')
    ax4.set_facecolor('#f8f9fa')
    
    plt.tight_layout()
    return fig

def visualize_sales_by_product(comparison):
    """
    Visualisation des ventes par produit
    """
    fig, ax = plt.subplots(figsize=(14, 8))
    
    x = np.arange(len(comparison))
    width = 0.35
    
    bars1 = ax.bar(x - width/2, comparison['Current Year'], width, label='Année Courante', color='#2E86AB')
    bars2 = ax.bar(x + width/2, comparison['Previous Year'], width, label='Année Précédente', color='#A23B72')
    
    ax.set_xlabel('Produits', fontsize=12, fontweight='bold')
    ax.set_ylabel('Ventes ($)', fontsize=12, fontweight='bold')
    ax.set_title('📦 Ventes par Produit - Comparaison Annuelle', fontsize=16, fontweight='bold', pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(comparison.index, rotation=45, ha='right')
    ax.legend()
    ax.grid(axis='y', alpha=0.3)
    
    # Ajout des valeurs sur les barres
    for bar in bars1:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height,
                f'${height:,.0f}', ha='center', va='bottom', fontsize=8)
    
    plt.tight_layout()
    return fig

def visualize_sales_by_month(comparison):
    """
    Visualisation des ventes par mois
    """
    fig, ax = plt.subplots(figsize=(14, 8))
    
    months = ['Jan', 'Fév', 'Mar', 'Avr', 'Mai', 'Jun', 
              'Jul', 'Aoû', 'Sep', 'Oct', 'Nov', 'Déc']
    x = comparison.index
    
    ax.plot(x, comparison['Current Year'], marker='o', linewidth=2.5, 
            label='Année Courante', color='#2E86AB', markersize=8)
    ax.plot(x, comparison['Previous Year'], marker='s', linewidth=2.5, 
            label='Année Précédente', color='#F18F01', markersize=8)
    
    ax.set_xlabel('Mois', fontsize=12, fontweight='bold')
    ax.set_ylabel('Ventes ($)', fontsize=12, fontweight='bold')
    ax.set_title('📅 Ventes Mensuelles - Comparaison Annuelle', fontsize=16, fontweight='bold', pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels([months[i-1] for i in x])
    ax.legend(fontsize=11)
    ax.grid(True, alpha=0.3)
    
    plt.tight_layout()
    return fig

def visualize_top_cities(top_cities):
    """
    Visualisation des top 5 villes
    """
    fig, ax = plt.subplots(figsize=(12, 7))
    
    colors = plt.cm.viridis(np.linspace(0.3, 0.9, len(top_cities)))
    bars = ax.barh(range(len(top_cities)), top_cities.values, color=colors)
    
    ax.set_yticks(range(len(top_cities)))
    ax.set_yticklabels(top_cities.index)
    ax.set_xlabel('Ventes ($)', fontsize=12, fontweight='bold')
    ax.set_title('🏙️ Top 5 des Villes par Ventes', fontsize=16, fontweight='bold', pad=20)
    ax.grid(axis='x', alpha=0.3)
    
    # Ajout des valeurs
    for i, (bar, value) in enumerate(zip(bars, top_cities.values)):
        ax.text(value, i, f' ${value:,.0f}', va='center', fontsize=10, fontweight='bold')
    
    plt.tight_layout()
    return fig

def visualize_profit_by_channel(comparison):
    """
    Visualisation du profit par canal
    """
    fig, ax = plt.subplots(figsize=(12, 7))
    
    x = np.arange(len(comparison))
    width = 0.35
    
    bars1 = ax.bar(x - width/2, comparison['Current Year'], width, label='Année Courante', color='#06A77D')
    bars2 = ax.bar(x + width/2, comparison['Previous Year'], width, label='Année Précédente', color='#D81159')
    
    ax.set_xlabel('Canal de Vente', fontsize=12, fontweight='bold')
    ax.set_ylabel('Profit ($)', fontsize=12, fontweight='bold')
    ax.set_title('📡 Profit par Canal de Vente - Comparaison Annuelle', fontsize=16, fontweight='bold', pad=20)
    ax.set_xticks(x)
    ax.set_xticklabels(comparison.index)
    ax.legend()
    ax.grid(axis='y', alpha=0.3)
    
    # Ajout des valeurs
    for bar in bars1:
        height = bar.get_height()
        ax.text(bar.get_x() + bar.get_width()/2., height,
                f'${height:,.0f}', ha='center', va='bottom', fontsize=9)
    
    plt.tight_layout()
    return fig

def visualize_customers(top_comparison, bottom_comparison):
    """
    Visualisation des top et last 5 clients
    """
    fig, (ax1, ax2) = plt.subplots(1, 2, figsize=(16, 7))
    
    # Top 5 clients
    x1 = np.arange(len(top_comparison))
    width = 0.35
    ax1.barh(x1 - width/2, top_comparison['Current Year'], width, label='Année Courante', color='#2E86AB')
    ax1.barh(x1 + width/2, top_comparison['Previous Year'], width, label='Année Précédente', color='#F18F01')
    ax1.set_yticks(x1)
    ax1.set_yticklabels(top_comparison.index)
    ax1.set_xlabel('Ventes ($)', fontsize=11, fontweight='bold')
    ax1.set_title('🏆 Top 5 Clients', fontsize=14, fontweight='bold')
    ax1.legend()
    ax1.grid(axis='x', alpha=0.3)
    
    # Last 5 clients
    x2 = np.arange(len(bottom_comparison))
    ax2.barh(x2 - width/2, bottom_comparison['Current Year'], width, label='Année Courante', color='#C73E1D')
    ax2.barh(x2 + width/2, bottom_comparison['Previous Year'], width, label='Année Précédente', color='#7209B7')
    ax2.set_yticks(x2)
    ax2.set_yticklabels(bottom_comparison.index)
    ax2.set_xlabel('Ventes ($)', fontsize=11, fontweight='bold')
    ax2.set_title('📉 Bottom 5 Clients', fontsize=14, fontweight='bold')
    ax2.legend()
    ax2.grid(axis='x', alpha=0.3)
    
    plt.tight_layout()
    return fig

# ============================================================================
# FONCTION PRINCIPALE
# ============================================================================

def main():
    """
    Fonction principale pour exécuter l'analyse complète
    """
    print("="*80)
    print("🎯 ANALYSE ET VISUALISATION DES DONNÉES DE VENTE")
    print("="*80)
    
    # Étape 1: Chargement et nettoyage
    df = load_and_clean_data('Sales.xlsx')
    
    # Étape 2: Création de la table de dates
    date_table = create_date_table(df)
    
    # Étape 3: Calcul des KPI
    kpis, df_current, df_previous = calculate_kpis(df, date_table)
    
    # Étape 4: Analyses comparatives
    print("\n--- Analyses comparatives ---")
    sales_by_product = analyze_sales_by_product(df_current, df_previous)
    sales_by_month = analyze_sales_by_month(df_current, df_previous)
    top_cities = analyze_top_cities(df_current)
    profit_by_channel = analyze_profit_by_channel(df_current, df_previous)
    top_customers, bottom_customers = analyze_customers(df_current, df_previous)
    
    print("✓ Analyses terminées")
    
    # Étape 5: Création des visualisations
    print("\n--- Création des visualisations ---")
    
    fig1 = create_kpi_cards(kpis)
    fig2 = visualize_sales_by_product(sales_by_product)
    fig3 = visualize_sales_by_month(sales_by_month)
    fig4 = visualize_top_cities(top_cities)
    fig5 = visualize_profit_by_channel(profit_by_channel)
    fig6 = visualize_customers(top_customers, bottom_customers)
    
    plt.show()
    
    print("\n✓ Visualisations créées avec succès!")
    print("\n" + "="*80)
    print("✅ ANALYSE TERMINÉE")
    print("="*80)
    
    return df, kpis, date_table

# Exécution
if __name__ == "__main__":
    df, kpis, date_table = main()
