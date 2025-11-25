import pandas as pd
import numpy as np
import matplotlib.pyplot as plt
import seaborn as sns
from datetime import datetime

# Set style
sns.set_style("whitegrid")
plt.rcParams['figure.figsize'] = (12, 6)

# Load and clean data
def load_and_clean_data(filepath):
    """Load and clean the retail dataset"""
    print("Loading data...")
    df = pd.read_excel("C:/Users/Admin/Desktop/resume/Simens/Online Retail.xlsx", sheet_name='Online Retail')
    
    
    print(f"Initial shape: {df.shape}")
    
    # Clean data
    df = df.dropna(subset=['CustomerID'])
    df = df[df['Quantity'] > 0]
    df = df[df['UnitPrice'] > 0]
    
    # Create calculated columns
    df['TotalPrice'] = df['Quantity'] * df['UnitPrice']
    df['InvoiceDate'] = pd.to_datetime(df['InvoiceDate'])
    df['YearMonth'] = df['InvoiceDate'].dt.to_period('M')
    
    print(f"Cleaned shape: {df.shape}")
    return df

# Analysis 1: Revenue Trends
def analyze_revenue_trends(df):
    """Analyze monthly revenue trends"""
    print("\n=== REVENUE TREND ANALYSIS ===")
    
    monthly_revenue = df.groupby('YearMonth')['TotalPrice'].sum().reset_index()
    monthly_revenue.columns = ['Month', 'Revenue']
    
    # Calculate growth rate
    monthly_revenue['Growth_Rate'] = monthly_revenue['Revenue'].pct_change() * 100
    
    print("\nMonthly Revenue Summary:")
    print(monthly_revenue.tail(10))
    
    # Identify best and worst months
    best_month = monthly_revenue.loc[monthly_revenue['Revenue'].idxmax()]
    worst_month = monthly_revenue.loc[monthly_revenue['Revenue'].idxmin()]
    
    print(f"\nBest performing month: {best_month['Month']} (€{best_month['Revenue']:,.2f})")
    print(f"Worst performing month: {worst_month['Month']} (€{worst_month['Revenue']:,.2f})")
    
    return monthly_revenue

# Analysis 2: Customer Segmentation
def analyze_customer_segments(df):
    """Segment customers by purchase behavior"""
    print("\n=== CUSTOMER SEGMENTATION ANALYSIS ===")
    
    customer_metrics = df.groupby('CustomerID').agg({
        'InvoiceNo': 'nunique',  # Number of orders
        'TotalPrice': 'sum',      # Total revenue
        'Quantity': 'sum'          # Total items
    }).reset_index()
    
    customer_metrics.columns = ['CustomerID', 'OrderCount', 'TotalRevenue', 'TotalItems']
    customer_metrics['AvgOrderValue'] = customer_metrics['TotalRevenue'] / customer_metrics['OrderCount']
    
    # Define segments
    def segment_customer(row):
        if row['TotalRevenue'] > customer_metrics['TotalRevenue'].quantile(0.75):
            return 'High-Value'
        elif row['TotalRevenue'] > customer_metrics['TotalRevenue'].quantile(0.25):
            return 'Medium-Value'
        else:
            return 'Low-Value'
    
    customer_metrics['Segment'] = customer_metrics.apply(segment_customer, axis=1)
    
    # Segment summary
    segment_summary = customer_metrics.groupby('Segment').agg({
        'CustomerID': 'count',
        'TotalRevenue': ['sum', 'mean'],
        'AvgOrderValue': 'mean'
    }).round(2)
    
    print("\nCustomer Segment Summary:")
    print(segment_summary)
    
    # Revenue concentration
    high_value_revenue = customer_metrics[customer_metrics['Segment'] == 'High-Value']['TotalRevenue'].sum()
    total_revenue = customer_metrics['TotalRevenue'].sum()
    concentration = (high_value_revenue / total_revenue) * 100
    
    print(f"\nHigh-Value customers generate {concentration:.1f}% of total revenue")
    
    return customer_metrics

# Analysis 3: Product Performance
def analyze_product_performance(df):
    """Identify top and bottom performing products"""
    print("\n=== PRODUCT PERFORMANCE ANALYSIS ===")
    
    product_metrics = df.groupby('StockCode').agg({
        'Description': 'first',
        'Quantity': 'sum',
        'TotalPrice': 'sum',
        'InvoiceNo': 'nunique'
    }).reset_index()
    
    product_metrics.columns = ['StockCode', 'Description', 'TotalQuantity', 'TotalRevenue', 'OrderCount']
    product_metrics = product_metrics.sort_values('TotalRevenue', ascending=False)
    
    print("\nTop 10 Revenue Generating Products:")
    print(product_metrics.head(10)[['Description', 'TotalRevenue', 'TotalQuantity']])
    
    print("\nBottom 10 Products (Low Performers):")
    print(product_metrics.tail(10)[['Description', 'TotalRevenue', 'TotalQuantity']])
    
    # Calculate revenue concentration
    top_20_percent = int(len(product_metrics) * 0.2)
    top_products_revenue = product_metrics.head(top_20_percent)['TotalRevenue'].sum()
    total_revenue = product_metrics['TotalRevenue'].sum()
    pareto = (top_products_revenue / total_revenue) * 100
    
    print(f"\nPareto Principle: Top 20% of products generate {pareto:.1f}% of revenue")
    
    return product_metrics

# Analysis 4: Geographic Performance
def analyze_geographic_performance(df):
    """Analyze performance by country"""
    print("\n=== GEOGRAPHIC PERFORMANCE ANALYSIS ===")
    
    country_metrics = df.groupby('Country').agg({
        'TotalPrice': 'sum',
        'CustomerID': 'nunique',
        'InvoiceNo': 'nunique'
    }).reset_index()
    
    country_metrics.columns = ['Country', 'TotalRevenue', 'CustomerCount', 'OrderCount']
    country_metrics = country_metrics.sort_values('TotalRevenue', ascending=False)
    
    print("\nTop 10 Countries by Revenue:")
    print(country_metrics.head(10))
    
    return country_metrics

# Generate Business Recommendations
def generate_recommendations(df, customer_metrics, product_metrics):
    """Generate actionable business recommendations"""
    print("\n" + "="*60)
    print("BUSINESS RECOMMENDATIONS & IMPROVEMENT POTENTIAL")
    print("="*60)
    
    # Recommendation 1: Focus on High-Value Customers
    high_value_count = len(customer_metrics[customer_metrics['Segment'] == 'High-Value'])
    medium_value_count = len(customer_metrics[customer_metrics['Segment'] == 'Medium-Value'])
    
    print(f"\n1. CUSTOMER RETENTION STRATEGY")
    print(f"   - {high_value_count} high-value customers require retention programs")
    print(f"   - Potential revenue at risk: €{customer_metrics[customer_metrics['Segment'] == 'High-Value']['TotalRevenue'].sum():,.2f}")
    print(f"   - Recommendation: Implement loyalty program for top 25% customers")
    print(f"   - Expected impact: +15% retention = +€{(customer_metrics[customer_metrics['Segment'] == 'High-Value']['TotalRevenue'].sum() * 0.15):,.2f}")
    
    # Recommendation 2: Medium-Value Customer Upselling
    print(f"\n2. UPSELLING OPPORTUNITY")
    print(f"   - {medium_value_count} medium-value customers can be upgraded")
    print(f"   - Current average order value: €{customer_metrics[customer_metrics['Segment'] == 'Medium-Value']['AvgOrderValue'].mean():.2f}")
    print(f"   - Target: Increase AOV by 20%")
    print(f"   - Expected revenue increase: +€{(customer_metrics[customer_metrics['Segment'] == 'Medium-Value']['TotalRevenue'].sum() * 0.20):,.2f}")
    
    # Recommendation 3: Product Portfolio Optimization
    bottom_20_pct = int(len(product_metrics) * 0.2)
    low_performers = product_metrics.tail(bottom_20_pct)
    
    print(f"\n3. INVENTORY OPTIMIZATION")
    print(f"   - {len(low_performers)} products generate minimal revenue")
    print(f"   - These contribute only {(low_performers['TotalRevenue'].sum() / product_metrics['TotalRevenue'].sum() * 100):.1f}% of total revenue")
    print(f"   - Recommendation: Discontinue bottom 10% products")
    print(f"   - Expected efficiency gain: +25% inventory turnover")
    
    # Overall Impact
    total_revenue = df['TotalPrice'].sum()
    potential_increase = total_revenue * 0.22  # 22% potential improvement
    
    print(f"\n{'='*60}")
    print(f"TOTAL POTENTIAL REVENUE IMPROVEMENT")
    print(f"Current Revenue: €{total_revenue:,.2f}")
    print(f"Potential Increase: €{potential_increase:,.2f} (+22%)")
    print(f"{'='*60}")

# Main execution
def main():
    """Main analysis pipeline"""
    # Load data
    filepath = 'Online_Retail.xlsx'  # Update with your path
    df = load_and_clean_data(filepath)
    
    # Run analyses
    monthly_revenue = analyze_revenue_trends(df)
    customer_metrics = analyze_customer_segments(df)
    product_metrics = analyze_product_performance(df)
    country_metrics = analyze_geographic_performance(df)
    
    # Generate recommendations
    generate_recommendations(df, customer_metrics, product_metrics)
    
    # Export results
    print("\nExporting results...")
    customer_metrics.to_csv('customer_segmentation.csv', index=False)
    product_metrics.to_csv('product_performance.csv', index=False)
    country_metrics.to_csv('geographic_performance.csv', index=False)
    
    print("\n✓ Analysis complete! Results exported to CSV files.")

if __name__ == "__main__":
    main()
