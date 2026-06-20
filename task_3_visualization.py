import os
import matplotlib.pyplot as plt
import seaborn as sns
import pandas as pd

def generate_visualizations(df):
    """
    Task 3: Data Visualization
    Uses Matplotlib and Seaborn to visualize data trends.
    """
    print("Generating visualizations with Matplotlib & Seaborn...")
    
    # Ensure there is an output directory for charts
    assets_dir = "charts_output"
    os.makedirs(assets_dir, exist_ok=True)
    
    # Set visual style for matplotlib/seaborn
    sns.set_theme(style="whitegrid")
    
    # Chart 1: Price Distribution
    plt.figure(figsize=(10, 6))
    sns.histplot(data=df, x='Price', bins=10, kde=True, color='skyblue')
    plt.title('Distribution of Book Prices', fontsize=14, fontweight='bold')
    plt.xlabel('Price (£)', fontsize=12)
    plt.ylabel('Frequency', fontsize=12)
    plt.tight_layout()
    chart1_path = os.path.join(assets_dir, 'chart_price_dist.png')
    plt.savefig(chart1_path, dpi=300)
    plt.close()
    print(f"Saved: {chart1_path}")
    
    # Chart 2: Average Price by Rating
    plt.figure(figsize=(8, 6))
    avg_price_df = df.groupby('Numeric Rating')['Price'].mean().reset_index()
    sns.barplot(data=avg_price_df, x='Numeric Rating', y='Price', hue='Numeric Rating', palette='viridis', legend=False)
    plt.title('Average Book Price by Star Rating', fontsize=14, fontweight='bold')
    plt.xlabel('Star Rating', fontsize=12)
    plt.ylabel('Average Price (£)', fontsize=12)
    plt.tight_layout()
    chart2_path = os.path.join(assets_dir, 'chart_avg_price_rating.png')
    plt.savefig(chart2_path, dpi=300)
    plt.close()
    print(f"Saved: {chart2_path}")

if __name__ == "__main__":
    # Create some mock data to test the visualization script standalone
    mock_data = {
        'Title': ['Book A', 'Book B', 'Book C', 'Book D', 'Book E', 'Book F'],
        'Numeric Rating': [1, 2, 3, 4, 5, 3],
        'Price': [10.5, 20.0, 15.0, 45.0, 50.0, 25.0]
    }
    df_mock = pd.DataFrame(mock_data)
    generate_visualizations(df_mock)
