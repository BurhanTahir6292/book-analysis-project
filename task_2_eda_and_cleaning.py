import pandas as pd

def clean_and_prepare_data(file_path="Task_1_Web_Scraping.xlsx"):
    """
    Task 2: Exploratory Data Analysis (EDA) and Data Cleaning
    Loads data and formats it into a meaningful table/context.
    """
    print(f"Loading data from {file_path}...")
    try:
        # Load existing data from the web scraping output
        df = pd.read_excel(file_path, sheet_name="Task 1 - Web Scraping", skiprows=6)
        
        # Clean up column names if needed
        if 'Title' not in df.columns:
            df.rename(columns={
                df.columns[1]: 'Title', 
                df.columns[3]: 'Numeric Rating', 
                df.columns[4]: 'Price'
            }, inplace=True)
            
    except Exception as e:
        print(f"Error loading Excel file: {e}")
        print("Fallback to simulated data for demonstration...")
        # Fallback raw data structure
        mock_data = [
            {'title': 'A Light in the Attic', 'rating_num': 3, 'price': 51.77},
            {'title': 'Tipping the Velvet', 'rating_num': 1, 'price': 53.74},
            {'title': 'Soumission', 'rating_num': 1, 'price': 50.10},
            {'title': 'Sharp Objects', 'rating_num': 4, 'price': 47.82}
        ]
        df = pd.DataFrame(mock_data)
        df.rename(columns={
            'title': 'Title', 
            'rating_num': 'Numeric Rating', 
            'price': 'Price'
        }, inplace=True)

    # Keep relevant columns and drop rows with missing data
    df = df[['Title', 'Numeric Rating', 'Price']].dropna()
    
    # Ensure proper data types for numerical analysis
    df['Price'] = df['Price'].astype(float)
    df['Numeric Rating'] = df['Numeric Rating'].astype(int)
    
    # Basic EDA stats
    print("\n--- Basic Statistics ---")
    print(df.describe())
    
    print("\n--- Top 5 Rows ---")
    print(df.head())
    
    return df

if __name__ == "__main__":
    cleaned_df = clean_and_prepare_data()
