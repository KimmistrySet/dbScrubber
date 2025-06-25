import pandas as pd

def load_data(file_path):
    """Load Excel file into a DataFrame."""
    return pd.read_excel(file_path)

def remove_duplicates(df):
    """Remove duplicate rows from the DataFrame."""
    return df.drop_duplicates()

def fill_missing_values(df):
    """Fill missing values with mean for numerical columns and mode for categorical columns."""
    for column in df.columns:
        if df[column].dtype == 'object':
            df[column].fillna(df[column].mode()[0], inplace=True) # Categorical
        else: 
            df[column].fillna(df[column].mean(), inplace=True) # Numerical
        return df

def correct_inconsistencies(df):
    """Correct inconsistencies, e.g., standardizing text data."""
    # Example: standardizing case for a specific column
    if 'Name' in df.columns:
        df['Name'] = df['Name'].str.title() # Capitalize names
    return df

def clean_data(file_path, output_path):
    """Main function to clean the data."""
    # Load Excel file data
    df = load_data(file_path)

    # Clean data
    df = remove_duplicates(df)
    df = fill_missing_values(df)
    df = correct_inconsistencies(df)

    # Save the cleaned data to a new Excel file
    df.to_excel(output_path, index=False)
    print(f"Cleaned data saved to {output_path}")

# Example usage
if __name__ == "__main__":
    input_file = 'input_data.xlsx' # Specify your input file here
    output_file = 'cleaned_data.xlsx' # Specify your output file here
    clean_data(input_file, output_file)

    