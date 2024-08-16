import pandas as pd
import re


def read_excel(file_path: str) -> pd.DataFrame:
    """
    Reads an Excel file and returns a pandas DataFrame.

    Args:
    - file_path: str : Path to the Excel file.

    Returns:
    - pd.DataFrame : DataFrame with the Excel data.
    """
    return pd.read_excel(file_path)


def clean_column(df: pd.DataFrame, column_name: str) -> pd.DataFrame:
    """
    Cleans invisible characters from a specific column in the DataFrame.

    Args:
    - df: pd.DataFrame : The input DataFrame.
    - column_name: str : The name of the column to be cleaned.

    Returns:
    - pd.DataFrame : The DataFrame with the specified column cleaned.
    """
    if column_name in df.columns:
        df[column_name] = df[column_name].apply(remove_invisible_chars)
    else:
        print(f"Warning: Column '{column_name}' not found in the Excel file.")
    return df


def clean_columns(df: pd.DataFrame, columns: list) -> pd.DataFrame:
    """
    Cleans invisible characters from multiple specified columns in the DataFrame.

    Args:
    - df: pd.DataFrame : The input DataFrame.
    - columns: list : List of column names to clean.

    Returns:
    - pd.DataFrame : The DataFrame with the specified columns cleaned.
    """
    for column in columns:
        df = clean_column(df, column)
    return df


def remove_invisible_chars(value: str) -> str:
    """
    Removes invisible characters such as zero-width spaces from the input string.

    Args:
    - value: str : The string to be cleaned.

    Returns:
    - str : The cleaned string with invisible characters removed.
    """
    invisible_chars = re.compile(r'[\u200B-\u200D\uFEFF]')  # Regex for invisible chars
    return invisible_chars.sub('', value) if isinstance(value, str) else value


def save_excel(df: pd.DataFrame, output_file: str) -> None:
    """
    Saves a DataFrame to an Excel file.

    Args:
    - df: pd.DataFrame : The DataFrame to be saved.
    - output_file: str : The path where the Excel file will be saved.
    """
    df.to_excel(output_file, index=False)
    print(f"Cleaned file saved as {output_file}")


def process_excel_file(file_path: str, columns: list, output_file: str) -> None:
    """
    Orchestrates the process of reading an Excel file, cleaning specified columns,
    and saving the cleaned DataFrame to a new file.

    Args:
    - file_path: str : Path to the Excel file.
    - columns: list : List of column names to clean.
    - output_file: str : The path where the cleaned Excel file will be saved.
    """
    # Read Excel file
    df = read_excel(file_path)

    # Clean specified columns
    df = clean_columns(df, columns)

    # Save cleaned DataFrame to a new Excel file
    save_excel(df, output_file)


# Example usage
if __name__ == "__main__":
    file_path = 'input_file.xlsx'  # Path to the Excel file to be processed
    columns_to_clean = ['Column1', 'Column2']  # List of columns to clean
    output_file = 'cleaned_output_file.xlsx'  # Output file for the cleaned data

    # Process the Excel file
    process_excel_file(file_path, columns_to_clean, output_file)
