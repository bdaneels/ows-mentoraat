import pandas as pd
import logging
from typing import List, Optional

# Configure logging
logging.basicConfig(level=logging.INFO, format='%(levelname)s: %(message)s')

def load_excel(file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame:
    """Load an Excel file into a DataFrame."""
    try:
        df = pd.read_excel(file_path, sheet_name=sheet_name)
        if isinstance(df, dict):
            raise ValueError(f"Multiple sheets found in {file_path}. Please specify a sheet name.")
        return df
    except FileNotFoundError:
        logging.error(f"File not found: {file_path}")
        raise
    except Exception as e:
        logging.error(f"Error loading file {file_path}: {e}")
        raise

def check_duplicates(df: pd.DataFrame, column: str) -> List[str]:
    """Check for duplicate values in a specified column."""
    duplicates = df[column].astype(str)[df[column].duplicated()]
    return duplicates.tolist()

def find_missing_ids(df1: pd.DataFrame, df2: pd.DataFrame, column: str) -> List[str]:
    """Find IDs in df2 that are not in df1."""
    ids1 = df1[column].astype(str)
    ids2 = df2[column].astype(str)
    missing_ids = ids2[~ids2.isin(ids1)]
    return missing_ids.tolist()

def append_missing_ids(reinoud_df: pd.DataFrame, sisa_df: pd.DataFrame, column: str, reinoud_file: str) -> pd.DataFrame:
    """Append missing IDs and corresponding Naam, Voornaam, Plan, and Campus emailadres to reinoud_df."""
    missing_ids = find_missing_ids(reinoud_df, sisa_df, column)
    if missing_ids:
        missing_rows = sisa_df[sisa_df[column].astype(str).isin(missing_ids)]
        # Select only the specified columns
        selected_columns = ['Rolnummer', 'Naam', 'Voornaam', 'Plan', 'Campus emailadres']
        missing_rows = missing_rows[selected_columns]
        
        # Rename 'Campus emailadres' to 'mail' for reinoud_df
        missing_rows = missing_rows.rename(columns={'Campus emailadres': 'mail'})
        
        # Append missing rows to reinoud_df
        reinoud_df = pd.concat([reinoud_df, missing_rows], ignore_index=True)
        
        logging.info(f"Appended missing IDs to {reinoud_file}:")
        for _, row in missing_rows.iterrows():
            logging.info(f"ID: {row[column]}, Naam: {row['Naam']}, Voornaam: {row['Voornaam']}, Plan: {row['Plan']}, mail: {row['mail']}")
    else:
        logging.info("No missing IDs to append.")
    return reinoud_df

def main(reinoud_file: str, sisa_file: str, column: str, reinoud_sheet: Optional[str] = None, sisa_sheet: Optional[str] = None):
    # Load the Excel files
    reinoud_df = load_excel(reinoud_file, sheet_name=reinoud_sheet)
    sisa_df = load_excel(sisa_file, sheet_name=sisa_sheet)

    # Debug: Print columns of sisa_df
    logging.info(f"Columns in {sisa_file}: {sisa_df.columns.tolist()}")

    # Check for duplicates in reinoud
    duplicates = check_duplicates(reinoud_df, column)
    if duplicates:
        logging.info("Duplicate IDs in reinoud.xlsx:")
        logging.info(duplicates)
    else:
        logging.info("No duplicates found in reinoud.xlsx.")

    # Append missing IDs from sisa to reinoud
    reinoud_df = append_missing_ids(reinoud_df, sisa_df, column, reinoud_file)

    # Save the updated reinoud_df back to the Excel file
    reinoud_df.to_excel(reinoud_file, sheet_name=reinoud_sheet, index=False)
    logging.info(f"Updated {reinoud_file} saved.")

if __name__ == "__main__":
    # Example usage
    # change the file names, column name, and sheet names as needed
    main('reinoud.xlsx', 'sisa.xlsx', 'Rolnummer', reinoud_sheet='Actief', sisa_sheet='sheet1')