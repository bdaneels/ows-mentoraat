# Script Documentation

## Overview

This script processes two Excel files (

reinoud.xlsx

and

sisa.xlsx

) to find and append missing IDs from

sisa.xlsx

to

reinoud.xlsx

. It also checks for duplicate IDs in

reinoud.xlsx

.

## Functions

### [`load_excel(file_path: str, sheet_name: Optional[str] = None) -> pd.DataFrame`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A7%2C%22character%22%3A4%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition")

Loads an Excel file into a DataFrame.

### [`check_duplicates(df: pd.DataFrame, column: str) -> List[str]`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A21%2C%22character%22%3A4%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition")

Checks for duplicate values in a specified column.

### [`find_missing_ids(df1: pd.DataFrame, df2: pd.DataFrame, column: str) -> List[str]`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A26%2C%22character%22%3A4%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition")

Finds IDs in [`df2`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A26%2C%22character%22%3A40%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition") that are not in [`df1`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A26%2C%22character%22%3A21%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition").

### [`append_missing_ids(reinoud_df: pd.DataFrame, sisa_df: pd.DataFrame, column: str, reinoud_file: str) -> pd.DataFrame`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A33%2C%22character%22%3A4%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition")

Appends missing IDs and corresponding details from [`sisa_df`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A33%2C%22character%22%3A49%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition") to [`reinoud_df`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A33%2C%22character%22%3A23%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition").

### [`main(reinoud_file: str, sisa_file: str, column: str, reinoud_sheet: Optional[str] = None, sisa_sheet: Optional[str] = None)`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A55%2C%22character%22%3A4%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition")

Main function to load the Excel files, check for duplicates, append missing IDs, and save the updated DataFrame back to the Excel file.

## Usage

Run the script with the following command:

```sh
python script.py
```

Example usage within the script:

```python
if __name__ == "__main__":
    main('reinoud.xlsx', 'sisa.xlsx', 'Rolnummer', reinoud_sheet='Actief', sisa_sheet='sheet1')
```

## Logging

The script uses the [`logging`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A1%2C%22character%22%3A7%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition") module to log information and errors. The log level is set to [`INFO`](command:_github.copilot.openSymbolFromReferences?%5B%22%22%2C%5B%7B%22uri%22%3A%7B%22scheme%22%3A%22file%22%2C%22authority%22%3A%22%22%2C%22path%22%3A%22%2Fc%3A%2FUsers%2Fbrech%2FDocuments%2FlocalReps%2Fows-mentoraat%2Fscript.py%22%2C%22query%22%3A%22%22%2C%22fragment%22%3A%22%22%7D%2C%22pos%22%3A%7B%22line%22%3A5%2C%22character%22%3A34%7D%7D%5D%2C%22723f84da-7613-47af-a432-459fba37ba55%22%5D "Go to definition").

## File Structure

```
.gitignore
reinoud.xlsx
script.py
sisa.xlsx
```

## Dependencies

- pandas
- logging

Install dependencies using:

```sh
pip install pandas
```

## License

This script is provided "as-is" without any warranty. Use at your own risk.
