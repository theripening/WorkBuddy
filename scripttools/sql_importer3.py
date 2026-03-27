import pandas as pd
from sqlalchemy import create_engine
import os

def read_csv_with_fallback(file_path, delimiter=",",
                           encodings=['utf-8', 'latin1', 'cp1252', 'utf-16']):
    for enc in encodings:
        try:
            return pd.read_csv(file_path, dtype=str, encoding=enc, delimiter=delimiter)
        except UnicodeDecodeError as e:
            print(f"⚠️ Encoding '{enc}' failed: {e}")
    raise UnicodeDecodeError("All encoding attempts failed.")

def import_file_to_sqlserver(file_path, table_name=None, server='localhost',
                             database='YourDatabase', if_exists='replace',
                             delimiter=","):
    """
    Imports a CSV or Excel file into SQL Server.
    If Excel: imports *each worksheet* as a separate table.
    """

    base_name = os.path.splitext(os.path.basename(file_path))[0]

    try:
        ext = os.path.splitext(file_path)[1].lower()

        # Create connection string
        conn_str = f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        engine = create_engine(conn_str)

        # CSV / TXT
        if ext in ['.csv', '.txt']:
            df = read_csv_with_fallback(file_path, delimiter=delimiter)
            final_table = table_name or base_name
            df.to_sql(final_table, con=engine, if_exists=if_exists, index=False)
            print(f"✅ Imported '{file_path}' into table '{final_table}'")
            return True

        # Excel: import ALL worksheets
        elif ext in ['.xls', '.xlsx']:
            sheets = pd.read_excel(file_path, dtype=str, sheet_name=None)

            for sheet_name, df in sheets.items():
                safe_sheet = sheet_name.replace(" ", "_").replace("-", "_")
                final_table = f"{base_name}_{safe_sheet}"

                df.to_sql(final_table, con=engine, if_exists=if_exists, index=False)
                print(f"✅ Imported worksheet '{sheet_name}' into table '{final_table}'")

            return True

        else:
            raise ValueError(f"Unsupported file type: {ext}")

    except Exception as e:
        print(f"❌ Failed to import file: {e}")
        return False


if __name__ == "__main__":
    import argparse

    parser = argparse.ArgumentParser(description="Import CSV or Excel file into SQL Server.")
    parser.add_argument("file_path", help="Path to the .csv or .xlsx file")
    parser.add_argument("--table_name", help="Target SQL Server table name (CSV only)", default=None)
    parser.add_argument("--server", default="MPOPE-11V\\SQLEXPRESS", help="SQL Server instance name")
    parser.add_argument("--database", default="ebppsupport", help="Target database name")
    parser.add_argument("--if_exists", default="replace", choices=["fail", "replace", "append"],
                        help="Behavior if the table already exists")

    # NEW: delimiter options
    parser.add_argument("--delimiter", default=",",
                        help="Delimiter for CSV/TXT files (default is comma)")
    parser.add_argument("--pipe", action="store_true",
                        help="Use pipe '|' as the delimiter")

    args = parser.parse_args()

    # Resolve delimiter
    delimiter = "|" if args.pipe else args.delimiter

    success = import_file_to_sqlserver(
        file_path=args.file_path,
        table_name=args.table_name,
        server=args.server,
        database=args.database,
        if_exists=args.if_exists,
        delimiter=delimiter
    )

    if not success:
        exit(1)
