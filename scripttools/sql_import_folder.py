#USAGE: python sql_importer.py C:\TEMP --server MPOPE-11V\SQLEXPRESS --database ebppsupport
import pandas as pd
from sqlalchemy import create_engine
import os
import argparse

def read_csv_with_fallback(file_path, encodings=['utf-8', 'latin1', 'cp1252', 'utf-16']):
    for enc in encodings:
        try:
            return pd.read_csv(file_path, dtype=str, encoding=enc)
        except UnicodeDecodeError as e:
            print(f"⚠️ Encoding '{enc}' failed for {file_path}: {e}")
    raise UnicodeDecodeError("All encoding attempts failed.")

def import_file_to_sqlserver(file_path, table_name, server='localhost', database='YourDatabase', if_exists='replace'):
    """
    Imports a CSV or Excel file into a SQL Server table.
    """
    try:
        ext = os.path.splitext(file_path)[1].lower()

        # Load file into DataFrame
        if ext in ['.csv', '.txt']:
            df = read_csv_with_fallback(file_path)
        elif ext in ['.xls', '.xlsx']:
            df = pd.read_excel(file_path, dtype=str)
        else:
            print(f"⚠️ Skipping unsupported file type: {file_path}")
            return False

        # Create connection string
        conn_str = f"mssql+pyodbc://{server}/{database}?driver=ODBC+Driver+17+for+SQL+Server"
        engine = create_engine(conn_str)

        # Write to SQL Server
        df.to_sql(table_name, con=engine, if_exists=if_exists, index=False)
        print(f"✅ Imported '{file_path}' into table '{table_name}'")
        return True

    except Exception as e:
        print(f"❌ Failed to import {file_path}: {e}")
        return False

if __name__ == "__main__":
    parser = argparse.ArgumentParser(description="Import all CSV/Excel files in a folder into SQL Server.")
    parser.add_argument("folder_path", help="Path to the folder containing .csv/.xlsx files")
    parser.add_argument("--server", default="MPOPE-11V\\SQLEXPRESS", help="SQL Server instance name")
    parser.add_argument("--database", default="ebppsupport", help="Target database name")
    parser.add_argument("--if_exists", default="replace", choices=["fail", "replace", "append"],
                        help="Behavior if the table already exists")

    args = parser.parse_args()

    # Loop through all files in folder
    for filename in os.listdir(args.folder_path):
        file_path = os.path.join(args.folder_path, filename)
        if os.path.isfile(file_path) and filename.lower().endswith(('.csv', '.xlsx')):
            table_name = os.path.splitext(filename)[0]  # base filename as table name
            success = import_file_to_sqlserver(
                file_path=file_path,
                table_name=table_name,
                server=args.server,
                database=args.database,
                if_exists=args.if_exists
            )
            if not success:
                print(f"⚠️ Skipped {filename}")
