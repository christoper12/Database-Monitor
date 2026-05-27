import pandas as pd
from openpyxl import Workbook
from openpyxl.utils.dataframe import dataframe_to_rows
from datetime import datetime
from sqlalchemy import create_engine
import time
import os
import configparser

# =========================
# EXPORT FUNCTION
# =========================
def export_excel_file(open_file=False):

    # =========================
    # LOAD CONFIG
    # =========================
    config = configparser.ConfigParser()
    config.read('config.ini')

    HOST = config.get('MYSQL', 'HOST')
    USER = config.get('MYSQL', 'USER')
    PASSWORD = config.get('MYSQL', 'PASSWORD')
    PORT = config.getint('MYSQL', 'PORT')

    print("=" * 60)
    print("START MYSQL DATABASE EXPORT")
    print("=" * 60)

    start_time = time.time()

    # =========================
    # CONNECT MYSQL
    # =========================
    print("[1/5] Connecting to MySQL...")

    engine = create_engine(
        f"mysql+pymysql://{USER}:{PASSWORD}@{HOST}:{PORT}"
    )

    print("SUCCESS: Connected to MySQL")

    # =========================
    # GET DATABASE LIST
    # =========================
    print("[2/5] Fetching database list...")

    db_query = """
    SELECT SCHEMA_NAME
    FROM information_schema.SCHEMATA
    WHERE SCHEMA_NAME NOT IN (
        'information_schema',
        'performance_schema',
        'mysql',
        'sys'
    )
    ORDER BY SCHEMA_NAME;
    """

    databases = pd.read_sql(db_query, engine)

    total_db = len(databases)

    print(f"SUCCESS: Found {total_db} databases")

    # =========================
    # CREATE EXCEL
    # =========================
    print("[3/5] Creating Excel workbook...")

    wb = Workbook()

    default_sheet = wb.active
    wb.remove(default_sheet)

    print("SUCCESS: Workbook created")

    # =========================
    # LOOP DATABASES
    # =========================
    print("[4/5] Processing databases...")
    print("-" * 60)

    for index, db in enumerate(databases['SCHEMA_NAME'], start=1):

        db_start = time.time()

        print(f"[{index}/{total_db}] Processing DB: {db}")

        query = f"""
        SELECT
          TABLE_NAME AS `Table`,
          ROUND((DATA_LENGTH + INDEX_LENGTH) / 1024 / 1024, 2) AS `Size (MB)`
        FROM
          information_schema.TABLES
        WHERE
          TABLE_SCHEMA = '{db}'
        ORDER BY
          (DATA_LENGTH + INDEX_LENGTH) DESC;
        """

        try:

            df = pd.read_sql(query, engine)

            total_tables = len(df)

            print(f"     Tables found : {total_tables}")

            for i, table_name in enumerate(df['Table'], start=1):
                print(f"        [{i}/{total_tables}] {table_name}")

            # Create Sheet
            ws = wb.create_sheet(title=db[:31])

            # Header
            ws.append([f"Database: {db}"])
            ws.append([])

            # Insert Data
            for row in dataframe_to_rows(
                df,
                index=False,
                header=True
            ):
                ws.append(row)

            elapsed = round(
                time.time() - db_start,
                2
            )

            print(f"     SUCCESS ({elapsed} sec)")
            print("-" * 60)

        except Exception as e:

            print(f"     ERROR: {e}")
            print("-" * 60)

    # =========================
    # SAVE FILE
    # =========================
    print("[5/5] Saving Excel file...")

    # Create output folder
    output_dir = "output"

    os.makedirs(output_dir, exist_ok=True)

    # Filename
    now = datetime.now().strftime(
        "%d-%m-%Y_%H-%M-%S"
    )

    output_file = os.path.join(
        output_dir,
        f"mysql_table_sizes_{now}.xlsx"
    )

    # Save
    wb.save(output_file)

    # Open file if requested
    if open_file:
        os.startfile(output_file)
        print("Opening Excel file...")

    total_time = round(
        time.time() - start_time,
        2
    )

    print("=" * 60)
    print("SUCCESS: File saved")
    print(f"FILE    : {output_file}")
    print(f"DURATION: {total_time} seconds")
    print("=" * 60)

    # Return file path
    return output_file


# =========================
# RUN MANUALLY
# =========================
if __name__ == "__main__":

    open_excel = input(
        "Open Excel file now? (y/n): "
    )

    result = export_excel_file(
        open_file=(open_excel.lower() == "y")
    )

    print(f"Generated: {result}")