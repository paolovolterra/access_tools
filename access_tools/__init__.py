# === access_tools/__init__.py ===

import os
import shutil
import pandas as pd
import pyodbc


# === 1. CREAZIONE DATABASE ===
def create_access_db_from_template(target_path, template_path):
    """
    Crea un file .accdb copiando un template vuoto.
    """
    if os.path.exists(target_path):
        print(f"✅ File già esistente: {target_path}")
        return

    if not os.path.exists(template_path):
        raise FileNotFoundError(f"Template non trovato: {template_path}")

    os.makedirs(os.path.dirname(target_path), exist_ok=True)
    shutil.copyfile(template_path, target_path)
    print(f"📁 Creato nuovo .accdb da template: {target_path}")


# === 2. CONNESSIONE ===
def get_connection(db_path):
    conn_str = fr"DRIVER={{Microsoft Access Driver (*.mdb, *.accdb)}};DBQ={db_path};"
    return pyodbc.connect(conn_str)


# === 3. CONVERSIONE TIPI ===
def infer_access_type(series):
    if pd.api.types.is_integer_dtype(series):
        return "INT"
    elif pd.api.types.is_float_dtype(series):
        return "DOUBLE"
    elif pd.api.types.is_datetime64_any_dtype(series):
        return "DATETIME"
    else:
        return "TEXT(255)"


# === 4. SCRITTURA TABELLA ===
def df_to_access_table(df, conn, table_name, overwrite=True):
    cursor = conn.cursor()
    try:
        tables = [row.table_name for row in cursor.tables(tableType='TABLE')]
        if table_name in tables:
            if overwrite:
                print(f"🗑️ Tabella '{table_name}' già esistente: la elimino.")
                cursor.execute(f"DROP TABLE [{table_name}]")
            else:
                raise ValueError(f"La tabella '{table_name}' esiste già e overwrite=False")

        col_defs = [f"[{col}] {infer_access_type(df[col])}" for col in df.columns]
        create_sql = f"CREATE TABLE [{table_name}] ({', '.join(col_defs)})"
        print(f"🛠️ Creo tabella '{table_name}'...")
        cursor.execute(create_sql)

        colnames = [f"[{c}]" for c in df.columns]
        insert_sql = f"INSERT INTO [{table_name}] ({', '.join(colnames)}) VALUES ({', '.join(['?'] * len(df.columns))})"
        print(f"📤 Inserisco {len(df)} righe...")
        for _, row in df.iterrows():
            cursor.execute(insert_sql, tuple(row))

        conn.commit()
        print("✅ Completato.")
    except Exception as e:
        conn.rollback()
        raise e
    finally:
        cursor.close()


# === 5. LETTURA TABELLA ===
def read_table(conn, table_name):
    """Legge una tabella Access in un DataFrame."""
    query = f"SELECT * FROM [{table_name}]"
    return pd.read_sql(query, conn)


# === 6. CSV → ACCESS ===
def csv_to_access(csv_path, conn, table_name, overwrite=True, sep=","):
    """Carica un CSV in una tabella Access."""
    df = pd.read_csv(csv_path, sep=sep)
    df_to_access_table(df, conn, table_name, overwrite=overwrite)
