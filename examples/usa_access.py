from access_tools import (
    create_access_db_from_template,
    get_connection,
    df_to_access_table,
    read_table,
    csv_to_access
)
import pandas as pd

# === Percorsi ===
db_path = "D:/clienti.accdb"
template_path = "D:/access_template/template.accdb"

# === Crea DB se non esiste
create_access_db_from_template(db_path, template_path)

# === Connessione
conn = get_connection(db_path)

# === DataFrame di esempio
df = pd.DataFrame({
    "Nome": ["Paolo", "Giulia"],
    "Età": [40, 35]
})
df_to_access_table(df, conn, "Clienti")

# === Leggi tabella e mostra
df2 = read_table(conn, "Clienti")
print(df2)

conn.close()
