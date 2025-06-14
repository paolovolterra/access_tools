# access_tools

`access_tools` è una libreria Python pensata per facilitare l'interazione con database Microsoft Access (`.accdb`), sfruttando `pandas`, `pyodbc` e strumenti di automazione.

---

## 🎯 Obiettivo

Permettere a chi lavora con dati in Excel, CSV o DataFrame Python di esportare (e importare) facilmente questi dati in database Access, senza aprire manualmente Microsoft Access.

Questo è utile per:
- Analisi che richiedono esportazione verso Access per uso legacy
- Backup o scambio dati in formato `.accdb`
- Uffici che gestiscono report e archivi in Access ma elaborano i dati in Python

---

## 🚀 Funzionalità incluse

| Funzione | Descrizione |
|---------|-------------|
| `create_access_db_from_template(path, template)` | Crea un nuovo file `.accdb` copiando un file template esistente |
| `get_connection(path)` | Apre una connessione ODBC con un database `.accdb` |
| `df_to_access_table(df, conn, table_name)` | Scrive un `DataFrame` come tabella nel database |
| `read_table(conn, table_name)` | Legge una tabella da Access e la restituisce come `DataFrame` |
| `csv_to_access(csv_path, conn, table_name)` | Carica un file `.csv` in una tabella Access |

---

## 🧱 Requisiti

- Sistema operativo: **Windows**
- Driver ODBC Microsoft Access installato (incluso con Microsoft 365 o `AccessDatabaseEngine`)
- Python 3.8+
- Librerie Python:
  - `pandas`
  - `pyodbc`
  - `pywin32`

---

## 🛠️ Esempio di utilizzo

```python
from access_tools import (
    create_access_db_from_template,
    get_connection,
    df_to_access_table,
    read_table
)
import pandas as pd

# Crea il file Access da template vuoto (una sola volta)
create_access_db_from_template("D:/demo.accdb", "D:/access_template/template.accdb")

# Connessione al database
conn = get_connection("D:/demo.accdb")

# Scrittura di una tabella
df = pd.DataFrame({
    "Nome": ["Mario", "Lucia"],
    "Età": [45, 37]
})
df_to_access_table(df, conn, "Anagrafica")

# Lettura della tabella
df2 = read_table(conn, "Anagrafica")
print(df2)

conn.close()
```

---

## 📁 Struttura del progetto

```
access_tools/
├── access_tools/         # Modulo principale con tutte le funzioni
│   └── __init__.py
├── examples/             # Script d’esempio
│   └── usa_access.py
├── README.md             # Questa guida
├── setup.py              # Setup per installazione pip
└── pyproject.toml        # Configurazione build moderna
```

⚠️ **Importante**: il file `template.accdb` (un database Access vuoto) deve essere creato una volta manualmente e salvato in una posizione come `D:/access_template/template.accdb`.

---

## 🧩 Idee per estensioni future

- Scrittura incrementale o aggiornamenti condizionati (`UPSERT`)
- Visualizzazione schema tabelle Access
- Supporto per tipi booleani e chiavi primarie
- GUI minimale per selezionare CSV e salvare in Access

---

## 👤 Autore

Creato da [Il Tuo Nome] • Contatti e contributi benvenuti!

