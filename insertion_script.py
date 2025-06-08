import pandas as pd
import sqlite3
import os

# === CONFIGURATION ===
CSV_FILE = 'data/customer_churn_dataset.csv'
DB_FILE = 'data.db'
TABLE_NAME = 'churn_table'

# === STEP 1: LOAD CSV WITH PANDAS ===
if not os.path.exists(CSV_FILE):
    raise FileNotFoundError(f"CSV file '{CSV_FILE}' not found.")

df = pd.read_csv(CSV_FILE)

# === STEP 2: CONNECT TO SQLITE DB ===
conn = sqlite3.connect(DB_FILE)

# === STEP 3: INSERT DATAFRAME INTO TABLE ===
# This will create the table if it doesn't exist, or replace it
df.to_sql(TABLE_NAME, conn, if_exists='replace', index=False)

# === STEP 4: CLEANUP ===
conn.close()

print(f"✅ Successfully imported '{CSV_FILE}' into table '{TABLE_NAME}' in '{DB_FILE}'.")
