import sqlite3

DB = "database/database.db"

conn = sqlite3.connect(DB)
cursor = conn.cursor()

# Get existing columns for sales
cursor.execute("PRAGMA table_info(sales)")
cols = [row[1] for row in cursor.fetchall()]

needed = {
    "front_price": "INTEGER",
    "rear_price": "INTEGER",
    "tube": "INTEGER",
    "alignment": "INTEGER",
    "balancing": "INTEGER",
    "discount": "INTEGER",
    "gst": "INTEGER"
}

for col, col_type in needed.items():
    if col not in cols:
        try:
            cursor.execute(f"ALTER TABLE sales ADD COLUMN {col} {col_type}")
            print(f"Added column {col}")
        except Exception as e:
            print(f"Could not add {col}: {e}")

conn.commit()
conn.close()

print("Migration complete.")
