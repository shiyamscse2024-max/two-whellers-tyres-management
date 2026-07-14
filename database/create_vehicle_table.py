import sqlite3

conn = sqlite3.connect("database/database.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS vehicles (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    company TEXT NOT NULL,

    model TEXT NOT NULL,

    front_tyre TEXT NOT NULL,

    rear_tyre TEXT NOT NULL,

    tube_type TEXT NOT NULL

)
""")

conn.commit()
conn.close()

print("Vehicle Table Created Successfully!")