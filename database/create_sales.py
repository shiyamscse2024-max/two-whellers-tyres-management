import sqlite3

conn = sqlite3.connect("database/database.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS sales (

    id INTEGER PRIMARY KEY AUTOINCREMENT,

    customer_name TEXT,

    vehicle_model TEXT,

    tyre_name TEXT,

    quantity INTEGER,

    price INTEGER,

    total INTEGER,

    sold_on TEXT
)
""")

conn.commit()
conn.close()

print("Sales Table Created Successfully ✅")