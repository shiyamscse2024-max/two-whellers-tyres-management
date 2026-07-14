import sqlite3

conn = sqlite3.connect("database/database.db")
cursor = conn.cursor()

# Tyres
cursor.execute("""
CREATE TABLE IF NOT EXISTS tyres (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    brand TEXT NOT NULL,
    tyre_size TEXT NOT NULL,
    tyre_name TEXT NOT NULL,
    price INTEGER NOT NULL,
    stock INTEGER NOT NULL
)
""")

# Price History
cursor.execute("""
CREATE TABLE IF NOT EXISTS price_history (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    tyre_id INTEGER,
    old_price INTEGER,
    new_price INTEGER,
    updated_on TEXT
)
""")

# Customers
cursor.execute("""
CREATE TABLE IF NOT EXISTS customers (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    name TEXT NOT NULL,
    mobile TEXT NOT NULL,
    vehicle_number TEXT NOT NULL,
    vehicle_model TEXT NOT NULL,
    created_at TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
""")

# Sales
cursor.execute("""
CREATE TABLE IF NOT EXISTS sales (
    id INTEGER PRIMARY KEY AUTOINCREMENT,
    customer_name TEXT,
    vehicle_model TEXT,
    front_tyre TEXT,
    rear_tyre TEXT,
    front_qty INTEGER,
    rear_qty INTEGER,
    total_amount INTEGER,
    sale_date TIMESTAMP DEFAULT CURRENT_TIMESTAMP
)
""")

conn.commit()
conn.close()

print("Database Created Successfully!")