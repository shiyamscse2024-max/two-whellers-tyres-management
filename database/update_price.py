import sqlite3

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

cursor.execute("""
UPDATE tyres
SET purchase_price = 2100,
    selling_price = price
WHERE purchase_price IS NULL
""")

conn.commit()
conn.close()

print("Prices updated!")