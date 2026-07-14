import sqlite3

conn = sqlite3.connect("database.db")
cursor = conn.cursor()

try:
    cursor.execute("ALTER TABLE tyres ADD COLUMN purchase_price INTEGER")
except:
    print("purchase_price already exists")

try:
    cursor.execute("ALTER TABLE tyres ADD COLUMN selling_price INTEGER")
except:
    print("selling_price already exists")

conn.commit()
conn.close()

print("Database updated successfully!")