import sqlite3

conn = sqlite3.connect("database/database.db")
cursor = conn.cursor()

cursor.execute("""
CREATE TABLE IF NOT EXISTS price_history(

id INTEGER PRIMARY KEY AUTOINCREMENT,

tyre_id INTEGER,

old_price INTEGER,

new_price INTEGER,

updated_on TEXT

)

""")

conn.commit()
conn.close()

print("Price History Table Created Successfully!")