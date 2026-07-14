import sqlite3

conn = sqlite3.connect("database/database.db")
cursor = conn.cursor()

tyres = [

    ("CEAT", "100/90-17", "Zoom Plus", 2350, 15),
    ("CEAT", "90/90-17", "Milaze", 1950, 20),
    ("CEAT", "80/100-18", "Gripp XL", 2100, 18),

    ("TVS", "100/90-17", "Eurogrip Pro", 2400, 12),
    ("TVS", "90/90-17", "Duratrail", 2050, 10),
    ("TVS", "80/100-18", "RoadHound", 2200, 8)

]

cursor.executemany("""
INSERT INTO tyres
(brand, tyre_size, tyre_name, price, stock)
VALUES (?, ?, ?, ?, ?)
""", tyres)

conn.commit()
conn.close()

print("Tyre data inserted successfully!")