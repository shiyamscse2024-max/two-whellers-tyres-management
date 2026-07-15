"""
Copy SQLite data to a Postgres database. Usage:

    export DATABASE_URL=postgres://user:pass@host:port/dbname
    python database/sqlite_to_postgres.py

This script will create tables in Postgres (if they don't exist) and copy rows from:
- tyres
- price_history
- customers
- vehicles
- sales

Be careful: it does not attempt to merge or deduplicate — it's a straightforward copy.
"""
import os
import sqlite3
import sys

try:
    import psycopg2
    from psycopg2.extras import execute_values
except Exception as e:
    print("psycopg2 is required. Install with: pip install psycopg2-binary")
    raise

SQLITE_DB = os.path.join(os.path.dirname(__file__), "database.db")
PG_DSN = os.environ.get("DATABASE_URL") or os.environ.get("DATABASE")

if not PG_DSN:
    print("Please set the DATABASE_URL environment variable to your Postgres DSN.")
    sys.exit(1)

print("Connecting to SQLite:", SQLITE_DB)
sql_conn = sqlite3.connect(SQLITE_DB)
sql_cur = sql_conn.cursor()

print("Connecting to Postgres:", PG_DSN)
pg_conn = psycopg2.connect(PG_DSN)
pg_cur = pg_conn.cursor()

# Helper to copy table
def copy_table(table, cols):
    print(f"Copying table {table}")
    sql_cur.execute(f"SELECT {', '.join(cols)} FROM {table}")
    rows = sql_cur.fetchall()
    if not rows:
        print("  (no rows)")
        return
    placeholders = ','.join(['%s'] * len(cols))
    insert_sql = f"INSERT INTO {table} ({', '.join(cols)}) VALUES ({placeholders})"
    try:
        execute_values(pg_cur, f"INSERT INTO {table} ({', '.join(cols)}) VALUES %s", rows)
        pg_conn.commit()
        print(f"  inserted {len(rows)} rows into {table}")
    except Exception as e:
        pg_conn.rollback()
        print(f"  failed to insert into {table}: {e}")

# Ensure tables exist in Postgres - create minimal schemas
print("Ensuring target tables exist in Postgres (creating if missing)")
pg_cur.execute("""
CREATE TABLE IF NOT EXISTS tyres (
    id SERIAL PRIMARY KEY,
    brand TEXT,
    tyre_size TEXT,
    tyre_name TEXT,
    price INTEGER,
    stock INTEGER
)
""")
pg_cur.execute("""
CREATE TABLE IF NOT EXISTS price_history (
    id SERIAL PRIMARY KEY,
    tyre_id INTEGER,
    old_price INTEGER,
    new_price INTEGER,
    updated_on TEXT
)
""")
pg_cur.execute("""
CREATE TABLE IF NOT EXISTS customers (
    id SERIAL PRIMARY KEY,
    name TEXT,
    mobile TEXT,
    vehicle_number TEXT,
    vehicle_model TEXT,
    created_at TIMESTAMP
)
""")
pg_cur.execute("""
CREATE TABLE IF NOT EXISTS vehicles (
    id SERIAL PRIMARY KEY,
    company TEXT,
    model TEXT,
    front_tyre TEXT,
    rear_tyre TEXT,
    tube_type TEXT
)
""")
pg_cur.execute("""
CREATE TABLE IF NOT EXISTS sales (
    id SERIAL PRIMARY KEY,
    customer_name TEXT,
    vehicle_model TEXT,
    front_tyre TEXT,
    rear_tyre TEXT,
    front_qty INTEGER,
    rear_qty INTEGER,
    total_amount INTEGER,
    front_price INTEGER,
    rear_price INTEGER,
    tube INTEGER,
    alignment INTEGER,
    balancing INTEGER,
    discount INTEGER,
    gst INTEGER,
    sale_date TIMESTAMP
)
""")
pg_conn.commit()

# Copy tables
copy_table('tyres', ['id','brand','tyre_size','tyre_name','price','stock'])
copy_table('price_history', ['id','tyre_id','old_price','new_price','updated_on'])
copy_table('customers', ['id','name','mobile','vehicle_number','vehicle_model','created_at'])
copy_table('vehicles', ['id','company','model','front_tyre','rear_tyre','tube_type'])
copy_table('sales', ['id','customer_name','vehicle_model','front_tyre','rear_tyre','front_qty','rear_qty','total_amount','front_price','rear_price','tube','alignment','balancing','discount','gst','sale_date'])

sql_conn.close()
pg_cur.close()
pg_conn.close()

print('\nMigration finished. Verify data in Postgres before switching production.')
