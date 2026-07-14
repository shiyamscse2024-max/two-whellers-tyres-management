from flask import Flask, render_template, request, redirect, url_for, session, flash
import sqlite3
import pandas as pd
from collections import defaultdict
from pathlib import Path
from openpyxl import Workbook

app = Flask(__name__)
app.secret_key = "saravana-auto-parts-secret"
print("MY APP.PY IS RUNNING")

DATABASE = "database/database.db"
MASTER_VEHICLE_FILE = r"c:\Users\s8162\OneDrive\Desktop\PROJECT NO 1.xlsx"


def _require_login():
    if not session.get("user"):
        flash("Please log in to continue.", "warning")
        return redirect(url_for("login"))
    return None


def _normalize_header(value):
    if pd.isna(value):
        return ""
    return str(value).strip().lower().replace(" ", "_")


def _sync_vehicles_to_excel(file_path):
    conn = sqlite3.connect(DATABASE)
    rows = conn.execute("""
        SELECT company, model, front_tyre, rear_tyre, tube_type
        FROM vehicles
        ORDER BY id
    """).fetchall()
    conn.close()

    path = Path(file_path)
    path.parent.mkdir(parents=True, exist_ok=True)

    wb = Workbook()
    ws = wb.active
    ws.title = "Vehicles"
    ws.append(["Company", "Model", "Front Tyre", "Rear Tyre", "Tube Type"])

    for company, model, front_tyre, rear_tyre, tube_type in rows:
        ws.append([company, model, front_tyre, rear_tyre, tube_type])

    wb.save(path)


def _import_master_vehicle_file(file_path):
    df = pd.read_excel(file_path, header=None)

    header_row = None
    for idx, row in df.iterrows():
        normalized = [_normalize_header(value) for value in row.tolist()]
        if any(name in normalized for name in ["company", "model", "front_tyre", "rear_tyre", "tube_type"]):
            header_row = idx
            break

    if header_row is None:
        raise ValueError("No valid header row found in the master Excel file")

    raw_headers = [str(value).strip() if pd.notna(value) else "" for value in df.iloc[header_row].tolist()]
    data_rows = df.iloc[header_row + 1:].copy()
    data_rows.columns = [f"col_{i}" for i in range(len(raw_headers))]

    renamed_columns = {}
    for index, header in enumerate(raw_headers):
        normalized = _normalize_header(header)
        if normalized in ["company", "model", "front_tyre", "rear_tyre", "tube_type"]:
            renamed_columns[f"col_{index}"] = normalized

    data_rows = data_rows.rename(columns=renamed_columns)
    data_rows = data_rows[[col for col in data_rows.columns if col in renamed_columns.values()]]

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()
    cursor.execute("DELETE FROM vehicles")

    for _, row in data_rows.iterrows():
        company = str(row.get("company", "")).strip() if not pd.isna(row.get("company", "")) else ""
        model = str(row.get("model", "")).strip() if not pd.isna(row.get("model", "")) else ""
        front_tyre = str(row.get("front_tyre", "")).strip() if not pd.isna(row.get("front_tyre", "")) else ""
        rear_tyre = str(row.get("rear_tyre", "")).strip() if not pd.isna(row.get("rear_tyre", "")) else ""
        tube_type = str(row.get("tube_type", "")).strip() if not pd.isna(row.get("tube_type", "")) else ""

        if not company or not model:
            continue

        cursor.execute("""
            INSERT INTO vehicles
            (company, model, front_tyre, rear_tyre, tube_type)
            VALUES (?, ?, ?, ?, ?)
        """, (company, model, front_tyre, rear_tyre, tube_type))

    conn.commit()
    conn.close()
    _sync_vehicles_to_excel(file_path)


# -----------------------------
# HOME (SMART SEARCH)
# -----------------------------
@app.route("/", methods=["GET", "POST"])
def home():

    tyres = []
    query = ""

    if request.method == "POST":
        query = request.form.get("query", "").strip()
        search = "%" + query + "%"

        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()

        cursor.execute("""
            SELECT * FROM tyres
            WHERE
                brand LIKE ?
                OR tyre_size LIKE ?
                OR tyre_name LIKE ?
        """, (search, search, search))

        tyres = cursor.fetchall()

        conn.close()

    return render_template("index.html", tyres=tyres, query=query)


# -----------------------------
# DASHBOARD
# -----------------------------
@app.route("/dashboard")
def dashboard():

    login_check = _require_login()
    if login_check is not None:
        return login_check

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("SELECT COUNT(*) FROM tyres")
    total_tyres = cursor.fetchone()[0]

    cursor.execute("SELECT SUM(stock) FROM tyres")
    total_stock = cursor.fetchone()[0] or 0

    cursor.execute("SELECT COUNT(*) FROM tyres WHERE stock < 5")
    low_stock = cursor.fetchone()[0]

    cursor.execute("SELECT COUNT(*) FROM sales")
    total_sales = cursor.fetchone()[0]

    cursor.execute("SELECT SUM(total_amount) FROM sales")
    total_revenue = cursor.fetchone()[0] or 0

    cursor.execute("""
        SELECT substr(sale_date,1,7), SUM(total_amount)
        FROM sales
        GROUP BY substr(sale_date,1,7)
        ORDER BY substr(sale_date,1,7)
    """)
    monthly_sales = cursor.fetchall()

    months = [row[0] for row in monthly_sales]
    revenues = [row[1] for row in monthly_sales]

    cursor.execute("""
        SELECT brand, COUNT(*)
        FROM tyres
        GROUP BY brand
        ORDER BY COUNT(*) DESC
    """)
    brands = cursor.fetchall()

    brand_names = [row[0] for row in brands]
    brand_count = [row[1] for row in brands]

    cursor.execute("""
        SELECT customer_name, total_amount, sale_date
        FROM sales
        ORDER BY id DESC
        LIMIT 5
    """)
    recent_sales = cursor.fetchall()

    cursor.execute("""
        SELECT id, brand, tyre_size, tyre_name, price, stock
        FROM tyres
        WHERE stock < 5
        ORDER BY stock ASC
    """)
    low_stock_items = cursor.fetchall()

    conn.close()

    return render_template(
        "dashboard.html",
        total_tyres=total_tyres,
        total_stock=total_stock,
        low_stock=low_stock,
        total_sales=total_sales,
        total_revenue=total_revenue,
        months=months,
        revenues=revenues,
        brand_names=brand_names,
        brand_count=brand_count,
        recent_sales=recent_sales,
        low_stock_items=low_stock_items
    )
# -----------------------------
# VIEW ALL TYRES
# -----------------------------
@app.route("/all-tyres")
def all_tyres():

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM tyres")
    tyres = cursor.fetchall()

    print("Tyres:", tyres)
    print("Count:", len(tyres))
    print("DATABASE:", DATABASE)

    conn.close()

    return render_template("all_tyres.html", tyres=tyres)

# -----------------------------
# ADD TYRE
# -----------------------------
@app.route("/add-tyre", methods=["GET", "POST"])
def add_tyre():

    if request.method == "POST":

        brand = request.form["brand"]
        size = request.form["size"]
        name = request.form["name"]
        price = request.form["price"]
        stock = request.form["stock"]

        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()

        cursor.execute("""
            INSERT INTO tyres
            (brand, tyre_size, tyre_name, price, stock)
            VALUES (?, ?, ?, ?, ?)
        """, (brand, size, name, price, stock))

        conn.commit()
        conn.close()

        return redirect(url_for("all_tyres"))

    return render_template("add_tyre.html")


# -----------------------------
# UPDATE TYRE
# -----------------------------
@app.route("/update/<int:id>", methods=["GET", "POST"])
def update_tyre(id):

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    if request.method == "POST":

        new_price = request.form["price"]

        cursor.execute(
            "SELECT price FROM tyres WHERE id=?",
            (id,)
        )

        old_price = cursor.fetchone()[0]

        cursor.execute(
            "UPDATE tyres SET price=? WHERE id=?",
            (new_price, id)
        )

        cursor.execute("""
            INSERT INTO price_history
            (tyre_id, old_price, new_price, updated_on)
            VALUES (?, ?, ?, datetime('now'))
        """, (id, old_price, new_price))

        conn.commit()

        conn.close()

        return redirect(url_for("all_tyres"))

    cursor.execute(
        "SELECT * FROM tyres WHERE id=?",
        (id,)
    )

    tyre = cursor.fetchone()

    conn.close()

    return render_template(
        "update_tyre.html",
        tyre=tyre
    )


# -----------------------------
# PRICE HISTORY
# -----------------------------
@app.route("/history/<int:id>")
def history(id):

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT *
        FROM price_history
        WHERE tyre_id=?
        ORDER BY id DESC
    """, (id,))

    history = cursor.fetchall()

    conn.close()

    return render_template(
        "history.html",
        history=history
    )
@app.route("/delete/<int:id>")
def delete_tyre(id):

    conn = sqlite3.connect("database/database.db")
    cursor = conn.cursor()

    cursor.execute("DELETE FROM tyres WHERE id=?", (id,))

    conn.commit()
    conn.close()

    return redirect("/all-tyres")
@app.route("/increase-stock/<int:id>")
def increase_stock(id):

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute(
        "UPDATE tyres SET stock = stock + 1 WHERE id=?",
        (id,)
    )

    conn.commit()
    conn.close()

    return redirect(url_for("all_tyres"))
@app.route("/decrease-stock/<int:id>")
def decrease_stock(id):

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute(
        """
        UPDATE tyres
        SET stock =
        CASE
            WHEN stock > 0 THEN stock - 1
            ELSE 0
        END
        WHERE id=?
        """,
        (id,)
    )

    conn.commit()
    conn.close()

    return redirect(url_for("all_tyres"))

@app.route("/add-customer", methods=["GET", "POST"])
def add_customer():

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    if request.method == "POST":

        name = request.form["name"]
        mobile = request.form["mobile"]
        vehicle_number = request.form["vehicle_number"]
        vehicle_model = request.form["vehicle_model"]

        cursor.execute("""
            INSERT INTO customers
            (name, mobile, vehicle_number, vehicle_model)
            VALUES (?, ?, ?, ?)
        """, (name, mobile, vehicle_number, vehicle_model))

        conn.commit()
        conn.close()

        return redirect("/customers")

    conn.close()
    return render_template("add_customer.html")

@app.route("/import-excel", methods=["GET", "POST"])
def import_excel():

    if request.method == "POST":

        file = request.files["excel_file"]

        df = pd.read_excel(file)

        conn = sqlite3.connect("database/database.db")
        cursor = conn.cursor()

        for index, row in df.iterrows():

            cursor.execute("""
            INSERT INTO tyres
            (brand, tyre_size, tyre_name, price, stock)
            VALUES (?, ?, ?, ?, ?)
            """,
            (
                row["brand"],
                row["tyre_size"],
                row["tyre_name"],
                row["price"],
                row["stock"]
            ))

        conn.commit()
        conn.close()

        return "Excel Imported Successfully! ✅"

    return render_template("import_excel.html")

@app.route("/vehicles", methods=["GET", "POST"])
def vehicles():

    vehicle = None

    if request.method == "POST":

        model = "%" + request.form["model"] + "%"

        conn = sqlite3.connect("database/database.db")
        cursor = conn.cursor()

        cursor.execute("""
        SELECT *
        FROM vehicles
        WHERE model LIKE ?
        """, (model,))

        vehicle = cursor.fetchone()

        conn.close()

    return render_template(
        "vehicles.html",
        vehicle=vehicle
    )

@app.route("/import-vehicles")
def import_vehicles():

    try:
        _import_master_vehicle_file(MASTER_VEHICLE_FILE)
        return "Vehicles Imported Successfully from master Excel file! ✅"
    except FileNotFoundError:
        return f"Master Excel file not found: {MASTER_VEHICLE_FILE}", 404
    except Exception as exc:
        return f"Import failed: {exc}", 500

@app.route("/all-vehicles")
def all_vehicles():

    conn = sqlite3.connect("database/database.db")
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM vehicles")

    vehicles = cursor.fetchall()

    conn.close()

    return render_template(
        "all_vehicles.html",
        vehicles=vehicles
    )

@app.route("/add-vehicle", methods=["GET", "POST"])
def add_vehicle():

    if request.method == "POST":

        company = request.form["company"]
        model = request.form["model"]
        front = request.form["front_tyre"]
        rear = request.form["rear_tyre"]
        tube = request.form["tube_type"]

        conn = sqlite3.connect("database/database.db")
        cursor = conn.cursor()

        cursor.execute("""
        INSERT INTO vehicles
        (company, model, front_tyre, rear_tyre, tube_type)
        VALUES (?, ?, ?, ?, ?)
        """, (company, model, front, rear, tube))

        conn.commit()
        conn.close()
        _sync_vehicles_to_excel(MASTER_VEHICLE_FILE)

        return redirect(url_for("all_vehicles"))

    return render_template("add_vehicle.html")

@app.route("/edit-vehicle/<int:id>", methods=["GET", "POST"])
def edit_vehicle(id):

    conn = sqlite3.connect("database/database.db")
    cursor = conn.cursor()

    if request.method == "POST":

        cursor.execute("""
        UPDATE vehicles
        SET company=?,
            model=?,
            front_tyre=?,
            rear_tyre=?,
            tube_type=?
        WHERE id=?
        """,(
            request.form["company"],
            request.form["model"],
            request.form["front_tyre"],
            request.form["rear_tyre"],
            request.form["tube_type"],
            id
        ))

        conn.commit()
        conn.close()
        _sync_vehicles_to_excel(MASTER_VEHICLE_FILE)

        return redirect(url_for("all_vehicles"))

    cursor.execute("SELECT * FROM vehicles WHERE id=?", (id,))
    vehicle = cursor.fetchone()

    conn.close()

    return render_template("edit_vehicle.html", vehicle=vehicle)

@app.route("/delete-vehicle/<int:id>")
def delete_vehicle(id):

    conn = sqlite3.connect("database/database.db")
    cursor = conn.cursor()

    cursor.execute(
        "DELETE FROM vehicles WHERE id=?",
        (id,)
    )

    conn.commit()
    conn.close()
    _sync_vehicles_to_excel(MASTER_VEHICLE_FILE)

    return redirect(url_for("all_vehicles"))


@app.route("/vehicle-search", methods=["GET", "POST"])
def vehicle_search():

    print("vehicle_search called")

    vehicle = None
    front_tyres = []
    rear_tyres = []

    if request.method == "POST":

        model = "%" + request.form["model"] + "%"

        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM vehicles")
        print("All vehicles:", cursor.fetchall())

        cursor.execute(
            "SELECT * FROM vehicles WHERE model LIKE ?",
            (model,)
        )

        vehicle = cursor.fetchone()
        print("Matched vehicle:", vehicle)

        if vehicle:

            front = vehicle[3]
            rear = vehicle[4]

            print("Front:", front)
            print("Rear:", rear)

            cursor.execute(
                "SELECT * FROM tyres WHERE tyre_size=?",
                (front,)
            )
            front_tyres = cursor.fetchall()
            print("Front tyres:", front_tyres)

            cursor.execute(
                "SELECT * FROM tyres WHERE tyre_size=?",
                (rear,)
            )
            rear_tyres = cursor.fetchall()
            print("Rear tyres:", rear_tyres)

        conn.close()

    return render_template(
        "vehicle_result.html",
        vehicle=vehicle,
        front_tyres=front_tyres,
        rear_tyres=rear_tyres
    )
@app.route("/customers")
def customers():

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM customers")
    customers = cursor.fetchall()

    conn.close()

    return render_template("customers.html", customers=customers)
@app.route("/billing", methods=["GET", "POST"])
def billing():

    login_check = _require_login()
    if login_check is not None:
        return login_check

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("""
        SELECT id, name, mobile, vehicle_number, vehicle_model
        FROM customers
        ORDER BY name
    """)
    customers = cursor.fetchall()

    cursor.execute("""
        SELECT id, company, model, front_tyre, rear_tyre
        FROM vehicles
        ORDER BY company, model
    """)
    vehicles = cursor.fetchall()

    cursor.execute("""
        SELECT id, brand, tyre_size, tyre_name, price, stock
        FROM tyres
        ORDER BY brand, tyre_size
    """)
    tyres = cursor.fetchall()

    if request.method == "POST":

        customer_id = request.form.get("customer_id")
        vehicle_id = request.form.get("vehicle_id")
        customer_name = request.form.get("customer_name", "")
        vehicle_model = request.form.get("vehicle_model", "")
        front_tyre = request.form.get("front_tyre", "")
        rear_tyre = request.form.get("rear_tyre", "")
        front_qty = int(request.form.get("front_qty", 1) or 1)
        rear_qty = int(request.form.get("rear_qty", 1) or 1)
        discount = int(request.form.get("discount", 0) or 0)

        cursor.execute("SELECT name, mobile, vehicle_model FROM customers WHERE id=?", (customer_id,))
        customer_row = cursor.fetchone()
        if customer_row:
            customer_name = customer_row[0]
            vehicle_model = customer_row[2]

        cursor.execute("SELECT model FROM vehicles WHERE id=?", (vehicle_id,))
        vehicle_row = cursor.fetchone()
        if vehicle_row:
            vehicle_model = vehicle_row[0]

        front = cursor.execute("SELECT id, price, stock FROM tyres WHERE tyre_size=?", (front_tyre,)).fetchone()
        rear = cursor.execute("SELECT id, price, stock FROM tyres WHERE tyre_size=?", (rear_tyre,)).fetchone()

        if not front or not rear:
            flash("Please select valid front and rear tyres.", "danger")
            conn.close()
            return render_template("billing.html", customers=customers, vehicles=vehicles, tyres=tyres)

        if front[2] < front_qty or rear[2] < rear_qty:
            flash("Not enough stock available for one or more tyres.", "danger")
            conn.close()
            return render_template("billing.html", customers=customers, vehicles=vehicles, tyres=tyres)

        subtotal = (front[1] * front_qty) + (rear[1] * rear_qty)
        discount_amount = round(subtotal * discount / 100)
        total = subtotal - discount_amount

        cursor.execute("UPDATE tyres SET stock = stock - ? WHERE id=?", (front_qty, front[0]))
        cursor.execute("UPDATE tyres SET stock = stock - ? WHERE id=?", (rear_qty, rear[0]))

        cursor.execute("""
            INSERT INTO sales
            (
                customer_name,
                vehicle_model,
                front_tyre,
                rear_tyre,
                front_qty,
                rear_qty,
                total_amount
            )
            VALUES (?, ?, ?, ?, ?, ?, ?)
        """, (
            customer_name,
            vehicle_model,
            front_tyre,
            rear_tyre,
            front_qty,
            rear_qty,
            total
        ))

        sale_id = cursor.lastrowid
        conn.commit()
        conn.close()

        flash("Invoice generated successfully.", "success")
        return redirect(url_for("invoice", sale_id=sale_id))

    conn.close()

    return render_template("billing.html", customers=customers, vehicles=vehicles, tyres=tyres)
@app.route("/invoice/<int:sale_id>")
def invoice(sale_id):

    login_check = _require_login()
    if login_check is not None:
        return login_check

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("SELECT * FROM sales WHERE id=?", (sale_id,))
    sale = cursor.fetchone()

    if not sale:
        conn.close()
        return redirect(url_for("sales"))

    front_price = None
    rear_price = None
    if sale[3]:
        front_row = cursor.execute("SELECT price FROM tyres WHERE tyre_size=?", (sale[3],)).fetchone()
        if front_row:
            front_price = front_row[0]
    if sale[4]:
        rear_row = cursor.execute("SELECT price FROM tyres WHERE tyre_size=?", (sale[4],)).fetchone()
        if rear_row:
            rear_price = rear_row[0]

    conn.close()

    return render_template("invoice.html", sale=sale, front_price=front_price, rear_price=rear_price)


@app.route("/login", methods=["GET", "POST"])
def login():
    if request.method == "POST":
        username = request.form.get("username", "").strip().lower()
        password = request.form.get("password", "").strip()

        valid_users = {
            "admin": ["1234", "admin"],
            "employee": ["1234", "employee"],
            "saravana": ["saravana"]
        }

        if username in valid_users and password in valid_users[username]:
            session["user"] = username
            flash("Welcome back.", "success")
            if username == "employee":
                return redirect(url_for("billing"))
            return redirect(url_for("dashboard"))

        flash("Invalid username or password. Try admin/1234 or employee/1234.", "danger")
    return render_template("login.html")


@app.route("/logout")
def logout():
    session.pop("user", None)
    flash("You have been logged out.", "info")
    return redirect(url_for("login"))


@app.route("/search", methods=["GET", "POST"])
def search():
    query = ""
    tyres = []
    customers = []
    vehicles = []

    if request.method == "POST":
        query = request.form.get("query", "").strip()
        search = "%" + query + "%"
        conn = sqlite3.connect(DATABASE)
        cursor = conn.cursor()

        cursor.execute("SELECT * FROM tyres WHERE brand LIKE ? OR tyre_size LIKE ? OR tyre_name LIKE ?", (search, search, search))
        tyres = cursor.fetchall()

        cursor.execute("SELECT * FROM customers WHERE name LIKE ? OR mobile LIKE ? OR vehicle_model LIKE ?", (search, search, search))
        customers = cursor.fetchall()

        cursor.execute("SELECT * FROM vehicles WHERE company LIKE ? OR model LIKE ?", (search, search))
        vehicles = cursor.fetchall()

        conn.close()

    return render_template("search.html", query=query, tyres=tyres, customers=customers, vehicles=vehicles)


@app.route("/guide")
def guide():
    login_check = _require_login()
    if login_check is not None:
        return login_check
    return render_template("guide.html")


@app.route("/sales")
def sales():

    conn = sqlite3.connect(DATABASE)
    cursor = conn.cursor()

    cursor.execute("""
    SELECT *
    FROM sales
    ORDER BY id DESC
    """)

    sales = cursor.fetchall()

    conn.close()

    return render_template(
        "sales.html",
        sales=sales
    )
# -----------------------------
# RUN APP
# -----------------------------
if __name__ == "__main__":
    app.run(debug=False, host="0.0.0.0", port=5000)