# Saravana Auto Parts

A Flask-based inventory and sales management application for an auto parts store. This repository includes tyre inventory management, customer tracking, sales reporting, vehicle data import/export, and dashboard analytics.

## Features

- Tyre inventory search and listing
- Add, update, and delete tyre records
- Customer management
- Sales recording and invoice generation
- Dashboard with sales metrics and low-stock alerts
- Vehicle master data import/export from Excel
- Login-protected admin routes

## Requirements

- Python 3.11+ recommended
- Flask
- pandas
- openpyxl
- SQLite (built-in with Python)

## Setup

1. Clone or copy the repository.
2. Create and activate a virtual environment:

```bash
python -m venv venv
venv\Scripts\activate
```

3. Install dependencies:

```bash
pip install -r requirements.txt
```

4. Create the SQLite database and tables:

```bash
python database/create_db.py
```

5. Optionally populate initial data:

```bash
python database/insert_data.py
```

6. Run the Flask application:

```bash
python app.py
```

7. Open the app in your browser:

```text
http://127.0.0.1:5000/
```

## Database

The app uses `database/database.db` with the following tables:

- `tyres`
- `price_history`
- `customers`
- `sales`

## Important Files

- `app.py` - Main Flask application
- `requirements.txt` - Python dependencies
- `database/create_db.py` - Creates SQLite tables
- `database/insert_data.py` - Seed sample data
- `templates/` - HTML templates
- `static/` - CSS, JS, and images

## Notes

- The app uses a hardcoded secret key in `app.py`. For production use, replace it with a secure environment variable.
- The master vehicle Excel file path is currently set to `c:\Users\s8162\OneDrive\Desktop\PROJECT NO 1.xlsx`. Update this path if needed.

## Licence

This repository does not include a license file. Add one if you want to publish or share it publicly.
