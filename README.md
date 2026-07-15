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

## Local start script

A helper script is available to start the app on Windows.

- Double-click `start_app.bat` to launch the server.

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

## Deployment

This app can be deployed online using a platform such as Render, Railway, or PythonAnywhere.

### Recommended steps

1. Move the master Excel file into the repository folder or upload it to the host, then update `MASTER_VEHICLE_FILE` in `app.py` to a relative path, e.g. `"PROJECT NO 1.xlsx"`.
2. Create the SQLite database before the first run:

```bash
python database/create_db.py
```

If you already have an existing `database/database.db`, run the migration to add billing fields:

```bash
python database/migrate_sales_schema.py
```

3. For services like Render/Railway, add `gunicorn` to `requirements.txt` and use a deployment command such as:

```bash
gunicorn app:app
```

4. For PythonAnywhere, configure the web app to use the Flask app from `app.py` and set the working directory to the repository root.

### Deploy to Vercel

1. Install Vercel CLI if needed:

```bash
npm install -g vercel
```

2. Log in to Vercel:

```bash
vercel login
```

3. Deploy from the repository root:

```bash
vercel
```

4. For a production deployment, run:

```bash
vercel --prod
```

### Important Vercel notes

- Vercel uses serverless functions and an ephemeral filesystem.
- This means `database/database.db` can be deployed, but runtime changes are not reliable long-term.
- For a production-ready app, move to a hosted database such as PostgreSQL or MySQL.
- Ensure `PROJECT NO 1.xlsx` is in the repo root before deploying.

### Deploy to Render with Postgres (recommended for persistence)

1. Create a new Web Service on Render and link your GitHub repository.
2. Create a new Postgres database on Render (Managed Databases) and copy the `DATABASE_URL`.
3. In the Render service environment variables, set `DATABASE_URL` to the Postgres DSN.
4. Set the start command to:

```bash
gunicorn app:app
```

5. Deploy. After first deploy, run the migration script to copy existing SQLite data into Postgres (run from server shell or locally with `DATABASE_URL` set):

```bash
python database/sqlite_to_postgres.py
```

6. Verify data in Postgres and then point your app to the Postgres `DATABASE_URL` in Render.

### SQLite and file-based storage

- `database/database.db` is a file-based SQLite database. This is fine for small deployments, but it is not ideal for large, production use.
- If you deploy to a service with ephemeral storage, the database file may be reset when the instance restarts. For long-term use, consider moving to PostgreSQL or MySQL.

## Notes

- The app uses a hardcoded secret key in `app.py`. For production use, replace it with a secure environment variable.
- The master vehicle Excel file path is currently set to `c:\Users\s8162\OneDrive\Desktop\PROJECT NO 1.xlsx`. Update this path to a relative or hosted file path.

## Licence

This repository does not include a license file. Add one if you want to publish or share it publicly.
