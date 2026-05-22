# Outsource Attendance App

Standalone Streamlit app for outsource attendance login, observer approval, admin override, IP capture, and Excel export.

## Run Locally

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## Streamlit Secrets

Add these in Streamlit Cloud secrets:

```toml
MONGODB_URI = "mongodb+srv://YOUR_DB_USER:YOUR_DB_PASSWORD@YOUR_CLUSTER.mongodb.net/?appName=Cluster0"
MONGODB_DATABASE = "attendance_db"
DATALENS_ATTENDANCE_ADMIN_PASSWORD = "admin123"
```

If `MONGODB_URI` is not provided, the app stores data in local SQLite at `data/outsource_attendance.sqlite`.

## Pages

- `Outsource Login`: outsource user selects name, verifies mobile number, enters PC name, and submits login.
- `Observer Desk`: observer accepts or rejects pending entries.
- `Admin Panel`: admin creates users, overrides decisions, views attendance, sees captured IP address, and exports Excel.

## GitHub Push

This folder is a separate Git repository. After creating a new empty repository on GitHub, run:

```powershell
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
git push -u origin main
```
