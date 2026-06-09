# Outsource Attendance App

Standalone Streamlit app for outsource attendance login, observer approval, admin override, IP capture, and Excel export.
The app uses a professional light theme with a white background.

## Run Locally

```powershell
pip install -r requirements.txt
streamlit run app.py
```

## Build Windows EXE

Build on a Windows machine:

```powershell
.\build_exe.ps1
```

The finished app is created at:

```text
dist\OutsourceAttendance\OutsourceAttendance.exe
```

Copy the whole `dist\OutsourceAttendance` folder to another PC and run
`OutsourceAttendance.exe`. The other PC does not need Python or pip packages
installed. The app opens in the user's browser and runs locally.

Packaged EXE builds require MongoDB Atlas. If `MONGODB_URI` is missing, the app
stops with setup instructions instead of saving new attendance data to SQLite.

Logs are written to:

```text
%LOCALAPPDATA%\OutsourceAttendance\logs\launcher.log
```

Before using the EXE in production, set a strong admin password on the PC that
runs it:

```powershell
[Environment]::SetEnvironmentVariable("DATALENS_ATTENDANCE_ADMIN_PASSWORD", "CHANGE_THIS_PASSWORD", "User")
```

For one shared attendance database across multiple PCs, configure MongoDB Atlas
or another central database, then set these values on each PC:

```powershell
[Environment]::SetEnvironmentVariable("MONGODB_URI", "mongodb+srv://USER:PASSWORD@CLUSTER.mongodb.net/?appName=Cluster0", "User")
[Environment]::SetEnvironmentVariable("MONGODB_DATABASE", "attendance_db", "User")
```

Do not hard-code real database passwords into the EXE.

## Streamlit Secrets

Add these in Streamlit Cloud secrets:

```toml
MONGODB_URI = "mongodb+srv://YOUR_DB_USER:YOUR_DB_PASSWORD@YOUR_CLUSTER.mongodb.net/?appName=Cluster0"
MONGODB_DATABASE = "attendance_db"
DATALENS_ATTENDANCE_ADMIN_PASSWORD = "admin123"
```

If `MONGODB_URI` is not provided, source runs can still use local SQLite at
`data/outsource_attendance.sqlite` for development. Packaged EXE runs do not use
the SQLite fallback.

## Pages

- `Outsource Login`: outsource user selects name, verifies mobile number, and submits login with the PC name captured automatically.
- `Observer Desk`: observer accepts or rejects pending entries.
- `Admin Panel`: admin creates users, overrides decisions, views attendance, sees captured IP address, and exports Excel.

## Attendance Percentage

Monthly attendance percentage is calculated from a fixed 26-day base:

```text
Attendance % = Total Present Days / 26 * 100
```

The value is capped at `100%`.

## GitHub Push

This folder is a separate Git repository. After creating a new empty repository on GitHub, run:

```powershell
git remote add origin https://github.com/YOUR_USERNAME/YOUR_REPOSITORY.git
git push -u origin main
```
