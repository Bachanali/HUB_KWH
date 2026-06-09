# VOLTGRID — Energy Monitoring Dashboard

A Python (Flask) web app that visualises the `All_Department_Monthly_Kwh.xlsx`
workbook, with login, user management, role-based editing, live Excel sync and
Excel export. Runs on your local network so others can open it in a browser.

## Run

```bash
cd energy_app
pip install -r requirements.txt
python app.py
```

Then open:
- on this machine: **http://localhost:5010**
- from other devices on your network: **http://<this-PC-IP>:5010**
  (find your IP with `ipconfig` on Windows or `ip addr` / `ifconfig` on Linux/Mac)

> First run creates `data/users_config.xlsx` (an Excel file holding all user accounts) with a default admin:
> **username `admin` · password `admin123`** — change it after logging in
> (User Management → Reset PW).

## How data works (live Excel sync)

- The dashboard reads `data/All_Department_Monthly_Kwh.xlsx` **live**.
- Edit that file directly in Excel on the server, **or** use the in-app
  **“+ Add / Edit Data”** button — either way the dashboard refreshes
  automatically (it checks the file every 20 seconds).
- Sheet names = production units. Each sheet keeps its own column layout;
  the app auto-detects the total feed (Main LT / Sub LT), solar, HVAC and
  process loads per sheet.

## Roles & permissions

| Role / flag        | View | Add / edit data | Manage users |
|--------------------|:----:|:---------------:|:------------:|
| Admin              |  ✅  |       ✅        |      ✅      |
| User + “Can edit”  |  ✅  |       ✅        |      —       |
| User (view-only)   |  ✅  |       —         |      —       |

Admins create users in **User Management** and grant/revoke edit access per user.

## Features

- 3 KPI cards: Total Consumption, Solar Share, Peak Department
- Trend line, full per-department donut (every department named — no “Other”),
  Solar-vs-Grid stacked bars, and a ranked raw-data bar chart above the table
- Month filter for charts, year filter + search for the table
- Excel export for both the table (by year) and the chart breakdown (by month)

## Production note

For real deployment behind a proper server, set a strong secret and use a WSGI
server, e.g.:

```bash
export VOLTGRID_SECRET="a-long-random-string"
pip install waitress
waitress-serve --host=0.0.0.0 --port=5010 app:app
```
