# Revit Products Page

A Flask web application for browsing and managing Revit and Autodesk product catalogs, allowing users to submit YES/NO selections per item and administrators to view aggregated results via a dashboard.

## Project Structure

```
revit_products_page/
├── app.py                        # Main app (Revit 3D + Autodesk routes)
├── autodesk.py                   # Standalone Autodesk-only app (port 5011)
├── .env                          # Environment variables (not committed)
├── static/
│   └── A.jpg
└── templates/
    ├── index.html                # Revit catalog page
    ├── autodesk_index.html       # Autodesk catalog page
    ├── dashboard.html            # Revit dashboard
    └── autodesk_dashboard.html   # Autodesk dashboard
```

## Prerequisites

- Python 3.8+
- [ODBC Driver 17 for SQL Server](https://learn.microsoft.com/en-us/sql/connect/odbc/download-odbc-driver-for-sql-server)
- A SQL Server database with the following tables:
  - `dbo.RevitItems`
  - `dbo.RevitItems_byuser`
  - `dbo.AccItems`
  - `dbo.ACCItems_byuser`

## Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/Ashma1122/revit_products_page.git
   cd revit_products_page
   ```

2. **Create and activate a virtual environment**
   ```bash
   python -m venv .venv

   # Windows
   .venv\Scripts\activate

   # macOS/Linux
   source .venv/bin/activate
   ```

3. **Install dependencies**
   ```bash
   pip install flask pyodbc python-dotenv
   ```

4. **Set up environment variables**

   Create a `.env` file in the root directory:
   ```env
   DB_SERVER=your_server_name
   DB_NAME=your_database_name
   DB_USER=your_username
   DB_PASSWORD=your_password
   ```

## Running the App

**Main app** (Revit + Autodesk, port 5010):
```bash
python app.py
```

**Autodesk-only app** (port 5011):
```bash
python autodesk.py
```

## Routes

### Main App (`app.py`)

| Route | Description |
|---|---|
| `/` | Revit product catalog |
| `/autodesk` | Autodesk product catalog |
| `/dashboard` | Revit selections dashboard |
| `/dashboard/summary` | Revit summary dashboard |
| `/autodesk/dashboard/summary` | Autodesk summary dashboard |

### Revit API

| Method | Endpoint | Description |
|---|---|---|
| GET | `/api/revititems` | Paginated + filtered Revit items |
| GET | `/api/user-selections` | Fetch a user's previous selections |
| POST | `/api/submit-selections` | Submit YES/NO selections |
| GET | `/api/management/selections` | Admin view of all selections |
| GET | `/api/management/summary` | Aggregated YES/NO counts per item |
| GET | `/api/management/summary/export` | Export summary as CSV |

### Autodesk API

| Method | Endpoint | Description |
|---|---|---|
| GET | `/api/revititems/autodesk` | Paginated + filtered Autodesk items |
| GET | `/api/user-selections/autodesk` | Fetch a user's previous selections |
| POST | `/api/submit-selections/autodesk` | Submit YES/NO selections |
| GET | `/api/management/selections/autodesk` | Admin view of all selections |
| GET | `/api/management/summary/autodesk` | Aggregated YES/NO counts per item |
| GET | `/api/management/summary/export/autodesk` | Export summary as CSV |

## API Query Parameters

### Catalog endpoints (`/api/revititems`, `/api/revititems/autodesk`)

| Param | Example | Description |
|---|---|---|
| `page` | `1` | Page number |
| `per_page` | `25` | Items per page (max 1000) |
| `versions` | `2023,2024` | Filter by Revit version (OR logic) |
| `units` | `Metric,Imperial` | Filter by units system (OR logic) |
| `categories` | `Doors,Walls` | Filter by category (OR logic) |
| `parametric` | `Yes` | Filter parametric items |
| `dynamo` | `No` | Filter Dynamo Build items |
| `q` | `window frame` | Free-text search (AND logic) |

### Submit selections (`POST /api/submit-selections`)

```json
{
  "U_id": "username",
  "selections": [
    { "item_id": 123, "Y_N": "YES" },
    { "item_id": 456, "Y_N": "NO" }
  ]
}
```

## Notes

- The app uses SQL Server's `MERGE` statement for upserts, requiring SQL Server 2008+.
- Summary/export queries use `FOR XML PATH` for string aggregation, compatible with SQL Server versions before 2022.
- Never commit your `.env` file — it is listed in `.gitignore`.
