from flask import Flask, render_template, request, jsonify, Response
import pyodbc
from datetime import datetime
import math
from datetime import datetime
from dotenv import load_dotenv
import csv
import io,os

app = Flask(__name__)

# Database configuration - replace with actual credentials
load_dotenv()

DB_SERVER   = os.getenv("DB_SERVER")
DB_NAME     = os.getenv("DB_NAME")
DB_USER     = os.getenv("DB_USER")
DB_PASSWORD = os.getenv("DB_PASSWORD")


def get_db_connection():
    conn = pyodbc.connect(
        f'DRIVER={{ODBC Driver 17 for SQL Server}};'
        f'SERVER={DB_SERVER};'
        f'DATABASE={DB_NAME};'
        f'UID={DB_USER};'
        f'PWD={DB_PASSWORD}'
    )
    return conn

@app.route("/")
def index():
    # Renders templates/catalog.html
    return render_template("index.html")


# ---------- API: Paginated Revit Items ----------
@app.route("/api/revititems")
def api_revit_items():
    """
    Paginated + filtered list from dbo.RevitItems for the catalog page.
    Accepts query params: page, per_page, versions, units, categories, parametric, dynamo, q
    """
    import math

    page = max(int(request.args.get("page", 1)), 1)
    per_page = max(min(int(request.args.get("per_page", 1000)), 1000), 1)
    offset = (page - 1) * per_page

    where_sql, params = build_filter_clause(request.args)

    base_from = " FROM dbo.RevitItems AS ri "

    count_sql = f"SELECT COUNT(*) {base_from} {where_sql}"

    data_sql = f"""
        SELECT ri.item_id,
               ri.[Revit Categories],
               ri.[name],
               ri.[product_url],
               ri.[image_urls],
               ri.[raw_tags],
               ri.[Revit Version],
               ri.[Units System],
               ri.[Parametric],
               ri.[Dynamo Build],
               ri.[More filters],
               ri.[created_at]
        {base_from}
        {where_sql}
        ORDER BY ISNULL(ri.[created_at], '1900-01-01') DESC, ri.item_id DESC
        OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
    """

    conn = get_db_connection()
    cur = conn.cursor()

    # total
    cur.execute(count_sql, params)
    total = cur.fetchone()[0] or 0

    # page rows
    cur.execute(data_sql, params + [offset, per_page])
    rows = cur.fetchall()
    conn.close()

    items = []
    for r in rows:
        items.append({
            "item_id": int(r[0]),
            "Revit Categories": r[1],
            "name": r[2],
            "product_url": r[3],
            "image_urls": r[4] or "",
            "raw_tags": r[5],
            "Revit Version": r[6],
            "Units System": r[7],
            "Parametric": r[8],
            "Dynamo Build": r[9],
            "More filters": r[10],
            "created_at": r[11].strftime("%Y-%m-%d") if r[11] else None
        })

    return jsonify({
        "items": items,
        "total": total,
        "page": page,
        "per_page": per_page,
        "total_pages": math.ceil(total / per_page) if per_page else 1
    })
# ---------- API: Fetch User's Previous Selections ----------
@app.route("/api/user-selections")
def api_user_selections():
    """
    Query string: ?U_id=...
    Returns the user's previous submissions joined to item names.
    """
    user_id = (request.args.get("U_id") or "").strip()
    if not user_id:
        return jsonify({"items": []})

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT ui.item_id,
               ui.Y_N,
               ui.U_id,
               ri.[name],
               ri.[product_url]
        FROM dbo.RevitItems_byuser AS ui
        INNER JOIN dbo.RevitItems AS ri ON ri.item_id = ui.item_id
        WHERE ui.U_id = ?
        ORDER BY ui.id DESC
        """,
        (user_id,)
    )
    rows = cur.fetchall()
    conn.close()

    items = []
    for r in rows:
        items.append({
            "item_id": int(r[0]),
            "Y_N": (r[1] or "").upper(),
            "U_id": r[2],
            "name": r[3],
            "product_url": r[4]
        })

    return jsonify({"items": items})


# ---------- API: Submit Selections ----------
@app.route("/api/submit-selections", methods=["POST"])
def api_submit_selections():
    """
    Bulk UPSERT selections:
      - Dedupes by (U_id, item_id) client-side (last value wins)
      - MERGE into dbo.RevitItems_byuser in a single transaction
      - Returns inserted/updated counts

    Payload:
      {
        "U_id": "username",
        "selections": [
          {"item_id": 123, "Y_N": "YES"|"NO"},
          ...
        ]
      }
    """
    payload = request.get_json(silent=True) or {}
    user_id = (payload.get("U_id") or "").strip()
    raw = payload.get("selections") or []

    if not user_id:
        return jsonify({"success": False, "error": "Missing user ID"}), 400
    if not isinstance(raw, list) or len(raw) == 0:
        return jsonify({"success": False, "error": "No selections provided"}), 400

    # Deduplicate per (user_id, item_id) with last occurrence winning
    dedup = {}
    for sel in raw:
        try:
            item_id = int(sel.get("item_id"))
        except (TypeError, ValueError):
            continue  # skip bad item_id
        yn = (sel.get("Y_N") or "").strip().upper()
        yn = "YES" if yn == "YES" else "NO"
        dedup[item_id] = yn

    if not dedup:
        return jsonify({"success": False, "error": "No valid item_id entries found"}), 400

    # Prepare rows for temp table insert
    rows = [(user_id, iid, yn) for iid, yn in dedup.items()]

    try:
        conn = get_db_connection()
        conn.autocommit = False
        cur = conn.cursor()

        # Create a temp table for bulk payload
        cur.execute("""
            IF OBJECT_ID('tempdb..#sel') IS NOT NULL DROP TABLE #sel;
            CREATE TABLE #sel (
                U_id     NVARCHAR(255) NOT NULL,
                item_id  INT           NOT NULL,
                Y_N      NVARCHAR(3)   NOT NULL
            );
        """)

        # Bulk insert into #sel
        cur.fast_executemany = True
        cur.executemany("INSERT INTO #sel (U_id, item_id, Y_N) VALUES (?, ?, ?)", rows)

        # Optional: filter to existing items only (join RevitItems)
        # Comment this WHERE line if you want to allow any item_id.
        cur.execute("""
            ;WITH valid_sel AS (
                SELECT s.U_id, s.item_id, s.Y_N
                FROM #sel s
                INNER JOIN dbo.RevitItems ri ON ri.item_id = s.item_id
            )
            MERGE dbo.RevitItems_byuser AS T
            USING valid_sel AS S
              ON T.U_id = S.U_id AND T.item_id = S.item_id
            WHEN MATCHED AND ISNULL(T.Y_N,'') <> S.Y_N THEN
              UPDATE SET T.Y_N = S.Y_N
            WHEN NOT MATCHED BY TARGET THEN
              INSERT (item_id, Y_N, U_id)
              VALUES (S.item_id, S.Y_N, S.U_id)
            OUTPUT
              $action AS MergeAction;
        """)

        # Count results from OUTPUT (fetchall reads all rows)
        # Each row is ('INSERT'|'UPDATE')
        res = cur.fetchall() if cur.description else []
        inserted = sum(1 for r in res if str(r[0]).upper() == 'INSERT')
        updated  = sum(1 for r in res if str(r[0]).upper() == 'UPDATE')

        conn.commit()
        conn.close()

        return jsonify({
            "success": True,
            "message": f"Selections saved. Inserted: {inserted}, Updated: {updated}.",
            "inserted": inserted,
            "updated": updated
        })
    except pyodbc.Error as e:
        # Handle duplicate insert race if unique index exists
        try:
            conn.rollback()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500

def build_filter_clause(args, *, item_alias="ri", user_alias=None):
    """
    Multi-select filters:
      - OR within a single group (e.g., Units: Metric OR Imperial)
      - AND across groups (Version AND Units AND Categories AND Parametric AND Dynamo AND Q …)

    CSV-aware for columns that store comma-separated values:
      [Revit Version], [Units System], [Revit Categories]

    Query params (all optional):
      - versions=2023,2024,2025
      - units=Metric,Imperial
      - categories=Doors,Walls
      - parametric=Yes|No|all
      - dynamo=Yes|No|all
      - q=<free text>   (tokens ANDed; each token checks name OR raw_tags)
      - U_id=<username> (only if user_alias provided)
      - y_n=YES|NO      (only if user_alias provided)
    """
    where = []
    params = []

    def csv_list(key):
        return [v.strip() for v in (args.get(key) or "").split(",") if v.strip()]

    # Build a normalized, comma-padded UPPER() expression so we can token-match with LIKE
    # e.g. ",METRIC,IMPERIAL," and then LIKE '%,IMPERIAL,%'
    def csv_norm_expr(col, alias):
        return (
            "(',' + "
            f"REPLACE(REPLACE(UPPER(LTRIM(RTRIM({alias}.[{col}] ))), ', ', ','), ' ,', ',')"
            " + ',')"
        )

    # Adds (expr LIKE ? OR expr LIKE ? ...) to WHERE for multi-values
    def add_or_group_for_csv(col, values, alias):
        if not values:
            return
        expr = csv_norm_expr(col, alias)
        ors = []
        for v in values:
            ors.append(f"{expr} LIKE ?")
            params.append(f"%,{v.upper()},%")
        where.append("(" + " OR ".join(ors) + ")")

    # ----- User / selection filters (for management views) -----
    if user_alias:
        u_id = (args.get("U_id") or "").strip()
        if u_id:
            where.append(f"UPPER(LTRIM(RTRIM({user_alias}.[U_id]))) = UPPER(LTRIM(RTRIM(?)))")
            params.append(u_id)

        y_n = (args.get("y_n") or "").strip().upper()
        if y_n in ("YES", "NO"):
            where.append(f"UPPER(LTRIM(RTRIM({user_alias}.[Y_N]))) = ?")
            params.append(y_n)

    # ----- Item filters (CSV-aware) -----
    add_or_group_for_csv("Revit Version",  csv_list("versions"),   item_alias)
    add_or_group_for_csv("Units System",   csv_list("units"),      item_alias)
    add_or_group_for_csv("Revit Categories", csv_list("categories"), item_alias)

    # Parametric
    parametric = (args.get("parametric") or "").strip()
    if parametric and parametric.lower() != "all":
        where.append(f"UPPER(LTRIM(RTRIM({item_alias}.[Parametric]))) = UPPER(LTRIM(RTRIM(?)))")
        params.append(parametric)

    # Dynamo Build
    dynamo = (args.get("dynamo") or "").strip()
    if dynamo and dynamo.lower() != "all":
        where.append(f"UPPER(LTRIM(RTRIM({item_alias}.[Dynamo Build]))) = UPPER(LTRIM(RTRIM(?)))")
        params.append(dynamo)

    # Free text: AND all tokens; each token checks name OR raw_tags
    q = (args.get("q") or "").strip()
    if q:
        tokens = [t for t in q.split() if t]
        for tok in tokens:
            where.append(
                f"(UPPER({item_alias}.[name]) LIKE UPPER(?) OR UPPER({item_alias}.[raw_tags]) LIKE UPPER(?))"
            )
            like = f"%{tok}%"
            params.extend([like, like])

    where_sql = " WHERE " + " AND ".join(where) if where else ""
    return where_sql, params



@app.route("/dashboard")
def dashboard():
    # Renders templates/dashboard.html defined below
    return render_template("dashboard.html")
@app.route("/api/management/selections")
def api_management_selections():
    """
    Returns joined data:
      ui: dbo.RevitItems_byuser
      ri: dbo.RevitItems
    Query params:
      - page (default 1), per_page (default 25)
      - filters: U_id, versions, units, categories, parametric, dynamo, q, y_n
    """
    import math  # ensure available

    page = max(int(request.args.get("page", 1)), 1)
    per_page = max(min(int(request.args.get("per_page", 25)), 200), 1)
    offset = (page - 1) * per_page

    where_sql, params = build_filter_clause(request.args)

    # Base SELECT/JOIN
    base_from = """
      FROM dbo.RevitItems_byuser AS ui
      INNER JOIN dbo.RevitItems AS ri ON ri.item_id = ui.item_id
    """

    # Count
    count_sql = f"SELECT COUNT(*) {base_from} {where_sql}"

    # Page data
    data_sql = f"""
      SELECT ui.id,
             ui.U_id,
             ui.item_id,
             ui.Y_N,
             ri.[name],
             ri.[product_url],
             ri.[image_urls],
             ri.[raw_tags],
             ri.[Revit Version],
             ri.[Units System],
             ri.[Revit Categories],
             ri.[Parametric],
             ri.[Dynamo Build],
             ri.[created_at]
      {base_from}
      {where_sql}
      ORDER BY COALESCE(ri.[created_at], '1900-01-01') DESC, ui.id DESC
      OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
    """

    conn = get_db_connection()
    cur = conn.cursor()

    # total
    cur.execute(count_sql, params)
    total = cur.fetchone()[0] or 0

    # page data
    page_params = params + [offset, per_page]
    cur.execute(data_sql, page_params)
    rows = cur.fetchall()
    conn.close()

    # ---- safe serializers ----
    def as_int(v):
        try:
            return int(v) if v is not None and str(v).strip() != "" else None
        except Exception:
            return None

    def as_date_str(v):
        try:
            return v.strftime("%Y-%m-%d") if v else None
        except Exception:
            return None

    # ---- build response items ----
    items = []
    for r in rows:
        items.append({
            "id": as_int(r[0]),                    # ui.id may be NULL on some legacy rows
            "U_id": r[1],
            "item_id": as_int(r[2]),
            "Y_N": (r[3] or "").upper(),
            "name": r[4],
            "product_url": r[5],
            "image_urls": r[6] or "",
            "raw_tags": r[7],
            "Revit Version": r[8],
            "Units System": r[9],
            "Revit Categories": r[10],
            "Parametric": r[11],
            "Dynamo Build": r[12],
            "created_at": as_date_str(r[13]),
        })

    return jsonify({
        "items": items,
        "total": total,
        "page": page,
        "per_page": per_page,
        "total_pages": math.ceil(total / per_page) if per_page else 1
    })

@app.route("/api/management/summary/export")
def api_management_summary_export():
    """
    Export the current aggregated summary (with filters) to CSV (no pagination).
    Compatible with SQL Server < 2022.
    """
    where_sql, params = build_filter_clause(request.args)

    sql = f"""
WITH base AS (
  SELECT
    ui.U_id,
    ui.item_id,
    UPPER(LTRIM(RTRIM(ui.Y_N))) AS YN,
    ri.[name],
    ri.[product_url],
    ri.[created_at]
  FROM dbo.RevitItems_byuser AS ui
  INNER JOIN dbo.RevitItems      AS ri ON ri.item_id = ui.item_id
  {where_sql}
),
grouped AS (
  SELECT
    b.item_id,
    MAX(b.[name])          AS name,
    MAX(b.[product_url])   AS product_url,
    MAX(b.[created_at])    AS created_at,
    SUM(CASE WHEN b.YN='YES' THEN 1 ELSE 0 END) AS yes_count,
    SUM(CASE WHEN b.YN='NO'  THEN 1 ELSE 0 END) AS no_count
  FROM base b
  GROUP BY b.item_id
),
final AS (
  SELECT
    g.item_id,
    g.name,
    g.product_url,
    g.yes_count,
    g.no_count,
    STUFF((
      SELECT ', ' + b2.U_id
      FROM base b2
      WHERE b2.item_id = g.item_id AND b2.YN = 'YES'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS yes_users,
    STUFF((
      SELECT ', ' + b3.U_id
      FROM base b3
      WHERE b3.item_id = g.item_id AND b3.YN = 'NO'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS no_users,
    g.created_at
  FROM grouped g
)
SELECT item_id, name, product_url, yes_count, no_count, yes_users, no_users
FROM final
ORDER BY ISNULL(created_at, '1900-01-01') DESC, item_id DESC
"""

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()

    import io, csv
    from flask import Response
    from datetime import datetime

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["item_id", "name", "product_url", "yes_count", "no_count", "yes_users", "no_users"])
    for r in rows:
       writer.writerow([r[0], r[1], r[2] or "", r[3] or 0, r[4] or 0, (r[5] or ""), (r[6] or "")])

    csv_data = output.getvalue()
    output.close()

    filename = f"summary_export_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.route("/dashboard/summary")
def dashboard_summary():
    return render_template("dashboard.html")

@app.route("/api/management/summary")
def api_management_summary():
    """
    Aggregated view per item with YES/NO counts and user lists.
    Compatible with SQL Server < 2022 (no STRING_AGG WITHIN GROUP, no OFFSET/FETCH).
    Filters: U_id, versions, units, categories, parametric, dynamo, q, y_n
    Pagination: page, per_page
    """
    import math

    page = max(int(request.args.get("page", 1)), 1)
    per_page = max(min(int(request.args.get("per_page", 25)), 200), 1)
    start_row = (page - 1) * per_page + 1
    end_row = start_row + per_page - 1

    where_sql, params = build_filter_clause(request.args)

    # Base filtered set once; reuse it in subsequent CTEs/subqueries
    base_cte = f"""
WITH base AS (
  SELECT
    ui.U_id,
    ui.item_id,
    UPPER(LTRIM(RTRIM(ui.Y_N))) AS YN,
    ri.[name],
    ri.[product_url],
    ri.[created_at]
  FROM dbo.RevitItems_byuser AS ui
  INNER JOIN dbo.RevitItems      AS ri ON ri.item_id = ui.item_id
  {where_sql}
),
grouped AS (
  SELECT
    b.item_id,
    MAX(b.[name])          AS name,
    MAX(b.[product_url])   AS product_url,
    MAX(b.[created_at])    AS created_at,
    SUM(CASE WHEN b.YN='YES' THEN 1 ELSE 0 END) AS yes_count,
    SUM(CASE WHEN b.YN='NO'  THEN 1 ELSE 0 END) AS no_count
  FROM base b
  GROUP BY b.item_id
),
final AS (
  SELECT
    g.item_id,
    g.name,
    g.product_url,
    g.yes_count,
    g.no_count,
    STUFF((
      SELECT ', ' + b2.U_id
      FROM base b2
      WHERE b2.item_id = g.item_id AND b2.YN = 'YES'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS yes_users,
    STUFF((
      SELECT ', ' + b3.U_id
      FROM base b3
      WHERE b3.item_id = g.item_id AND b3.YN = 'NO'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS no_users,
    g.created_at
  FROM grouped g
),
numbered AS (
  SELECT
    item_id, name, product_url, yes_count, no_count, yes_users, no_users,
    ROW_NUMBER() OVER (ORDER BY ISNULL(created_at, '1900-01-01') DESC, item_id DESC) AS rn
  FROM final
)
SELECT item_id, name, product_url, yes_count, no_count, yes_users, no_users
FROM numbered
WHERE rn BETWEEN ? AND ?
"""

    # Count distinct items for pagination
    count_sql = f"""
    WITH base AS (
      SELECT
        ui.U_id, ui.item_id, UPPER(LTRIM(RTRIM(ui.Y_N))) AS YN, ri.[name], ri.[created_at]
      FROM dbo.RevitItems_byuser AS ui
      INNER JOIN dbo.RevitItems      AS ri ON ri.item_id = ui.item_id
      {where_sql}
    )
    SELECT COUNT(*) FROM (SELECT item_id FROM base GROUP BY item_id) x
    """

    conn = get_db_connection()
    cur = conn.cursor()

    # total distinct items
    cur.execute(count_sql, params)
    total = cur.fetchone()[0] or 0

    # page data
    cur.execute(base_cte, params + [start_row, end_row])
    rows = cur.fetchall()
    conn.close()

    def as_int(v):
        try:
            return int(v) if v is not None and str(v).strip() != "" else None
        except Exception:
            return None

    items = []
    for r in rows:
        items.append({
        "item_id": int(r[0]) if r[0] is not None else None,
        "name": r[1],
        "product_url": r[2],
        "yes_count": int(r[3] or 0),
        "no_count":  int(r[4] or 0),
        "yes_users": (r[5] or "").strip(", ").strip(),
        "no_users":  (r[6] or "").strip(", ").strip(),
    })
    return jsonify({
        "items": items,
        "total": total,
        "page": page,
        "per_page": per_page,
        "total_pages": math.ceil(total / per_page) if per_page else 1
    })

@app.route("/autodesk")
def index1():
    # Renders templates/catalog.html
    return render_template("autodesk_index.html")

# ---------- API: Paginated Revit Items ----------
@app.route("/api/revititems/autodesk")
def api_revit_items1():
    """
    Paginated + filtered list from dbo.RevitItems for the catalog page.
    Accepts query params: page, per_page, versions, units, categories, parametric, dynamo, q
    """
    import math

    page = max(int(request.args.get("page", 1)), 1)
    per_page = max(min(int(request.args.get("per_page", 1000)), 1000), 1)
    offset = (page - 1) * per_page

    where_sql, params = build_filter_clause(request.args)

    base_from = " FROM dbo.AccItems AS ri "

    count_sql = f"SELECT COUNT(*) {base_from} {where_sql}"

    data_sql = f"""
        SELECT ri.item_id,
               ri.[Revit Categories],
               ri.[name],
               ri.[product_url],
               ri.[image_urls],
               ri.[raw_tags],
               ri.[Revit Version],
               ri.[Units System],
               ri.[Parametric],
               ri.[Dynamo Build],
               ri.[More filters],
               ri.[created_at]
        {base_from}
        {where_sql}
        ORDER BY ISNULL(ri.[created_at], '1900-01-01') DESC, ri.item_id DESC
        OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
    """

    conn = get_db_connection()
    cur = conn.cursor()

    # total
    cur.execute(count_sql, params)
    total = cur.fetchone()[0] or 0

    # page rows
    cur.execute(data_sql, params + [offset, per_page])
    rows = cur.fetchall()
    conn.close()

    items = []
    for r in rows:
        items.append({
            "item_id": int(r[0]),
            "Revit Categories": r[1],
            "name": r[2],
            "product_url": r[3],
            "image_urls": r[4] or "",
            "raw_tags": r[5],
            "Revit Version": r[6],
            "Units System": r[7],
            "Parametric": r[8],
            "Dynamo Build": r[9],
            "More filters": r[10],
            "created_at": r[11].strftime("%Y-%m-%d") if r[11] else None
        })

    return jsonify({
        "items": items,
        "total": total,
        "page": page,
        "per_page": per_page,
        "total_pages": math.ceil(total / per_page) if per_page else 1
    })
# ---------- API: Fetch User's Previous Selections ----------
@app.route("/api/user-selections/autodesk")
def api_user_selections1():
    """
    Query string: ?U_id=...
    Returns the user's previous submissions joined to item names.
    """
    user_id = (request.args.get("U_id") or "").strip()
    if not user_id:
        return jsonify({"items": []})

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(
        """
        SELECT ui.item_id,
               ui.Y_N,
               ui.U_id,
               ri.[name],
               ri.[product_url]
        FROM dbo.ACCItems_byuser AS ui
        INNER JOIN dbo.AccItems AS ri ON ri.item_id = ui.item_id
        WHERE ui.U_id = ?
        ORDER BY ui.id DESC
        """,
        (user_id,)
    )
    rows = cur.fetchall()
    conn.close()

    items = []
    for r in rows:
        items.append({
            "item_id": int(r[0]),
            "Y_N": (r[1] or "").upper(),
            "U_id": r[2],
            "name": r[3],
            "product_url": r[4]
        })

    return jsonify({"items": items})


# ---------- API: Submit Selections ----------
@app.route("/api/submit-selections/autodesk", methods=["POST"])
def api_submit_selections1():
    """
    Bulk UPSERT selections:
      - Dedupes by (U_id, item_id) client-side (last value wins)
      - MERGE into dbo.ACCItems_byuser in a single transaction
      - Returns inserted/updated counts

    Payload:
      {
        "U_id": "username",
        "selections": [
          {"item_id": 123, "Y_N": "YES"|"NO"},
          ...
        ]
      }
    """
    payload = request.get_json(silent=True) or {}
    user_id = (payload.get("U_id") or "").strip()
    raw = payload.get("selections") or []

    if not user_id:
        return jsonify({"success": False, "error": "Missing user ID"}), 400
    if not isinstance(raw, list) or len(raw) == 0:
        return jsonify({"success": False, "error": "No selections provided"}), 400

    # Deduplicate per (user_id, item_id) with last occurrence winning
    dedup = {}
    for sel in raw:
        try:
            item_id = int(sel.get("item_id"))
        except (TypeError, ValueError):
            continue  # skip bad item_id
        yn = (sel.get("Y_N") or "").strip().upper()
        yn = "YES" if yn == "YES" else "NO"
        dedup[item_id] = yn

    if not dedup:
        return jsonify({"success": False, "error": "No valid item_id entries found"}), 400

    # Prepare rows for temp table insert
    rows = [(user_id, iid, yn) for iid, yn in dedup.items()]

    try:
        conn = get_db_connection()
        conn.autocommit = False
        cur = conn.cursor()

        # Create a temp table for bulk payload
        cur.execute("""
            IF OBJECT_ID('tempdb..#sel') IS NOT NULL DROP TABLE #sel;
            CREATE TABLE #sel (
                U_id     NVARCHAR(255) NOT NULL,
                item_id  INT           NOT NULL,
                Y_N      NVARCHAR(3)   NOT NULL
            );
        """)

        # Bulk insert into #sel
        cur.fast_executemany = True
        cur.executemany("INSERT INTO #sel (U_id, item_id, Y_N) VALUES (?, ?, ?)", rows)

        # Optional: filter to existing items only (join RevitItems)
        # Comment this WHERE line if you want to allow any item_id.
        cur.execute("""
            ;WITH valid_sel AS (
                SELECT s.U_id, s.item_id, s.Y_N
                FROM #sel s
                INNER JOIN dbo.AccItems ri ON ri.item_id = s.item_id
            )
            MERGE dbo.ACCItems_byuser AS T
            USING valid_sel AS S
              ON T.U_id = S.U_id AND T.item_id = S.item_id
            WHEN MATCHED AND ISNULL(T.Y_N,'') <> S.Y_N THEN
              UPDATE SET T.Y_N = S.Y_N
            WHEN NOT MATCHED BY TARGET THEN
              INSERT (item_id, Y_N, U_id)
              VALUES (S.item_id, S.Y_N, S.U_id)
            OUTPUT
              $action AS MergeAction;
        """)

        # Count results from OUTPUT (fetchall reads all rows)
        # Each row is ('INSERT'|'UPDATE')
        res = cur.fetchall() if cur.description else []
        inserted = sum(1 for r in res if str(r[0]).upper() == 'INSERT')
        updated  = sum(1 for r in res if str(r[0]).upper() == 'UPDATE')

        conn.commit()
        conn.close()

        return jsonify({
            "success": True,
            "message": f"Selections saved. Inserted: {inserted}, Updated: {updated}.",
            "inserted": inserted,
            "updated": updated
        })
    except pyodbc.Error as e:
        # Handle duplicate insert race if unique index exists
        try:
            conn.rollback()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500
    except Exception as e:
        try:
            conn.rollback()
        except Exception:
            pass
        return jsonify({"success": False, "error": str(e)}), 500

def build_filter_clause(args, *, item_alias="ri", user_alias=None):
    """
    Multi-select filters:
      - OR within a single group (e.g., Units: Metric OR Imperial)
      - AND across groups (Version AND Units AND Categories AND Parametric AND Dynamo AND Q …)

    CSV-aware for columns that store comma-separated values:
      [Revit Version], [Units System], [Revit Categories]

    Query params (all optional):
      - versions=2023,2024,2025
      - units=Metric,Imperial
      - categories=Doors,Walls
      - parametric=Yes|No|all
      - dynamo=Yes|No|all
      - q=<free text>   (tokens ANDed; each token checks name OR raw_tags)
      - U_id=<username> (only if user_alias provided)
      - y_n=YES|NO      (only if user_alias provided)
    """
    where = []
    params = []

    def csv_list(key):
        return [v.strip() for v in (args.get(key) or "").split(",") if v.strip()]

    # Build a normalized, comma-padded UPPER() expression so we can token-match with LIKE
    # e.g. ",METRIC,IMPERIAL," and then LIKE '%,IMPERIAL,%'
    def csv_norm_expr(col, alias):
        return (
            "(',' + "
            f"REPLACE(REPLACE(UPPER(LTRIM(RTRIM({alias}.[{col}] ))), ', ', ','), ' ,', ',')"
            " + ',')"
        )

    # Adds (expr LIKE ? OR expr LIKE ? ...) to WHERE for multi-values
    def add_or_group_for_csv(col, values, alias):
        if not values:
            return
        expr = csv_norm_expr(col, alias)
        ors = []
        for v in values:
            ors.append(f"{expr} LIKE ?")
            params.append(f"%,{v.upper()},%")
        where.append("(" + " OR ".join(ors) + ")")

    # ----- User / selection filters (for management views) -----
    if user_alias:
        u_id = (args.get("U_id") or "").strip()
        if u_id:
            where.append(f"UPPER(LTRIM(RTRIM({user_alias}.[U_id]))) = UPPER(LTRIM(RTRIM(?)))")
            params.append(u_id)

        y_n = (args.get("y_n") or "").strip().upper()
        if y_n in ("YES", "NO"):
            where.append(f"UPPER(LTRIM(RTRIM({user_alias}.[Y_N]))) = ?")
            params.append(y_n)

    # ----- Item filters (CSV-aware) -----
    add_or_group_for_csv("Revit Version",  csv_list("versions"),   item_alias)
    add_or_group_for_csv("Units System",   csv_list("units"),      item_alias)
    add_or_group_for_csv("Revit Categories", csv_list("categories"), item_alias)

    # Parametric
    parametric = (args.get("parametric") or "").strip()
    if parametric and parametric.lower() != "all":
        where.append(f"UPPER(LTRIM(RTRIM({item_alias}.[Parametric]))) = UPPER(LTRIM(RTRIM(?)))")
        params.append(parametric)

    # Dynamo Build
    dynamo = (args.get("dynamo") or "").strip()
    if dynamo and dynamo.lower() != "all":
        where.append(f"UPPER(LTRIM(RTRIM({item_alias}.[Dynamo Build]))) = UPPER(LTRIM(RTRIM(?)))")
        params.append(dynamo)

    # Free text: AND all tokens; each token checks name OR raw_tags
    q = (args.get("q") or "").strip()
    if q:
        tokens = [t for t in q.split() if t]
        for tok in tokens:
            where.append(
                f"(UPPER({item_alias}.[name]) LIKE UPPER(?) OR UPPER({item_alias}.[raw_tags]) LIKE UPPER(?))"
            )
            like = f"%{tok}%"
            params.extend([like, like])

    where_sql = " WHERE " + " AND ".join(where) if where else ""
    return where_sql, params



# @app.route("/dashboard")
# def dashboard():
#     # Renders templates/dashboard.html defined below
#     return render_template("dashboard.html")
@app.route("/api/management/selections/autodesk")
def api_management_selections1():
    """
    Returns joined data:
      ui: dbo.ACCItems_byuser
      ri: dbo.AccItems
    Query params:
      - page (default 1), per_page (default 25)
      - filters: U_id, versions, units, categories, parametric, dynamo, q, y_n
    """
    import math  # ensure available

    page = max(int(request.args.get("page", 1)), 1)
    per_page = max(min(int(request.args.get("per_page", 25)), 200), 1)
    offset = (page - 1) * per_page

    where_sql, params = build_filter_clause(request.args)

    # Base SELECT/JOIN
    base_from = """
      FROM dbo.ACCItems_byuser AS ui
      INNER JOIN dbo.AccItems AS ri ON ri.item_id = ui.item_id
    """

    # Count
    count_sql = f"SELECT COUNT(*) {base_from} {where_sql}"

    # Page data
    data_sql = f"""
      SELECT ui.id,
             ui.U_id,
             ui.item_id,
             ui.Y_N,
             ri.[name],
             ri.[product_url],
             ri.[image_urls],
             ri.[raw_tags],
             ri.[Revit Version],
             ri.[Units System],
             ri.[Revit Categories],
             ri.[Parametric],
             ri.[Dynamo Build],
             ri.[created_at]
      {base_from}
      {where_sql}
      ORDER BY COALESCE(ri.[created_at], '1900-01-01') DESC, ui.id DESC
      OFFSET ? ROWS FETCH NEXT ? ROWS ONLY
    """

    conn = get_db_connection()
    cur = conn.cursor()

    # total
    cur.execute(count_sql, params)
    total = cur.fetchone()[0] or 0

    # page data
    page_params = params + [offset, per_page]
    cur.execute(data_sql, page_params)
    rows = cur.fetchall()
    conn.close()

    # ---- safe serializers ----
    def as_int(v):
        try:
            return int(v) if v is not None and str(v).strip() != "" else None
        except Exception:
            return None

    def as_date_str(v):
        try:
            return v.strftime("%Y-%m-%d") if v else None
        except Exception:
            return None

    # ---- build response items ----
    items = []
    for r in rows:
        items.append({
            "id": as_int(r[0]),                    # ui.id may be NULL on some legacy rows
            "U_id": r[1],
            "item_id": as_int(r[2]),
            "Y_N": (r[3] or "").upper(),
            "name": r[4],
            "product_url": r[5],
            "image_urls": r[6] or "",
            "raw_tags": r[7],
            "Revit Version": r[8],
            "Units System": r[9],
            "Revit Categories": r[10],
            "Parametric": r[11],
            "Dynamo Build": r[12],
            "created_at": as_date_str(r[13]),
        })

    return jsonify({
        "items": items,
        "total": total,
        "page": page,
        "per_page": per_page,
        "total_pages": math.ceil(total / per_page) if per_page else 1
    })

@app.route("/api/management/summary/export/autodesk")
def api_management_summary_export1():
    """
    Export the current aggregated summary (with filters) to CSV (no pagination).
    Compatible with SQL Server < 2022.
    """
    where_sql, params = build_filter_clause(request.args)

    sql = f"""
WITH base AS (
  SELECT
    ui.U_id,
    ui.item_id,
    UPPER(LTRIM(RTRIM(ui.Y_N))) AS YN,
    ri.[name],
    ri.[product_url],
    ri.[created_at]
  FROM ACCItems_byuser AS ui
  INNER JOIN dbo.AccItems     AS ri ON ri.item_id = ui.item_id
  {where_sql}
),
grouped AS (
  SELECT
    b.item_id,
    MAX(b.[name])          AS name,
    MAX(b.[product_url])   AS product_url,
    MAX(b.[created_at])    AS created_at,
    SUM(CASE WHEN b.YN='YES' THEN 1 ELSE 0 END) AS yes_count,
    SUM(CASE WHEN b.YN='NO'  THEN 1 ELSE 0 END) AS no_count
  FROM base b
  GROUP BY b.item_id
),
final AS (
  SELECT
    g.item_id,
    g.name,
    g.product_url,
    g.yes_count,
    g.no_count,
    STUFF((
      SELECT ', ' + b2.U_id
      FROM base b2
      WHERE b2.item_id = g.item_id AND b2.YN = 'YES'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS yes_users,
    STUFF((
      SELECT ', ' + b3.U_id
      FROM base b3
      WHERE b3.item_id = g.item_id AND b3.YN = 'NO'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS no_users,
    g.created_at
  FROM grouped g
)
SELECT item_id, name, product_url, yes_count, no_count, yes_users, no_users
FROM final
ORDER BY ISNULL(created_at, '1900-01-01') DESC, item_id DESC
"""

    conn = get_db_connection()
    cur = conn.cursor()
    cur.execute(sql, params)
    rows = cur.fetchall()
    conn.close()

    import io, csv
    from flask import Response
    from datetime import datetime

    output = io.StringIO()
    writer = csv.writer(output)
    writer.writerow(["item_id", "name", "product_url", "yes_count", "no_count", "yes_users", "no_users"])
    for r in rows:
       writer.writerow([r[0], r[1], r[2] or "", r[3] or 0, r[4] or 0, (r[5] or ""), (r[6] or "")])

    csv_data = output.getvalue()
    output.close()

    filename = f"summary_export_{datetime.utcnow().strftime('%Y%m%d_%H%M%S')}.csv"
    return Response(
        csv_data,
        mimetype="text/csv",
        headers={"Content-Disposition": f"attachment; filename={filename}"}
    )

@app.route("/autodesk/dashboard/summary")
def dashboard_summary1():
    return render_template("autodesk_dashboard.html")

@app.route("/api/management/summary/autodesk")
def api_management_summary1():
    """
    Aggregated view per item with YES/NO counts and user lists.
    Compatible with SQL Server < 2022 (no STRING_AGG WITHIN GROUP, no OFFSET/FETCH).
    Filters: U_id, versions, units, categories, parametric, dynamo, q, y_n
    Pagination: page, per_page
    """
    import math

    page = max(int(request.args.get("page", 1)), 1)
    per_page = max(min(int(request.args.get("per_page", 25)), 200), 1)
    start_row = (page - 1) * per_page + 1
    end_row = start_row + per_page - 1

    where_sql, params = build_filter_clause(request.args)

    # Base filtered set once; reuse it in subsequent CTEs/subqueries
    base_cte = f"""
WITH base AS (
  SELECT
    ui.U_id,
    ui.item_id,
    UPPER(LTRIM(RTRIM(ui.Y_N))) AS YN,
    ri.[name],
    ri.[product_url],
    ri.[created_at]
  FROM dbo.ACCItems_byuser AS ui
  INNER JOIN dbo.AccItems      AS ri ON ri.item_id = ui.item_id
  {where_sql}
),
grouped AS (
  SELECT
    b.item_id,
    MAX(b.[name])          AS name,
    MAX(b.[product_url])   AS product_url,
    MAX(b.[created_at])    AS created_at,
    SUM(CASE WHEN b.YN='YES' THEN 1 ELSE 0 END) AS yes_count,
    SUM(CASE WHEN b.YN='NO'  THEN 1 ELSE 0 END) AS no_count
  FROM base b
  GROUP BY b.item_id
),
final AS (
  SELECT
    g.item_id,
    g.name,
    g.product_url,
    g.yes_count,
    g.no_count,
    STUFF((
      SELECT ', ' + b2.U_id
      FROM base b2
      WHERE b2.item_id = g.item_id AND b2.YN = 'YES'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS yes_users,
    STUFF((
      SELECT ', ' + b3.U_id
      FROM base b3
      WHERE b3.item_id = g.item_id AND b3.YN = 'NO'
      FOR XML PATH(''), TYPE
    ).value('.', 'nvarchar(max)'), 1, 2, '') AS no_users,
    g.created_at
  FROM grouped g
),
numbered AS (
  SELECT
    item_id, name, product_url, yes_count, no_count, yes_users, no_users,
    ROW_NUMBER() OVER (ORDER BY ISNULL(created_at, '1900-01-01') DESC, item_id DESC) AS rn
  FROM final
)
SELECT item_id, name, product_url, yes_count, no_count, yes_users, no_users
FROM numbered
WHERE rn BETWEEN ? AND ?
"""

    # Count distinct items for pagination
    count_sql = f"""
    WITH base AS (
      SELECT
        ui.U_id, ui.item_id, UPPER(LTRIM(RTRIM(ui.Y_N))) AS YN, ri.[name], ri.[created_at]
      FROM dbo.ACCItems_byuser AS ui
      INNER JOIN dbo.AccItems AS ri ON ri.item_id = ui.item_id
      {where_sql}
    )
    SELECT COUNT(*) FROM (SELECT item_id FROM base GROUP BY item_id) x
    """

    conn = get_db_connection()
    cur = conn.cursor()

    # total distinct items
    cur.execute(count_sql, params)
    total = cur.fetchone()[0] or 0

    # page data
    cur.execute(base_cte, params + [start_row, end_row])
    rows = cur.fetchall()
    conn.close()

    def as_int(v):
        try:
            return int(v) if v is not None and str(v).strip() != "" else None
        except Exception:
            return None

    items = []
    for r in rows:
        items.append({
        "item_id": int(r[0]) if r[0] is not None else None,
        "name": r[1],
        "product_url": r[2],
        "yes_count": int(r[3] or 0),
        "no_count":  int(r[4] or 0),
        "yes_users": (r[5] or "").strip(", ").strip(),
        "no_users":  (r[6] or "").strip(", ").strip(),
    })
    return jsonify({
        "items": items,
        "total": total,
        "page": page,
        "per_page": per_page,
        "total_pages": math.ceil(total / per_page) if per_page else 1
    })
if __name__ == "__main__":
    app.run(debug=True, port=5010, host="0.0.0.0")