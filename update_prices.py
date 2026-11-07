import os
import subprocess
import tempfile
from datetime import datetime, timezone
from zoneinfo import ZoneInfo

import pandas as pd
import requests
from openpyxl import load_workbook
from openpyxl.chart import LineChart, Reference
from openpyxl.styles import Font, PatternFill
from openpyxl.utils import get_column_letter

API_URL = "https://echoes.mobi/api/v1/prices"
TYPE_IDS = [42001000000, 42001000001, 42001000002, 42001000003, 42001000004, 42001000005, 42001000006, 42001000007, 42001000008, 42001000009, 42001000010, 42001000011, 42001000018, 42001000019, 42001000020, 42001000021, 42001000022, 42001000023, 42001000024, 42001000025, 42001000026, 42001000027, 42001000028, 42001000029, 42001000030, 42001000031, 42001000032, 42001000033, 42002000012, 42002000013, 42002000014, 42002000015, 42002000016, 42002000017]
TIMEZONE = ZoneInfo("Europe/Berlin")
WORKBOOK = "Echoes_Market_History.xlsx"

def run_rclone(*args, check=True):
    subprocess.run(["rclone", *args], check=check)

def fetch_prices():
    """Consulta conjunta y reintento individual para cada identificador ausente."""
    def request_ids(ids):
        params = {
            "lang": "en",
            "type_ids": ",".join(map(str, ids)),
            "page": 1,
            "page_size": 50,
        }
        response = requests.get(API_URL, params=params, timeout=60)
        response.raise_for_status()
        payload = response.json()
        results = payload.get("results")
        if not isinstance(results, list):
            raise RuntimeError("La API no devolvió una lista de resultados válida.")
        return payload, results

    payload, results = request_ids(TYPE_IDS)
    requested = set(TYPE_IDS)
    items_by_id = {}
    for item in results:
        type_id = int(item["type_id"])
        if type_id not in requested:
            raise RuntimeError(f"La API devolvió un ID no solicitado: {type_id}")
        if type_id in items_by_id:
            raise RuntimeError(f"La API devolvió un ID duplicado: {type_id}")
        items_by_id[type_id] = (item, payload)

    missing = [type_id for type_id in TYPE_IDS if type_id not in items_by_id]
    print(f"Consulta conjunta: {len(items_by_id)} de {len(TYPE_IDS)}; "
          f"reintentando individualmente {len(missing)} identificadores.")
    for type_id in missing:
        try:
            individual_payload, individual_results = request_ids([type_id])
            matches = [item for item in individual_results if int(item["type_id"]) == type_id]
            if len(matches) > 1:
                raise RuntimeError(f"La consulta individual devolvió duplicados para {type_id}")
            if matches:
                items_by_id[type_id] = (matches[0], individual_payload)
                print(f"Recuperado individualmente: {type_id}")
            else:
                print(f"Sin datos de API: {type_id}")
        except (requests.RequestException, ValueError, KeyError, RuntimeError) as exc:
            # No se inventan precios: queda registrado en Missing_Items y se reintenta mañana.
            print(f"No se pudo recuperar {type_id}: {exc}")

    if not items_by_id:
        raise RuntimeError("La API no devolvió ningún producto; no se modificará el Excel.")

    captured_at = datetime.now(timezone.utc).isoformat(timespec="seconds")
    rows = []
    for type_id in TYPE_IDS:
        if type_id not in items_by_id:
            continue
        item, source_payload = items_by_id[type_id]
        rows.append({
            "prices_as_of": item.get("prices_as_of") or source_payload.get("prices_as_of"),
            "base_prices_version": item.get("base_prices_version") or source_payload.get("base_prices_version"),
            "type_id": type_id,
            "name": item.get("name"),
            "avg": item.get("avg"),
            "avg_no_outliers": item.get("avg_no_outliers"),
            "base_price": item.get("base_price"),
            "captured_at_utc": captured_at,
        })
    df = pd.DataFrame(rows)
    print(f"Total recuperado: {len(df)} de {len(TYPE_IDS)} productos.")
    return df, payload

def update_workbook(new_df, payload, path):
    if path.exists():
        try:
            old = pd.read_excel(path, sheet_name="Historical_Data")
        except Exception:
            old = pd.DataFrame()
    else:
        old = pd.DataFrame()

    combined = pd.concat([old, new_df], ignore_index=True)
    # Un registro por producto y snapshot oficial de precios; evita duplicar al reejecutar.
    if not combined.empty:
        combined["type_id"] = pd.to_numeric(combined["type_id"], errors="coerce").astype("Int64")
        combined = combined.drop_duplicates(subset=["prices_as_of", "type_id"], keep="last")
        combined = combined.sort_values(["prices_as_of", "name", "type_id"], na_position="last")

    # Control de cobertura: mostrar los IDs no devueltos sin crear precios ficticios.
    received_ids = set(new_df["type_id"].astype(int))
    missing_ids = [type_id for type_id in TYPE_IDS if type_id not in received_ids]
    known_names = {}
    if not combined.empty and "type_id" in combined.columns and "name" in combined.columns:
        for _, record in combined.dropna(subset=["type_id"]).iterrows():
            if pd.notna(record.get("name")):
                known_names[int(record["type_id"])] = str(record["name"])
    missing = pd.DataFrame([{
        "type_id": type_id,
        "last_known_name": known_names.get(type_id, ""),
        "status": "Sin datos de API",
        "prices_as_of": payload.get("prices_as_of"),
        "checked_at_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
    } for type_id in missing_ids], columns=[
        "type_id", "last_known_name", "status", "prices_as_of", "checked_at_utc"
    ])

    latest = new_df.copy()
    latest["prices_as_of"] = pd.to_datetime(latest["prices_as_of"], utc=True, errors="coerce")
    latest = latest.sort_values("prices_as_of").drop_duplicates("type_id", keep="last")
    latest["prices_as_of"] = latest["prices_as_of"].astype(str)

    # Resumen de evolución: primera y última observación por producto.
    analysis_rows = []
    if not combined.empty:
        for type_id, group in combined.groupby("type_id", dropna=True):
            group = group.sort_values("prices_as_of")
            first = group.iloc[0]
            last = group.iloc[-1]
            first_price = pd.to_numeric(pd.Series([first.get("avg_no_outliers")]), errors="coerce").iloc[0]
            last_price = pd.to_numeric(pd.Series([last.get("avg_no_outliers")]), errors="coerce").iloc[0]
            change = None
            if pd.notna(first_price) and first_price != 0 and pd.notna(last_price):
                change = (last_price / first_price) - 1
            analysis_rows.append({
                "type_id": int(type_id),
                "name": last.get("name"),
                "first_observation": first.get("prices_as_of"),
                "latest_observation": last.get("prices_as_of"),
                "first_avg_no_outliers": first_price,
                "latest_avg_no_outliers": last_price,
                "change_since_first_pct": change,
                "observations": len(group),
            })
    analysis = pd.DataFrame(analysis_rows)

    with pd.ExcelWriter(path, engine="openpyxl") as writer:
        combined.to_excel(writer, sheet_name="Historical_Data", index=False)
        latest.to_excel(writer, sheet_name="Latest_Prices", index=False)
        missing.to_excel(writer, sheet_name="Missing_Items", index=False)
        analysis.to_excel(writer, sheet_name="Price_Analysis", index=False)
        pd.DataFrame([{
            "last_successful_update_utc": datetime.now(timezone.utc).isoformat(timespec="seconds"),
            "prices_as_of": payload.get("prices_as_of"),
            "base_prices_version": payload.get("base_prices_version"),
            "products_requested": len(TYPE_IDS),
            "products_returned": len(new_df),
            "products_missing": len(missing_ids),
            "missing_type_ids": ", ".join(map(str, missing_ids)),
        }]).to_excel(writer, sheet_name="Run_Info", index=False)

    wb = load_workbook(path)
    for ws in wb.worksheets:
        ws.freeze_panes = "A2"
        ws.auto_filter.ref = ws.dimensions
        for cell in ws[1]:
            cell.font = Font(bold=True, color="FFFFFF")
            cell.fill = PatternFill("solid", fgColor="244062")
        for col in ws.columns:
            max_len = max((len(str(c.value)) if c.value is not None else 0) for c in list(col)[:200]) 
            ws.column_dimensions[get_column_letter(col[0].column)].width = min(max(max_len + 2, 12), 32)
    if "Dashboard" in wb.sheetnames:
        del wb["Dashboard"]
    dash = wb.create_sheet("Dashboard", 0)
    dash["A1"] = "Echoes Market — evolución de precios"
    dash["A1"].font = Font(bold=True, size=16, color="244062")
    dash["A3"] = "El gráfico utiliza avg_no_outliers del histórico diario."
    dash["A4"] = "Para ver un producto, filtra Historical_Data por name o type_id."
    dash["A5"] = f"Cobertura API: {len(new_df)} de {len(TYPE_IDS)} productos; {len(missing_ids)} sin datos."
    dash["A5"].font = Font(bold=True, color="244062")
    dash["A6"] = "Consulta la hoja Missing_Items para identificar los productos ausentes."
    # Crear un gráfico de ejemplo para el primer producto con datos históricos.
    if not combined.empty:
        first_id = int(combined.iloc[0]["type_id"])
        series = combined[combined["type_id"] == first_id].copy()
        series["avg_no_outliers"] = pd.to_numeric(series["avg_no_outliers"], errors="coerce")
        series = series.dropna(subset=["avg_no_outliers"]).tail(100)
        if len(series) >= 2:
            dash["A8"] = "Fecha de precios"
            dash["B8"] = f"{series.iloc[-1]['name']} — avg_no_outliers"
            for idx, (_, row) in enumerate(series.iterrows(), start=9):
                dash.cell(idx, 1, str(row["prices_as_of"]))
                dash.cell(idx, 2, float(row["avg_no_outliers"]))
            chart = LineChart()
            chart.title = f"Precio sin atípicos — {series.iloc[-1]['name']}"
            chart.y_axis.title = "Precio"
            chart.x_axis.title = "Fecha del snapshot"
            data = Reference(dash, min_col=2, min_row=8, max_row=8 + len(series))
            cats = Reference(dash, min_col=1, min_row=9, max_row=8 + len(series))
            chart.add_data(data, titles_from_data=True)
            chart.set_categories(cats)
            chart.height = 9
            chart.width = 20
            dash.add_chart(chart, "D8")
    wb.save(path)

def main():
    # GitHub Actions dispara a las 06:00 y 07:00 UTC para cubrir el cambio CET/CEST.
    # Solo continúa cuando son las 08:00 en Alemania. La ejecución manual siempre continúa.
    manual = os.getenv("GITHUB_EVENT_NAME") == "workflow_dispatch"
    local_now = datetime.now(TIMEZONE)
    if not manual and local_now.hour != 8:
        print(f"Se omite esta ejecución: hora local {local_now:%H:%M}; objetivo 08:00.")
        return

    remote = os.environ["RCLONE_REMOTE"]
    remote_file = f"{remote}:Echoes_Market_History.xlsx"
    with tempfile.TemporaryDirectory() as tmp:
        path = __import__("pathlib").Path(tmp) / WORKBOOK
        # Descargar si existe; si no existe, empezar un libro nuevo.
        check = subprocess.run(["rclone", "copyto", remote_file, str(path)], capture_output=True, text=True)
        if check.returncode != 0 and ("not found" not in (check.stderr + check.stdout).lower()
                                      and "doesn't exist" not in (check.stderr + check.stdout).lower()
                                      and "object not found" not in (check.stderr + check.stdout).lower()):
            raise RuntimeError("No se pudo descargar el Excel de OneDrive: " + check.stderr[-1000:])

        new_df, payload = fetch_prices()
        update_workbook(new_df, payload, path)
        run_rclone("copyto", str(path), remote_file)
        print(f"Actualización completada: {len(new_df)} de {len(TYPE_IDS)} productos; "
              f"faltan {len(TYPE_IDS) - len(new_df)}; snapshot {payload.get('prices_as_of')}.")

if __name__ == "__main__":
    main()
