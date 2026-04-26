"""
update_facturacion.py
─────────────────────────────────────────────────────────────────────────────
Lee todos los archivos "Cierres de cajas*.xlsx" de la carpeta OneDrive,
extrae los totales por fecha y turno (Desayuno->Manana, Comida->Mediodia,
Cena->Noche), y actualiza la hoja 'Facturacion 2026' del Excel principal
en aquellas fechas donde el total actual es 0 (sin datos todavia).

Variables de entorno necesarias (igual que update_horarios.py):
  MS_TENANT_ID, MS_CLIENT_ID, MS_CLIENT_SECRET
  ONEDRIVE_SHARE_URL  (share URL del Excel principal)
  [opcional] DRY_RUN=1  imprime cambios sin subir
"""
from __future__ import annotations
import base64, datetime, io, os, sys, time
import requests
from openpyxl import load_workbook

SHEET_FAC = "Facturacion 2026"
DRY_RUN   = os.environ.get("DRY_RUN", "").strip() == "1"
TURNO_COL = {"desayuno": 3, "comida": 4, "cena": 5}
COL_TOTAL = 6

def _graph_token():
    resp = requests.post(
        f"https://login.microsoftonline.com/{os.environ['MS_TENANT_ID']}/oauth2/v2.0/token",
        data={"client_id": os.environ["MS_CLIENT_ID"], "client_secret": os.environ["MS_CLIENT_SECRET"],
              "scope": "https://graph.microsoft.com/.default", "grant_type": "client_credentials"},
        timeout=30)
    resp.raise_for_status()
    return resp.json()["access_token"]

def _share_item_url():
    enc = "u!" + base64.urlsafe_b64encode(os.environ["ONEDRIVE_SHARE_URL"].encode()).decode().rstrip("=")
    return f"https://graph.microsoft.com/v1.0/shares/{enc}/driveItem"

def _get_main_item(token):
    r = requests.get(_share_item_url(), headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    return r.json()

def list_folder_files(token):
    item = _get_main_item(token)
    drive_id  = item["parentReference"]["driveId"]
    parent_id = item["parentReference"]["id"]
    headers   = {"Authorization": f"Bearer {token}"}
    files = []
    url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{parent_id}/children?$top=200"
    while url:
        r = requests.get(url, headers=headers, timeout=30)
        r.raise_for_status()
        d = r.json()
        files.extend(d.get("value", []))
        url = d.get("@odata.nextLink")
    return files, drive_id

def download_item(token, drive_id, item_id):
    r = requests.get(f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content",
        headers={"Authorization": f"Bearer {token}"}, allow_redirects=True, stream=True, timeout=60)
    r.raise_for_status()
    return r.content

def upload_excel(token, drive_id, item_id, data):
    put_url = f"https://graph.microsoft.com/v1.0/drives/{drive_id}/items/{item_id}/content"
    headers = {"Authorization": f"Bearer {token}",
               "Content-Type": "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"}
    for intento, espera in enumerate([15, 30, 60, 120, 300], start=1):
        r = requests.put(put_url, headers=headers, data=data, timeout=120)
        if r.status_code == 423:
            print(f"OneDrive bloqueado (423) intento {intento}/5. Esperando {espera}s...")
            time.sleep(espera)
            continue
        r.raise_for_status()
        return
    r.raise_for_status()

def parse_cierres(raw):
    wb = load_workbook(io.BytesIO(raw), read_only=True, data_only=True)
    ws = wb.active
    result = {}
    for row in ws.iter_rows(min_row=2, values_only=True):
        if not row or len(row) < 16: continue
        fecha_raw, _, _, turno_raw = row[0], row[1], row[2], row[3]
        total_raw = row[15]
        if isinstance(fecha_raw, datetime.datetime):
            fecha = fecha_raw.date()
        elif isinstance(fecha_raw, str):
            try: fecha = datetime.datetime.strptime(fecha_raw.strip(), "%d/%m/%Y").date()
            except: continue
        else: continue
        if not turno_raw or not isinstance(total_raw, (int, float)) or total_raw <= 0: continue
        col = TURNO_COL.get(str(turno_raw).strip().lower())
        if col is None: continue
        if fecha not in result: result[fecha] = {}
        result[fecha][col] = round(result[fecha].get(col, 0.0) + float(total_raw), 2)
    return result

def update_excel(excel_raw, cierres):
    wb_ro = load_workbook(io.BytesIO(excel_raw), read_only=True, data_only=True)
    ws_ro = wb_ro[SHEET_FAC]
    pendientes = {}
    for i, row in enumerate(ws_ro.iter_rows(min_row=4, values_only=True), start=4):
        if not isinstance(row[0], datetime.datetime): continue
        fecha = row[0].date()
        if fecha not in cierres: continue
        total = row[5]
        if not isinstance(total, (int, float)) or total == 0:
            pendientes[fecha] = i
    wb_ro.close()
    if not pendientes: return excel_raw, 0
    wb = load_workbook(io.BytesIO(excel_raw))
    ws = wb[SHEET_FAC]
    updated = 0
    for fecha, row_num in sorted(pendientes.items()):
        day = cierres[fecha]
        ma, md, no = day.get(3, 0.0), day.get(4, 0.0), day.get(5, 0.0)
        tot = round(ma + md + no, 2)
        ws.cell(row_num, 3).value = ma
        ws.cell(row_num, 4).value = md
        ws.cell(row_num, 5).value = no
        ws.cell(row_num, 6).value = tot
        print(f"   {fecha.strftime('%d/%m/%Y')}  MA={ma:.2f}  MD={md:.2f}  NO={no:.2f}  TOT={tot:.2f}")
        updated += 1
    buf = io.BytesIO()
    wb.save(buf)
    return buf.getvalue(), updated

def main():
    print("Autenticando con Microsoft Graph...")
    token = _graph_token()
    print("Listando carpeta OneDrive...")
    files, drive_id = list_folder_files(token)
    cierres_files = [f for f in files if "cierres de cajas" in f["name"].lower() and f["name"].lower().endswith(".xlsx")]
    print(f"   Archivos de cierres: {len(cierres_files)}")
    if not cierres_files:
        print("No hay archivos de cierres. Nada que hacer.")
        return 0
    all_cierres = {}
    for cf in cierres_files:
        print(f"   Leyendo: {cf['name']}")
        raw = download_item(token, drive_id, cf["id"])
        for fecha, turnos in parse_cierres(raw).items():
            if fecha not in all_cierres: all_cierres[fecha] = {}
            for col, total in turnos.items():
                all_cierres[fecha][col] = max(all_cierres[fecha].get(col, 0.0), total)
    print(f"   Total dias con cierres: {len(all_cierres)}")
    print("Descargando Excel principal...")
    item      = _get_main_item(token)
    item_id   = item["id"]
    excel_raw = download_item(token, drive_id, item_id)
    print(f"   {len(excel_raw)//1024} KB")
    print(f"Actualizando hoja '{SHEET_FAC}'...")
    new_excel, n = update_excel(excel_raw, all_cierres)
    if n == 0:
        print("Todas las fechas ya tienen datos.")
        return 0
    print(f"   {n} dia(s) actualizado(s).")
    if DRY_RUN:
        print("DRY_RUN=1 — no se sube.")
        return 0
    print(f"Subiendo Excel ({len(new_excel)//1024} KB)...")
    upload_excel(token, drive_id, item_id, new_excel)
    print("OK — facturacion actualizada.")
    return 0

if __name__ == "__main__":
    sys.exit(main())
