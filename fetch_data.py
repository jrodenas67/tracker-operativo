#!/usr/bin/env python3
"""
fetch_data.py — Descarga el Excel desde OneDrive/SharePoint (Microsoft Graph)
               y actualiza index.html. Mantiene compatibilidad con Google Drive
               como fallback legacy.

Estructura esperada del Excel (se detectan automáticamente por nombre de columna):
  Hoja "Diario"   → fecha | mañana | mediodía | noche | previsto | coste | evento
  Hoja "Personal" → nombre | €/hora | horas | total | pct | costeMa | costeMd | costeNo
  Hoja "Eventos"  → mes | fecha | evento | tipo | mult | prev | real | estado
  Hoja "Productos"→ producto | familia | uds | importe | pct

Variables de entorno (OneDrive/SharePoint vía Microsoft Graph — prioritario):
  MS_TENANT_ID         → Azure AD tenant ID
  MS_CLIENT_ID         → App (client) ID
  MS_CLIENT_SECRET     → Client secret
  ONEDRIVE_SHARE_URL   → URL compartida del Excel (recomendado)
  — o alternativamente —
  ONEDRIVE_USER        → UPN del propietario (ej. juan@company.onmicrosoft.com)
  ONEDRIVE_FILE_PATH   → Ruta relativa al archivo desde la raíz del drive

Variables de entorno (Google Drive — legacy, fallback):
  GOOGLE_DRIVE_FILE_ID → ID del archivo en Google Drive

Configuración:
  TARGET_HTML → Ruta al HTML (default: index.html)
"""

import os, sys, re, json, base64
from pathlib import Path
from datetime import datetime

# ── Dependencias ─────────────────────────────────────────────────────────────
def _pip(pkg):
    os.system(f"{sys.executable} -m pip install {pkg} -q")

try: import requests
except ImportError: _pip("requests"); import requests

try: import openpyxl
except ImportError: _pip("openpyxl"); import openpyxl

# ── Config ────────────────────────────────────────────────────────────────────
# Microsoft Graph (prioritario)
MS_TENANT_ID       = os.environ.get("MS_TENANT_ID", "")
MS_CLIENT_ID       = os.environ.get("MS_CLIENT_ID", "")
MS_CLIENT_SECRET   = os.environ.get("MS_CLIENT_SECRET", "")
ONEDRIVE_SHARE_URL = os.environ.get("ONEDRIVE_SHARE_URL", "")
ONEDRIVE_USER      = os.environ.get("ONEDRIVE_USER", "")
ONEDRIVE_FILE_PATH = os.environ.get("ONEDRIVE_FILE_PATH", "")

# Google Drive (legacy)
FILE_ID    = os.environ.get("GOOGLE_DRIVE_FILE_ID", "")

ROOT       = Path(__file__).parent
XLSX_PATH  = ROOT / "source.xlsx"
HTML_PATH  = ROOT / os.environ.get("TARGET_HTML", "index.html")

DIAS_ES = {"Monday":"lun","Tuesday":"mar","Wednesday":"mié",
           "Thursday":"jue","Friday":"vie","Saturday":"sáb","Sunday":"dom"}
MESES   = ["Enero","Febrero","Marzo","Abril","Mayo","Junio",
           "Julio","Agosto","Septiembre","Octubre","Noviembre","Diciembre"]

# ── Helpers ───────────────────────────────────────────────────────────────────
def flt(v, d=0.0):
    if v is None: return d
    if isinstance(v, (int, float)): return float(v)
    try: return float(str(v).replace(",",".").replace("€","").strip())
    except: return d

def parse_date(v):
    if isinstance(v, datetime): return v.date()
    if hasattr(v, "date"): return v.date()
    if isinstance(v, str):
        for fmt in ("%Y-%m-%d","%d/%m/%Y","%d-%m-%Y","%m/%d/%Y","%d/%m/%y"):
            try: return datetime.strptime(v.strip(), fmt).date()
            except: pass
    return None

def get_headers(sheet):
    for row in sheet.iter_rows(min_row=1, max_row=5, values_only=True):
        h = [str(c or "").strip().lower() for c in row]
        if sum(1 for x in h if x) >= 3: return h
    return []

def ci(hdrs, *kws):
    """Column index: first header matching any keyword."""
    for kw in kws:
        for i,h in enumerate(hdrs):
            if kw in h: return i
    return None

def find_sheet(wb, *keys):
    for k in keys:
        for n in wb.sheetnames:
            if k in n.lower(): return wb[n]
    return None

# ── Download: Microsoft Graph (OneDrive/SharePoint) ───────────────────────────
def _graph_token():
    """OAuth2 client credentials flow → access token."""
    url = f"https://login.microsoftonline.com/{MS_TENANT_ID}/oauth2/v2.0/token"
    data = {
        "client_id":     MS_CLIENT_ID,
        "client_secret": MS_CLIENT_SECRET,
        "scope":         "https://graph.microsoft.com/.default",
        "grant_type":    "client_credentials",
    }
    r = requests.post(url, data=data, timeout=30)
    r.raise_for_status()
    return r.json()["access_token"]

def _encode_share_url(share_url):
    """Codifica URL compartida al formato u!<base64url-sin-padding> que Graph requiere."""
    b64 = base64.urlsafe_b64encode(share_url.encode()).decode().rstrip("=")
    return "u!" + b64

def download_onedrive():
    if not (MS_TENANT_ID and MS_CLIENT_ID and MS_CLIENT_SECRET):
        return False
    if not (ONEDRIVE_SHARE_URL or (ONEDRIVE_USER and ONEDRIVE_FILE_PATH)):
        return False

    print("🔐  Autenticando con Microsoft Graph…")
    try:
        token = _graph_token()
    except Exception as e:
        print(f"✗  Error al obtener token: {e}")
        return False
    headers = {"Authorization": f"Bearer {token}"}

    if ONEDRIVE_SHARE_URL:
        enc = _encode_share_url(ONEDRIVE_SHARE_URL)
        graph_url = f"https://graph.microsoft.com/v1.0/shares/{enc}/driveItem/content"
        print("⬇  Descargando Excel desde OneDrive (via share URL)…")
    else:
        # Construir ruta: /users/{user}/drive/root:/{path}:/content
        path = ONEDRIVE_FILE_PATH.lstrip("/")
        graph_url = (
            f"https://graph.microsoft.com/v1.0/users/{ONEDRIVE_USER}"
            f"/drive/root:/{path}:/content"
        )
        print(f"⬇  Descargando Excel desde OneDrive (user={ONEDRIVE_USER})…")

    r = requests.get(graph_url, headers=headers, stream=True,
                     allow_redirects=True, timeout=60)
    if r.status_code != 200:
        print(f"✗  HTTP {r.status_code}: {r.text[:300]}")
        return False
    with open(XLSX_PATH, "wb") as f:
        for chunk in r.iter_content(32768): f.write(chunk)
    print(f"✓  {XLSX_PATH.stat().st_size//1024} KB guardados")
    return True

# ── Download: Google Drive (legacy fallback) ─────────────────────────────────
def download_gdrive():
    if not FILE_ID:
        return False
    url = f"https://drive.google.com/uc?export=download&id={FILE_ID}"
    print(f"⬇  Descargando Excel desde Drive (legacy)…")
    s = requests.Session()
    r = s.get(url, stream=True, allow_redirects=True, timeout=30)
    for k, v in r.cookies.items():
        if k.startswith("download_warning"):
            r = s.get(f"{url}&confirm={v}", stream=True, timeout=30)
            break
    if r.status_code != 200:
        print(f"✗  HTTP {r.status_code}"); return False
    with open(XLSX_PATH, "wb") as f:
        for chunk in r.iter_content(32768): f.write(chunk)
    print(f"✓  {XLSX_PATH.stat().st_size//1024} KB guardados")
    return True

def download():
    """Intenta OneDrive primero, luego Google Drive como fallback."""
    if download_onedrive():
        return True
    if download_gdrive():
        return True
    print("⚠  Sin credenciales válidas (ni OneDrive ni Drive) — conservando datos actuales.")
    return False

# ── Parsers ───────────────────────────────────────────────────────────────────
def parse_diario(wb, year=None):
    if year is None:
        year = datetime.now().year
    sheet = None
    for n in wb.sheetnames:
        nl = n.lower()
        if str(year) in n and ("factur" in nl or "diario" in nl):
            sheet = wb[n]
            break
    if sheet is None:
        sheet = find_sheet(wb, "diario","daily","factur","datos","venta")
    if sheet is None:
        sheet = wb.active
    hdrs = get_headers(sheet)
    if not hdrs: return []

    C = {
        "fecha":   ci(hdrs,"fecha","date","día","dia"),
        "man":     ci(hdrs,"mañana","manana","morning","m1","mañ"),
        "mid":     ci(hdrs,"mediodía","mediodia","noon","lunch","med"),
        "noch":    ci(hdrs,"noche","night","evening","noch"),
        "total":   ci(hdrs,"total","real","factur"),
        "prev":    ci(hdrs,"previsto","forecast","objetivo","prev"),
        "coste":   ci(hdrs,"coste","cost","personal","salari"),
        "evento":  ci(hdrs,"evento","event","nota","observ"),
    }

    records = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(c is None for c in row): continue
        d = parse_date(row[C["fecha"]]) if C["fecha"] is not None else None
        if d is None: continue
        if d.year != year: continue
        man   = flt(row[C["man"]])   if C["man"]   is not None else 0
        mid   = flt(row[C["mid"]])   if C["mid"]   is not None else 0
        noch  = flt(row[C["noch"]])  if C["noch"]  is not None else 0
        tot_r = flt(row[C["total"]]) if C["total"] is not None else 0
        total = tot_r if tot_r > 0 else round(man+mid+noch, 2)
        if total <= 0: continue
        prev  = flt(row[C["prev"]])  if C["prev"]  is not None else 0
        coste = flt(row[C["coste"]]) if C["coste"] is not None else 0
        ev    = str(row[C["evento"]] or "").strip() if C["evento"] is not None else ""
        records.append({
            "fecha":    d.strftime("%Y-%m-%d"),
            "dia":      DIAS_ES.get(d.strftime("%A"), "?"),
            "mes":      d.month,
            "manana":   round(man,2),
            "mediodia": round(mid,2),
            "noche":    round(noch,2),
            "total":    round(total,2),
            "previsto": round(prev,2),
            "coste":    round(coste,2),
            "pctCoste": round(coste/total,4) if total>0 else 0,
            "evento":   ev,
        })
    records.sort(key=lambda r: r["fecha"])
    print(f"  ✓ Diario {year}: {len(records)} días")
    return records


def parse_historico(wb, year):
    """Devuelve list de dicts por fecha del año historico (por defecto 2025)."""
    sheet = None
    for n in wb.sheetnames:
        nl = n.lower()
        if str(year) in n and ("factur" in nl or "diario" in nl):
            sheet = wb[n]; break
    if sheet is None: return []
    hdrs = get_headers(sheet)
    if not hdrs: return []
    C = {
        "fecha":   ci(hdrs,"fecha","date","día","dia"),
        "man":     ci(hdrs,"mañana","manana","morning"),
        "mid":     ci(hdrs,"mediodía","mediodia","noon","lunch"),
        "noch":    ci(hdrs,"noche","night","evening"),
        "total":   ci(hdrs,"total","real","factur"),
        "evento":  ci(hdrs,"evento","event","nota"),
    }
    out = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(c is None for c in row): continue
        d = parse_date(row[C["fecha"]]) if C["fecha"] is not None else None
        if d is None or d.year != year: continue
        man  = flt(row[C["man"]])   if C["man"]   is not None else 0
        mid  = flt(row[C["mid"]])   if C["mid"]   is not None else 0
        noch = flt(row[C["noch"]])  if C["noch"]  is not None else 0
        totr = flt(row[C["total"]]) if C["total"] is not None else 0
        total = totr if totr > 0 else round(man+mid+noch, 2)
        ev = str(row[C["evento"]] or "").strip() if C["evento"] is not None else ""
        out.append({
            "fecha":    d.strftime("%Y-%m-%d"),
            "mes":      d.month,
            "manana":   round(man,2),
            "mediodia": round(mid,2),
            "noche":    round(noch,2),
            "total":    round(total,2),
            "evento":   ev,
        })
    out.sort(key=lambda r: r["fecha"])
    print(f"  ✓ Histórico {year}: {len(out)} días")
    return out


def parse_caja(wb):
    """Agrega datos de Caja 1 + Caja 2 por (fecha, turno).
    Devuelve (pax_map, caja_detalle):
      pax_map       : {(fecha_str, turno_letra): pax_total}   ← compatibilidad
      caja_detalle  : {(fecha_str, turno_nombre): {pax, ventas, coste}}
    Turno_nombre = "Mañana" | "Mediodía" | "Noche" para usar directo en frontend.
    """
    TURNO_NORM   = {"mañana":"M", "manana":"M", "mediodía":"D", "mediodia":"D",
                    "noche":"N", "tarde":"T"}
    TURNO_NOMBRE = {"M":"Mañana", "D":"Mediodía", "N":"Noche", "T":"Tarde"}
    pax = {}
    caja = {}
    targets = {"caja1", "caja2"}
    procesadas = set()
    for n in wb.sheetnames:
        norm = n.lower().replace(" ","")
        if norm not in targets or norm in procesadas: continue
        procesadas.add(norm)
        sheet = wb[n]
        hdrs = get_headers(sheet)
        if not hdrs: continue
        Cf = ci(hdrs,"fecha","date","día","dia")
        Ct = ci(hdrs,"turno","shift")
        Cp = ci(hdrs,"comensal","pax","cubierto","cliente")
        Cv = ci(hdrs,"subtotal","ventas","venta","factur","importe","total")
        Cc = ci(hdrs,"coste personal","coste","cost personal","cost")
        if Cf is None or Ct is None or Cp is None:
            print(f"  ⚠ {sheet.title}: encabezados no reconocidos")
            continue
        n_filas = 0
        for row in sheet.iter_rows(min_row=2, values_only=True):
            if all(c is None for c in row): continue
            d = parse_date(row[Cf])
            if d is None: continue
            t_raw = str(row[Ct] or "").strip().lower()
            t = TURNO_NORM.get(t_raw)
            if t is None: continue
            p = int(flt(row[Cp])) if Cp is not None else 0
            v = flt(row[Cv]) if Cv is not None else 0
            c_ = flt(row[Cc]) if Cc is not None else 0
            if p <= 0 and v <= 0: continue
            fstr = d.strftime("%Y-%m-%d")
            # pax_map legacy (para previsión)
            if p > 0:
                key = (fstr, t)
                pax[key] = pax.get(key, 0) + p
            # caja_detalle agregado
            key2 = (fstr, TURNO_NOMBRE[t])
            if key2 not in caja:
                caja[key2] = {"pax":0, "ventas":0.0, "coste":0.0}
            caja[key2]["pax"]    += p
            caja[key2]["ventas"] += v
            caja[key2]["coste"]  += c_
            n_filas += 1
        print(f"  ✓ {sheet.title}: {n_filas} filas")
    return pax, caja


def caja_to_list(caja):
    """Convierte dict caja_detalle a lista para JSON."""
    out = []
    for (fecha, turno), v in caja.items():
        out.append({
            "fecha":    fecha,
            "turno":    turno,
            "pax":      int(v["pax"]),
            "ventas":   round(v["ventas"], 2),
            "coste":    round(v["coste"], 2),
        })
    out.sort(key=lambda r: (r["fecha"], r["turno"]))
    return out


def parse_horarios_personal(wb):
    """Intenta extraer horas de camareros y cocina por (fecha, turno).
    Busca hojas con nombre que contenga 'horario' o 'horas'.
    Espera columnas: fecha | turno | (nombre|cargo) + horas.
    Devuelve dict {(fecha, turno_nombre): {horasCam, horasCoc}}.
    """
    TURNO_NORM   = {"mañana":"M", "manana":"M", "mediodía":"D", "mediodia":"D", "noche":"N"}
    TURNO_NOMBRE = {"M":"Mañana", "D":"Mediodía", "N":"Noche"}
    CARGO_COCINA = ("cocin", "chef", "ayudante cocina")
    sheet = None
    for n in wb.sheetnames:
        nl = n.lower()
        if "horario" in nl or "horas" in nl:
            sheet = wb[n]; break
    if sheet is None: return {}
    hdrs = get_headers(sheet)
    if not hdrs: return {}
    Cf = ci(hdrs,"fecha","date","día","dia")
    Ct = ci(hdrs,"turno","shift")
    Cn = ci(hdrs,"nombre","empleado","trabajador","name")
    Cg = ci(hdrs,"cargo","puesto","rol","categ","funci")
    Ch = ci(hdrs,"horas","hours","total horas","total h")
    if Cf is None or Ct is None or Ch is None:
        print(f"  ⚠ {sheet.title}: encabezados no reconocidos para horarios")
        return {}
    out = {}
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(c is None for c in row): continue
        d = parse_date(row[Cf])
        if d is None: continue
        t_raw = str(row[Ct] or "").strip().lower()
        t = TURNO_NORM.get(t_raw)
        if t is None: continue
        h = flt(row[Ch])
        if h <= 0: continue
        cargo = (str(row[Cg] or "") if Cg is not None else "").lower()
        nombre = (str(row[Cn] or "") if Cn is not None else "").lower()
        es_cocina = any(k in cargo for k in CARGO_COCINA) or any(k in nombre for k in CARGO_COCINA)
        key = (d.strftime("%Y-%m-%d"), TURNO_NOMBRE[t])
        if key not in out: out[key] = {"horasCam":0.0, "horasCoc":0.0}
        if es_cocina: out[key]["horasCoc"] += h
        else:         out[key]["horasCam"] += h
    print(f"  ✓ Horarios Personal: {len(out)} turnos con horas")
    return out


def horarios_to_list(h):
    return [{"fecha":f, "turno":t,
             "horasCam": round(v["horasCam"],2),
             "horasCoc": round(v["horasCoc"],2)}
            for (f,t),v in sorted(h.items())]


def build_prevision(daily_cur, historico_prev, events, pax_map, weeks=8):
    """Construye el data-structure de la pestaña Previsión.
    daily_cur       : lista de dicts del año actual (2026)
    historico_prev  : lista de dicts del año anterior (2025)
    events          : lista de dicts de la hoja Eventos
    pax_map         : dict {(fecha, turno): pax} de Caja 1+2
    weeks           : ventana desde el lunes de esta semana
    """
    from datetime import timedelta
    hoy = datetime.now().date()
    lunes_esta_sem = hoy - timedelta(days=hoy.weekday())
    fin = lunes_esta_sem + timedelta(days=weeks*7 - 1)

    hprev = {r["fecha"]: r for r in historico_prev}
    dcur  = {r["fecha"]: r for r in daily_cur}

    def _norm(s): return str(s or "").strip().lower()
    ev_prev = {}; ev_cur = {}
    for e in events:
        f = e.get("fecha","")
        if not f or len(f) < 10 or "-" not in f: continue
        y = f[:4]
        name = _norm(e.get("evento",""))
        if not name: continue
        if y == "2025": ev_prev[name] = f
        elif y == "2026": ev_cur.setdefault(name, []).append(f)
    evt_en_cur = {}
    for name, fechas in ev_cur.items():
        for f in fechas:
            evt_en_cur[f] = name

    def _percentiles(values):
        v = sorted(x for x in values if x > 0)
        if not v: return (0, 0)
        n = len(v)
        return (v[min(int(n*0.33), n-1)], v[min(int(n*0.66), n-1)])
    umbrales = {
        "manana":   _percentiles([r["manana"]   for r in historico_prev]),
        "mediodia": _percentiles([r["mediodia"] for r in historico_prev]),
        "noche":    _percentiles([r["noche"]    for r in historico_prev]),
    }
    def _carga(val, turno):
        p33, p66 = umbrales.get(turno, (0,0))
        if val >= p66 and p66 > 0: return "ALTA"
        if val >= p33 and p33 > 0: return "MEDIA"
        return "BAJA"

    def _equiv(fecha_cur_str):
        d = datetime.strptime(fecha_cur_str, "%Y-%m-%d").date()
        name = _norm(evt_en_cur.get(fecha_cur_str,""))
        # 1) Evento → evento (puede cruzar días de la semana)
        if name and name in ev_prev:
            return ev_prev[name], "evento"
        # 2) Cerrado los lunes (0) y martes (1) salvo evento
        if d.weekday() in (0, 1):
            return None, "cerrado"
        # 3) Mismo día de la semana, mismo mes, día más cercano
        cand = []
        for k in hprev:
            dk = datetime.strptime(k, "%Y-%m-%d").date()
            if dk.month == d.month and dk.weekday() == d.weekday():
                cand.append((abs(dk.day - d.day), k))
        if cand:
            cand.sort()
            return cand[0][1], "dia_semana"
        return None, None

    filas = []
    n_dias = (fin - lunes_esta_sem).days + 1
    for i in range(n_dias):
        d = lunes_esta_sem + timedelta(days=i)
        fcur = d.strftime("%Y-%m-%d")
        evtcur = evt_en_cur.get(fcur, "")
        fprev, metodo = _equiv(fcur)
        ref = hprev.get(fprev, {}) if fprev else {}
        evtprev = ref.get("evento","")
        pax_prev = 0
        if fprev:
            pax_prev = (pax_map.get((fprev,"M"),0)
                      + pax_map.get((fprev,"D"),0)
                      + pax_map.get((fprev,"N"),0))
        pm = ref.get("manana", 0); pd = ref.get("mediodia", 0); pn = ref.get("noche", 0)
        pt = round(pm + pd + pn, 2)
        real = dcur.get(fcur, {})
        rm = real.get("manana", 0); rd = real.get("mediodia", 0); rn = real.get("noche", 0)
        rt = real.get("total", 0)
        def _d(a,b): return round(a-b, 2)
        def _p(a,b): return round((a-b)/b, 4) if b > 0 else 0
        filas.append({
            "fechaCur":  fcur,
            "dia":       DIAS_ES.get(d.strftime("%A"), "?"),
            "evtCur":    evtcur,
            "fechaPrev": fprev or "",
            "evtPrev":   evtprev,
            "metodo":    metodo or "",
            "prev":      { "manana": round(pm,2), "mediodia": round(pd,2),
                           "noche": round(pn,2), "total": pt },
            "paxPrev":   pax_prev,
            "carga": {
                "manana":   _carga(pm, "manana"),
                "mediodia": _carga(pd, "mediodia"),
                "noche":    _carga(pn, "noche"),
            },
            "real": {
                "manana":   rm, "mediodia": rd, "noche": rn, "total": rt,
            },
            "diff": {
                "manana":   _d(rm, pm), "mediodia": _d(rd, pd),
                "noche":    _d(rn, pn), "total":    _d(rt, pt),
            },
            "pct": {
                "manana":   _p(rm, pm), "mediodia": _p(rd, pd),
                "noche":    _p(rn, pn), "total":    _p(rt, pt),
            },
            "cerrado":  bool(real),
        })

    DOW_NAMES = ["Lunes","Martes","Miércoles","Jueves","Viernes","Sábado","Domingo"]
    dow_agg = {i: {"tot":0,"count":0} for i in range(7)}
    for r in historico_prev:
        dd = datetime.strptime(r["fecha"], "%Y-%m-%d").date()
        dow_agg[dd.weekday()]["tot"]   += r["total"]
        dow_agg[dd.weekday()]["count"] += 1
    dow = []
    for i in range(7):
        c = dow_agg[i]["count"]
        avg = (dow_agg[i]["tot"] / c) if c else 0
        dow.append({"dia": DOW_NAMES[i], "avg": round(avg,2), "n": c})

    return {
        "ventana":  {"inicio": lunes_esta_sem.strftime("%Y-%m-%d"),
                     "fin":    fin.strftime("%Y-%m-%d"),
                     "dias":   n_dias},
        "umbrales": {k: {"p33": round(v[0],2), "p66": round(v[1],2)}
                     for k,v in umbrales.items()},
        "filas":    filas,
        "dow":      dow,
    }


def monthly_from_daily(daily, historico_prev=None, html_text=""):
    """Calcula resumen mensual usando histórico parseado si se pasa;
    si no, cae al regex del HTML previo (compatibilidad hacia atrás)."""
    r25 = {}
    if historico_prev:
        for r in historico_prev:
            r25.setdefault(r["mes"], 0)
            r25[r["mes"]] += r["total"]
    else:
        try:
            m = re.search(r'"monthly"\s*:\s*(\[[\s\S]*?\])', html_text)
            if m:
                prev_monthly = json.loads(m.group(1))
                MES_TO_NUM = {n:i for i,n in enumerate(MESES, 1)}
                for row in prev_monthly:
                    mn = MES_TO_NUM.get(row.get("mes",""))
                    if mn: r25[mn] = row.get("real2025", 0)
        except: pass

    agg = {}
    for d in daily:
        mn = d["mes"]
        if mn not in agg: agg[mn] = {"dias":0,"total":0,"prev":0,"coste":0}
        agg[mn]["dias"]  += 1
        agg[mn]["total"] += d["total"]
        agg[mn]["prev"]  += d["previsto"]
        agg[mn]["coste"] += d["coste"]

    result = []
    for i, mes in enumerate(MESES, 1):
        a = agg.get(i, {})
        total = a.get("total", 0)
        coste = a.get("coste", 0)
        result.append({
            "mes":      mes,
            "real2026": round(total, 2),
            "prev2026": round(a.get("prev", 0), 2),
            "real2025": round(r25.get(i, 0), 2),
            "dias":     a.get("dias", 0),
            "pctCoste": round(coste/total, 4) if total > 0 else 0,
        })
    return result


def parse_employees(wb, key="personal"):
    sheet = find_sheet(wb, key, "empleado", "equipo", "staff", "costes")
    if not sheet: return [], []
    hdrs = get_headers(sheet)
    C = {
        "nombre": ci(hdrs,"nombre","empleado","name","trabajador"),
        "euroh":  ci(hdrs,"€/hora","euro","tarifa","rate","hora"),
        "horas":  ci(hdrs,"horas","hours","h"),
        "total":  ci(hdrs,"total","coste","importe"),
        "pct":    ci(hdrs,"pct","%","porcentaje"),
        "ma":     ci(hdrs,"mañana","manana","ma"),
        "mid":    ci(hdrs,"mediodía","mediodia","mid","md"),
        "no":     ci(hdrs,"noche","no"),
    }
    emps = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(c is None for c in row): continue
        nombre = str(row[C["nombre"]] or "").strip() if C["nombre"] is not None else ""
        if not nombre or nombre.lower() in ("total","suma","totales"): continue
        total = flt(row[C["total"]]) if C["total"] is not None else 0
        if total <= 0: continue
        emps.append({
            "nombre":   nombre,
            "euroHora": flt(row[C["euroh"]]) if C["euroh"] is not None else 0,
            "horas":    flt(row[C["horas"]]) if C["horas"] is not None else 0,
            "total":    round(total,2),
            "pct":      flt(row[C["pct"]])  if C["pct"]  is not None else 0,
            "costeMa":  flt(row[C["ma"]])   if C["ma"]   is not None else 0,
            "costeMd":  flt(row[C["mid"]])  if C["mid"]  is not None else 0,
            "costeNo":  flt(row[C["no"]])   if C["no"]   is not None else 0,
        })
    tot = sum(e["total"] for e in emps)
    for e in emps:
        if e["pct"] == 0 and tot > 0: e["pct"] = round(e["total"]/tot, 4)
    print(f"  ✓ Personal: {len(emps)} empleados")
    return emps, []


def parse_events(wb):
    sheet = find_sheet(wb, "event","evento","acto","agenda")
    if not sheet: return []
    hdrs = get_headers(sheet)
    C = {
        "mes":    ci(hdrs,"mes","month"),
        "fecha":  ci(hdrs,"fecha","date","día"),
        "evento": ci(hdrs,"evento","event","nombre","descripción","desc"),
        "tipo":   ci(hdrs,"tipo","type","categ"),
        "mult":   ci(hdrs,"mult","factor","multiplicador"),
        "prev":   ci(hdrs,"previsto","forecast","objetivo"),
        "real":   ci(hdrs,"real","total","importe"),
        "estado": ci(hdrs,"estado","status","state"),
    }
    evs = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(c is None for c in row): continue
        ev = str(row[C["evento"]] or "").strip() if C["evento"] is not None else ""
        if not ev: continue
        fd = parse_date(row[C["fecha"]]) if C["fecha"] is not None else None
        evs.append({
            "mes":    str(row[C["mes"]] or "").strip() if C["mes"] is not None else "",
            "fecha":  fd.strftime("%Y-%m-%d") if fd else str(row[C["fecha"]] or ""),
            "evento": ev,
            "tipo":   str(row[C["tipo"]] or "").strip() if C["tipo"] is not None else "",
            "mult":   flt(row[C["mult"]], 1.0) if C["mult"] is not None else 1.0,
            "prev":   flt(row[C["prev"]]) if C["prev"] is not None else 0,
            "real":   flt(row[C["real"]]) if C["real"] is not None else 0,
            "estado": str(row[C["estado"]] or "").strip() if C["estado"] is not None else "",
        })
    print(f"  ✓ Eventos: {len(evs)}")
    return evs


def parse_products(wb):
    sheet = find_sheet(wb, "prod","mix","carta","venta","articulo")
    if not sheet: return [], []
    hdrs = get_headers(sheet)
    C = {
        "prod": ci(hdrs,"producto","product","artículo","item","nombre","descrip"),
        "fam":  ci(hdrs,"familia","family","categ","tipo","grupo"),
        "uds":  ci(hdrs,"uds","unidades","qty","cantidad","cant"),
        "imp":  ci(hdrs,"importe","total","€","ventas","venta"),
        "pct":  ci(hdrs,"pct","%","porcentaje"),
    }
    prods = []
    for row in sheet.iter_rows(min_row=2, values_only=True):
        if all(c is None for c in row): continue
        p = str(row[C["prod"]] or "").strip() if C["prod"] is not None else ""
        if not p: continue
        imp = flt(row[C["imp"]]) if C["imp"] is not None else 0
        if imp <= 0: continue
        prods.append({
            "producto": p,
            "familia":  str(row[C["fam"]] or "Sin familia").strip() if C["fam"] is not None else "Sin familia",
            "uds":      int(flt(row[C["uds"]])) if C["uds"] is not None else 0,
            "importe":  round(imp,2),
            "pct":      flt(row[C["pct"]]) if C["pct"] is not None else 0,
        })
    if not prods: return [], []
    prods.sort(key=lambda p: p["importe"], reverse=True)
    tot = sum(p["importe"] for p in prods)
    for p in prods:
        if p["pct"] == 0 and tot > 0: p["pct"] = round(p["importe"]/tot, 4)

    fam_map = {}
    for p in prods:
        f = p["familia"]
        if f not in fam_map: fam_map[f] = {"uds":0,"importe":0}
        fam_map[f]["uds"]    += p["uds"]
        fam_map[f]["importe"] += p["importe"]
    families = sorted(
        [{"familia":f,"uds":v["uds"],"importe":round(v["importe"],2),
          "pct":round(v["importe"]/tot,4)} for f,v in fam_map.items()],
        key=lambda x: x["importe"], reverse=True
    )
    print(f"  ✓ Productos: {len(prods)} → top30, {len(families)} familias")
    return prods[:30], families


# ── Inject into HTML ──────────────────────────────────────────────────────────
def compact(obj):
    return json.dumps(obj, ensure_ascii=False, separators=(",",":"))

def inject(html, key, value_str):
    """Reemplaza 'const KEY = {…};' en el HTML."""
    pattern = rf'(const\s+{re.escape(key)}\s*=\s*)\{{[\s\S]*?\}};'
    new_html, n = re.subn(pattern, rf'\g<1>{value_str};', html, count=1)
    if n: print(f"  ✓ Actualizado: {key}")
    else: print(f"  ⚠  No encontrado en HTML: {key} (se conserva el existente)")
    return new_html

def update_header(html, n_dias, last_date):
    try:
        d = datetime.strptime(last_date, "%Y-%m-%d")
        fecha_fmt = d.strftime("%-d %b %Y")
    except: fecha_fmt = last_date
    new_meta = f"Tracker Operativo 2026 · {n_dias} días · Actualizado {fecha_fmt}"
    return re.sub(
        r"Tracker Operativo 2026 · \d+ días · Actualizado [^<\"']+",
        new_meta, html, count=1
    )


# ── Main ──────────────────────────────────────────────────────────────────────
def main():
    if not HTML_PATH.exists():
        print(f"✗ No existe {HTML_PATH}"); sys.exit(1)

    html = HTML_PATH.read_text(encoding="utf-8")
    print(f"📄 HTML cargado: {len(html)//1024} KB")

    ok = download()
    if not ok or not XLSX_PATH.exists():
        print("ℹ Sin Excel — actualizando solo timestamp.")
        today = datetime.utcnow().strftime("%Y-%m-%d")
        html = update_header(html, html.count('"fecha"'), today)
        HTML_PATH.write_text(html, encoding="utf-8")
        return

    print("📊 Procesando Excel…")
    wb = openpyxl.load_workbook(XLSX_PATH, read_only=True, data_only=True)
    print(f"  Hojas: {wb.sheetnames}")

    year_cur  = datetime.now().year
    year_prev = year_cur - 1

    # --- Diario (año actual) + Histórico (año anterior) ---
    daily = parse_diario(wb, year=year_cur)
    historico = parse_historico(wb, year=year_prev)

    if daily:
        monthly = monthly_from_daily(daily, historico_prev=historico, html_text=html)
        emps, emps_mes = parse_employees(wb)
        events = parse_events(wb)

        # --- Caja (Caja 1 + Caja 2) por (fecha, turno) ---
        pax_map, caja_detalle = parse_caja(wb)
        caja_list = caja_to_list(caja_detalle)
        print(f"  ✓ Caja agregada: {len(caja_list)} turnos con datos")

        # --- Reconciliación: completar daily con fechas presentes en cajas
        # pero ausentes en 'Facturación 2026' (hoja resumen desfasada) ---
        fechas_daily = {d["fecha"] for d in daily}
        TURNO2KEY = {"Mañana": "manana", "Mediodía": "mediodia", "Noche": "noche"}
        por_fecha = {}
        for (fstr, turno_nombre), v in caja_detalle.items():
            if fstr in fechas_daily:
                continue
            if fstr not in por_fecha:
                por_fecha[fstr] = {"manana": 0, "mediodia": 0, "noche": 0, "coste": 0}
            k = TURNO2KEY.get(turno_nombre)
            if k:
                por_fecha[fstr][k]     += v.get("ventas", 0) or 0
                por_fecha[fstr]["coste"] += v.get("coste", 0) or 0
        added = 0
        for fstr in sorted(por_fecha.keys()):
            vals = por_fecha[fstr]
            total = round(vals["manana"] + vals["mediodia"] + vals["noche"], 2)
            if total <= 0:
                continue
            d = datetime.strptime(fstr, "%Y-%m-%d").date()
            daily.append({
                "fecha":    fstr,
                "dia":      DIAS_ES.get(d.strftime("%A"), "?"),
                "mes":      d.month,
                "manana":   round(vals["manana"], 2),
                "mediodia": round(vals["mediodia"], 2),
                "noche":    round(vals["noche"], 2),
                "total":    total,
                "previsto": 0,
                "coste":    round(vals["coste"], 2),
                "pctCoste": round(vals["coste"] / total, 4) if total > 0 else 0,
                "evento":   "",
            })
            added += 1
        if added:
            daily.sort(key=lambda r: r["fecha"])
            # recomputar monthly con los nuevos días
            monthly = monthly_from_daily(daily, historico_prev=historico, html_text=html)
            print(f"  ✓ Daily reconciliado: +{added} días desde cajas (total {len(daily)})")

        # --- Horas de personal por (fecha, turno) ---
        horarios = parse_horarios_personal(wb)
        horarios_list = horarios_to_list(horarios)

        data_obj = {
            "daily":        daily,
            "monthly":      monthly,
            "employees":    emps if emps else json.loads(
                re.search(r'"employees"\s*:\s*(\[[\s\S]*?\])', html).group(1)
                if re.search(r'"employees"\s*:\s*\[', html) else "[]"
            ),
            "employeesMes": emps_mes if emps_mes else json.loads(
                re.search(r'"employeesMes"\s*:\s*(\[[\s\S]*?\])', html).group(1)
                if re.search(r'"employeesMes"\s*:\s*\[', html) else "[]"
            ),
            "events": events if events else json.loads(
                re.search(r'"events"\s*:\s*(\[[\s\S]*?\])', html).group(1)
                if re.search(r'"events"\s*:\s*\[', html) else "[]"
            ),
            "caja":              caja_list,
            "horariosPersonal":  horarios_list,
        }
        html = inject(html, "DATA", compact(data_obj))

        # --- Previsión (semana actual + 7 semanas) ---
        prevision = build_prevision(daily, historico, events, pax_map, weeks=8)
        html = inject(html, "PREVISION", compact(prevision))

    # --- Productos ---
    top_prod, families = parse_products(wb)
    if top_prod:
        html = inject(html, "PM", compact({"topProd": top_prod, "families": families}))

    wb.close()
    XLSX_PATH.unlink(missing_ok=True)

    # --- Header ---
    if daily:
        html = update_header(html, len(daily), daily[-1]["fecha"])

    HTML_PATH.write_text(html, encoding="utf-8")
    print(f"🎉 Dashboard actualizado correctamente → {HTML_PATH.name}")


if __name__ == "__main__":
    main()
