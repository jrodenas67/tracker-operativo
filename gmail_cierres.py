"""
gmail_cierres.py
─────────────────────────────────────────────────────────────────────────────
Descarga los PDFs "cierreCaja.pdf" adjuntos a los correos con asunto
"Cierre de caja" del Gmail del usuario, extrae el `Número de cierre` y la
`Fecha inicio` de cada uno, y devuelve un dict {numero_cierre: fecha_inicio}.

Variables de entorno requeridas:
  GMAIL_CLIENT_ID, GMAIL_CLIENT_SECRET, GMAIL_REFRESH_TOKEN
"""
from __future__ import annotations

import base64
import datetime
import io
import os
import re

import requests

GMAIL_TOKEN_URL = "https://oauth2.googleapis.com/token"
GMAIL_API       = "https://gmail.googleapis.com/gmail/v1/users/me"


def _access_token() -> str:
    resp = requests.post(GMAIL_TOKEN_URL, data={
        "client_id":     os.environ["GMAIL_CLIENT_ID"],
        "client_secret": os.environ["GMAIL_CLIENT_SECRET"],
        "refresh_token": os.environ["GMAIL_REFRESH_TOKEN"],
        "grant_type":    "refresh_token",
    }, timeout=20)
    resp.raise_for_status()
    return resp.json()["access_token"]


def _list_message_ids(token: str, query: str) -> list[str]:
    """Devuelve hasta 200 IDs de mensajes que coincidan con `query`."""
    ids: list[str] = []
    page = None
    headers = {"Authorization": f"Bearer {token}"}
    while True:
        params = {"q": query, "maxResults": 100}
        if page:
            params["pageToken"] = page
        r = requests.get(f"{GMAIL_API}/messages", headers=headers, params=params, timeout=20)
        r.raise_for_status()
        data = r.json()
        ids.extend(m["id"] for m in data.get("messages", []))
        page = data.get("nextPageToken")
        if not page or len(ids) >= 200:
            break
    return ids


def _get_message(token: str, msg_id: str) -> dict:
    r = requests.get(f"{GMAIL_API}/messages/{msg_id}",
                     headers={"Authorization": f"Bearer {token}"},
                     params={"format": "full"}, timeout=20)
    r.raise_for_status()
    return r.json()


def _find_pdf_attachment(payload: dict) -> str | None:
    """Devuelve el attachmentId del primer PDF encontrado en el mensaje."""
    parts = [payload]
    while parts:
        p = parts.pop()
        for sub in p.get("parts", []) or []:
            parts.append(sub)
        if (p.get("mimeType") == "application/pdf"
                and p.get("body", {}).get("attachmentId")):
            return p["body"]["attachmentId"]
    return None


def _download_attachment(token: str, msg_id: str, att_id: str) -> bytes:
    r = requests.get(f"{GMAIL_API}/messages/{msg_id}/attachments/{att_id}",
                     headers={"Authorization": f"Bearer {token}"}, timeout=30)
    r.raise_for_status()
    data = r.json()["data"]
    return base64.urlsafe_b64decode(data)


_RE_NUMERO = re.compile(r"Número\s+de\s+cierre[:\s]+(\d+)", re.IGNORECASE)
_RE_FECHA  = re.compile(r"Fecha\s+inicio[:\s]+(\d{2}/\d{2}/\d{4})", re.IGNORECASE)
_RE_TOTAL  = re.compile(r"Total\s+ventas\s+([0-9.,]+)\s*€", re.IGNORECASE)


def _parse_pdf(raw: bytes) -> dict | None:
    """Extrae numero_cierre, fecha_inicio y total_ventas del PDF."""
    try:
        from pypdf import PdfReader  # type: ignore
    except ImportError:
        from PyPDF2 import PdfReader  # type: ignore  # fallback

    text = ""
    try:
        reader = PdfReader(io.BytesIO(raw))
        for page in reader.pages:
            text += page.extract_text() or ""
            text += "\n"
    except Exception as e:
        print(f"⚠  PDF ilegible: {e}")
        return None

    m_num = _RE_NUMERO.search(text)
    m_fec = _RE_FECHA.search(text)
    m_tot = _RE_TOTAL.search(text)
    if not (m_num and m_fec):
        return None
    try:
        fecha = datetime.datetime.strptime(m_fec.group(1), "%d/%m/%Y").date()
    except ValueError:
        return None
    total = 0.0
    if m_tot:
        try:
            total = float(m_tot.group(1).replace(".", "").replace(",", "."))
        except ValueError:
            total = 0.0
    return {"numero": int(m_num.group(1)), "fecha": fecha, "total": total}


def fetch_cierres(days: int = 90) -> list[dict]:
    """
    Devuelve [{numero, fecha, total}, ...] con todos los PDFs de cierre
    de los últimos `days` días. Deduplica por numero_cierre.
    """
    if not all(os.environ.get(k) for k in
               ("GMAIL_CLIENT_ID", "GMAIL_CLIENT_SECRET", "GMAIL_REFRESH_TOKEN")):
        print("⚠  Gmail OAuth no configurado.")
        return []

    try:
        token = _access_token()
    except Exception as e:
        print(f"⚠  No se pudo obtener token Gmail: {e}")
        return []

    since = (datetime.date.today() - datetime.timedelta(days=days)).strftime("%Y/%m/%d")
    query = f'subject:"Cierre de caja" has:attachment after:{since}'
    print(f"   Gmail query: {query}")

    try:
        ids = _list_message_ids(token, query)
    except Exception as e:
        print(f"⚠  Gmail list error: {e}")
        return []
    print(f"   Correos de cierre encontrados: {len(ids)}")

    seen: set[int] = set()
    result: list[dict] = []
    for mid in ids:
        try:
            msg = _get_message(token, mid)
            att_id = _find_pdf_attachment(msg.get("payload", {}))
            if not att_id:
                continue
            pdf_raw = _download_attachment(token, mid, att_id)
            parsed = _parse_pdf(pdf_raw)
            if not parsed:
                continue
            if parsed["numero"] in seen:
                continue
            seen.add(parsed["numero"])
            result.append(parsed)
        except Exception as e:
            print(f"⚠  Error procesando mensaje {mid}: {e}")
            continue

    print(f"   PDFs parseados: {len(result)} cierres")
    return result


def fetch_fechas_inicio(days: int = 90) -> dict[int, datetime.date]:
    """Compat — solo {numero: fecha_inicio}."""
    return {c["numero"]: c["fecha"] for c in fetch_cierres(days)}


if __name__ == "__main__":
    for c in sorted(fetch_cierres(), key=lambda x: x["fecha"]):
        print(f"  #{c['numero']}: {c['fecha'].strftime('%d/%m/%Y')}  →  {c['total']:.2f} €")
