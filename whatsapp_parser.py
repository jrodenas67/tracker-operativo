"""Parser de mensajes de WhatsApp con horarios de personal.

Formato de entrada esperado:
    [DD/MM/YY, HH:MM:SS] REMITENTE: DiaSemana Num turno
    Nombre de HH[:MM] a HH[:MM] [y de HH[:MM] a HH[:MM]]
    ...

Robusto ante:
  - Ausencia de newlines (split por timestamps)
  - Separadores raros en horas: 12:30, 12;30, 1230
  - Turnos partidos (genera dos entradas)
  - Lineas "ruido" tras el ultimo turno parseable
  - Horas ambiguas AM/PM (heuristica por turno)

Uso:
    from whatsapp_parser import parse_mensajes
    turnos = parse_mensajes(texto)
    # -> List[Turno] con .fecha (date), .turno ("Mañana"|"Noche"),
    #    .nombre (str), .entrada (time), .salida (time)
"""

from __future__ import annotations
import re
import datetime
from dataclasses import dataclass
from typing import Iterator

DIAS_SEMANA = {
    "lunes": 0, "martes": 1, "miércoles": 2, "miercoles": 2,
    "jueves": 3, "viernes": 4, "sábado": 5, "sabado": 5, "domingo": 6,
}
TURNOS = {"mañana": "Mañana", "manana": "Mañana", "noche": "Noche", "tarde": "Tarde"}

# Timestamp que abre cada bloque: [12/4/26, 8:40:46]
# Header = exactamente "DiaSemana Numero Turno" (3 tokens), para no depender
# de que haya newline. GitHub workflow_dispatch input de tipo string es
# single-line y colapsa saltos de linea al pegar.
RE_TIMESTAMP = re.compile(
    r"\[(?P<dia>\d{1,2})/(?P<mes>\d{1,2})/(?P<ano>\d{2,4}),\s*"
    r"(?P<h>\d{1,2}):(?P<m>\d{2}):(?P<s>\d{2})\]\s*"
    r"[^:]+:\s*"  # remitente
    r"(?P<header>\S+\s+\d{1,2}\s+\S+)"  # "Sábado 11 noche"
)

# Linea de turno personal: "Nombre de 7 a 22:45" o "Nombre de 7 a 10 y de 12:30 a 3"
# Acepta 1-2 palabras en el nombre (ej. "Mireia M.").
# SIN flag IGNORECASE: el nombre debe empezar con mayuscula para que "de lunes" no
# pase por nombre valido.
RE_TURNO = re.compile(
    r"(?P<nombre>[A-ZÁÉÍÓÚÑ][A-Za-zÁÉÍÓÚÑáéíóúñ]+(?:\s+[A-ZÁÉÍÓÚÑ][\.A-Za-zÁÉÍÓÚÑáéíóúñ]*)?)"
    r"\s+(?:de\s+)?"
    r"(?P<h1>\d{1,2}(?:[:;]\d{1,2})?|\d{3,4})"
    r"\s+a\s+"
    r"(?P<h2>\d{1,2}(?:[:;]\d{1,2})?|\d{3,4})"
    r"(?:\s+[yY]\s+de\s+"
    r"(?P<h3>\d{1,2}(?:[:;]\d{1,2})?|\d{3,4})"
    r"\s+a\s+"
    r"(?P<h4>\d{1,2}(?:[:;]\d{1,2})?|\d{3,4}))?"
)

# Palabras reservadas que no pueden ser primera palabra de un nombre
RESERVED_FIRST_WORD = {"De", "A", "Y", "En", "Con", "Lunes", "Martes", "Miércoles",
                      "Miercoles", "Jueves", "Viernes", "Sábado", "Sabado", "Domingo"}


@dataclass(frozen=True)
class Turno:
    fecha: datetime.date
    turno: str         # "Mañana" | "Noche" | "Tarde"
    nombre: str
    entrada: datetime.time
    salida: datetime.time

    def key(self) -> tuple:
        return (self.fecha, self.turno)


def _parse_hora(raw: str, turno: str, es_salida: bool, hora_ref: int | None = None) -> datetime.time:
    """Convierte '7', '22:45', '1230', '12;30' a datetime.time con heuristica AM/PM.

    Reglas:
      - Normaliza separadores ; -> :
      - Si 3-4 digitos sin separador: HMMM o HHMM -> HH:MM (ej. 1230 -> 12:30)
      - turno 'Mañana': si es salida y h < hora_ref -> +12 (ej. "8 a 4" -> 8:00-16:00)
      - turno 'Noche': h < 12 -> +12 (ej. "7 a 11" -> 19:00-23:00)
      - turno 'Tarde': similar a noche para horas < 8
    """
    s = raw.replace(";", ":").strip()
    if ":" not in s and len(s) >= 3:
        # 1230 -> 12:30, 830 -> 8:30
        s = s[:-2] + ":" + s[-2:]
    if ":" in s:
        h, m = s.split(":", 1)
        h = int(h); m = int(m)
    else:
        h = int(s); m = 0

    if not (0 <= h <= 23 and 0 <= m <= 59):
        # Fallback: intentar clamp
        h = max(0, min(23, h))
        m = max(0, min(59, m))

    if turno == "Mañana":
        # h < 5 -> probablemente PM (ej. "Aina de 2 a 3:30" = 14:00 a 15:30)
        if h < 5:
            h += 12
        elif es_salida and hora_ref is not None and h < hora_ref and h < 12:
            h += 12
    elif turno == "Noche":
        # Horas < 12 se interpretan como PM
        if h < 12:
            h += 12
    elif turno == "Tarde":
        if h < 8:
            h += 12

    return datetime.time(h, m)


def _infer_fecha(dia_mes: int, dia_semana: str, ts_day: int, ts_month: int, ts_year: int) -> datetime.date:
    """Dada la cabecera 'Sabado 11' y el timestamp [12/4/26], devuelve date(2026,4,11).

    Si dia_mes > ts_day, asumimos mes anterior."""
    if ts_year < 100:
        ts_year += 2000
    month = ts_month
    year = ts_year
    if dia_mes > ts_day:
        month -= 1
        if month == 0:
            month = 12
            year -= 1
    # Validacion: si el dia_semana no coincide, aviso por stderr pero seguimos
    try:
        d = datetime.date(year, month, dia_mes)
    except ValueError:
        # dia_mes invalido para ese mes -> intentar mes +/-1
        d = datetime.date(ts_year, ts_month, min(dia_mes, 28))
    return d


def _iter_bloques(text: str) -> Iterator[tuple[datetime.date, str, str]]:
    """Itera (fecha_bloque, turno_bloque, cuerpo_bloque).

    Splittea por timestamps. El cuerpo es todo el texto hasta el siguiente timestamp."""
    matches = list(RE_TIMESTAMP.finditer(text))
    for i, m in enumerate(matches):
        header = m.group("header").strip()
        # extraer "Sabado 11 noche" del header
        parts = header.lower().split()
        dia_semana = None
        dia_mes = None
        turno = None
        for p in parts:
            p_norm = p.replace(",", "").replace(".", "").strip()
            if p_norm in DIAS_SEMANA and dia_semana is None:
                dia_semana = p_norm
            elif p_norm in TURNOS and turno is None:
                turno = TURNOS[p_norm]
            elif p_norm.isdigit() and dia_mes is None:
                dia_mes = int(p_norm)
        if dia_mes is None or turno is None:
            # Bloque malformado: lo saltamos
            continue
        fecha = _infer_fecha(
            dia_mes=dia_mes,
            dia_semana=dia_semana or "",
            ts_day=int(m.group("dia")),
            ts_month=int(m.group("mes")),
            ts_year=int(m.group("ano")),
        )
        # Cuerpo: desde final del header hasta antes del siguiente timestamp
        start = m.end()
        end = matches[i + 1].start() if i + 1 < len(matches) else len(text)
        cuerpo = text[start:end]
        yield fecha, turno, cuerpo


def parse_mensajes(text: str) -> list[Turno]:
    """Parsea el texto completo y devuelve la lista de turnos encontrados."""
    turnos: list[Turno] = []
    for fecha, turno_bloque, cuerpo in _iter_bloques(text):
        for t in RE_TURNO.finditer(cuerpo):
            nombre = t.group("nombre").strip()
            primera = nombre.split()[0]
            if primera in RESERVED_FIRST_WORD:
                continue
            h1_raw = t.group("h1")
            h2_raw = t.group("h2")
            h1 = _parse_hora(h1_raw, turno_bloque, es_salida=False)
            h2 = _parse_hora(h2_raw, turno_bloque, es_salida=True, hora_ref=h1.hour)
            turnos.append(Turno(fecha, turno_bloque, nombre, h1, h2))
            # Turno partido
            if t.group("h3"):
                h3 = _parse_hora(t.group("h3"), turno_bloque, es_salida=False)
                h4 = _parse_hora(t.group("h4"), turno_bloque, es_salida=True, hora_ref=h3.hour)
                turnos.append(Turno(fecha, turno_bloque, nombre, h3, h4))
    return turnos


if __name__ == "__main__":
    import sys
    raw = sys.stdin.read() if not sys.stdin.isatty() else __doc__
    for t in parse_mensajes(raw):
        print(f"{t.fecha} {t.turno:<7} {t.nombre:<12} {t.entrada} - {t.salida}")
