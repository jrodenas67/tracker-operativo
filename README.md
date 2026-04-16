# Taperia de Caldes — Tracker Operativo 2026

Dashboard interno de operaciones con actualización automática diaria desde Google Drive.

**URL:** https://tracker-operativo.netlify.app/

---

## Estructura del proyecto

```
/
├── index.html                    ← Dashboard completo (diseño + datos)
├── fetch_data.py                 ← Script: descarga Excel → actualiza index.html
├── netlify.toml                  ← Config despliegue Netlify
├── package.json                  ← Metadata proyecto
├── README.md
└── .github/workflows/
    └── update.yml                ← GitHub Actions: ejecución diaria 00:00 UTC
```

---

## Qué hace el sistema

```
Google Drive (Excel)
       ↓  (fetch_data.py)
index.html actualizado
       ↓  (git push automático)
GitHub repo
       ↓  (Netlify auto-deploy)
https://tracker-operativo.netlify.app/
```

Cada noche a medianoche:
1. GitHub Actions descarga el Excel desde Drive
2. `fetch_data.py` parsea los datos y reemplaza los objetos `DATA` y `PM` en el HTML
3. Hace commit y push automático
4. Netlify detecta el push y despliega en segundos

---

## Puesta en marcha

### 1. Compartir el Excel en Google Drive
- Abre el archivo en Drive
- "Compartir" → "Cualquier persona con el enlace puede ver"
- Copia la URL: `https://drive.google.com/file/d/FILE_ID_AQUÍ/view`
- Anota el `FILE_ID` (la parte entre `/d/` y `/view`)

### 2. Subir el proyecto a GitHub
```bash
git init
git add .
git commit -m "Dashboard inicial"
git remote add origin https://github.com/TU_USUARIO/taperia-caldes.git
git push -u origin main
```

### 3. Añadir el secreto en GitHub
- Ve a: Settings → Secrets and variables → Actions → New repository secret
- Nombre: `GOOGLE_DRIVE_FILE_ID`
- Valor: tu FILE_ID

### 4. Conectar Netlify
- app.netlify.com → "Add new site" → "Import an existing project"
- Conecta tu repo de GitHub
- Build command: (vacío)
- Publish directory: `.`
- Deploy → listo

---

## Estructura del Excel esperada

El script detecta columnas automáticamente por nombre (flexible).

### Hoja "Diario"
| fecha | mañana | mediodía | noche | previsto | coste | evento |
|-------|--------|----------|-------|----------|-------|--------|

### Hoja "Personal"
| nombre | €/hora | horas | total | pct | costeMa | costeMd | costeNo |

### Hoja "Eventos"
| mes | fecha | evento | tipo | mult | prev | real | estado |

### Hoja "Productos" (opcional)
| producto | familia | uds | importe | pct |

---

## Uso local
```bash
pip install requests openpyxl

# Con datos reales:
export GOOGLE_DRIVE_FILE_ID="tu_id_aquí"
python fetch_data.py

# Servidor local:
python -m http.server 8080
# → http://localhost:8080
```

---

## Notas
- El script sólo actualiza `DATA` y `PM` en el HTML — todos los demás datos (EBITDA_DATA, OMNES, INCENTIVOS) se mantienen y deben editarse manualmente cuando cambien.
- Si el Excel no tiene alguna hoja, se conservan los datos actuales del HTML.
- `real2025` en los meses se preserva automáticamente del HTML anterior.
