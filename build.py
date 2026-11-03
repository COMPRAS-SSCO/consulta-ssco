"""Descarga la relación de SSCO de SUNAT y genera index.html.
Si la descarga falla, usa el último Excel guardado en data/ssco.xlsx."""
import json, datetime, urllib.request, pathlib, sys
import openpyxl

URL = "https://www.sunat.gob.pe/padronesnotificaciones/ssco/sujesincapacidadOperativa.xlsx"
XLSX = pathlib.Path("data/ssco.xlsx")
lima = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None) - datetime.timedelta(hours=5)

try:
    req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
    blob = urllib.request.urlopen(req, timeout=60).read()
    tmp = pathlib.Path("data/_nuevo.xlsx"); tmp.write_bytes(blob)
    openpyxl.load_workbook(tmp, read_only=True)  # valida que sea un Excel
    tmp.replace(XLSX)
    print("Descargado de SUNAT:", len(blob), "bytes")
    checked = lima.strftime("%d/%m/%Y")
except Exception as e:
    print("AVISO: no se pudo descargar de SUNAT, se usa el último archivo guardado:", e)
    checked = "sin conexión (" + lima.strftime("%d/%m/%Y") + ")"

def fmt(d):
    if isinstance(d, (datetime.date, datetime.datetime)):
        return d.strftime("%d/%m/%Y")
    return str(d).strip() if d is not None else ""

ws = openpyxl.load_workbook(XLSX, read_only=True).active
rows, started = [], False
for r in ws.iter_rows(values_only=True):
    if not started:
        if r and str(r[0] or "").strip().upper() == "RUC":
            started = True
        continue
    if not r or r[0] is None or not str(r[0]).strip():
        continue
    r = list(r) + [None] * 9
    ruc = str(r[0]).strip().split(".")[0]
    rows.append([ruc] + [fmt(x) for x in r[1:9]])

if len(rows) < 100:
    sys.exit(f"ERROR: solo se leyeron {len(rows)} registros; revisar formato del Excel de SUNAT")

def key(s):
    try: return datetime.datetime.strptime(s, "%d/%m/%Y")
    except Exception: return datetime.datetime.min
last = max((r[8] for r in rows), key=key)

html = (pathlib.Path("template.html").read_text(encoding="utf-8")
        .replace("__DATA__", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
        .replace("__LAST__", last).replace("__CHK__", checked))
pathlib.Path("site").mkdir(exist_ok=True)
pathlib.Path("site/index.html").write_text(html, encoding="utf-8")
print(f"OK: {len(rows)} SSCO, última publicación {last}")
