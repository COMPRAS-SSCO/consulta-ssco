"""Genera la página de consulta SSCO.
1) Descarga la relación de SSCO de SUNAT (si falla, usa data/ssco.xlsx).
2) Si existe padron.zip (Padrón Reducido del RUC de SUNAT), lo parte en fragmentos
   p/XXXXXX.gz (por los 6 primeros dígitos del RUC) para buscar la razón social.
   Si no existe, reutiliza los fragmentos guardados en padron/ (caché).
3) Escribe site/ con index.html, app instalable (manifest, sw.js, íconos)."""
import json, datetime, urllib.request, pathlib, sys, zipfile, io, gzip, shutil, collections
import openpyxl

URL = "https://www.sunat.gob.pe/padronesnotificaciones/ssco/sujesincapacidadOperativa.xlsx"
XLSX = pathlib.Path("data/ssco.xlsx")
XLSX.parent.mkdir(exist_ok=True)
SITE = pathlib.Path("site"); SITE.mkdir(exist_ok=True)
PAD = pathlib.Path("padron")  # fragmentos del padrón (se guardan en caché entre ejecuciones)
lima = datetime.datetime.now(datetime.timezone.utc).replace(tzinfo=None) - datetime.timedelta(hours=5)
hoy = lima.strftime("%d/%m/%Y")

# ---------- 1. SSCO ----------
try:
    req = urllib.request.Request(URL, headers={"User-Agent": "Mozilla/5.0"})
    blob = urllib.request.urlopen(req, timeout=60).read()
    tmp = pathlib.Path("data/_nuevo.xlsx"); tmp.write_bytes(blob)
    openpyxl.load_workbook(tmp, read_only=True)
    tmp.replace(XLSX)
    print("SSCO descargado de SUNAT:", len(blob), "bytes"); checked = hoy
except Exception as e:
    print("AVISO: no se pudo descargar SSCO de SUNAT, se usa el último archivo guardado:", e)
    checked = "sin conexión (" + hoy + ")"

def fmt(d):
    if isinstance(d, (datetime.date, datetime.datetime)): return d.strftime("%d/%m/%Y")
    return str(d).strip() if d is not None else ""

ws = openpyxl.load_workbook(XLSX, read_only=True).active
rows, started = [], False
for r in ws.iter_rows(values_only=True):
    if not started:
        if r and str(r[0] or "").strip().upper() == "RUC": started = True
        continue
    if not r or r[0] is None or not str(r[0]).strip(): continue
    r = list(r) + [None] * 9
    rows.append([str(r[0]).strip().split(".")[0]] + [fmt(x) for x in r[1:9]])
if len(rows) < 100:
    sys.exit(f"ERROR: solo se leyeron {len(rows)} registros SSCO; revisar formato del Excel de SUNAT")
def key(s):
    try: return datetime.datetime.strptime(s, "%d/%m/%Y")
    except Exception: return datetime.datetime.min
last = max((r[8] for r in rows), key=key)
print(f"SSCO: {len(rows)} registros, última publicación {last}")

# ---------- 2. Padrón reducido ----------
zp = pathlib.Path("padron.zip")
if zp.exists() and zp.stat().st_size > 10_000_000:
    try:
        shards = collections.defaultdict(list); n = 0
        with zipfile.ZipFile(zp) as z:
            name = max(z.namelist(), key=lambda x: z.getinfo(x).file_size)
            with z.open(name) as f:
                for raw in io.TextIOWrapper(f, encoding="latin-1", errors="replace", newline=""):
                    p = raw.rstrip("\r\n").split("|")
                    if len(p) < 4 or len(p[0]) != 11 or not p[0].isdigit(): continue
                    nom = p[1].strip().replace("\n", " ")
                    shards[p[0][:6]].append(f"{p[0][6:]}|{nom}|{p[2].strip()}|{p[3].strip()}")
                    n += 1
        if n < 1_000_000: raise ValueError(f"solo {n} registros")
        nuevo = pathlib.Path("padron_nuevo"); shutil.rmtree(nuevo, ignore_errors=True); (nuevo / "p").mkdir(parents=True)
        for k, v in shards.items():
            (nuevo / "p" / f"{k}.gz").write_bytes(gzip.compress(("\n".join(v) + "\n").encode("utf-8"), 9))
        (nuevo / "fecha.txt").write_text(hoy)
        shutil.rmtree(PAD, ignore_errors=True); nuevo.rename(PAD)
        print(f"Padrón: {n} RUC en {len(shards)} fragmentos")
    except Exception as e:
        print("AVISO: no se pudo procesar el padrón, se usan los fragmentos anteriores:", e)
else:
    print("AVISO: no se descargó el padrón; se usan los fragmentos anteriores si existen")
pfecha = (PAD / "fecha.txt").read_text().strip() if (PAD / "fecha.txt").exists() else ""
if pfecha:
    shutil.rmtree(SITE / "p", ignore_errors=True); shutil.copytree(PAD / "p", SITE / "p")
print("Padrón disponible:", pfecha or "NO")

# ---------- 3. Página + app instalable ----------
html = (pathlib.Path("template.html").read_text(encoding="utf-8")
        .replace("__DATA__", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
        .replace("__LAST__", last).replace("__CHK__", checked).replace("__PFECHA__", pfecha))
(SITE / "index.html").write_text(html, encoding="utf-8")
(SITE / "manifest.webmanifest").write_text(json.dumps({
    "name": "Consulta SSCO – Compras", "short_name": "Consulta SSCO", "start_url": "./", "scope": "./",
    "display": "standalone", "background_color": "#f4f6f9", "theme_color": "#0b4f8a", "lang": "es",
    "icons": [{"src": "icon-192.png", "sizes": "192x192", "type": "image/png"},
              {"src": "icon-512.png", "sizes": "512x512", "type": "image/png", "purpose": "any maskable"}]},
    ensure_ascii=False), encoding="utf-8")
(SITE / "sw.js").write_text("""const C='ssco-v1';
self.addEventListener('install',e=>self.skipWaiting());
self.addEventListener('activate',e=>e.waitUntil(self.clients.claim()));
self.addEventListener('fetch',e=>{const r=e.request;if(r.method!=='GET'||new URL(r.url).origin!==location.origin)return;
e.respondWith(fetch(r).then(res=>{if(res.ok){const c=res.clone();caches.open(C).then(x=>x.put(r,c));}return res;})
.catch(()=>caches.match(r,{ignoreSearch:true}).then(m=>m||caches.match('./'))));});
""")
try:
    from PIL import Image, ImageDraw, ImageFont
    for s in (192, 512, 180):
        im = Image.new("RGB", (s, s), "#0b4f8a"); d = ImageDraw.Draw(im)
        d.rounded_rectangle([s*.18, s*.2, s*.82, s*.62], radius=s*.06, fill="white")
        try: f = ImageFont.truetype("DejaVuSans-Bold.ttf", int(s*.2))
        except Exception: f = ImageFont.load_default()
        d.text((s/2, s*.41), "SSCO", fill="#0b4f8a", font=f, anchor="mm")
        try: f2 = ImageFont.truetype("DejaVuSans-Bold.ttf", int(s*.13))
        except Exception: f2 = f
        d.text((s/2, s*.78), "RUC", fill="white", font=f2, anchor="mm")
        im.save(SITE / ("apple-touch-icon.png" if s == 180 else f"icon-{s}.png"))
except Exception as e:
    print("AVISO: no se generaron íconos:", e)
print("OK")
