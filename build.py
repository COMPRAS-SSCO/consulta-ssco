"""Genera la página de consulta SSCO.
1) Descarga la relación de SSCO de SUNAT (si falla, usa data/ssco.xlsx).
2) Si existe padron.zip (Padrón Reducido del RUC de SUNAT), lo parte en fragmentos
   p/XXXXXX.gz (por los 6 primeros dígitos del RUC) para buscar la razón social.
   Si no existe, reutiliza los fragmentos guardados en padron/ (caché).
3) Escribe site/ con index.html, app instalable (manifest, sw.js, íconos)."""
import json, datetime, urllib.request, pathlib, sys, zipfile, io, gzip, shutil, collections, csv
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

# ---------- 2a. Actividad económica (Padrón RUC - datos abiertos SUNAT) ----------
ACT = {}; ACTS = []; actidx = {}; TPER = ""
dz = pathlib.Path("padronruc/PadronRUC.zip")
if dz.exists() and dz.stat().st_size > 1_000_000:
    try:
        csv.field_size_limit(10**7)
        with zipfile.ZipFile(dz) as z:
            nm = max(z.namelist(), key=lambda x: z.getinfo(x).file_size)
            with z.open(nm) as f:
                tf = io.TextIOWrapper(f, encoding="latin-1", errors="replace", newline="")
                head = tf.readline()
                dl = "|" if head.count("|") > head.count(",") else (";" if head.count(";") > head.count(",") else ",")
                H = [h.strip().strip('"').upper() for h in head.rstrip("\r\n").split(dl)]
                print("Padrón RUC (datos abiertos) columnas:", H)
                ri = next(i for i, h in enumerate(H) if h == "RUC" or ("RUC" in h and "TIPO" not in h))
                ac = [i for i, h in enumerate(H) if "CIIU" in h or "ACTIVIDAD" in h]
                fx = lambda *ks: next((i for i in ac if all(k in H[i] for k in ks)), None)
                ip = [x for x in (fx("REVISION4", "PRINCIPAL"), fx("PRINCIPAL")) if x is not None]
                isec = [x for x in (fx("REVISION4", "SECUNDARIA"), fx("SECUNDARIA")) if x is not None]
                it = next((i for i, h in enumerate(H) if "TRAB" in h), None)
                print("  columna RUC:", H[ri], "| trabajadores:", H[it] if it is not None else None, "| principal:", [H[i] for i in ip], "| secundaria:", [H[i] for i in isec])
                rd = csv.reader(tf, delimiter=dl)
                for k, r in enumerate(rd):
                    if k < 2: print("  ejemplo:", r)
                    if len(r) <= ri: continue
                    ruc = r[ri].strip()
                    if len(ruc) != 11: continue
                    def first(ix):
                        for i in ix:
                            v = r[i].strip() if i < len(r) else ""
                            if v and v.upper() not in ("-", "0", "NULL", "NO DISPONIBLE"): return v
                        return ""
                    pr, se = first(ip), first(isec)
                    tr = first([it]) if it is not None else ""
                    if not pr and not se and not tr: continue
                    j = ""
                    if pr or se:
                        s = (pr or "-") + ((" · " + se) if se and se != pr else "")
                        j = actidx.get(s)
                        if j is None: j = actidx[s] = len(ACTS); ACTS.append(s)
                    ACT[ruc] = f"{j}|{tr}"
        try:  # trabajadores (PLAME) = 2 meses antes del mes de publicación
            mp = pathlib.Path("padronruc/mes.txt").read_text().strip()
            y, m = int(mp[:4]), int(mp[4:6]) - 2
            if m < 1: y, m = y - 1, m + 12
            TPER = f"{y}-{m:02d}"
        except Exception: TPER = ""
        print(f"Actividad económica: {len(ACT)} RUC, {len(ACTS)} descripciones distintas; trabajadores periodo {TPER}")
    except Exception as e:
        print("AVISO: no se pudo leer el Padrón RUC de datos abiertos:", repr(e)); ACT = {}; ACTS = []
else:
    print("AVISO: no hay Padrón RUC de datos abiertos (actividad económica)")

# ---------- 2b. Padrón reducido ----------
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
                    a = ACT.get(p[0], "|")
                    shards[p[0][:6]].append(f"{p[0][6:]}|{nom}|{p[2].strip()}|{p[3].strip()}|{a}")
                    n += 1
        if n < 1_000_000: raise ValueError(f"solo {n} registros")
        nuevo = pathlib.Path("padron_nuevo"); shutil.rmtree(nuevo, ignore_errors=True); (nuevo / "p").mkdir(parents=True)
        for k, v in shards.items():
            (nuevo / "p" / f"{k}.gz").write_bytes(gzip.compress(("\n".join(v) + "\n").encode("utf-8"), 9))
        (nuevo / "fecha.txt").write_text(hoy)
        (nuevo / "act.json").write_text(json.dumps(ACTS, ensure_ascii=False), encoding="utf-8")
        (nuevo / "tper.txt").write_text(TPER)
        shutil.rmtree(PAD, ignore_errors=True); nuevo.rename(PAD)
        print(f"Padrón: {n} RUC en {len(shards)} fragmentos")
    except Exception as e:
        print("AVISO: no se pudo procesar el padrón, se usan los fragmentos anteriores:", e)
else:
    print("AVISO: no se descargó el padrón; se usan los fragmentos anteriores si existen")
pfecha = (PAD / "fecha.txt").read_text().strip() if (PAD / "fecha.txt").exists() else ""
tper = (PAD / "tper.txt").read_text().strip() if (PAD / "tper.txt").exists() else ""
if pfecha:
    shutil.rmtree(SITE / "p", ignore_errors=True); shutil.copytree(PAD / "p", SITE / "p")
    shutil.copy(PAD / "act.json", SITE / "act.json") if (PAD / "act.json").exists() else (SITE / "act.json").write_text("[]")
print("Padrón disponible:", pfecha or "NO")

# ---------- 3. Página + app instalable ----------
html = (pathlib.Path("template.html").read_text(encoding="utf-8")
        .replace("__DATA__", json.dumps(rows, ensure_ascii=False, separators=(",", ":")))
        .replace("__LAST__", last).replace("__CHK__", checked).replace("__PFECHA__", pfecha).replace("__TPER__", tper))
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
