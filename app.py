#!/usr/bin/env python3
"""
Servicio generador de ofertas Q-SUN / SEG  (v10)
- /generar : crea el .docx y el .pdf (base64) a partir de los datos (como antes).
- Reserva ATOMICA de numero de oferta y estado por pasos (idempotencia) en Redis (Upstash).
  Endpoints: /reserve, /oferta/{num} (GET), /oferta/{num}/step (POST).
Pensado para llamarse desde n8n (nodo HTTP Request).

Variables de entorno:
  API_KEY       (opcional)  -> si se define, n8n debe enviar la cabecera x-api-key.
  REDIS_URL     (opcional)  -> rediss://... de Upstash. Si falta, los endpoints de
                              reserva/estado responden 503, pero /generar sigue OK.
  START_NUMBER  (def 20980) -> primer numero de oferta.
"""
import os, base64, tempfile, subprocess, shutil, uuid
from fastapi import FastAPI, Request, Header, HTTPException
from fastapi.responses import JSONResponse
import gen_offer

app = FastAPI(title="Generador de ofertas Q-SUN / SEG", version="10.0")

API_KEY = os.environ.get("API_KEY", "").strip()
START_NUMBER = int(os.environ.get("START_NUMBER", "20980"))
REDIS_URL = os.environ.get("REDIS_URL", "").strip()

# --- Redis (Upstash). Si no hay URL, r=None y los endpoints de estado dan 503. ---
r = None
if REDIS_URL:
    import redis
    r = redis.from_url(REDIS_URL, decode_responses=True)

COUNTER = "oferta:counter"


def _need_redis():
    if r is None:
        raise HTTPException(status_code=503, detail="Redis no configurado (define REDIS_URL en Render).")


def _init_counter():
    # deja el contador en START_NUMBER-1 la primera vez (INCR dara START_NUMBER)
    r.setnx(COUNTER, START_NUMBER - 1)


def _bump_counter_to(n):
    # sube el contador a n si va por detras (p.ej. numero puesto a mano mas alto)
    with r.pipeline() as p:
        while True:
            try:
                p.watch(COUNTER)
                c = int(p.get(COUNTER) or 0)
                if n > c:
                    p.multi(); p.set(COUNTER, n); p.execute()
                else:
                    p.unwatch()
                return
            except Exception:
                continue


STEP_FIELDS = ("generated", "drive", "emailed", "email_review", "historial", "cliente")


@app.get("/")
def health():
    return {"ok": True, "service": "generador-ofertas", "version": "10.0", "redis": bool(r)}


# ---------------- Reserva atomica de numero ----------------
@app.post("/reserve")
async def reserve(request: Request, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")
    _need_redis(); _init_counter()
    body = await request.json()

    # 1) si la fila YA trae número (reintento) -> se conserva; el contador sube si va por delante.
    existing = body.get("existing")
    try:
        ex = int(str(existing).strip()) if existing not in (None, "", "null") else 0
    except ValueError:
        ex = 0
    if ex > 0:
        _bump_counter_to(ex)
        return {"num": ex, "reused": True}

    # 2) fila nueva -> reserva IDEMPOTENTE por clave de contenido (fingerprint):
    #    dos llamadas de la MISMA oferta nueva devuelven el MISMO número.
    key = str(body.get("key") or "").strip()
    if key:
        m = r.get("reserve:" + key)
        if m:
            return {"num": int(m), "reused": True}
        num = r.incr(COUNTER)
        if not r.setnx("reserve:" + key, num):     # carrera con la MISMA clave
            num = int(r.get("reserve:" + key))     # -> usa el que ganó (hueco tolerado, sin duplicar)
        return {"num": num, "reused": False}

    # 3) sin clave (no debería ocurrir desde n8n) -> INCR simple
    return {"num": r.incr(COUNTER), "reused": False}


# ---------------- Bloqueo por oferta (con dueño/token) ----------------
LOCK_TTL = int(os.environ.get("LOCK_TTL", "600"))   # 10 min: cubre de sobra una oferta

@app.post("/lock/{num}")
def lock(num: str, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")
    _need_redis()
    token = uuid.uuid4().hex
    ok = r.set("lock:" + str(num), token, nx=True, ex=LOCK_TTL)
    return {"locked": bool(ok), "token": token if ok else ""}


@app.post("/lock/{num}/refresh")
async def lock_refresh(num: str, request: Request, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")
    _need_redis()
    body = await request.json()
    token = str(body.get("token") or "")
    k = "lock:" + str(num)
    with r.pipeline() as p:
        while True:
            try:
                p.watch(k)
                cur = p.get(k)
                if cur == token and token != "":
                    p.multi(); p.expire(k, LOCK_TTL); p.execute()   # renueva solo el dueño
                    return {"ok": True, "refreshed": True}
                else:
                    p.unwatch()
                    return {"ok": True, "refreshed": False}
            except Exception:
                continue


@app.post("/unlock/{num}")
async def unlock(num: str, request: Request, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")
    _need_redis()
    body = await request.json()
    token = str(body.get("token") or "")
    k = "lock:" + str(num)
    # solo el DUEÑO libera (check-and-del atómico con WATCH)
    with r.pipeline() as p:
        while True:
            try:
                p.watch(k)
                cur = p.get(k)
                if cur == token and token != "":
                    p.multi(); p.delete(k); p.execute()
                    return {"ok": True, "released": True}
                else:
                    p.unwatch()
                    return {"ok": True, "released": False}   # no era el dueño: no toca el bloqueo ajeno
            except Exception:
                continue


# ---------------- Estado por pasos (idempotencia) ----------------
def _estado(num):
    st = r.hgetall("oferta:" + str(num)) or {}
    out = {"num": str(num)}
    for f in STEP_FIELDS:
        out[f] = st.get(f) in ("1", "true", "True")
    out["drive_link"] = st.get("drive_link", "")
    out["total"] = st.get("total", "")
    out["filename"] = st.get("filename", "")
    return out


@app.get("/oferta/{num}")
def get_estado(num: str, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")
    _need_redis()
    return _estado(num)


@app.post("/oferta/{num}/step")
async def set_step(num: str, request: Request, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")
    _need_redis()
    body = await request.json()
    step = str(body.get("step") or "").strip()
    if step not in STEP_FIELDS:
        raise HTTPException(status_code=400, detail="step debe ser uno de: " + ", ".join(STEP_FIELDS))
    key = "oferta:" + str(num)
    val = body.get("value", True)
    r.hset(key, step, "1" if val else "0")
    data = body.get("data") or {}
    for k in ("drive_link", "total", "filename"):
        if k in data and data[k] is not None:
            r.hset(key, k, str(data[k]))
    return _estado(num)


# ---------------- Generacion del documento (como antes) ----------------
def _num(v, default=0):
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(" ", "")
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


def _build_lines(data):
    if data.get("lineas"):
        out = []
        for ln in data["lineas"]:
            out.append({
                "item": ln.get("item", ""), "peak": _num(ln.get("peak")),
                "pcs": int(_num(ln.get("pcs"))), "precio": _num(ln.get("precio")),
                "incoterm": ln.get("incoterm", ""),
            })
        return out
    modelo = data.get("modelo", ""); wps = _num(data.get("wps")); pcs = int(_num(data.get("pcs")))
    lines = []
    for i in range(1, 9):
        inc = data.get(f"incoterm{i}") or data.get(f"incoterm_{i}")
        pre = data.get(f"precio{i}") or data.get(f"precio_{i}")
        if inc and pre not in (None, "", 0):
            lines.append({"item": modelo, "peak": wps, "pcs": pcs,
                          "precio": _num(pre), "incoterm": str(inc).strip()})
    return lines


@app.post("/generar")
async def generar(request: Request, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")

    raw = await request.json()
    data = {
        "marca": (raw.get("marca") or "Q-SUN").strip(),
        "tipo": (raw.get("tipo") or "quotation").strip(),
        "idioma": (raw.get("idioma") or "es").strip(),
        "descripcion": (raw.get("descripcion") or raw.get("tecnologia") or "").strip(),
        "garantia_material": raw.get("garantia_material", 15),
        "garantia_potencia": raw.get("garantia_potencia", 30),
        "ficha_adjunta": bool(raw.get("ficha_adjunta")),
        "fecha": str(raw.get("fecha") or "").strip(),
        "num_oferta": str(raw.get("num_oferta") or raw.get("num") or "").strip(),
        "comprador": {
            "name": raw.get("empresa") or (raw.get("comprador", {}) or {}).get("name", ""),
            "add": raw.get("direccion") or (raw.get("comprador", {}) or {}).get("add", ""),
            "contact": raw.get("contacto") or (raw.get("comprador", {}) or {}).get("contact", ""),
            "tel": raw.get("tel") or (raw.get("comprador", {}) or {}).get("tel", ""),
            "cif": raw.get("cif") or (raw.get("comprador", {}) or {}).get("cif", ""),
            "email": raw.get("email") or (raw.get("comprador", {}) or {}).get("email", ""),
        },
        "lineas": _build_lines(raw),
    }
    if not data["lineas"]:
        raise HTTPException(status_code=400, detail="No hay lineas de producto (incoterm + precio)")

    total = None
    if len(data["lineas"]) == 1:
        ln = data["lineas"][0]
        p = round(float(ln["precio"]), 4)              # mismo redondeo que el documento
        base = ln["pcs"] * ln["peak"] * p
        iva = round(base * 0.21, 2)
        total = round(base + iva, 2)                   # base + IVA (igual que gen_offer)

    workdir = tempfile.mkdtemp()
    try:
        safe = f"Oferta_{data['marca']}_{data['num_oferta']}".replace(" ", "_").replace("/", "-")
        docx_path = os.path.join(workdir, safe + ".docx")
        try:
            gen_offer.generate(data, docx_path)
        except ValueError as e:
            raise HTTPException(status_code=400, detail=str(e))

        subprocess.run(["soffice", "--headless", "--convert-to", "pdf", "--outdir", workdir, docx_path],
                       check=True, timeout=120, stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL)
        pdf_path = docx_path[:-5] + ".pdf"
        if not os.path.exists(pdf_path):
            raise HTTPException(status_code=500, detail="No se pudo generar el PDF")

        with open(docx_path, "rb") as f: docx_b64 = base64.b64encode(f.read()).decode()
        with open(pdf_path, "rb") as f: pdf_b64 = base64.b64encode(f.read()).decode()

        # marca "generated" en el estado (idempotencia) si hay Redis y numero
        if r is not None and data["num_oferta"]:
            k = "oferta:" + data["num_oferta"]
            r.hset(k, "generated", "1")
            r.hset(k, "total", "" if total is None else str(total))
            r.hset(k, "filename", safe)

        return JSONResponse({
            "ok": True, "filename": safe, "marca": data["marca"],
            "num_oferta": data["num_oferta"], "total": total,
            "docx_base64": docx_b64, "pdf_base64": pdf_b64,
        })
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
