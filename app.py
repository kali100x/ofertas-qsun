#!/usr/bin/env python3
"""
Servicio generador de ofertas Q-SUN / SEG.
Recibe los datos por HTTP (JSON) y devuelve el .docx y el .pdf en base64.
Pensado para llamarse desde n8n (nodo HTTP Request).
"""
import os, base64, tempfile, subprocess, shutil
from fastapi import FastAPI, Request, Header, HTTPException
from fastapi.responses import JSONResponse
import gen_offer

app = FastAPI(title="Generador de ofertas Q-SUN / SEG")

# Clave opcional: si defines la variable de entorno API_KEY en Render,
# n8n debe enviar la cabecera  x-api-key  con ese mismo valor.
API_KEY = os.environ.get("API_KEY", "").strip()


@app.get("/")
def health():
    return {"ok": True, "service": "generador-ofertas", "version": "1.0"}


def _num(v, default=0):
    """Convierte texto tipo '7.344' o '0,108' a numero."""
    if v is None or v == "":
        return default
    if isinstance(v, (int, float)):
        return v
    s = str(v).strip().replace(" ", "")
    # quitar separador de miles y usar punto decimal
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except ValueError:
        return default


def _build_lines(data):
    """Acepta 'lineas' ya montadas, o pares incoterm/precio sueltos."""
    if data.get("lineas"):
        out = []
        for ln in data["lineas"]:
            out.append({
                "item": ln.get("item", ""),
                "peak": _num(ln.get("peak")),
                "pcs": int(_num(ln.get("pcs"))),
                "precio": _num(ln.get("precio")),
                "incoterm": ln.get("incoterm", ""),
            })
        return out
    # modo tabla: modelo/wps/pcs comunes + pares incoterm_i / precio_i
    modelo = data.get("modelo", "")
    wps = _num(data.get("wps"))
    pcs = int(_num(data.get("pcs")))
    lines = []
    for i in range(1, 9):
        inc = data.get(f"incoterm{i}") or data.get(f"incoterm_{i}")
        pre = data.get(f"precio{i}") or data.get(f"precio_{i}")
        if inc and pre not in (None, "", 0):
            lines.append({
                "item": modelo, "peak": wps, "pcs": pcs,
                "precio": _num(pre), "incoterm": str(inc).strip(),
            })
    return lines


@app.post("/generar")
async def generar(request: Request, x_api_key: str = Header(default="")):
    if API_KEY and x_api_key != API_KEY:
        raise HTTPException(status_code=401, detail="API key incorrecta")

    raw = await request.json()

    data = {
        "marca": (raw.get("marca") or "Q-SUN").strip(),
        "tipo": (raw.get("tipo") or "quotation").strip(),
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

    # total (solo si hay un unico incoterm)
    total = None
    if len(data["lineas"]) == 1:
        ln = data["lineas"][0]
        base_amount = ln["pcs"] * ln["peak"] * ln["precio"]
        total = round(base_amount * 1.21, 2)

    workdir = tempfile.mkdtemp()
    try:
        safe = f"Oferta_{data['marca']}_{data['num_oferta']}".replace(" ", "_").replace("/", "-")
        docx_path = os.path.join(workdir, safe + ".docx")
        gen_offer.generate(data, docx_path)

        subprocess.run(
            ["soffice", "--headless", "--convert-to", "pdf", "--outdir", workdir, docx_path],
            check=True, timeout=120,
            stdout=subprocess.DEVNULL, stderr=subprocess.DEVNULL,
        )
        pdf_path = docx_path[:-5] + ".pdf"
        if not os.path.exists(pdf_path):
            raise HTTPException(status_code=500, detail="No se pudo generar el PDF")

        with open(docx_path, "rb") as f:
            docx_b64 = base64.b64encode(f.read()).decode()
        with open(pdf_path, "rb") as f:
            pdf_b64 = base64.b64encode(f.read()).decode()

        return JSONResponse({
            "ok": True,
            "filename": safe,
            "marca": data["marca"],
            "num_oferta": data["num_oferta"],
            "total": total,
            "docx_base64": docx_b64,
            "pdf_base64": pdf_b64,
        })
    finally:
        shutil.rmtree(workdir, ignore_errors=True)
