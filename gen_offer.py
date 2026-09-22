#!/usr/bin/env python3
# Generador de ofertas Q-SUN / SEG  (bilingue ES/EN, formato numerico espanol)
import sys, json, os
from docx import Document
from docx.shared import Pt, Cm, RGBColor
from docx.enum.text import WD_ALIGN_PARAGRAPH
from docx.enum.table import WD_ALIGN_VERTICAL
from docx.oxml.ns import qn
from docx.oxml import OxmlElement

_BASE = os.path.dirname(os.path.abspath(__file__))
ASSETS = os.path.join(_BASE, "assets")
if not os.path.isdir(ASSETS):
    ASSETS = _BASE

GREY = "EDEDED"        # (reservado) fondo cabeceras
GREY_TOTAL = "DFF3D6"  # (reservado) verde claro
TOTAL_BG   = "F0F0F0"  # fondo suave (gris) para la fila Total
BORDER = "000000"      # bordes negros finos, bien definidos
FONT = "Liberation Sans"  # profesional, estable en Render (equivale a Arial)

BRANDS = {
 "Q-SUN": {
   "logo": "qsun_image1.png", "logo_w": 6.0,
   "seller": {"name":"JORESP ENERGY, S.L.","add":"Independencia, 24-26\n50004 Zaragoza",
              "contact":"Santiago Mediano García","tel":"+34 606 423 097","cif":"B99551350",
              "email":"s.mediano@joresp.net\nsantiago.mediano@q-sunsolar.group"},
   "bank": [("Banco","Banco Santander"),("Titular","Joresp Energy, S.L."),
            ("IBAN","ES80 0049 4470 3523 1003 7467"),("BIC/SWIFT","BSCHESMMXXX"),
            ("Dirección banco","Camino del Pilón, 61, 50011 Zaragoza"),("Divisa","EUR")],
 },
 "SEG": {
   "logo": "seg_image5.png", "logo_w": 4.0,
   "seller": {"name":"JORESP ENERGY, S.L.","add":"Independencia, 24-26\n50004 Zaragoza",
              "contact":"Santiago Mediano García","tel":"+34 606 423 097","cif":"B99551350",
              "email":"s.mediano@joresp.net"},
   "bank": [("Banco","Banco Santander"),("Titular","Joresp Energy, S.L."),
            ("IBAN","ES80 0049 4470 3523 1003 7467"),("BIC/SWIFT","BSCHESMMXXX"),
            ("Dirección banco","Camino del Pilón, 61, 50011 Zaragoza"),("Divisa","EUR")],
 },
}

# ---------- textos por idioma ----------
T = {
 "es": {
  "title":"Oferta comercial", "offer_no":"N.º de oferta", "date":"Fecha",
  "title_proforma":"Factura proforma", "offer_no_proforma":"N.º de proforma", "foot_proforma":"Proforma",
  "buyer":"Comprador", "seller":"Vendedor",
  "add":"Dirección","phone":"Teléfono","taxid":"NIF","contact":"Contacto","email":"Email",
  "supply":"Detalle del suministro",
  "h_item":"Artículo","h_modp":"Potencia módulo (Wp)","h_qty":"Cantidad (uds)",
  "h_totp":"Potencia total (Wp)","h_price":"Precio (€/Wp)","h_inco":"Incoterm","h_amount":"Importe sin IVA (€)",
  "t_base":"Importe sin IVA","t_iva":"IVA 21%","t_total":"Total",
  "t_adv":"Anticipo 10%","t_bal":"Saldo 90% contra B/L",
  "pay_title":"Condiciones de pago",
  "pay1":"El comprador pagará el 10% del importe total por adelantado y el 90% restante contra presentación del conocimiento de embarque (B/L).",
  "pay2":"El comprador asume los gastos bancarios del país emisor. El vendedor asume los gastos bancarios del país receptor.",
  "pay3":"Datos bancarios:",
  "ship_title":"Condiciones de entrega","ship_arr":"Incoterms acordados:","ship_one":"Incoterm acordado:",
  "ship_t1":"Plazo de entrega estimado: 10 a 12 semanas desde la confirmación del pedido.",
  "ship_t2":"La fecha concreta se confirmará por escrito para cada pedido, conforme al Incoterm acordado.",
  "validity":"Validez de la oferta: 10 días desde la fecha de emisión.",
  "prod_title":"Producto y garantías","w_material":"Garantía de material","w_power":"Garantía de potencia lineal","w_years":"años",
  "w_note":"Garantías conforme a las condiciones del fabricante aplicables al modelo ofertado.",
  "datasheet":"Ficha técnica disponible bajo petición.","datasheet_att":"Ficha técnica adjunta.",
  "accept":"Aceptación del comprador","a_name":"Nombre","a_pos":"Cargo","a_date":"Fecha","a_sign":"Firma",
  "page":"Página","of":"de",
 },
 "en": {
  "title":"Quotation", "offer_no":"Quotation No.", "date":"Date",
  "title_proforma":"Proforma Invoice", "offer_no_proforma":"Proforma No.", "foot_proforma":"Proforma",
  "buyer":"Buyer", "seller":"Seller",
  "add":"Address","phone":"Phone","taxid":"Tax ID","contact":"Contact","email":"Email",
  "supply":"Supply details",
  "h_item":"Item","h_modp":"Module power (Wp)","h_qty":"Quantity (pcs)",
  "h_totp":"Total power (Wp)","h_price":"Price (€/Wp)","h_inco":"Incoterm","h_amount":"Amount excl. VAT (€)",
  "t_base":"Amount excl. VAT","t_iva":"VAT 21%","t_total":"Total",
  "t_adv":"10% advance payment","t_bal":"90% balance against B/L",
  "pay_title":"Payment terms",
  "pay1":"The Buyer shall pay 10% of the total amount in advance and the remaining 90% against presentation of the Bill of Lading (B/L).",
  "pay2":"The Buyer shall bear all bank charges incurred in the remitting country. The Seller shall bear all bank charges incurred in the receiving country.",
  "pay3":"Bank details:",
  "ship_title":"Shipment terms","ship_arr":"Agreed Incoterms:","ship_one":"Agreed Incoterm:",
  "ship_t1":"Estimated delivery time: 10 to 12 weeks from order confirmation.",
  "ship_t2":"The exact date will be confirmed in writing for each order, in accordance with the agreed Incoterm.",
  "validity":"Offer validity: 10 days from the issue date.",
  "prod_title":"Product and warranties","w_material":"Material warranty","w_power":"Linear power warranty","w_years":"years",
  "w_note":"Warranties subject to the manufacturer's conditions applicable to the quoted model.",
  "datasheet":"Datasheet available on request.","datasheet_att":"Datasheet attached.",
  "accept":"Buyer's acceptance","a_name":"Name","a_pos":"Position","a_date":"Date","a_sign":"Signature",
  "page":"Page","of":"of",
 },
}
BANK_LABELS = {
 "Banco":{"es":"Banco","en":"Bank"},"Titular":{"es":"Titular","en":"Account name"},
 "IBAN":{"es":"IBAN","en":"IBAN"},"BIC/SWIFT":{"es":"BIC/SWIFT","en":"BIC/SWIFT"},
 "Dirección banco":{"es":"Dirección banco","en":"Bank address"},"Divisa":{"es":"Divisa","en":"Currency"},
}

# ---------- formato numerico espanol (en ambos idiomas) ----------
def nmoney(n):
    s=f"{n:,.2f}"                       # 1,234.56
    return s.replace(",","§").replace(".",",").replace("§",".")
def nint(n):
    return f"{int(round(n)):,}".replace(",",".")
def nprice(n):
    s=f"{float(n):.4f}"
    if s.endswith("0"): s=s[:-1]        # 0,1080 -> 0,108 ; 0,1085 -> 4 dec
    return s.replace(".",",")

# ---------- helpers docx ----------
def _run(p, text, bold=False, size=10.5, align=None, color=None):
    if align is not None: p.alignment=align
    r=p.add_run(text); r.bold=bold; r.font.size=Pt(size); r.font.name=FONT
    if color: r.font.color.rgb=color
    return r

def set_cell(cell, text, bold=False, size=10, align="left", color=None, shade=None, underline=False):
    cell.text=""
    p=cell.paragraphs[0]
    p.alignment={"left":WD_ALIGN_PARAGRAPH.LEFT,"center":WD_ALIGN_PARAGRAPH.CENTER,"right":WD_ALIGN_PARAGRAPH.RIGHT}[align]
    p.paragraph_format.space_after=Pt(0); p.paragraph_format.space_before=Pt(0)
    p.paragraph_format.line_spacing=1.0
    for i,line in enumerate(str(text).split("\n")):
        if i>0: p.add_run().add_break()
        r=p.add_run(line); r.bold=bold; r.underline=underline; r.font.size=Pt(size); r.font.name=FONT
        if color: r.font.color.rgb=color
    cell.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
    if shade: _shade(cell, shade)

def cell_hide_borders(cell, edges):
    """Quita solo los bordes indicados (el resto se hereda de la tabla)."""
    tcPr=cell._tc.get_or_add_tcPr()
    for old in tcPr.findall(qn('w:tcBorders')): tcPr.remove(old)
    b=OxmlElement('w:tcBorders')
    for edge in edges:
        e=OxmlElement('w:'+edge); e.set(qn('w:val'),'nil'); b.append(e)
    shd=tcPr.find(qn('w:shd'))
    if shd is not None: shd.addprevious(b)   # tcBorders debe ir antes de shd
    else: tcPr.append(b)

def _shade(cell,hexc):
    tcPr=cell._tc.get_or_add_tcPr()
    sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),hexc); tcPr.append(sh)

def cell_margins(tb, top=60, bottom=60, left=90, right=90):
    m=OxmlElement('w:tblCellMar')
    for side,val in (('top',top),('bottom',bottom),('start',left),('end',right),('left',left),('right',right)):
        e=OxmlElement('w:'+side); e.set(qn('w:w'),str(val)); e.set(qn('w:type'),'dxa'); m.append(e)
    tb._tbl.tblPr.append(m)

def row_height(row, h=300):
    trPr=row._tr.get_or_add_trPr()
    tr=OxmlElement('w:trHeight'); tr.set(qn('w:val'),str(h)); tr.set(qn('w:hRule'),'atLeast'); trPr.append(tr)

def cell_borders(cell, color=BORDER, sz=4):
    tcPr=cell._tc.get_or_add_tcPr()
    b=OxmlElement('w:tcBorders')
    for edge in ('top','left','bottom','right'):
        e=OxmlElement('w:'+edge); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),str(sz))
        e.set(qn('w:space'),'0'); e.set(qn('w:color'),color); b.append(e)
    tcPr.append(b)

def tbl_borders(tb, color=BORDER, sz=4):
    borders=OxmlElement('w:tblBorders')
    for edge in ('top','left','bottom','right','insideH','insideV'):
        e=OxmlElement('w:'+edge); e.set(qn('w:val'),'single'); e.set(qn('w:sz'),str(sz))
        e.set(qn('w:space'),'0'); e.set(qn('w:color'),color); borders.append(e)
    tb._tbl.tblPr.append(borders)

def set_default_font(doc, name=FONT):
    st=doc.styles['Normal']; st.font.name=name; st.font.size=Pt(10.5)
    rpr=st.element.get_or_add_rPr(); rf=rpr.get_or_add_rFonts()
    for a in ('w:ascii','w:hAnsi','w:cs'): rf.set(qn(a),name)

def norm_marca(v):
    mk=str(v or "").strip().upper().replace(" ","")
    if mk in ("Q-SUN","QSUN"): return "Q-SUN"
    if mk=="SEG": return "SEG"
    raise ValueError(f"Marca desconocida: '{v}'. Debe ser Q-SUN o SEG.")

def norm_idioma(v):
    x=str(v or "es").strip().lower()
    return "en" if x.startswith("en") or x.startswith("in") else "es"

def _field(p, code, size=9):
    r=p.add_run(); r.font.size=Pt(size); r.font.name=FONT
    b=OxmlElement('w:fldChar'); b.set(qn('w:fldCharType'),'begin'); r._r.append(b)
    i=OxmlElement('w:instrText'); i.set(qn('xml:space'),'preserve'); i.text=' '+code+' '; r._r.append(i)
    e=OxmlElement('w:fldChar'); e.set(qn('w:fldCharType'),'end'); r._r.append(e)

def add_footer(section, prefix, page_word, of_word):
    p=section.footer.paragraphs[0]; p.text=""
    p.alignment=WD_ALIGN_PARAGRAPH.CENTER
    _run(p, f"{prefix}   ·   {page_word} ", size=9)
    _field(p, "PAGE", size=9)
    _run(p, f" {of_word} ", size=9)
    _field(p, "NUMPAGES", size=9)

def para(d, text, bold=False, size=10.5, after=6, before=0, align=WD_ALIGN_PARAGRAPH.LEFT, indent=0):
    p=d.add_paragraph(); p.alignment=align
    p.paragraph_format.space_after=Pt(after); p.paragraph_format.space_before=Pt(before)
    if indent: p.paragraph_format.left_indent=Cm(indent)
    if text: _run(p, text, bold=bold, size=size)
    return p

def section_title(d, text, size=12, after=8, before=0, page_break=False):
    """Titulo de seccion con regla fina inferior (aspecto profesional)."""
    p=d.add_paragraph()
    if page_break: p.paragraph_format.page_break_before=True
    p.paragraph_format.space_after=Pt(after); p.paragraph_format.space_before=Pt(before)
    p.paragraph_format.right_indent=Cm(0.2)   # la regla acaba a la altura de las tablas (17,6 cm)
    _run(p, text, bold=True, size=size)
    pPr=p._p.get_or_add_pPr()
    pbdr=OxmlElement('w:pBdr')
    bottom=OxmlElement('w:bottom')
    bottom.set(qn('w:val'),'single'); bottom.set(qn('w:sz'),'8'); bottom.set(qn('w:space'),'4'); bottom.set(qn('w:color'),'000000')
    pbdr.append(bottom); pPr.append(pbdr)
    return p

def hide_vert_2col(tb):
    """En una tabla de 2 columnas, quita la vertical interna (mantiene horizontales y recuadro)."""
    for r in tb.rows:
        cell_hide_borders(r.cells[0], ('right',))
        cell_hide_borders(r.cells[1], ('left',))

def set_widths(tb, widths, fixed=False):
    tb.autofit=False; tb.allow_autofit=False
    if fixed:
        tblPr=tb._tbl.tblPr
        lay=OxmlElement('w:tblLayout'); lay.set(qn('w:type'),'fixed'); tblPr.append(lay)
        grid=tb._tbl.find(qn('w:tblGrid'))
        if grid is not None:
            for gc,w in zip(grid.findall(qn('w:gridCol')),widths):
                gc.set(qn('w:w'), str(int(round(Cm(w).pt*20))))
    for j,w in enumerate(widths):
        for r in tb.rows:
            r.cells[j].width=Cm(w)

def generate(data, out_path):
    data["marca"]=norm_marca(data.get("marca"))
    lang=norm_idioma(data.get("idioma"))
    tr=T[lang]; b=BRANDS[data["marca"]]
    RED=RGBColor(0xC0,0x00,0x00)

    # tipo: proforma cambia titulo, subtitulo y pie
    is_prof = str(data.get("tipo","")).strip().lower().startswith("proforma")
    doc_title   = tr["title_proforma"]    if is_prof else tr["title"]
    doc_offerno = tr["offer_no_proforma"] if is_prof else tr["offer_no"]
    foot_word   = tr["foot_proforma"]     if is_prof else ("Oferta" if lang=="es" else "Offer")

    d=Document(); set_default_font(d)
    sec=d.sections[0]
    for m in ("top_margin","bottom_margin"): setattr(sec,m,Cm(1.6))
    for m in ("left_margin","right_margin"): setattr(sec,m,Cm(1.0))   # margenes laterales menores: cuadros mas anchos

    # pie: Oferta/Proforma 20982 · Página 1 de 2
    add_footer(sec, f"{foot_word} {data['num_oferta']}", tr['page'], tr['of'])

    # logo
    p=d.add_paragraph(); p.paragraph_format.space_after=Pt(2)
    try: p.add_run().add_picture(os.path.join(ASSETS,b["logo"]), width=Cm(b["logo_w"]))
    except Exception: pass

    # titulo + (nº y fecha justo debajo)
    para(d, doc_title, bold=True, size=21, after=2, align=WD_ALIGN_PARAGRAPH.CENTER)
    sub=d.add_paragraph(); sub.alignment=WD_ALIGN_PARAGRAPH.CENTER
    sub.paragraph_format.space_after=Pt(12)   # aire antes de los cuadros
    _run(sub, f"{doc_offerno}: ", size=11); _run(sub, str(data['num_oferta']), bold=True, size=11)
    _run(sub, f"      {tr['date']}: ", size=11); _run(sub, str(data['fecha']), bold=True, size=11)

    # comprador / vendedor
    # comprador y vendedor en UNA sola tabla (celdas idénticas), con hueco central
    pt=d.add_table(rows=6,cols=5)
    set_widths(pt, [2.7, 6.75, 0.1, 2.7, 6.75], fixed=True)   # 19.0 cm; columnas de datos mas anchas
    cell_margins(pt, top=80, bottom=80, left=180, right=150)   # bastante aire dentro de las celdas
    bu=data["comprador"]; se=b["seller"]
    labL=[tr["buyer"],tr["add"],tr["contact"],tr["phone"],tr["taxid"],tr["email"]]
    labR=[tr["seller"],tr["add"],tr["contact"],tr["phone"],tr["taxid"],tr["email"]]
    valL=[bu["name"],bu["add"],bu["contact"],bu["tel"],bu["cif"],bu["email"]]
    valR=[se["name"],se["add"],se["contact"],se["tel"],se["cif"],se["email"]]
    for i in range(6):
        set_cell(pt.rows[i].cells[0], labL[i], size=10, bold=True)         # etiquetas en negrita
        set_cell(pt.rows[i].cells[1], valL[i], size=10, bold=(i==0))       # todo mismo tamano (10)
        set_cell(pt.rows[i].cells[3], labR[i], size=10, bold=True)         # etiquetas en negrita
        set_cell(pt.rows[i].cells[4], valR[i], size=10, bold=(i==0))
        for c in (0,1,3,4):
            cell_borders(pt.rows[i].cells[c])                              # bordes negros solo en los dos cuadros
            for p in pt.rows[i].cells[c].paragraphs: p.paragraph_format.line_spacing=1.2   # lineas mas separadas
        row_height(pt.rows[i], 560 if i in (1,5) else 380)    # mas altura: filas mas holgadas
    para(d,"",after=8)    # separacion con Detalle del suministro
    section_title(d, tr["supply"], after=6)                     # titulo del detalle con regla

    # tabla de producto + resumen economico INTEGRADO (una sola tabla, protagonista)
    lines=data["lineas"]; single=(len(lines)==1)
    n_items=len(lines)
    pad=0                            # sin filas en blanco (la tabla ya es protagonista)
    n_prod=n_items+pad
    n_sum=5                          # importe sin IVA, IVA, total, anticipo, saldo
    tb=d.add_table(rows=1+n_prod+n_sum, cols=7)
    hdr=[tr["h_item"],tr["h_modp"],tr["h_qty"],tr["h_totp"],tr["h_price"],tr["h_inco"],tr["h_amount"]]
    for j,h in enumerate(hdr):
        set_cell(tb.rows[0].cells[j],h,bold=True,size=10,align="center")   # cabecera clara (sin fondo)
    aligns=["center","right","right","right","right","center","right"]
    base=0.0
    for i,ln in enumerate(lines):
        precio=round(float(ln["precio"]),4)
        totp=ln["pcs"]*ln["peak"]; amount=totp*precio; base+=amount
        vals=[ln["item"], nint(ln["peak"]), nint(ln["pcs"]), nint(totp),
              nprice(precio), ln["incoterm"], nmoney(amount)]
        bolds=[True, False, False, False, False, False, False]  # solo el modelo en negrita
        sizes=[11.5, 11.5, 10, 10, 11.5, 10, 10]               # modelo/potencia/precio algo mayores
        for j,v in enumerate(vals):
            set_cell(tb.rows[1+i].cells[j],v,size=sizes[j],align=aligns[j],bold=bolds[j])
    # resumen economico
    if single:
        iva=round(base*0.21,2); total=round(base+iva,2)
        adv=round(total*0.10,2); bal=round(total-adv,2)     # saldo = total - anticipo
        svals=[nmoney(base),nmoney(iva),nmoney(total),nmoney(adv),nmoney(bal)]
    else:
        svals=["—","—","—","—","—"]
    slabels=[tr["t_base"],tr["t_iva"],tr["t_total"],tr["t_adv"],tr["t_bal"]]
    set_widths(tb, [5.15, 2.1, 2.05, 2.3, 1.7, 2.2, 3.5], fixed=True)   # 19.0 cm
    r0=1+n_prod
    for i in range(n_sum):
        r=tb.rows[r0+i]; istotal=(i==2)
        tshade = TOTAL_BG if istotal else None                                    # fondo suave (gris) solo en Total
        set_cell(r.cells[0], slabels[i], size=(11 if istotal else 10), bold=istotal, align="left", shade=tshade)
        set_cell(r.cells[6], (svals[i]+" €" if svals[i]!="—" else svals[i]), size=(11 if istotal else 10), bold=istotal, align="right", shade=tshade)
        r.cells[1].merge(r.cells[5])   # zona central en blanco, importe alineado a la derecha
        if istotal: _shade(r.cells[1], TOTAL_BG)                                   # fondo tambien en la zona central
        # quitar SOLO las verticales internas (columnas); se mantienen horizontales y recuadro
        cell_hide_borders(r.cells[0], ('right',))
        cell_hide_borders(r.cells[1], ('left','right'))
        cell_hide_borders(r.cells[6], ('left',))
    cell_margins(tb, top=85, bottom=85, left=110, right=110)
    for k,rr in enumerate(tb.rows):
        if k==0: h=340                       # cabecera
        elif k<1+n_items: h=500              # fila(s) de producto (protagonista)
        elif k<1+n_prod: h=380               # fila(s) en blanco
        else: h=250                          # resumen economico
        row_height(rr, h)
    tbl_borders(tb)

    # producto y garantias (primera hoja, tras el resumen)
    para(d,"",after=6)                                       # pequeno espacio con la tabla superior
    section_title(d, tr["prod_title"], after=4, before=8)    # separacion respecto a la tabla
    prod_prefix = f"{lines[0]['item']} ({nint(lines[0]['peak'])} W) " if lines else ""
    prod_line = (prod_prefix + str(data.get("descripcion") or "")).strip()
    if prod_line: para(d, prod_line, size=10.5, after=4)     # modelo (Wp) + descripcion, misma linea, sin negrita
    gm = data.get("garantia_material", 15); gp = data.get("garantia_potencia", 30)
    wr=[(tr["w_material"],gm),(tr["w_power"],gp)]
    wr=[(lab,y) for lab,y in wr if y not in (None,"",0)]
    if wr:
        wt=d.add_table(rows=1,cols=len(wr))                     # garantias una al lado de otra
        for i,(lab,y) in enumerate(wr):
            c=wt.rows[0].cells[i]; c.text=""
            p=c.paragraphs[0]; p.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p.paragraph_format.space_after=Pt(0); p.paragraph_format.space_before=Pt(0)
            _run(p, str(y), bold=True, size=22); _run(p, " "+tr["w_years"], bold=True, size=12)
            p2=c.add_paragraph(); p2.alignment=WD_ALIGN_PARAGRAPH.CENTER
            p2.paragraph_format.space_before=Pt(2); p2.paragraph_format.space_after=Pt(0)
            _run(p2, lab, size=10)
            c.vertical_alignment=WD_ALIGN_VERTICAL.CENTER
        set_widths(wt, [19.0/len(wr)]*len(wr), fixed=True)
        cell_margins(wt, top=55, bottom=55, left=110, right=110)
        for rr in wt.rows: row_height(rr, 360)
        tbl_borders(wt)
    para(d,"",after=2)
    para(d, tr["w_note"], size=10, after=2)                                   # nota de garantias

    # pagina 2: empieza en Condiciones de pago (salto antes del titulo)
    section_title(d, tr["pay_title"], after=10, page_break=True)
    para(d, "1.1   "+tr["pay1"], size=10.5, after=10)
    para(d, "1.2   "+tr["pay2"], size=10.5, after=10)
    para(d, "1.3   "+tr["pay3"], size=10.5, after=6)
    bt=d.add_table(rows=len(b["bank"]),cols=2)
    for i,(k,v) in enumerate(b["bank"]):
        set_cell(bt.rows[i].cells[0], BANK_LABELS[k][lang], size=10, bold=True)
        set_cell(bt.rows[i].cells[1], v, size=10)
    set_widths(bt, [4.6, 14.4], fixed=True)   # recuadro a todo el ancho
    cell_margins(bt, top=55, bottom=55, left=110, right=110)
    for rr in bt.rows: row_height(rr, 300)
    tbl_borders(bt); hide_vert_2col(bt)        # recuadro con horizontales, sin vertical interna

    para(d,"",after=12)
    # condiciones de entrega
    section_title(d, tr["ship_title"], after=10)
    if single:
        para(d, "2.1   "+tr["ship_one"]+" "+str(lines[0]["incoterm"]), size=10.5, after=10)
    else:
        para(d, "2.1   "+tr["ship_arr"], size=10.5, after=4)
        for ln in lines:
            para(d, "        • "+str(ln["incoterm"]), size=10.5, after=2)
        para(d,"",after=4)
    para(d, "2.2   "+tr["ship_t1"], size=10.5, after=10)
    para(d, "2.3   "+tr["validity"], size=10.5, after=16)

    # separacion antes de la aceptacion (hacia la parte inferior)
    esp=d.add_paragraph(); esp.paragraph_format.space_before=Pt(72)

    # aceptacion: casillas de firma enmarcadas (aspecto profesional)
    section_title(d, tr["accept"], after=14)
    at=d.add_table(rows=2,cols=2)
    cells=[(tr["a_name"],tr["a_sign"]),(tr["a_pos"],tr["a_date"])]
    for i,(k1,k2) in enumerate(cells):
        set_cell(at.rows[i].cells[0], k1, size=10.5, bold=True)
        set_cell(at.rows[i].cells[1], k2, size=10.5, bold=True)
        for c in (at.rows[i].cells[0], at.rows[i].cells[1]):
            c.vertical_alignment=WD_ALIGN_VERTICAL.TOP
    set_widths(at, [9.5, 9.5], fixed=True)
    cell_margins(at, top=120, bottom=520, left=140, right=140)   # espacio amplio para firmar
    for rr in at.rows: row_height(rr, 1000)
    tbl_borders(at)

    d.save(out_path)

if __name__=="__main__":
    data=json.load(open(sys.argv[1],encoding="utf-8"))
    generate(data, sys.argv[2]); print("OK ->", sys.argv[2])
