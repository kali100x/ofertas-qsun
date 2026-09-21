#!/usr/bin/env python3
# Generador de ofertas Q-SUN / SEG
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
    ASSETS = _BASE   # imagenes en la misma carpeta (raiz del repo)

BRANDS = {
 "Q-SUN": {
   "logo": "qsun_image1.png",
   "seller": {"name":"JORESP ENERGY, S.L.","add":"Independencia, 24-26\n50004 Zaragoza",
              "contact":"Santiago Mediano García","tel":"+34 606 423 097","cif":"B99551350",
              "email":"s.mediano@joresp.net\nsantiago.mediano@q-sunsolar.group"},
   "bank": ["Bank's name:\tBanco Santander","Account Name:\tJoresp Energy, S.L.",
            "Account No:\tES80 0049 4470 3523 1003 7467","Bic/Swift:\tBSCHESMMXXX",
            "Bank Address:\tCamino del Pilón, 61, 50011 Zaragoza","Account Currency:\tEUR"],
   "shipment_tail": [],
   "warranty": None,
   "numfmt": "eu",
   "footer": None, "footer_img": None,
 },
 "SEG": {
   "logo": "seg_image5.png",
   "seller": {"name":"SEG IBERIA S.L.","add":"Paseo Independencia, 24-26\n50004 Zaragoza",
              "contact":"Santiago Mediano García","tel":"+34 606 423 097","cif":"B01961143",
              "email":"s.mediano@joresp.net"},
   "bank": ["Bank's name:\tBanco Santander","Account Name:\tJoresp Energy, S.L.",
            "Account No:\tES80 0049 4470 3523 1003 7467","Bic/Swift:\tBSCHESMMXXX",
            "Bank Address:\tCamino del Pilón, 61, 50011 Zaragoza","Account Currency:\tEUR"],
   "shipment_tail": ["- 10 weeks of order confirmation."],
   "warranty": None,
   "numfmt": "us",
   "footer": None, "footer_img": None,
 },
}

def fmt(n, kind, money=False):
    if kind=="eu":
        s=f"{n:,.2f}".replace(",","X").replace(".",",").replace("X",".")
        if not money:  # integer-ish
            s=f"{n:,.0f}".replace(",",".")
    else:
        s=f"{n:,.2f}" if money else f"{n:,.0f}"
    return s

def fmt3(n, kind):
    if kind=="eu":
        return f"{n:,.3f}".replace(",","X").replace(".",",").replace("X",".")
    return f"{n:,.3f}"

def fmt_price(n, kind):
    # muestra el precio con 3 o 4 decimales, exactamente el valor usado en el calculo
    s=f"{n:.4f}"
    if s.endswith("0"): s=s[:-1]   # 0,1080 -> 0,108 ; 0,1085 se queda con 4
    if kind=="eu": s=s.replace(".",",")
    return s

def norm_marca(v):
    mk=str(v or "").strip().upper().replace(" ","")
    if mk in ("Q-SUN","QSUN"): return "Q-SUN"
    if mk=="SEG": return "SEG"
    raise ValueError(f"Marca desconocida: '{v}'. Debe ser Q-SUN o SEG.")

def set_cell(cell, text, bold=False, size=10, align="left", color=None):
    cell.text=""
    p=cell.paragraphs[0]
    p.alignment={"left":WD_ALIGN_PARAGRAPH.LEFT,"center":WD_ALIGN_PARAGRAPH.CENTER,"right":WD_ALIGN_PARAGRAPH.RIGHT}[align]
    for i,line in enumerate(str(text).split("\n")):
        if i>0: p.add_run().add_break()
        r=p.add_run(line); r.bold=bold; r.font.size=Pt(size)
        if color: r.font.color.rgb=color
    cell.vertical_alignment=WD_ALIGN_VERTICAL.CENTER

def shade(cell,hexc):
    tcPr=cell._tc.get_or_add_tcPr()
    sh=OxmlElement('w:shd'); sh.set(qn('w:val'),'clear'); sh.set(qn('w:fill'),hexc)
    tcPr.append(sh)


def cell_margins(tb, top=60, bottom=60, left=90, right=90):
    tblPr=tb._tbl.tblPr
    m=OxmlElement('w:tblCellMar')
    for side,val in (('top',top),('bottom',bottom),('start',left),('end',right),('left',left),('right',right)):
        e=OxmlElement('w:'+side); e.set(qn('w:w'),str(val)); e.set(qn('w:type'),'dxa'); m.append(e)
    tblPr.append(m)

def row_height(row, h=340):
    trPr=row._tr.get_or_add_trPr()
    tr=OxmlElement('w:trHeight'); tr.set(qn('w:val'),str(h)); tr.set(qn('w:hRule'),'atLeast'); trPr.append(tr)

def set_default_font(doc, name="Arial"):
    st=doc.styles['Normal']
    st.font.name=name
    rpr=st.element.get_or_add_rPr(); rf=rpr.get_or_add_rFonts()
    for a in ('w:ascii','w:hAnsi','w:cs'): rf.set(qn(a),name)

def generate(data, out_path):
    data["marca"]=norm_marca(data.get("marca"))   # acepta 'q-sun', ' SEG ', etc.
    b=BRANDS[data["marca"]]
    nf=b["numfmt"]; RED=RGBColor(0xC0,0x00,0x00)
    d=Document()
    set_default_font(d, "Times New Roman")   # tipografia formal (serif)
    sec=d.sections[0]
    for m in ("top_margin","bottom_margin","left_margin","right_margin"): setattr(sec,m,Cm(1.6))
    # logo
    logo=os.path.join(ASSETS,b["logo"])
    p=d.add_paragraph(); r=p.add_run()
    try: r.add_picture(logo, width=Cm(6) if data["marca"]=="Q-SUN" else Cm(4))
    except Exception: pass
    # contract / date / no
    for txt in ["Contract",f"Date: {data['fecha']}",f"Contract NO.: {data['num_oferta']}"]:
        pp=d.add_paragraph(); rr=pp.add_run(txt); rr.font.size=Pt(10)
        pp.paragraph_format.space_after=Pt(0)
    # title
    t=d.add_paragraph(); t.alignment=WD_ALIGN_PARAGRAPH.CENTER
    titulo="PROFORMA INVOICE" if str(data.get("tipo","quotation")).strip().lower().startswith("pro") else "QUOTATION"
    rt=t.add_run(titulo); rt.bold=True; rt.font.size=Pt(22)
    d.add_paragraph()
    # buyer/seller side by side (one 1x2 table, each cell holds a sub-table)
    outer=d.add_table(rows=1,cols=2); outer.autofit=True
    def party(cell, title, info):
        tb=cell.add_table(rows=6,cols=2)
        tb.style='Table Grid'
        rows=[("The "+title+":",info["name"]),("Add:",info["add"]),("Contact:",info["contact"]),
              ("Tel:",info["tel"]),("Cif:",info["cif"]),("Email:",info["email"])]
        for i,(k,v) in enumerate(rows):
            set_cell(tb.rows[i].cells[0],k,bold=False,size=9)
            set_cell(tb.rows[i].cells[1],v,bold=(i==0),size=9)
        tb.columns[0].width=Cm(2.2); tb.columns[1].width=Cm(6.3)
        cell_margins(tb, top=90, bottom=90, left=110, right=110)
        for rr in tb.rows: row_height(rr, 420)
    party(outer.rows[0].cells[0],"Buyer",data["comprador"])
    party(outer.rows[0].cells[1],"Seller",b["seller"])
    d.add_paragraph()
    # product table (misma estructura para Q-SUN y SEG)
    lines=data["lineas"]; single=(len(lines)==1)
    body=max(6,len(lines))                 # filas de items + espaciadoras
    tb=d.add_table(rows=2+body+4,cols=7); tb.style='Table Grid'

    hdr_amount="Amount\n+IVA 21%\n(EURO€)" if data["marca"]=="SEG" else "Amount\n(EURO€)"
    # fila superior de cabecera
    set_cell(tb.rows[0].cells[0],"Item",bold=True,size=9,align="center")
    set_cell(tb.rows[0].cells[1],"Peak Power (Wps/PC)",bold=True,size=9,align="center")
    set_cell(tb.rows[0].cells[2],"Quantity",bold=True,size=9,align="center")
    set_cell(tb.rows[0].cells[4],"Unit price (EURO/W)",bold=True,size=9,align="center")
    set_cell(tb.rows[0].cells[5],"Incoterm",bold=True,size=9,align="center")
    set_cell(tb.rows[0].cells[6],hdr_amount,bold=True,size=9,align="center")
    # fila inferior de cabecera (solo PCS / Wps bajo Quantity)
    set_cell(tb.rows[1].cells[2],"PCS",bold=True,size=9,align="center")
    set_cell(tb.rows[1].cells[3],"Wps",bold=True,size=9,align="center")
    # fusiones: cabeceras que ocupan las dos filas (una sola vez)
    for c in (0,1,4,5,6):
        tb.cell(0,c).merge(tb.cell(1,c))
    # "Quantity" ocupa PCS+Wps en la fila superior
    tb.cell(0,2).merge(tb.cell(0,3))
    cell_margins(tb, top=80, bottom=80, left=90, right=90)
    for prow in tb.rows: row_height(prow, 420)
    total_amount=0
    for i,ln in enumerate(lines):
        row=tb.rows[2+i].cells
        precio=round(float(ln["precio"]),4)     # mismo valor para mostrar y calcular
        wps=ln["pcs"]*ln["peak"]
        amount=wps*precio
        total_amount+=amount
        set_cell(row[0],ln["item"],align="center",size=9)
        set_cell(row[1],fmt(ln["peak"],nf),align="center",size=9)
        set_cell(row[2],fmt(ln["pcs"],nf),align="center",size=9)
        set_cell(row[3],fmt(wps,nf),align="center",size=9)
        set_cell(row[4],fmt_price(precio,nf)+" €",align="center",size=9)
        set_cell(row[5],ln["incoterm"],align="center",size=9)
        set_cell(row[6],fmt(amount,nf,money=True)+" €",align="center",size=9)
    # bloque de totales: SIEMPRE presente (mismos apartados en ambas marcas)
    t0=2+body
    ivalabel="21% IVA" if data["marca"]=="SEG" else "IVA 21%"
    if single:
        iva=total_amount*0.21; total=total_amount+iva; d10=total*0.10; d90=total*0.90
        vals=[(ivalabel,iva),("Total Amount",total),
              ("10% Down Payment",d10),("90% Against BL",d90)]
    else:
        # varios incoterms: apartados presentes pero sin valor
        vals=[(ivalabel,None),("Total Amount",None),
              ("10% Down Payment",None),("90% Against BL",None)]
    for k,(label,val) in enumerate(vals):
        r=tb.rows[t0+k].cells
        set_cell(r[0],label,size=9)
        if val is not None:
            set_cell(r[-1],fmt(val,nf,money=True)+" €",size=9)
        else:
            set_cell(r[-1],"-€",size=9)   # varios incoterms
    d.add_page_break()
    # payment terms
    def h(txt):
        pp=d.add_paragraph(); rr=pp.add_run(txt); rr.bold=True; rr.font.size=Pt(10)
    def line(txt,size=9):
        pp=d.add_paragraph(); rr=pp.add_run(txt); rr.font.size=Pt(size); pp.paragraph_format.space_after=Pt(0)
    def spaced(txt,size=10,bold=False,after=6):
        pp=d.add_paragraph(); rr=pp.add_run(txt); rr.font.size=Pt(size); rr.bold=bold
        pp.paragraph_format.space_after=Pt(after); pp.paragraph_format.space_before=Pt(0)
        return pp
    spaced("Payment terms:",size=12,bold=True,after=14)
    spaced("1.1   The Buyer shall pay 10% of the total amount in advance and the remaining 90% against presentation of the Bill of Lading (B/L).",after=12)
    spaced("1.2   The Buyer shall bear all bank charges incurred in the remitting country. The Seller shall bear all bank charges incurred in the receiving country.",after=12)
    spaced("1.3   Please use the account:",after=10)
    # datos bancarios en tabla invisible para alinear etiqueta / valor
    bt=d.add_table(rows=len(b["bank"]),cols=2)
    for i,bl in enumerate(b["bank"]):
        parts=bl.split("\t")
        set_cell(bt.rows[i].cells[0], parts[0], size=10)
        set_cell(bt.rows[i].cells[1], parts[1] if len(parts)>1 else "", size=10)
    bt.columns[0].width=Cm(3.4); bt.columns[1].width=Cm(11)
    cell_margins(bt, top=10, bottom=10, left=0, right=0)
    d.add_paragraph(); d.add_paragraph()
    spaced("Shipment Terms:",size=12,bold=True,after=14)
    spaced("2.1   Delivery Arrangement: Incoterms",after=10)
    # incoterms reales de esta oferta (una linea por incoterm)
    for ln in lines:
        spaced(ln["incoterm"], after=6)
    d.add_paragraph()
    spaced("The estimated delivery date will be confirmed in writing for each order, in accordance with the agreed Incoterm.",after=10)
    for extra in b.get("shipment_tail", []):
        spaced(extra, after=10)
    d.add_paragraph()
    spaced("- Validity of this offer is 10 days.",after=24)
    spaced("Acceptance by Buyer & Date (Signature):",after=6)
    if b.get("footer_img"):
        pf=d.add_paragraph(); pf.alignment=WD_ALIGN_PARAGRAPH.CENTER
        try: pf.add_run().add_picture(os.path.join(ASSETS,b["footer_img"]), width=Cm(17))
        except Exception: pass
    elif b["footer"]:
        pf=d.add_paragraph(); pf.alignment=WD_ALIGN_PARAGRAPH.RIGHT
        rr=pf.add_run(b["footer"]); rr.bold=True; rr.font.size=Pt(12)
    d.save(out_path)

if __name__=="__main__":
    data=json.load(open(sys.argv[1],encoding="utf-8"))
    generate(data, sys.argv[2])
    print("OK ->", sys.argv[2])
