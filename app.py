
import io
import re
from pathlib import Path

import cv2
import numpy as np
import pandas as pd
import pytesseract
import streamlit as st
from openpyxl import load_workbook
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="Presupuestos OCR", page_icon="📄", layout="wide")

APP_NAME = "Presupuestos OCR"
APP_VERSION = "2.0"

st.markdown("""
<style>
.block-container {max-width: 1500px; padding-top: 2rem;}
.hero-title {font-size: 2.35rem; font-weight: 800;}
.hero-subtitle {color:#6b7280; margin-bottom:1.2rem;}
div[data-testid="stMetric"] {border:1px solid #e5e7eb; padding:.8rem 1rem; border-radius:12px;}
</style>
""", unsafe_allow_html=True)

def clean_text(value):
    value = str(value or "").replace("\x0c", " ")
    return re.sub(r"[ \t]+", " ", value).strip()

def preprocess(img_bgr):
    h, w = img_bgr.shape[:2]
    if w < 1800:
        scale = 1800 / w
        img_bgr = cv2.resize(img_bgr, None, fx=scale, fy=scale, interpolation=cv2.INTER_CUBIC)
    gray = cv2.cvtColor(img_bgr, cv2.COLOR_BGR2GRAY)
    gray = cv2.GaussianBlur(gray, (3,3), 0)
    bw = cv2.adaptiveThreshold(gray,255,cv2.ADAPTIVE_THRESH_GAUSSIAN_C,cv2.THRESH_BINARY,31,11)
    return img_bgr, gray, bw

def ocr_image(img, psm=6):
    for lang in ("spa+eng","spa","eng"):
        try:
            txt = pytesseract.image_to_string(img, lang=lang, config=f"--psm {psm}")
            if txt.strip():
                return txt
        except Exception:
            pass
    try:
        return pytesseract.image_to_string(img, config=f"--psm {psm}")
    except Exception:
        return ""

def cluster(values, tolerance=10):
    values = sorted(int(v) for v in values)
    if not values:
        return []
    groups = [[values[0]]]
    for v in values[1:]:
        if v - groups[-1][-1] <= tolerance:
            groups[-1].append(v)
        else:
            groups.append([v])
    return [int(round(sum(g)/len(g))) for g in groups]

def detect_table(binary):
    inv = 255 - binary
    hk = cv2.getStructuringElement(cv2.MORPH_RECT,(max(30,binary.shape[1]//25),1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT,(1,max(30,binary.shape[0]//25)))
    hlines = cv2.morphologyEx(inv,cv2.MORPH_OPEN,hk)
    vlines = cv2.morphologyEx(inv,cv2.MORPH_OPEN,vk)
    grid = cv2.add(hlines,vlines)
    contours,_ = cv2.findContours(grid,cv2.RETR_EXTERNAL,cv2.CHAIN_APPROX_SIMPLE)
    H,W = binary.shape
    candidates=[]
    for c in contours:
        x,y,w,h = cv2.boundingRect(c)
        if w > .42*W and h > .055*H and w*h > .035*W*H:
            candidates.append((x,y,w,h))
    return max(candidates,key=lambda r:r[2]*r[3]) if candidates else None

def extract_grid(img_bgr, rect, binary):
    x,y,w,h = rect
    roi = binary[y:y+h,x:x+w]
    inv = 255-roi
    hk = cv2.getStructuringElement(cv2.MORPH_RECT,(max(20,w//20),1))
    vk = cv2.getStructuringElement(cv2.MORPH_RECT,(1,max(20,h//20)))
    hs = cv2.morphologyEx(inv,cv2.MORPH_OPEN,hk)
    vs = cv2.morphologyEx(inv,cv2.MORPH_OPEN,vk)
    hp = np.sum(hs>0,axis=1)
    vp = np.sum(vs>0,axis=0)
    ys = cluster(np.where(hp>.35*w)[0],10)
    xs = cluster(np.where(vp>.35*h)[0],10)
    xs = cluster([0]+xs+[w-1],12)
    ys = cluster([0]+ys+[h-1],12)
    xs = [v for i,v in enumerate(xs) if i==0 or v-xs[i-1]>15]
    ys = [v for i,v in enumerate(ys) if i==0 or v-ys[i-1]>12]
    if len(xs)<3 or len(ys)<3:
        return pd.DataFrame()
    rows=[]
    for r in range(len(ys)-1):
        row=[]
        for c in range(len(xs)-1):
            x1,x2=xs[c],xs[c+1]; y1,y2=ys[r],ys[r+1]
            px=max(4,int((x2-x1)*.04)); py=max(4,int((y2-y1)*.08))
            cell=img_bgr[y+y1+py:y+y2-py,x+x1+px:x+x2-px]
            if cell.size==0:
                row.append(""); continue
            g=cv2.cvtColor(cell,cv2.COLOR_BGR2GRAY)
            g=cv2.resize(g,None,fx=2,fy=2,interpolation=cv2.INTER_CUBIC)
            row.append(clean_text(ocr_image(g,6).replace("\n"," ")))
        rows.append(row)
    df=pd.DataFrame(rows)
    if df.empty:
        return df
    return df.replace(r"^\s*$",np.nan,regex=True).dropna(how="all").dropna(axis=1,how="all").fillna("").reset_index(drop=True)

PATTERNS = {
"Fecha": r"Fecha\s*:\s*([0-9]{1,2}/[0-9]{1,2}/[0-9]{4})",
"Cliente": r"Señor\(es\)\s*:\s*(.+?)(?=\s+RUC\b|\s+Direcci[oó]n\b|\n|$)",
"RUC": r"RUC\s*:\s*([0-9-]+)",
"Código de Cliente": r"C[oó]digo de Cliente\s*:\s*([A-Za-z0-9-]+)",
"Presupuesto N°": r"Presupuesto N[°ºo]?\s*:\s*([A-Za-z0-9-]+)",
"Validez": r"Validez de la oferta\s*:\s*(.+?)(?=\s+Condici[oó]n|\n|$)",
"SUB TOTAL": r"SUB TOTAL\s*:\s*([0-9.,]+)",
"IVA 10%": r"I\.V\.A\.\s*10%\s*:\s*([0-9.,]+)",
"Condición de venta": r"Condici[oó]n Venta\s*:\s*(.+?)(?=\s+Plazo de Entrega|\n|$)",
"Plazo de Entrega": r"Plazo de Entrega\s*:\s*(.+?)(?=\s+Entrega\s*:|\n|$)",
"Entrega": r"Entrega\s*:\s*(.+?)(?=\s+Modo de Pago|\n|$)",
"Modo de Pago": r"Modo de Pago\s*:\s*(.+?)(?=\s+Observaci[oó]n|\n|$)",
"Vendedor": r"Vendedor\s*:\s*(.+?)(?=\s+Correo\s*:|\n|$)",
"Correo": r"Correo\s*:\s*([^\s]+)",
"Teléfono": r"Tel[eé]fono\s*:\s*([0-9 ]+)",
}

def parse_fields(text):
    text=clean_text(text)
    out={}
    for name,pat in PATTERNS.items():
        m=re.search(pat,text,re.I|re.S)
        out[name]=clean_text(m.group(1)) if m else ""
    return out

def load_pages(data, filename):
    if Path(filename).suffix.lower()==".pdf":
        import fitz
        doc=fitz.open(stream=data,filetype="pdf")
        pages=[]
        for page in doc:
            pix=page.get_pixmap(matrix=fitz.Matrix(2,2),alpha=False)
            arr=np.frombuffer(pix.samples,dtype=np.uint8).reshape(pix.height,pix.width,3)
            pages.append(cv2.cvtColor(arr,cv2.COLOR_RGB2BGR))
        return pages
    arr=np.frombuffer(data,np.uint8)
    img=cv2.imdecode(arr,cv2.IMREAD_COLOR)
    return [img] if img is not None else []

def process_document(data, filename):
    results=[]; tables=[]
    for i,img in enumerate(load_pages(data,filename),1):
        work,gray,bw=preprocess(img)
        raw=ocr_image(gray,6)
        fields=parse_fields(raw)
        rect=detect_table(bw)
        table=extract_grid(work,rect,bw) if rect else pd.DataFrame()
        if not table.empty:
            tables.append(table)
        results.append({"page":i,"fields":fields,"table":table,"raw":raw,"preview":work})
    combined=pd.concat(tables,ignore_index=True) if tables else pd.DataFrame()
    return results,combined

def make_excel(fields, table_df, source_name):
    bio=io.BytesIO()
    with pd.ExcelWriter(bio,engine="openpyxl") as writer:
        pd.DataFrame({"Campo":list(fields.keys()),"Valor":list(fields.values())}).to_excel(writer,index=False,sheet_name="Datos")
        table_df.to_excel(writer,index=False,header=False,sheet_name="Productos")
    bio.seek(0)
    wb=load_workbook(bio)
    fill=PatternFill("solid",fgColor="1F2937")
    font=Font(color="FFFFFF",bold=True)
    thin=Side(style="thin",color="D1D5DB")
    ws=wb["Datos"]
    for cell in ws[1]:
        cell.fill=fill; cell.font=font; cell.alignment=Alignment(horizontal="center")
    ws.column_dimensions["A"].width=28; ws.column_dimensions["B"].width=65
    for row in ws.iter_rows():
        for cell in row:
            cell.border=Border(bottom=thin); cell.alignment=Alignment(vertical="top",wrap_text=True)
    wp=wb["Productos"]
    for row in wp.iter_rows():
        for cell in row:
            cell.border=Border(bottom=thin); cell.alignment=Alignment(vertical="top",wrap_text=True)
    for col in range(1,wp.max_column+1):
        vals=[len(str(wp.cell(r,col).value or "")) for r in range(1,min(wp.max_row,50)+1)]
        wp.column_dimensions[get_column_letter(col)].width=min(55,max(12,max(vals or [12])+2))
    wi=wb.create_sheet("Info")
    wi["A1"]="Archivo original"; wi["B1"]=source_name
    wi["A2"]="Aplicación"; wi["B2"]=f"{APP_NAME} v{APP_VERSION}"
    wi["A3"]="Aviso"; wi["B3"]="Revisar los datos antes de utilizarlos comercial o contablemente."
    wi.column_dimensions["A"].width=24; wi.column_dimensions["B"].width=85
    out=io.BytesIO(); wb.save(out); return out.getvalue()

if "results" not in st.session_state: st.session_state.results=None
if "fields" not in st.session_state: st.session_state.fields={}
if "table" not in st.session_state: st.session_state.table=pd.DataFrame()

with st.sidebar:
    st.markdown(f"## 📄 {APP_NAME}")
    st.caption(f"Versión {APP_VERSION}")
    st.markdown("---")
    st.write("1. Subir documento")
    st.write("2. Procesar OCR")
    st.write("3. Revisar datos")
    st.write("4. Descargar Excel")

st.markdown(f'<div class="hero-title">📄 {APP_NAME}</div>',unsafe_allow_html=True)
st.markdown('<div class="hero-subtitle">Convierte presupuestos escaneados en datos editables y Excel.</div>',unsafe_allow_html=True)

uploaded=st.file_uploader("Subí un presupuesto",type=["pdf","jpg","jpeg","png"])

if uploaded is None:
    c1,c2,c3=st.columns(3)
    c1.markdown("### Cargar"); c1.write("Subí el presupuesto escaneado.")
    c2.markdown("### Revisar"); c2.write("Corregí campos o celdas.")
    c3.markdown("### Exportar"); c3.write("Descargá el resultado en Excel.")
    st.info("Optimizada inicialmente para presupuestos con tablas visibles.")
    st.stop()

if st.button("🔎 Procesar presupuesto",type="primary",use_container_width=True):
    with st.spinner("Procesando OCR y tabla..."):
        results,combined=process_document(uploaded.getvalue(),uploaded.name)
        st.session_state.results=results
        st.session_state.fields=results[0]["fields"] if results else {}
        st.session_state.table=combined.copy()
        st.success("Documento procesado.")

if st.session_state.results is None:
    st.info("Presioná **Procesar presupuesto**.")
    st.stop()

results=st.session_state.results
m1,m2,m3,m4=st.columns(4)
m1.metric("Páginas",len(results))
m2.metric("Campos detectados",sum(bool(v) for v in st.session_state.fields.values()))
m3.metric("Filas de tabla",len(st.session_state.table))
m4.metric("Columnas",len(st.session_state.table.columns) if not st.session_state.table.empty else 0)

tab1,tab2,tab3,tab4=st.tabs(["📋 Datos","🧾 Productos","🖼️ Vista previa","🔍 OCR bruto"])

with tab1:
    df=pd.DataFrame({"Campo":list(st.session_state.fields.keys()),"Valor":list(st.session_state.fields.values())})
    edited=st.data_editor(df,use_container_width=True,hide_index=True,num_rows="fixed",
        column_config={"Campo":st.column_config.TextColumn(disabled=True),"Valor":st.column_config.TextColumn(width="large")},
        key="fields_editor")
    st.session_state.fields=dict(zip(edited["Campo"],edited["Valor"]))

with tab2:
    if st.session_state.table.empty:
        st.warning("No se detectó automáticamente una tabla.")
    else:
        st.session_state.table=st.data_editor(st.session_state.table,use_container_width=True,hide_index=True,num_rows="dynamic",key="table_editor")

with tab3:
    p=st.selectbox("Página",list(range(1,len(results)+1)),format_func=lambda x:f"Página {x}")
    st.image(cv2.cvtColor(results[p-1]["preview"],cv2.COLOR_BGR2RGB),use_container_width=True)

with tab4:
    p=st.selectbox("Página",list(range(1,len(results)+1)),format_func=lambda x:f"Página {x}",key="raw_page")
    st.text_area("Texto reconocido",results[p-1]["raw"],height=420)

excel=make_excel(st.session_state.fields,st.session_state.table,uploaded.name)
st.download_button("⬇️ Descargar Excel corregido",data=excel,
    file_name=f"{Path(uploaded.name).stem}_extraido.xlsx",
    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
    type="primary",use_container_width=True)
st.caption("El OCR puede cometer errores; revisá los datos antes de utilizarlos.")
