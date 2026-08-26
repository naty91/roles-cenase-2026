
import io
import re
import unicodedata
from datetime import datetime

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(page_title="CENASE | Roles vs IESS", page_icon="📊", layout="wide")

st.markdown("""
<style>
.stApp {background:#f5f7fb}
.block-container{padding-top:1.25rem;max-width:1550px}
.hero{background:linear-gradient(135deg,#073b6f,#0d5fa6);padding:22px 26px;border-radius:18px;color:white;margin-bottom:18px;box-shadow:0 8px 24px rgba(13,95,166,.15)}
.hero h1{color:white;margin:0;font-size:2rem}.hero p{margin:5px 0 0;color:#e9f4ff}
div[data-testid="stMetric"]{background:white;border:1px solid #e5eaf0;padding:13px 15px;border-radius:14px;box-shadow:0 4px 14px rgba(0,0,0,.04)}
.ok{background:#ecfdf3;color:#166534;border:1px solid #bbf7d0;border-radius:12px;padding:10px 13px;margin:8px 0}
.warn{background:#fff7ed;color:#9a3412;border:1px solid #fed7aa;border-radius:12px;padding:10px 13px;margin:8px 0}
.bad{background:#fef2f2;color:#991b1b;border:1px solid #fecaca;border-radius:12px;padding:10px 13px;margin:8px 0}
.small{font-size:.87rem;color:#64748b}
</style>
<div class="hero">
<h1>Reporte Consolidado de Roles + Conciliación IESS</h1>
<p>Gerentes · Administración · Operativos · IESS | Consulta, diferencias, cuadre y descarga mensual</p>
</div>
""", unsafe_allow_html=True)

def norm(x):
    if x is None:
        return ""
    s = str(x).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    return re.sub(r"\s+"," ",s).upper()

def as_num(s):
    return pd.to_numeric(s, errors="coerce").fillna(0.0)

def excel_date(v):
    if pd.isna(v) or v == "":
        return pd.NaT
    if isinstance(v,(pd.Timestamp,datetime)):
        return pd.to_datetime(v)
    try:
        n=float(v)
        if 20000<n<80000:
            return pd.Timestamp("1899-12-30")+pd.to_timedelta(n,unit="D")
    except Exception:
        pass
    return pd.to_datetime(v,errors="coerce")

CANONICAL = [
"Tipo Rol","Mes","Cédula","Nombre","Fecha Ingreso","Cargo","Puesto / Cliente","Días Laborados",
"Base","Sueldo","Horas Suplementarias 50%","Horas Extraordinarias 100%","Recargo 25%",
"Décimo Tercero","Décimo Cuarto","Fondo Reserva","Movilización","Otros Ingresos","Total Ingresos",
"Préstamo Quirografario","Préstamo Hipotecario","Anticipos","Faltas / Pérdida Remuneración",
"Otros Egresos","IESS","Multa","Impuesto Renta","Total Egresos","Neto a Recibir","Observaciones","Email"
]

ALIASES = {
"MAIL":"Email","EMAIL":"Email","MES":"Mes","C.I":"Cédula","CI":"Cédula","CEDULA":"Cédula",
"NOMBRE":"Nombre","F.INGRESO":"Fecha Ingreso","F INGRESO":"Fecha Ingreso","CARGO":"Cargo",
"PUESTO":"Puesto / Cliente","D.LAB":"Días Laborados","DIAS":"Días Laborados","DIAS LABORADOS":"Días Laborados",
"BASE":"Base","SUELDO":"Sueldo","HORAS EXTRAS":"Horas Extraordinarias 100%","H.S. 50%":"Horas Suplementarias 50%",
"HS 50%":"Horas Suplementarias 50%","H.E. 100":"Horas Extraordinarias 100%","HE 100":"Horas Extraordinarias 100%",
"RECARGO 25%":"Recargo 25%","DECIMO TERCER":"Décimo Tercero","13 AVO":"Décimo Tercero",
"DECIMO CUARTO":"Décimo Cuarto","14 AVO":"Décimo Cuarto","FONDO RESER":"Fondo Reserva","F.R.":"Fondo Reserva",
"MOVILIZACION":"Movilización","OTROS ING.":"Otros Ingresos","OTROS ING":"Otros Ingresos",
"TOTAL INGRESOS":"Total Ingresos","T. ING":"Total Ingresos","PTMO-QUIROG":"Préstamo Quirografario",
"P. QUIR":"Préstamo Quirografario","PTMO HIPOT":"Préstamo Hipotecario","ANTICIPO":"Anticipos","ANTICIPOS":"Anticipos",
"FALTAS":"Faltas / Pérdida Remuneración","OTROS EGRESOS":"Otros Egresos","OTROS EGR":"Otros Egresos",
"IESS":"IESS","MULTA":"Multa","I.R":"Impuesto Renta","TOTAL EGRESOS":"Total Egresos","T. DESC":"Total Egresos",
"NETO A RECIBIR":"Neto a Recibir","NETO":"Neto a Recibir","OBSERVACIONES / PENDIENTES":"Observaciones"
}

NUMERIC = [c for c in CANONICAL if c not in ["Tipo Rol","Mes","Cédula","Nombre","Fecha Ingreso","Cargo","Puesto / Cliente","Observaciones","Email"]]

def find_role_sheet(f):
    xls=pd.ExcelFile(f)
    for s in xls.sheet_names:
        if norm(s)=="LISTA": return s
    for s in xls.sheet_names:
        if "LISTA" in norm(s): return s
    return xls.sheet_names[0]

def role_header(raw):
    for i in range(min(15,len(raw))):
        vals=[norm(v) for v in raw.iloc[i].tolist()]
        if "NOMBRE" in vals and any(v in ("C.I","CI","CEDULA") for v in vals) and any(v in ("NETO","NETO A RECIBIR") for v in vals):
            return i
    raise ValueError("No se encontró el encabezado del rol.")

def read_role(f,tipo):
    sheet=find_role_sheet(f)
    raw=pd.read_excel(f,sheet_name=sheet,header=None,dtype=object)
    h=role_header(raw)
    headers=[norm(x) for x in raw.iloc[h].tolist()]
    d=raw.iloc[h+1:].copy()
    d.columns=headers
    d=d.rename(columns={c:ALIASES[norm(c)] for c in d.columns if norm(c) in ALIASES})

    # collapse duplicated columns
    out=pd.DataFrame(index=d.index)
    for c in list(dict.fromkeys(d.columns)):
        sub=d.loc[:,d.columns==c]
        out[c]=sub.bfill(axis=1).iloc[:,0] if sub.shape[1]>1 else sub.iloc[:,0]
    d=out

    d["Cédula"]=d["Cédula"].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
    d["Nombre"]=d["Nombre"].fillna("").astype(str).str.strip()
    d=d[d["Cédula"].str.fullmatch(r"\d{8,13}",na=False)&d["Nombre"].ne("")].copy()
    d["Tipo Rol"]=tipo
    for c in CANONICAL:
        if c not in d.columns: d[c]=np.nan
    d["Fecha Ingreso"]=d["Fecha Ingreso"].apply(excel_date)
    mes=d["Mes"].apply(excel_date)
    if mes.notna().any(): d["Mes"]=mes.dt.strftime("%Y-%m")
    for c in NUMERIC: d[c]=as_num(d[c])
    if tipo=="Operativos":
        d["Cargo"]=d["Cargo"].replace("",np.nan).fillna("Guardia")
    d["Puesto / Cliente"]=d["Puesto / Cliente"].fillna("")
    return d[CANONICAL].reset_index(drop=True)

IESS_CANON = ["Periodo","Cédula","Nombre IESS","Rel. Trabajo","Sueldo IESS","Días IESS",
              "Patronal IESS","Individual IESS","Aporte Adic","Cesantía","% CCC","Valor CCC","Total Aporte IESS"]

def read_iess(f):
    # Supports .xls and .xlsx. xlrd is included in requirements for .xls.
    xls=pd.ExcelFile(f)
    candidates=[]
    for sheet in xls.sheet_names:
        raw=pd.read_excel(f,sheet_name=sheet,header=None,dtype=object)
        for i in range(min(20,len(raw))):
            vals=[norm(v) for v in raw.iloc[i].tolist()]
            if "CEDULA" in vals and "NOMBRE" in vals and "SUELDO" in vals and "DIAS" in vals:
                candidates.append((sheet,i,raw))
                break
    if not candidates:
        raise ValueError("No pude identificar las columnas Cédula, Nombre, Sueldo y Días en el consolidado IESS.")

    sheet,h,raw=candidates[0]
    orig=[norm(x) for x in raw.iloc[h].tolist()]
    d=raw.iloc[h+1:].copy()
    d.columns=orig

    def pick(keys, required=False):
        for col in d.columns:
            n=norm(col)
            if any(k in n for k in keys):
                return col
        if required:
            raise ValueError("Falta una columna necesaria en el consolidado IESS: "+"/".join(keys))
        return None

    mapping={}
    mapping[pick(["PERIODO"],False)]="Periodo" if pick(["PERIODO"],False) else None
    mapping[pick(["CEDULA"],True)]="Cédula"
    mapping[pick(["NOMBRE"],True)]="Nombre IESS"
    rel=pick(["REL. TRABAJO","REL TRABAJO","RELACION"],False)
    if rel: mapping[rel]="Rel. Trabajo"
    mapping[pick(["SUELDO"],True)]="Sueldo IESS"
    mapping[pick(["DIAS"],True)]="Días IESS"

    patronal=None
    individual=None
    for col in d.columns:
        n=norm(col)
        if "PATRONAL" in n or "11.15" in n: patronal=col
        if "INDIVIDUAL" in n or "9.45" in n: individual=col
    if patronal: mapping[patronal]="Patronal IESS"
    if individual: mapping[individual]="Individual IESS"

    ad=pick(["APORTE ADIC"],False)
    ces=pick(["CESANTIA"],False)
    pct=pick(["% CCC","CCC %"],False)
    val=pick(["VALOR CCC"],False)
    total=pick(["TOTAL APORTE"],False)
    for c,new in [(ad,"Aporte Adic"),(ces,"Cesantía"),(pct,"% CCC"),(val,"Valor CCC"),(total,"Total Aporte IESS")]:
        if c: mapping[c]=new

    d=d.rename(columns={k:v for k,v in mapping.items() if k is not None and v is not None})
    d["Cédula"]=d["Cédula"].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
    d["Nombre IESS"]=d["Nombre IESS"].fillna("").astype(str).str.strip()
    d=d[d["Cédula"].str.fullmatch(r"\d{8,13}",na=False)&d["Nombre IESS"].ne("")].copy()
    for c in IESS_CANON:
        if c not in d.columns: d[c]=0.0 if c in ["Sueldo IESS","Días IESS","Patronal IESS","Individual IESS","Aporte Adic","Cesantía","% CCC","Valor CCC","Total Aporte IESS"] else ""
    for c in ["Sueldo IESS","Días IESS","Patronal IESS","Individual IESS","Aporte Adic","Cesantía","% CCC","Valor CCC","Total Aporte IESS"]:
        d[c]=as_num(d[c])
    return d[IESS_CANON].reset_index(drop=True)

def make_compare(roles,iess):
    # Consolidate roles by employee in case an ID appears more than once.
    r=roles.groupby("Cédula",as_index=False).agg({
        "Nombre":"first","Tipo Rol":"first","Cargo":"first","Puesto / Cliente":"first",
        "Días Laborados":"sum","Base":"sum","Sueldo":"sum","IESS":"sum",
        "Total Ingresos":"sum","Total Egresos":"sum","Neto a Recibir":"sum"
    })
    # Prefer Base when populated; otherwise Sueldo.
    r["Base Rol IESS"]=np.where(r["Base"].abs()>0.001,r["Base"],r["Sueldo"])

    i=iess.groupby("Cédula",as_index=False).agg({
        "Nombre IESS":"first","Rel. Trabajo":"first","Sueldo IESS":"sum","Días IESS":"sum",
        "Patronal IESS":"sum","Individual IESS":"sum","Aporte Adic":"sum","Cesantía":"sum",
        "Valor CCC":"sum","Total Aporte IESS":"sum"
    })
    c=r.merge(i,on="Cédula",how="outer",indicator=True)
    c["Nombre"]=c["Nombre"].fillna(c["Nombre IESS"])
    c["Tipo Rol"]=c["Tipo Rol"].fillna("SOLO IESS")
    c["Cargo"]=c["Cargo"].fillna("")
    c["Puesto / Cliente"]=c["Puesto / Cliente"].fillna("")
    numeric=["Días Laborados","Base Rol IESS","IESS","Sueldo IESS","Días IESS","Patronal IESS","Individual IESS",
             "Aporte Adic","Cesantía","Valor CCC","Total Aporte IESS"]
    for col in numeric: c[col]=as_num(c[col])

    c["Diferencia Días"]=c["Días Laborados"]-c["Días IESS"]
    c["Diferencia Base"]=c["Base Rol IESS"]-c["Sueldo IESS"]

    # Employee contribution: Role deduction vs IESS individual.
    c["% Individual Rol"]=np.where(c["Base Rol IESS"].abs()>0.001,c["IESS"]/c["Base Rol IESS"]*100,0)
    c["% Individual IESS"]=np.where(c["Sueldo IESS"].abs()>0.001,c["Individual IESS"]/c["Sueldo IESS"]*100,0)
    c["Diferencia Aporte Individual"]=c["IESS"]-c["Individual IESS"]

    # Employer contribution: expected from role base vs IESS patronal.
    c["Patronal Esperado Rol 11.15%"]=c["Base Rol IESS"]*0.1115
    c["% Patronal IESS"]=np.where(c["Sueldo IESS"].abs()>0.001,c["Patronal IESS"]/c["Sueldo IESS"]*100,0)
    c["Diferencia Aporte Patronal"]=c["Patronal Esperado Rol 11.15%"]-c["Patronal IESS"]

    # Percentage deviations from statutory values.
    c["Dif. % Individual vs 9.45"]=c["% Individual IESS"]-9.45
    c["Dif. % Patronal vs 11.15"]=c["% Patronal IESS"]-11.15

    c["Presencia"]=c["_merge"].map({"both":"ROL + IESS","left_only":"SOLO ROL","right_only":"SOLO IESS"}).astype(str)
    tol_money=0.05
    tol_pct=0.02
    c["Estado"]=np.where(
        (c["Presencia"]=="ROL + IESS") &
        (c["Diferencia Días"].abs()<=0.01) &
        (c["Diferencia Base"].abs()<=tol_money) &
        (c["Diferencia Aporte Individual"].abs()<=tol_money) &
        (c["Diferencia Aporte Patronal"].abs()<=tol_money) &
        (c["Dif. % Individual vs 9.45"].abs()<=tol_pct) &
        (c["Dif. % Patronal vs 11.15"].abs()<=tol_pct),
        "CUADRA","REVISAR"
    )
    return c

def fmt_money(v): return f"${v:,.2f}"

def export_excel(roles,summary,compare,diffs):
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="xlsxwriter",datetime_format="dd/mm/yyyy") as writer:
        wb=writer.book
        title=wb.add_format({"bold":True,"font_size":16,"font_color":"#FFFFFF","bg_color":"#0B4F88","align":"center"})
        head=wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#0D5FA6","border":1,"align":"center","text_wrap":True})
        money=wb.add_format({"num_format":"$#,##0.00","border":1})
        pct=wb.add_format({"num_format":"0.00%","border":1})
        num=wb.add_format({"num_format":"0.00","border":1})
        text=wb.add_format({"border":1})
        warn=wb.add_format({"bg_color":"#FDECEC","font_color":"#991B1B","border":1})
        ok=wb.add_format({"bg_color":"#ECFDF3","font_color":"#166534","border":1})

        summary.to_excel(writer,sheet_name="Resumen",index=False,startrow=2)
        ws=writer.sheets["Resumen"]; ws.merge_range(0,0,0,max(5,len(summary.columns)-1),"RESUMEN CONSOLIDADO DE ROLES",title)
        for j,c in enumerate(summary.columns): ws.write(2,j,c,head)
        ws.set_column(0,0,24); ws.set_column(1,1,12); ws.set_column(2,5,18,money)

        for sheet,df in [("Consolidado",roles),
                         ("Gerentes",roles[roles["Tipo Rol"]=="Gerentes"]),
                         ("Administrativos",roles[roles["Tipo Rol"]=="Administrativos"]),
                         ("Operativos",roles[roles["Tipo Rol"]=="Operativos"])]:
            df.to_excel(writer,sheet_name=sheet,index=False,startrow=2)
            ws=writer.sheets[sheet]
            ws.merge_range(0,0,0,max(1,len(df.columns)-1),f"REPORTE DE ROLES - {sheet.upper()}",title)
            for j,c in enumerate(df.columns): ws.write(2,j,c,head)
            ws.freeze_panes(3,0); ws.autofilter(2,0,2+len(df),len(df.columns)-1)
            ws.set_column(0,0,18); ws.set_column(1,1,10); ws.set_column(2,2,14); ws.set_column(3,3,36)
            ws.set_column(4,6,24); ws.set_column(7,28,16,money); ws.set_column(29,30,28,text)

        for sheet,df in [("Rol vs IESS",compare),("Diferencias",diffs)]:
            export=df.drop(columns=["_merge"],errors="ignore").copy()
            export.to_excel(writer,sheet_name=sheet,index=False,startrow=2)
            ws=writer.sheets[sheet]
            ws.merge_range(0,0,0,max(1,len(export.columns)-1),sheet.upper(),title)
            for j,c in enumerate(export.columns): ws.write(2,j,c,head)
            ws.freeze_panes(3,0); ws.autofilter(2,0,2+len(export),len(export.columns)-1)
            ws.set_column(0,0,15); ws.set_column(1,1,35); ws.set_column(2,4,22)
            ws.set_column(5,len(export.columns)-1,18)
            # conditional formatting on Estado
            if "Estado" in export.columns and len(export):
                ec=export.columns.get_loc("Estado")
                ws.conditional_format(3,ec,2+len(export),ec,{"type":"text","criteria":"containing","value":"REVISAR","format":warn})
                ws.conditional_format(3,ec,2+len(export),ec,{"type":"text","criteria":"containing","value":"CUADRA","format":ok})
    out.seek(0)
    return out.getvalue()

# ---------- UPLOADS ----------
with st.sidebar:
    st.header("Carga mensual")
    st.caption("Sube los 4 archivos del mismo período.")
    fger=st.file_uploader("1. Rol de Gerentes",type=["xlsx","xls"],key="ger")
    fad=st.file_uploader("2. Rol de Administración",type=["xlsx","xls"],key="adm")
    fop=st.file_uploader("3. Rol de Operativos",type=["xlsx","xls"],key="ope")
    fiess=st.file_uploader("4. Consolidado IESS",type=["xlsx","xls"],key="iess")
    st.divider()
    st.caption("El cruce Rol vs IESS se realiza por cédula.")

if not (fger and fad and fop and fiess):
    st.info("Carga los cuatro archivos para generar el reporte y la conciliación Rol vs IESS.")
    st.stop()

try:
    ger=read_role(fger,"Gerentes")
    adm=read_role(fad,"Administrativos")
    ope=read_role(fop,"Operativos")
    roles=pd.concat([ger,adm,ope],ignore_index=True)
    iess=read_iess(fiess)
    comp=make_compare(roles,iess)
except Exception as e:
    st.error(f"No pude procesar los archivos: {e}")
    st.stop()

# ---------- SUMMARY ----------
grp=roles.groupby("Tipo Rol",as_index=False).agg(
    Empleados=("Cédula","count"),Total_Ingresos=("Total Ingresos","sum"),
    Total_Egresos=("Total Egresos","sum"),Neto=("Neto a Recibir","sum")
)
total=pd.DataFrame([{"Tipo Rol":"TOTAL GENERAL","Empleados":len(roles),
                     "Total_Ingresos":roles["Total Ingresos"].sum(),
                     "Total_Egresos":roles["Total Egresos"].sum(),
                     "Neto":roles["Neto a Recibir"].sum()}])
summary=pd.concat([grp,total],ignore_index=True)
summary.columns=["Tipo de Rol","Empleados","Total Ingresos","Total Egresos","Neto a Recibir"]

diffs=comp[comp["Estado"]=="REVISAR"].copy()
missing_role=(comp["Presencia"]=="SOLO IESS").sum()
missing_iess=(comp["Presencia"]=="SOLO ROL").sum()

a,b,c,d,e,f=st.columns(6)
a.metric("Trabajadores Rol",roles["Cédula"].nunique())
b.metric("Trabajadores IESS",iess["Cédula"].nunique())
c.metric("Diferencias",len(diffs))
d.metric("Solo en Rol",int(missing_iess))
e.metric("Solo en IESS",int(missing_role))
f.metric("Neto Roles",fmt_money(roles["Neto a Recibir"].sum()))

tabs=st.tabs(["📊 Resumen Roles","🏛️ Rol vs IESS","⚠️ Diferencias","🔎 Consulta Roles","✅ Cuadre"])

with tabs[0]:
    show=summary.copy()
    for col in ["Total Ingresos","Total Egresos","Neto a Recibir"]: show[col]=show[col].map(fmt_money)
    st.dataframe(show,use_container_width=True,hide_index=True)
    st.subheader("Neto por tipo de rol")
    st.bar_chart(grp.set_index("Tipo Rol")[["Neto"]],use_container_width=True)

with tabs[1]:
    st.subheader("Conciliación individual Rol vs IESS")
    st.caption("Base Rol IESS usa la columna BASE del rol; si está vacía usa SUELDO. Aporte individual Rol corresponde al descuento IESS del rol.")

    c1,c2,c3,c4=st.columns(4)
    estado=c1.multiselect("Estado",["CUADRA","REVISAR"],default=["CUADRA","REVISAR"])
    pres=c2.multiselect("Presencia",["ROL + IESS","SOLO ROL","SOLO IESS"],default=["ROL + IESS","SOLO ROL","SOLO IESS"])
    rolopt=sorted(comp["Tipo Rol"].dropna().astype(str).unique())
    rolfil=c3.multiselect("Tipo de rol",rolopt,default=rolopt)
    q=c4.text_input("Nombre o cédula")

    view=comp[comp["Estado"].isin(estado)&comp["Presencia"].isin(pres)&comp["Tipo Rol"].isin(rolfil)].copy()
    if q.strip():
        nq=norm(q)
        view=view[view["Nombre"].map(norm).str.contains(nq,na=False)|view["Cédula"].astype(str).str.contains(q.strip(),na=False)]

    money_cols=["Base Rol IESS","Sueldo IESS","IESS","Individual IESS","Diferencia Aporte Individual",
                "Patronal Esperado Rol 11.15%","Patronal IESS","Diferencia Aporte Patronal","Diferencia Base"]
    pct_cols=["% Individual Rol","% Individual IESS","% Patronal IESS","Dif. % Individual vs 9.45","Dif. % Patronal vs 11.15"]
    display_cols=["Estado","Presencia","Tipo Rol","Cédula","Nombre","Cargo","Puesto / Cliente",
                  "Días Laborados","Días IESS","Diferencia Días",
                  "Base Rol IESS","Sueldo IESS","Diferencia Base",
                  "IESS","Individual IESS","Diferencia Aporte Individual",
                  "% Individual Rol","% Individual IESS","Dif. % Individual vs 9.45",
                  "Patronal Esperado Rol 11.15%","Patronal IESS","Diferencia Aporte Patronal",
                  "% Patronal IESS","Dif. % Patronal vs 11.15","Rel. Trabajo"]
    st.dataframe(view[display_cols],use_container_width=True,hide_index=True,
        column_config={
            **{x:st.column_config.NumberColumn(x,format="$ %.2f") for x in money_cols},
            **{x:st.column_config.NumberColumn(x,format="%.2f %%") for x in pct_cols},
            "Días Laborados":st.column_config.NumberColumn(format="%.0f"),
            "Días IESS":st.column_config.NumberColumn(format="%.0f"),
            "Diferencia Días":st.column_config.NumberColumn(format="%.0f"),
        })

with tabs[2]:
    st.subheader("Solo registros que requieren revisión")
    x1,x2,x3,x4=st.columns(4)
    x1.metric("Registros a revisar",len(diffs))
    x2.metric("Dif. base total",fmt_money(diffs["Diferencia Base"].sum()))
    x3.metric("Dif. aporte individual",fmt_money(diffs["Diferencia Aporte Individual"].sum()))
    x4.metric("Dif. aporte patronal",fmt_money(diffs["Diferencia Aporte Patronal"].sum()))

    if diffs.empty:
        st.success("Todo cuadra entre Roles e IESS.")
    else:
        st.dataframe(diffs[["Presencia","Tipo Rol","Cédula","Nombre","Días Laborados","Días IESS","Diferencia Días",
                            "Base Rol IESS","Sueldo IESS","Diferencia Base","IESS","Individual IESS",
                            "Diferencia Aporte Individual","Patronal Esperado Rol 11.15%","Patronal IESS",
                            "Diferencia Aporte Patronal","% Individual IESS","% Patronal IESS"]],
                     use_container_width=True,hide_index=True)

with tabs[3]:
    st.subheader("Consulta detallada de Roles")
    r1,r2,r3,r4=st.columns(4)
    rols=sorted(roles["Tipo Rol"].unique())
    fr=r1.multiselect("Tipo de rol",rols,default=rols,key="fr")
    fq=r2.text_input("Nombre / cédula",key="fq")
    cargos=sorted([x for x in roles["Cargo"].dropna().astype(str).unique() if x.strip()])
    fc=r3.multiselect("Cargo",cargos,key="fc")
    puestos=sorted([x for x in roles["Puesto / Cliente"].dropna().astype(str).unique() if x.strip()])
    fp=r4.multiselect("Puesto / cliente",puestos,key="fp")

    rr=roles[roles["Tipo Rol"].isin(fr)].copy()
    if fq.strip():
        nq=norm(fq); rr=rr[rr["Nombre"].map(norm).str.contains(nq,na=False)|rr["Cédula"].astype(str).str.contains(fq.strip(),na=False)]
    if fc: rr=rr[rr["Cargo"].isin(fc)]
    if fp: rr=rr[rr["Puesto / Cliente"].isin(fp)]
    st.dataframe(rr,use_container_width=True,hide_index=True)

with tabs[4]:
    st.subheader("Cuadre general")
    q1,q2,q3,q4=st.columns(4)
    q1.metric("Base Rol",fmt_money(comp[comp["Presencia"]!="SOLO IESS"]["Base Rol IESS"].sum()))
    q2.metric("Base IESS",fmt_money(iess["Sueldo IESS"].sum()))
    q3.metric("Aporte individual Rol",fmt_money(roles["IESS"].sum()))
    q4.metric("Aporte individual IESS",fmt_money(iess["Individual IESS"].sum()))

    st.markdown("#### Totales IESS")
    t1,t2,t3,t4=st.columns(4)
    t1.metric("Patronal IESS 11,15%",fmt_money(iess["Patronal IESS"].sum()))
    t2.metric("Individual IESS 9,45%",fmt_money(iess["Individual IESS"].sum()))
    t3.metric("Valor CCC",fmt_money(iess["Valor CCC"].sum()))
    t4.metric("Total aporte",fmt_money(iess["Total Aporte IESS"].sum()))

# ---------- DOWNLOAD ----------
st.divider()
excel=export_excel(roles,summary,comp,diffs)
mes=roles["Mes"].dropna().astype(str)
mes=mes.iloc[0] if len(mes) else datetime.now().strftime("%Y-%m")
st.download_button("⬇️ Descargar reporte completo Roles + IESS",
                   data=excel,file_name=f"Roles_vs_IESS_{mes}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)
st.markdown('<p class="small">El archivo descargado incluye: Resumen, Consolidado, Gerentes, Administrativos, Operativos, Rol vs IESS y Diferencias.</p>',unsafe_allow_html=True)
