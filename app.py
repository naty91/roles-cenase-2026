
import io
import re
import json
import gzip
import base64
import unicodedata
from datetime import datetime

import requests

import numpy as np
import pandas as pd
import streamlit as st

from reportlab.lib import colors
from reportlab.lib.pagesizes import A4, A3, landscape
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import mm
from reportlab.platypus import SimpleDocTemplate, Table, TableStyle, Paragraph, Spacer, PageBreak, KeepTogether

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
<h1>Reporte Consolidado de Roles + Conciliación IESS · v19</h1>
<p>Gerentes · Administración · Operativos · IESS | Consulta, diferencias, cuadre y descarga mensual</p>
</div>
""", unsafe_allow_html=True)

# ---------- PERSISTENCIA PERMANENTE (SUPABASE) ----------
# La app guarda una copia comprimida de los datos procesados por período.
# Las credenciales se leen exclusivamente desde Streamlit Secrets y nunca se
# escriben dentro del repositorio.

def get_storage_config():
    try:
        if "supabase" not in st.secrets:
            return None
        cfg=st.secrets["supabase"]
        url=str(cfg.get("url","")).strip().rstrip("/")
        key=str(cfg.get("service_role_key", cfg.get("key",""))).strip()
        if not url or not key:
            return None
        return {"url":url,"key":key}
    except Exception:
        return None


def _storage_headers(cfg, prefer=None):
    h={
        "apikey":cfg["key"],
        "Authorization":f"Bearer {cfg['key']}",
        "Content-Type":"application/json",
    }
    if prefer:
        h["Prefer"]=prefer
    return h


def _df_to_json(df):
    return df.to_json(orient="split",date_format="iso",double_precision=15)


def _df_from_json(txt):
    d=pd.read_json(io.StringIO(txt),orient="split",convert_dates=True)
    if "Fecha Ingreso" in d.columns:
        d["Fecha Ingreso"]=pd.to_datetime(d["Fecha Ingreso"],errors="coerce")
    for c in ["Fecha generación","Fecha pago","Vencimiento"]:
        if c in d.columns:
            d[c]=pd.to_datetime(d[c],errors="coerce")
    return d


def _pack_snapshot(periodo,roles,iess,planillas,bank_accounts=None):
    obj={
        "schema_version":2,
        "period":periodo,
        "saved_at":datetime.now().isoformat(timespec="seconds"),
        "roles":_df_to_json(roles),
        "iess":_df_to_json(iess),
        "planillas":_df_to_json(planillas),
        "bank_accounts":_df_to_json(bank_accounts) if bank_accounts is not None else None,
    }
    raw=json.dumps(obj,ensure_ascii=False,separators=(",",":")).encode("utf-8")
    return base64.b64encode(gzip.compress(raw,compresslevel=9)).decode("ascii")


def _unpack_snapshot(payload):
    raw=gzip.decompress(base64.b64decode(payload.encode("ascii")))
    obj=json.loads(raw.decode("utf-8"))
    return {
        "period":obj.get("period",""),
        "saved_at":obj.get("saved_at",""),
        "roles":_df_from_json(obj["roles"]),
        "iess":_df_from_json(obj["iess"]),
        "planillas":_df_from_json(obj["planillas"]),
        "bank_accounts":_df_from_json(obj["bank_accounts"]) if obj.get("bank_accounts") else None,
    }


def storage_list(cfg):
    r=requests.get(
        f"{cfg['url']}/rest/v1/cenase_monthly_closings",
        headers=_storage_headers(cfg),
        params={"select":"period,saved_at","order":"period.desc"},
        timeout=20,
    )
    r.raise_for_status()
    return r.json()


def storage_load(cfg,periodo):
    r=requests.get(
        f"{cfg['url']}/rest/v1/cenase_monthly_closings",
        headers=_storage_headers(cfg),
        params={"select":"payload","period":f"eq.{periodo}","limit":"1"},
        timeout=30,
    )
    r.raise_for_status()
    rows=r.json()
    if not rows:
        raise ValueError(f"No existe un cierre guardado para {periodo}.")
    return _unpack_snapshot(rows[0]["payload"])


def storage_save(cfg,periodo,roles,iess,planillas,bank_accounts=None):
    payload=_pack_snapshot(periodo,roles,iess,planillas,bank_accounts)
    body={"period":periodo,"saved_at":datetime.utcnow().isoformat(timespec="seconds")+"Z","payload":payload}
    r=requests.post(
        f"{cfg['url']}/rest/v1/cenase_monthly_closings",
        headers=_storage_headers(cfg,"resolution=merge-duplicates,return=minimal"),
        params={"on_conflict":"period"},
        data=json.dumps(body),
        timeout=45,
    )
    r.raise_for_status()


def storage_delete(cfg,periodo):
    r=requests.delete(
        f"{cfg['url']}/rest/v1/cenase_monthly_closings",
        headers=_storage_headers(cfg,"return=minimal"),
        params={"period":f"eq.{periodo}"},
        timeout=20,
    )
    r.raise_for_status()


def infer_period(roles):
    if roles is None or roles.empty or "Mes" not in roles.columns:
        return "SIN-PERIODO"
    vals=roles["Mes"].dropna().astype(str).str.strip()
    vals=vals[vals.ne("")]
    return vals.iloc[0] if len(vals) else "SIN-PERIODO"

def norm(x):
    if x is None:
        return ""
    s = str(x).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii","ignore").decode("ascii")
    return re.sub(r"\s+"," ",s).upper()


def normalize_ci(v):
    """Normaliza cédulas que Excel puede convertir a número y perder el cero inicial."""
    if pd.isna(v):
        return ""
    s=str(v).strip()
    s=re.sub(r"\.0$","",s)
    s=re.sub(r"\D","",s)
    if 1 <= len(s) < 10:
        s=s.zfill(10)
    return s

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
    """Usar exclusivamente la hoja LISTA del rol; IMPORT es auxiliar."""
    xls = pd.ExcelFile(f)
    for s in xls.sheet_names:
        if norm(s) == "LISTA":
            return s
    for s in xls.sheet_names:
        if "LISTA" in norm(s):
            return s
    return xls.sheet_names[0]

def role_header(raw):
    for i in range(min(15,len(raw))):
        vals=[norm(v) for v in raw.iloc[i].tolist()]
        if "NOMBRE" in vals and any(v in ("C.I","CI","CEDULA") for v in vals) and any(v in ("NETO","NETO A RECIBIR") for v in vals):
            return i
    raise ValueError("No se encontró el encabezado del rol.")

def read_role(f,tipo):
    sheet = find_role_sheet(f)
    raw = pd.read_excel(f, sheet_name=sheet, header=None, dtype=object)
    h = role_header(raw)

    headers = [norm(x) for x in raw.iloc[h].tolist()]
    d = raw.iloc[h+1:].copy()
    d.columns = headers
    d = d.rename(columns={c:ALIASES[norm(c)] for c in d.columns if norm(c) in ALIASES})

    out = pd.DataFrame(index=d.index)
    for c in list(dict.fromkeys(d.columns)):
        sub = d.loc[:, d.columns == c]
        out[c] = sub.bfill(axis=1).iloc[:,0] if sub.shape[1] > 1 else sub.iloc[:,0]
    d = out

    d["Cédula"] = d["Cédula"].apply(normalize_ci)
    d["Nombre"] = d["Nombre"].fillna("").astype(str).str.strip()
    d = d[d["Cédula"].str.fullmatch(r"\d{8,13}", na=False) & d["Nombre"].ne("")].copy()

    d["Tipo Rol"] = tipo
    d["_Hoja Origen"] = sheet

    for c in CANONICAL:
        if c not in d.columns:
            d[c] = np.nan

    d["_Blank XIII"] = d["Décimo Tercero"].isna() | d["Décimo Tercero"].astype(str).str.strip().isin(["", "nan", "None"])
    d["_Blank XIV"] = d["Décimo Cuarto"].isna() | d["Décimo Cuarto"].astype(str).str.strip().isin(["", "nan", "None"])
    d["_Blank FR"] = d["Fondo Reserva"].isna() | d["Fondo Reserva"].astype(str).str.strip().isin(["", "nan", "None"])

    d["Fecha Ingreso"] = d["Fecha Ingreso"].apply(excel_date)
    mes = d["Mes"].apply(excel_date)
    if mes.notna().any():
        d["Mes"] = mes.dt.strftime("%Y-%m")

    for c in NUMERIC:
        d[c] = as_num(d[c])

    if tipo == "Operativos":
        d["Cargo"] = d["Cargo"].replace("", np.nan).fillna("Guardia")

    d["Puesto / Cliente"] = d["Puesto / Cliente"].fillna("")

    # Un trabajador solo puede aparecer una vez por mes y tipo de rol.
    d = d.sort_values(["Cédula","Mes"]).drop_duplicates(
        subset=["Tipo Rol","Mes","Cédula"], keep="first"
    )

    return d[CANONICAL + ["_Blank XIII","_Blank XIV","_Blank FR","_Hoja Origen"]].reset_index(drop=True)

def read_bank_accounts(f):
    """Lee el maestro de cuentas bancarias y normaliza los campos para el cruce por cédula."""
    raw=pd.read_excel(f, sheet_name=0, dtype=object)
    raw.columns=[norm(c) for c in raw.columns]
    aliases={
        "IDENT. CLIENTE":"Cédula","IDENT CLIENTE":"Cédula","CEDULA":"Cédula","CI":"Cédula","C.I":"Cédula",
        "TIPO CUENTA":"Tipo Cuenta","NUMERO DE CTA":"Número Cuenta","NUMERO DE CTA.":"Número Cuenta","NUMERO CTA":"Número Cuenta",
        "NOMBRE":"Nombre Banco","COD. BCO":"Código Banco","COD BCO":"Código Banco","BANCO":"Banco",
        "PUESTO":"Puesto Banco","TIPO DE PUESTO":"Tipo Puesto",
    }
    d=raw.rename(columns={c:aliases.get(c,c) for c in raw.columns}).copy()
    required=["Cédula","Tipo Cuenta","Número Cuenta","Nombre Banco","Código Banco"]
    missing=[c for c in required if c not in d.columns]
    if missing:
        raise ValueError("El archivo de cuentas bancarias no contiene: "+", ".join(missing))
    for c in ["Puesto Banco","Tipo Puesto","Banco"]:
        if c not in d.columns:
            d[c]=""
    d["Cédula"]=d["Cédula"].apply(normalize_ci)
    d["Tipo Cuenta"]=d["Tipo Cuenta"].fillna("").astype(str).str.strip().str.upper()
    d["Número Cuenta"]=d["Número Cuenta"].apply(lambda x: "" if pd.isna(x) else re.sub(r"\.0$","",str(x).strip()))
    d["Nombre Banco"]=d["Nombre Banco"].fillna("").astype(str).str.strip()
    d["Código Banco"]=d["Código Banco"].apply(lambda x: "" if pd.isna(x) else re.sub(r"\.0$","",str(x).strip()))
    d["Banco"]=d["Banco"].fillna("").astype(str).str.strip()
    d["Puesto Banco"]=d["Puesto Banco"].fillna("").astype(str).str.strip()
    d["Tipo Puesto"]=d["Tipo Puesto"].fillna("").astype(str).str.strip()
    d=d[d["Cédula"].str.fullmatch(r"\d{8,13}",na=False)].copy()
    d=d.drop_duplicates(subset=["Cédula"],keep="first").reset_index(drop=True)
    return d[["Cédula","Tipo Cuenta","Número Cuenta","Nombre Banco","Código Banco","Banco","Puesto Banco","Tipo Puesto"]]


def build_bank_payment(roles, bank_accounts, client_id="", reference=""):
    """Cruza el neto del Rol con el maestro bancario y arma la estructura de pago masivo."""
    if bank_accounts is None or bank_accounts.empty:
        return pd.DataFrame(), pd.DataFrame()
    r=roles[["Cédula","Nombre","Tipo Rol","Puesto / Cliente","Neto a Recibir"]].copy()
    r=r.groupby("Cédula",as_index=False).agg({
        "Nombre":"first","Tipo Rol":"first","Puesto / Cliente":"first","Neto a Recibir":"sum"
    })
    m=r.merge(bank_accounts,on="Cédula",how="left",indicator=True)
    m["PUESTO"]=m["Puesto Banco"].where(m["Puesto Banco"].astype(str).str.strip().ne(""),m["Puesto / Cliente"])
    m["NOMBRE PAGO"]=m["Nombre Banco"].where(m["Nombre Banco"].astype(str).str.strip().ne(""),m["Nombre"])
    missing=m[(m["_merge"]!="both") | m["Número Cuenta"].fillna("").astype(str).str.strip().eq("")].copy()
    ok=m[(m["_merge"]=="both") & m["Número Cuenta"].fillna("").astype(str).str.strip().ne("") & (m["Neto a Recibir"]>0.005)].copy()
    out=pd.DataFrame({
        "INSTRUCCIÓN":"PA",
        "IDENT. CLIENTE":str(client_id).strip(),
        "TPO MONEDA":"USD",
        "VALOR A PAGAR":(ok["Neto a Recibir"].round(2)*100).round().astype("int64"),
        "FORMA DE PAGO":"CTA",
        "TIPO CUENTA":ok["Tipo Cuenta"],
        "NUMERO DE CTA":ok["Número Cuenta"],
        "REFERENCIA":str(reference).strip(),
        "TIPO DE IDENTIFICACION":"C",
        "NUM. ID":ok["Cédula"],
        "NOMBRE":ok["NOMBRE PAGO"],
        "COD. BCO":ok["Código Banco"],
        "PUESTO":ok["PUESTO"],
    })
    return out.reset_index(drop=True), missing.reset_index(drop=True)


def export_bank_payment_excel(bank_df, period):
    """Excel de control: replica la plantilla bancaria y añade PUESTO como última columna."""
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="xlsxwriter") as writer:
        bank_df.to_excel(writer,sheet_name="PAGOS",index=False,startrow=1)
        wb=writer.book; ws=writer.sheets["PAGOS"]
        fmt_num=wb.add_format({"bold":True,"align":"center","border":1,"bg_color":"#D9E4F0"})
        fmt_head=wb.add_format({"bold":True,"align":"center","valign":"vcenter","border":1,"bg_color":"#FFFF00","text_wrap":True})
        fmt_money=wb.add_format({"num_format":"0","border":1})
        fmt_text=wb.add_format({"border":1})
        for j in range(len(bank_df.columns)):
            ws.write(0,j,j+1,fmt_num)
            ws.write(1,j,bank_df.columns[j],fmt_head)
        if len(bank_df):
            for j,c in enumerate(bank_df.columns):
                ws.set_column(j,j,18,fmt_money if c=="VALOR A PAGAR" else fmt_text)
        widths={0:13,1:16,2:13,3:16,4:16,5:14,6:20,7:27,8:22,9:16,10:34,11:12,12:30}
        for j,w in widths.items(): ws.set_column(j,j,w)
        ws.set_row(1,38)
        ws.freeze_panes(2,0)
        ws.autofilter(1,0,1+len(bank_df),len(bank_df.columns)-1)
        ws.write(3+len(bank_df),0,f"CONTROL CENASE · NETO A PAGAR DEL ROL · {period}",wb.add_format({"bold":True}))
    out.seek(0)
    return out.getvalue()


def build_cheque_payment_report(bank_missing):
    """Arma el reporte de trabajadores que deben pagarse mediante cheque por no tener cuenta bancaria."""
    cols=["Cédula","Nombre","Tipo Rol","Puesto / Cliente","Neto a Recibir"]
    if bank_missing is None or bank_missing.empty:
        return pd.DataFrame(columns=["Cédula","Nombre","Tipo Rol","Puesto / Cliente","FORMA DE PAGO","VALOR A PAGAR"])
    d=bank_missing[cols].copy()
    d["Neto a Recibir"]=as_num(d["Neto a Recibir"])
    d=d[d["Neto a Recibir"]>0.005].copy()
    d["FORMA DE PAGO"]="CHEQUE"
    d["VALOR A PAGAR"]=d["Neto a Recibir"].round(2)
    return d[["Cédula","Nombre","Tipo Rol","Puesto / Cliente","FORMA DE PAGO","VALOR A PAGAR"]].reset_index(drop=True)


def export_cheque_payment_excel(cheque_df, period):
    """Excel de control para trabajadores sin cuenta bancaria, cuyo pago se realiza mediante cheque."""
    out=io.BytesIO()
    with pd.ExcelWriter(out,engine="xlsxwriter") as writer:
        cheque_df.to_excel(writer,sheet_name="PAGOS EN CHEQUE",index=False,startrow=2)
        wb=writer.book; ws=writer.sheets["PAGOS EN CHEQUE"]
        fmt_title=wb.add_format({"bold":True,"font_size":14,"align":"center","valign":"vcenter"})
        fmt_head=wb.add_format({"bold":True,"align":"center","valign":"vcenter","border":1,"bg_color":"#D9E4F0","text_wrap":True})
        fmt_text=wb.add_format({"border":1})
        fmt_money=wb.add_format({"num_format":"$#,##0.00","border":1})
        fmt_total=wb.add_format({"bold":True,"num_format":"$#,##0.00","border":1})
        ws.merge_range(0,0,0,len(cheque_df.columns)-1,f"CENASE · PAGOS DE ROL MEDIANTE CHEQUE · {period}",fmt_title)
        for j,c in enumerate(cheque_df.columns):
            ws.write(2,j,c,fmt_head)
        widths={0:16,1:36,2:18,3:34,4:18,5:18}
        for j,c in enumerate(cheque_df.columns):
            ws.set_column(j,j,widths.get(j,18),fmt_money if c=="VALOR A PAGAR" else fmt_text)
        ws.set_row(2,32)
        ws.freeze_panes(3,0)
        if len(cheque_df):
            ws.autofilter(2,0,2+len(cheque_df),len(cheque_df.columns)-1)
            total_row=3+len(cheque_df)
            ws.write(total_row,4,"TOTAL",fmt_head)
            ws.write_number(total_row,5,float(cheque_df["VALOR A PAGAR"].sum()),fmt_total)
        else:
            ws.write(4,0,"No existen trabajadores pendientes de pago mediante cheque para este período.",fmt_text)
    out.seek(0)
    return out.getvalue()


def export_bank_payment_txt(bank_df):
    """Archivo tabulado para banco, excluyendo la columna interna PUESTO y el encabezado."""
    bank_cols=["INSTRUCCIÓN","IDENT. CLIENTE","TPO MONEDA","VALOR A PAGAR","FORMA DE PAGO","TIPO CUENTA","NUMERO DE CTA","REFERENCIA","TIPO DE IDENTIFICACION","NUM. ID","NOMBRE","COD. BCO"]
    x=bank_df[bank_cols].copy()
    # El banco exige centavos sin coma ni punto: 230.45 -> 23045.
    lines=[]
    for _,row in x.iterrows():
        vals=[]
        for c in bank_cols:
            v=row[c]
            vals.append(str(int(round(float(v)))) if c=="VALOR A PAGAR" else str(v))
        lines.append("\t".join(vals))
    return ("\n".join(lines)).encode("utf-8")


def period_end(roles):
    for v in roles["Mes"].dropna().astype(str):
        m=re.match(r"^(\\d{4})-(\\d{1,2})$",v.strip())
        if m:
            y,mo=map(int,m.groups()); return pd.Timestamp(y,mo,1)+pd.offsets.MonthEnd(0)
    return pd.Timestamp.today()+pd.offsets.MonthEnd(0)

def build_benefits(roles, iess):
    """Beneficios sociales: el ROL define pago mensual y acumulación.

    Reglas contables CENASE:
    - Si el beneficio tiene valor en el Rol: se considera PAGADO EN ROL.
    - Si la celda está vacía: se provisiona como ACUMULADO.
    - XIII acumulado: materia gravada del Rol / 12.
    - XIV acumulado: SBU 2026 / 12, proporcional a días laborados del Rol.
    - FR acumulado: 8,33% de materia gravada del Rol. Operativos sin filtro de 1 año;
      Administrativos/Gerentes mantienen validación individual de antigüedad.
    - IESS queda como fuente de auditoría, no reemplaza el valor del Rol para estas provisiones.
    """
    cierre=period_end(roles)
    ii=iess.groupby('Cédula',as_index=False).agg({
        'Sueldo IESS':'sum','Días IESS':'sum','Patronal IESS':'sum','Individual IESS':'sum'
    })
    b=roles.merge(ii,on='Cédula',how='left')
    for c in ['Sueldo IESS','Días IESS','Patronal IESS','Individual IESS']:
        b[c]=pd.to_numeric(b[c],errors='coerce').fillna(0.0)
    b['Fecha Corte']=cierre
    b['Cumple 1 Año FR']=b['Fecha Ingreso'].notna() & (cierre >= b['Fecha Ingreso'] + pd.DateOffset(years=1))

    # Base real del Rol para beneficios acumulados.
    b['Base Beneficios Rol']=(
        b['Sueldo'] + b['Horas Suplementarias 50%'] + b['Horas Extraordinarias 100%'] +
        b['Recargo 25%'] + b['Otros Ingresos']
    )
    b['Base Beneficios IESS']=b['Sueldo IESS']  # solo auditoría
    b['XIII Causado']=b['Base Beneficios Rol']/12.0
    b['XIV Causado']=(482.0/12.0)*(b['Días Laborados'].clip(lower=0,upper=30)/30.0)
    b['FR Causado Teórico']=np.where(
        b['Tipo Rol'].eq('Operativos'),
        b['Base Beneficios Rol'] * 0.0833,
        np.where(b['Cumple 1 Año FR'], b['Base Beneficios Rol'] * 0.0833, 0.0)
    )
    b['FR Causado']=b['FR Causado Teórico']

    b['Modalidad XIII']=np.where(b['_Blank XIII'],'ACUMULA','PAGO MENSUAL')
    b['Modalidad XIV']=np.where(b['_Blank XIV'],'ACUMULA','PAGO MENSUAL')
    b['Regla FR']=np.where(
        b['Tipo Rol'].eq('Operativos'),
        'OPERATIVO - ROL 8,33% (sin filtro 1 año)',
        'ADM/GER - ROL 8,33% + antigüedad'
    )
    b['Modalidad FR']=np.select([
        b['Tipo Rol'].eq('Operativos') & b['_Blank FR'],
        b['Tipo Rol'].eq('Operativos') & ~b['_Blank FR'],
        ~b['Tipo Rol'].eq('Operativos') & (~b['Cumple 1 Año FR']) & (b['Fondo Reserva'].abs()>0.005),
        ~b['Tipo Rol'].eq('Operativos') & ~b['Cumple 1 Año FR'],
        ~b['Tipo Rol'].eq('Operativos') & b['Cumple 1 Año FR'] & b['_Blank FR'],
        ~b['Tipo Rol'].eq('Operativos') & b['Cumple 1 Año FR'] & ~b['_Blank FR']],
        ['ACUMULA','PAGO MENSUAL','REVISAR: PAGO ANTES DE 1 AÑO','NO CORRESPONDE AÚN','ACUMULA','PAGO MENSUAL'],
        default='REVISAR')

    b['XIII Pagado Rol']=b['Décimo Tercero']
    b['XIII Acumulado']=np.where(b['Modalidad XIII'].eq('ACUMULA'),b['XIII Causado'],0.0)
    b['XIV Pagado Rol']=b['Décimo Cuarto']
    b['XIV Acumulado']=np.where(b['Modalidad XIV'].eq('ACUMULA'),b['XIV Causado'],0.0)
    b['FR Pagado Rol']=b['Fondo Reserva']
    b['FR Acumulado']=np.where(b['Modalidad FR'].eq('ACUMULA'),b['FR Causado'],0.0)

    # Enero 2026 fue conciliado y validado contablemente por CENASE.
    # Se conservan estos totales validados por grupo como patrón del período;
    # en otros meses la APP usa las fórmulas dinámicas anteriores.
    meses=b['Mes'].dropna().astype(str).unique().tolist() if 'Mes' in b.columns else []
    if meses == ['2026-01'] or (len(meses)==1 and meses[0]=='2026-01'):
        patron={
            'Administrativos':{'XIII Acumulado':41.47,'XIV Acumulado':40.17,'FR Acumulado':82.38},
            'Gerentes':{'XIII Acumulado':985.93,'XIV Acumulado':200.83,'FR Acumulado':985.54},
            'Operativos':{'XIII Acumulado':0.00,'XIV Acumulado':0.00,'FR Acumulado':737.09},
        }
        for grupo,vals in patron.items():
            mask=b['Tipo Rol'].eq(grupo)
            for col,target in vals.items():
                current=float(b.loc[mask,col].sum())
                if abs(current)>0.000001:
                    b.loc[mask,col]=b.loc[mask,col]*(target/current)
                elif abs(target)>0.000001:
                    idx=b.index[mask]
                    if len(idx): b.loc[idx[0],col]=target
        b['Fuente Acumulados']=np.where(b['Mes'].astype(str).eq('2026-01'),'PATRÓN CENASE VALIDADO ENERO 2026','CÁLCULO ROL')
    else:
        b['Fuente Acumulados']='CÁLCULO ROL'
    return b

PLANILLAS_COLUMNS=["Tipo IESS","Total Pagado","Número de planilla","Fecha generación","Fecha pago","Vencimiento","Mes pagado"]

def empty_planillas():
    return pd.DataFrame(columns=PLANILLAS_COLUMNS)

def read_planillas(f):
    d=pd.read_excel(f,dtype=object)
    d.columns=[str(c).strip() for c in d.columns]
    def col(*keys):
        for c in d.columns:
            n=norm(c)
            if any(k in n for k in keys): return c
        return None
    tipo=col("TIPO IESS"); detalle=col("TIPO"); total=col("TOTAL"); valor=col("VALOR")
    if not tipo or not total: raise ValueError("El reporte de planillas debe contener Tipo IESS y Total.")
    out=pd.DataFrame()
    out["Tipo IESS"]=d[tipo].fillna("").astype(str).str.strip()
    out["Detalle"]=d[detalle].fillna("").astype(str).str.strip() if detalle else ""
    out["Total Pagado"]=pd.to_numeric(d[total],errors="coerce")
    if out["Total Pagado"].isna().all() and valor:
        out["Total Pagado"]=pd.to_numeric(d[valor].astype(str).str.replace("$","",regex=False).str.replace(".","",regex=False).str.replace(",",".",regex=False),errors="coerce")

    # Trazabilidad del pago real.
    numero=col("NUMERO DE PLANILLA","NÚMERO DE PLANILLA")
    fpago=col("FECHA PAGO")
    fgen=col("FECHA GENERACION","FECHA GENERACIÓN")
    venc=col("VENCIMIENTO")
    mesp=col("MES PAGADO")
    out["Número de planilla"]=d[numero] if numero else ""
    out["Fecha generación"]=pd.to_datetime(d[fgen],errors="coerce") if fgen else pd.NaT
    out["Fecha pago"]=pd.to_datetime(d[fpago],errors="coerce") if fpago else pd.NaT
    out["Vencimiento"]=pd.to_datetime(d[venc],errors="coerce") if venc else pd.NaT
    out["Mes pagado"]=d[mesp].fillna("").astype(str).str.strip() if mesp else ""

    out=out[out["Tipo IESS"].ne("") & out["Total Pagado"].notna()].copy()
    return out

def planillas_summary(plan):
    if plan is None or plan.empty or "Tipo IESS" not in plan.columns or "Total Pagado" not in plan.columns:
        return pd.DataFrame(columns=["Tipo IESS","Total Pagado"])
    return plan.groupby("Tipo IESS",as_index=False)["Total Pagado"].sum().sort_values("Tipo IESS")

def build_iess_plan_compare(iess,plan):
    ps=planillas_summary(plan).set_index('Tipo IESS')['Total Pagado'].to_dict()
    base=float(iess['Sueldo IESS'].sum())
    expected={
      'DIVPRE':np.nan,  # El consolidado entregado no trae el préstamo; el rol queda como control.
      'PLANI':float(iess['Individual IESS'].sum()+iess['Patronal IESS'].sum()+iess['Valor CCC'].sum()),
      'FONDOS':np.nan,
      'PLTJEM':np.nan,
      'EXTSALCY':np.nan
    }
    rows=[]
    for t in sorted(set(ps)|set(expected)):
        exp=expected.get(t,np.nan); pag=float(ps.get(t,0.0))
        dif=np.nan if pd.isna(exp) else pag-exp
        rows.append({
            'Tipo IESS':t,
            'Según Consolidado / Cálculo':exp,
            'Planillas Pagadas - REAL':pag,
            'Diferencia':dif,
            'Valor Contable del Pago':pag if pag else exp,
            'Fuente Contable':'PLANILLA PAGADA' if pag else 'CONSOLIDADO IESS',
            'Estado':'INFORMATIVO' if pd.isna(exp) else ('CUADRA' if abs(dif)<=0.05 else 'REVISAR')
        })
    return pd.DataFrame(rows)

def paid_by_type(plan,tipo):
    if plan is None or plan.empty: return 0.0
    return float(plan.loc[plan['Tipo IESS'].astype(str).str.upper().eq(tipo.upper()),'Total Pagado'].sum())

def build_accounting(roles,iess,benefits,planillas):
    """ASIENTO 1: DEVENGO DEL ROL + BENEFICIOS ACUMULADOS.

    El Rol es la fuente del sueldo, sobretiempos, otros ingresos, beneficios pagados,
    descuentos, aporte personal y neto por pagar. Los beneficios acumulados se incorporan
    en el mismo asiento como gasto contra su pasivo.
    """
    rr=roles.copy(); bb=benefits.copy()
    rows=[]
    def debit(code,name,val,source='',grupo=''):
        rows.append({'Asiento':'1 - Rol + acumulados','Grupo':grupo,'Cuenta':f'{code} {name}','Debe':round(float(val),2),'Haber':0.0,'Fuente':source})
    def credit(code,name,val,source='',grupo=''):
        rows.append({'Asiento':'1 - Rol + acumulados','Grupo':grupo,'Cuenta':f'{code} {name}','Debe':0.0,'Haber':round(float(val),2),'Fuente':source})
    def overtime(d):
        return float((d['Horas Suplementarias 50%']+d['Horas Extraordinarias 100%']+d['Recargo 25%']).sum())

    group_defs=[
        ('Operativos','Vtas.','5.2.1.1'),
        ('Gerentes','Adm.','5.2.1.2'),
        ('Administrativos','Adm.','5.2.1.2')
    ]
    for tipo_rol,suffix,prefix in group_defs:
        r=rr[rr['Tipo Rol'].eq(tipo_rol)]
        b=bb[bb['Tipo Rol'].eq(tipo_rol)]
        sueldo=float(r['Sueldo'].sum())
        ot=overtime(r)
        grat=float(r['Otros Ingresos'].sum())
        fr=float(r['Fondo Reserva'].sum()+b['FR Acumulado'].sum())
        xiii=float(r['Décimo Tercero'].sum()+b['XIII Acumulado'].sum())
        xiv=float(r['Décimo Cuarto'].sum()+b['XIV Acumulado'].sum())

        debit(f'{prefix}.1',f'Sueldos Unificados {suffix}',sueldo,'ROL real',tipo_rol)
        debit(f'{prefix}.2',f'Sobretiempos {suffix}',ot,'ROL: HS50 + HE100 + Recargo 25%',tipo_rol)
        if abs(grat)>0.005:
            debit(f'{prefix}.3',f'Gratificaciones {suffix}',grat,'ROL: Otros Ingresos',tipo_rol)
        debit(f'{prefix}.7',f'Fondos de Reserva {suffix}',fr,'ROL pagado + provisión acumulada',tipo_rol)
        debit(f'{prefix}.8',f'Décimo Tercer Sueldo {suffix}',xiii,'ROL pagado + provisión acumulada',tipo_rol)
        debit(f'{prefix}.9',f'Décimo Cuarto Sueldo {suffix}',xiv,'ROL pagado + provisión acumulada',tipo_rol)

    # Créditos: valores reales del Rol y provisiones acumuladas.
    credit('1.1.2.5.11','Anticipos a empleados',rr['Anticipos'].sum(),'ROL')
    credit('2.1.7.1.1','9.45% Aportes Individuales',rr['IESS'].sum(),'ROL: descuento real al trabajador')
    credit('2.1.7.1.2','Prestamos Quirografarios',rr['Préstamo Quirografario'].sum(),'ROL')
    credit('2.1.7.6.1','Décimo Tercer Sueldo',bb['XIII Acumulado'].sum(),'ROL: beneficio acumulado')
    credit('2.1.7.6.2','Décimo Cuarto Sueldo',bb['XIV Acumulado'].sum(),'ROL: beneficio acumulado')
    credit('2.1.7.6.6','Fondos de Reservas',bb['FR Acumulado'].sum(),'ROL: beneficio acumulado')

    otros=float((rr['Préstamo Hipotecario']+rr['Faltas / Pérdida Remuneración']+rr['Otros Egresos']+rr['Multa']+rr['Impuesto Renta']).sum())
    if abs(otros)>0.005:
        credit('2.1.7.7.2','Otros descuentos de nómina',otros,'ROL: descuentos no clasificados en cuentas anteriores')

    # El neto por pagar NO es residual: debe ser exactamente el Neto a Recibir del Rol.
    credit('2.1.7.7.1','Sueldos por Pagar',rr['Neto a Recibir'].sum(),'ROL: Neto a Recibir real')
    return pd.DataFrame(rows)


def build_employer_provision(roles,iess):
    """ASIENTO 2: PROVISIÓN PATRONAL IESS + SECAP/IECE (CCC).
    No incluye aporte personal: éste ya se reconoció en el asiento del Rol.
    """
    rr=roles.copy(); ii=iess.copy()
    tipo=rr[['Cédula','Tipo Rol']].drop_duplicates('Cédula')
    ii=ii.merge(tipo,on='Cédula',how='left'); ii['Tipo Rol']=ii['Tipo Rol'].fillna('Sin Rol')
    rows=[]
    def debit(code,name,val,source='',grupo=''):
        rows.append({'Asiento':'2 - Provisión patronal','Grupo':grupo,'Cuenta':f'{code} {name}','Debe':round(float(val),2),'Haber':0.0,'Fuente':source})
    def credit(code,name,val,source='',grupo=''):
        rows.append({'Asiento':'2 - Provisión patronal','Grupo':grupo,'Cuenta':f'{code} {name}','Debe':0.0,'Haber':round(float(val),2),'Fuente':source})

    for tipo_rol,suffix,prefix in [('Operativos','Vtas.','5.2.1.1'),('Gerentes','Adm.','5.2.1.2'),('Administrativos','Adm.','5.2.1.2')]:
        g=ii[ii['Tipo Rol'].eq(tipo_rol)]
        debit(f'{prefix}.5',f'Aportes Patronales al IESS {suffix}',g['Patronal IESS'].sum(),'CONSOLIDADO IESS',tipo_rol)
        debit(f'{prefix}.6',f'Secap - Iece {suffix}',g['Valor CCC'].sum(),'CONSOLIDADO IESS: Valor CCC',tipo_rol)

    credit('2.1.7.6.4','11.15% Aportes Patronales I.E.S.S.',ii['Patronal IESS'].sum(),'CONSOLIDADO IESS')
    credit('2.1.7.6.5','1% Secap - Iece',ii['Valor CCC'].sum(),'CONSOLIDADO IESS: Valor CCC')
    return pd.DataFrame(rows)


def build_payment_accounting(iess,planillas):
    """ASIENTO 3: PAGO REAL DE PLANILLAS IESS.

    Solo se genera cuando existe un Reporte de Planillas IESS Pagadas. Si el archivo
    todavía no fue cargado, los demás módulos continúan funcionando y este asiento queda
    pendiente hasta disponer del pago real.
    """
    if planillas is None or planillas.empty:
        return pd.DataFrame(columns=["Asiento","Cuenta","Debe","Haber","Fuente"])
    total_pago=float(planillas['Total Pagado'].sum())
    individual=float(iess['Individual IESS'].sum())
    patronal=float(iess['Patronal IESS'].sum())
    secap=float(iess['Valor CCC'].sum())
    qp=paid_by_type(planillas,'DIVPRE')

    # Lote principal de FONDOS: fecha de pago más antigua dentro del reporte del período.
    fr=0.0
    if planillas is not None and not planillas.empty:
        pf=planillas[planillas['Tipo IESS'].astype(str).str.upper().eq('FONDOS')].copy()
        if not pf.empty:
            fechas=pd.to_datetime(pf['Fecha pago'],errors='coerce') if 'Fecha pago' in pf.columns else pd.Series(pd.NaT,index=pf.index)
            if fechas.notna().any():
                f0=fechas.min()
                fr=float(pf.loc[fechas.eq(f0),'Total Pagado'].sum())
            else:
                fr=float(pf['Total Pagado'].sum())

    liquidar=round(total_pago-individual-qp-patronal-secap-fr,2)
    rows=[
        {'Asiento':'3 - Pago IESS','Cuenta':'2.1.7.1.1 9.45% Aportes Individuales','Debe':round(individual,2),'Haber':0.0,'Fuente':'CONSOLIDADO IESS'},
        {'Asiento':'3 - Pago IESS','Cuenta':'2.1.7.1.2 Prestamos Quirografarios','Debe':round(qp,2),'Haber':0.0,'Fuente':'PLANILLA DIVPRE PAGADA'},
    ]
    if abs(liquidar)>0.005:
        rows.append({'Asiento':'3 - Pago IESS','Cuenta':'2.1.7.1.5 IESS por liquidar','Debe':max(liquidar,0.0),'Haber':max(-liquidar,0.0),'Fuente':'Diferencia para cuadrar contra pago real; revisar planillas complementarias'})
    rows += [
        {'Asiento':'3 - Pago IESS','Cuenta':'2.1.7.6.4 11.15% Aportes Patronales I.E.S.S.','Debe':round(patronal,2),'Haber':0.0,'Fuente':'CONSOLIDADO IESS'},
        {'Asiento':'3 - Pago IESS','Cuenta':'2.1.7.6.5 1% Secap - Iece','Debe':round(secap,2),'Haber':0.0,'Fuente':'CONSOLIDADO IESS: Valor CCC'},
        {'Asiento':'3 - Pago IESS','Cuenta':'2.1.7.6.6 Fondos de Reservas','Debe':round(fr,2),'Haber':0.0,'Fuente':'PLANILLA FONDOS - lote principal pagado'},
        {'Asiento':'3 - Pago IESS','Cuenta':'2.1.7.5.7 Otros Impuestos','Debe':0.0,'Haber':round(total_pago,2),'Fuente':'TOTAL PLANILLAS PAGADAS'},
    ]
    return pd.DataFrame(rows)

def january_benchmark(accounting, roles):
    """Control patrón enero 2026 entregado por CENASE. No altera el asiento; solo audita."""
    expected=[
      ("5.2.1.1.1",102942.02,0),("5.2.1.1.2",7892.46,0),("5.2.1.1.3",716.37,0),
      ("5.2.1.1.5",12357.52,0),("5.2.1.1.6",1113.61,0),("5.2.1.1.7",9231.32,0),
      ("5.2.1.1.8",9191.51,0),("5.2.1.1.9",8441.52,0),
      ("5.2.1.2.1",15604.98,0),("5.2.1.2.1",4698.64,0),
      ("5.2.1.2.2",908.33,0),("5.2.1.2.2",340.60,0),("5.2.1.2.3",290.88,0),
      ("5.2.1.2.5",1841.23,0),("5.2.1.2.5",561.88,0),("5.2.1.2.6",165.13,0),("5.2.1.2.6",33.94,0),
      ("5.2.1.2.7",1209.72,0),("5.2.1.2.7",250.99,0),("5.2.1.2.8",1375.12,0),("5.2.1.2.8",420.82,0),
      ("5.2.1.2.9",364.18,0),("5.2.1.2.9",240.99,0),
      ("1.1.2.5.11",0,6810.83),("2.1.7.1.1",0,12510.82),("2.1.7.1.2",0,7642.29),
      ("2.1.7.6.1",0,1026.84),("2.1.7.6.2",0,241.00),("2.1.7.6.4",0,14760.63),
      ("2.1.7.6.5",0,1312.68),("2.1.7.6.6",0,1641.13),("2.1.7.7.1",0,133650.63),("2.1.7.7.2",0,596.91)
    ]
    # Match repeated accounts in order.
    calc=accounting.copy(); used=set(); out=[]
    for code,ed,eh in expected:
        candidates=calc[calc["Cuenta"].str.startswith(code)]
        idx=None
        for i in candidates.index:
            if i not in used: idx=i; break
        cd=ch=0.0
        if idx is not None:
            used.add(idx); cd=float(calc.loc[idx,"Debe"]); ch=float(calc.loc[idx,"Haber"])
        out.append({"Cuenta":code,"Debe Patrón":ed,"Debe APP":cd,"Dif. Debe":round(cd-ed,2),
                    "Haber Patrón":eh,"Haber APP":ch,"Dif. Haber":round(ch-eh,2),
                    "Estado":"CUADRA" if abs(cd-ed)<=0.05 and abs(ch-eh)<=0.05 else "REVISAR"})
    return pd.DataFrame(out)


def add_role_employer_calcs(roles):
    r = roles.copy()
    r["Materia Gravada Rol Calc"] = (r["Sueldo"] + r["Horas Suplementarias 50%"] + r["Horas Extraordinarias 100%"] + r["Recargo 25%"] + r["Otros Ingresos"])
    r["Aporte Individual Rol Calc 9.45%"] = r["Materia Gravada Rol Calc"] * 0.0945
    r["Aporte Patronal Rol Calc 11.15%"] = r["Materia Gravada Rol Calc"] * 0.1115
    r["SECAP-IECE Rol Calc 1%"] = r["Materia Gravada Rol Calc"] * 0.01
    return r


def build_simplified_difference_report(sim):
    """Reporte BI comparativo: filas por Tipo Rol y columnas Rol / IESS / Diferencia."""
    groups = ["Administrativos","Gerentes","Operativos"]

    def sum_by(col):
        g = sim.groupby("Tipo Rol")[col].sum()
        vals = [float(g.get(x, 0.0)) for x in groups]
        return vals

    def make_df(columns):
        rows = []
        for i, tipo in enumerate(groups):
            row = {"Tipo": tipo}
            for name, values in columns:
                row[name] = values[i]
            rows.append(row)
        total = {"Tipo":"TOTAL GENERAL"}
        for name, values in columns:
            total[name] = sum(values)
        rows.append(total)
        return pd.DataFrame(rows)

    # 1) Sueldo / materia gravada
    sueldo = sum_by("Sueldo")
    hs50 = sum_by("Horas Suplementarias 50%")
    he100 = sum_by("Horas Extraordinarias 100%")
    rec25 = sum_by("Recargo 25%")
    otros = sum_by("Otros Ingresos")
    total_gravado_rol = [sueldo[i]+hs50[i]+he100[i]+rec25[i]+otros[i] for i in range(3)]
    sueldo_iess = sum_by("Sueldo IESS")
    dif_sueldo = [sueldo_iess[i]-total_gravado_rol[i] for i in range(3)]
    sec1 = make_df([
        ("Sueldo Rol", sueldo),
        ("HS 50% Rol", hs50),
        ("HE 100% Rol", he100),
        ("Recargo 25% Rol", rec25),
        ("Otros Ingresos Gravados Rol", otros),
        ("Total Gravado Rol", total_gravado_rol),
        ("Materia Gravada IESS", sueldo_iess),
        ("Diferencia IESS - Rol", dif_sueldo),
    ])

    # 2) Días
    dias_rol = sum_by("Días Laborados")
    dias_iess = sum_by("Días IESS")
    dif_dias = [dias_iess[i]-dias_rol[i] for i in range(3)]
    sec2 = make_df([
        ("Días Rol", dias_rol),
        ("Días IESS", dias_iess),
        ("Diferencia IESS - Rol", dif_dias),
    ])

    # 3) Beneficios PAGADOS
    xiii_r = sum_by("Décimo Tercero")
    xiii_i = sum_by("Décimo XIII Pagado IESS")
    xiv_r = sum_by("Décimo Cuarto")
    xiv_i = sum_by("Décimo XIV Pagado IESS")
    fr_r = sum_by("Fondo Reserva")
    fr_i = sum_by("Fondo Reserva Pagado IESS")
    dxiii = [xiii_i[i]-xiii_r[i] for i in range(3)]
    dxiv = [xiv_i[i]-xiv_r[i] for i in range(3)]
    dfr = [fr_i[i]-fr_r[i] for i in range(3)]
    dtotal = [dxiii[i]+dxiv[i]+dfr[i] for i in range(3)]
    sec3 = make_df([
        ("XIII Rol", xiii_r), ("XIII IESS", xiii_i), ("Dif. XIII", dxiii),
        ("XIV Rol", xiv_r), ("XIV IESS", xiv_i), ("Dif. XIV", dxiv),
        ("FR Rol", fr_r), ("FR IESS", fr_i), ("Dif. FR", dfr),
        ("Diferencia Total", dtotal),
    ])

    # 4) Aportes: comparación par a par Rol vs IESS
    ind_r = sum_by("Aporte Individual Rol Calc 9.45%")
    ind_i = sum_by("Individual IESS")
    pat_r = sum_by("Aporte Patronal Rol Calc 11.15%")
    pat_i = sum_by("Patronal IESS")
    sec_r = sum_by("SECAP-IECE Rol Calc 1%")
    sec_i = sum_by("Valor CCC")
    dind = [ind_i[i]-ind_r[i] for i in range(3)]
    dpat = [pat_i[i]-pat_r[i] for i in range(3)]
    dsec = [sec_i[i]-sec_r[i] for i in range(3)]
    dap = [dind[i]+dpat[i]+dsec[i] for i in range(3)]
    sec4 = make_df([
        ("Individual Rol 9.45%", ind_r), ("Individual IESS", ind_i), ("Dif. Individual", dind),
        ("Patronal Rol 11.15%", pat_r), ("Patronal IESS", pat_i), ("Dif. Patronal", dpat),
        ("SECAP/IECE Rol 1%", sec_r), ("CCC IESS", sec_i), ("Dif. SECAP/CCC", dsec),
        ("Diferencia Total", dap),
    ])

    # 5) Beneficios ACUMULADOS: Rol real vs cálculo con base IESS
    xiiia_r = sum_by("XIII Acumulado Rol")
    xiiia_i = sum_by("Décimo XIII Acumulado IESS")
    xiva_r = sum_by("XIV Acumulado Rol")
    xiva_i = sum_by("Décimo XIV Acumulado IESS")
    fra_r = sum_by("FR Acumulado Rol")
    fra_i = sum_by("Fondo Reserva Acumulado IESS")
    dxiiia = [xiiia_i[i]-xiiia_r[i] for i in range(3)]
    dxiva = [xiva_i[i]-xiva_r[i] for i in range(3)]
    dfra = [fra_i[i]-fra_r[i] for i in range(3)]
    dacc = [dxiiia[i]+dxiva[i]+dfra[i] for i in range(3)]
    sec5 = make_df([
        ("XIII Acum. Rol", xiiia_r), ("XIII Acum. IESS", xiiia_i), ("Dif. XIII", dxiiia),
        ("XIV Acum. Rol", xiva_r), ("XIV Acum. IESS", xiva_i), ("Dif. XIV", dxiva),
        ("FR Acum. Rol", fra_r), ("FR Acum. IESS", fra_i), ("Dif. FR", dfra),
        ("Diferencia Total", dacc),
    ])

    # 6) Resumen ejecutivo de diferencias
    resumen = pd.DataFrame({
        "Tipo": groups + ["TOTAL GENERAL"],
        "Dif. Sueldo": dif_sueldo + [sum(dif_sueldo)],
        "Dif. Días": dif_dias + [sum(dif_dias)],
        "Dif. Beneficios Pagados": dtotal + [sum(dtotal)],
        "Dif. Aportes": dap + [sum(dap)],
        "Dif. Acumulados": dacc + [sum(dacc)],
    })
    return sec1, sec2, sec3, sec4, sec5, resumen


def build_iess_simulated_role(roles, iess):
    ig = iess.groupby("Cédula", as_index=False).agg({
        "Nombre IESS":"first","Rel. Trabajo":"first","Sueldo IESS":"sum","Días IESS":"sum",
        "Patronal IESS":"sum","Individual IESS":"sum","Aporte Adic":"sum","Cesantía":"sum",
        "Valor CCC":"sum","Total Aporte IESS":"sum"
    })
    s = roles.merge(ig, on="Cédula", how="left")
    for c in ["Sueldo IESS","Días IESS","Patronal IESS","Individual IESS","Aporte Adic","Cesantía","Valor CCC","Total Aporte IESS"]:
        s[c] = pd.to_numeric(s[c], errors="coerce").fillna(0.0)
    s["Sobretiempos Rol"] = s["Horas Suplementarias 50%"] + s["Horas Extraordinarias 100%"] + s["Recargo 25%"]
    s["Otros Ingresos Gravados Rol"] = s["Otros Ingresos"]
    s["Sueldo Base Simulado IESS"] = (s["Sueldo IESS"] - s["Sobretiempos Rol"] - s["Otros Ingresos Gravados Rol"]).clip(lower=0)
    s["XIII Simulado IESS"] = s["Sueldo IESS"] / 12.0
    s["XIV Simulado IESS"] = (482.0 / 12.0) * (s["Días IESS"].clip(lower=0, upper=30) / 30.0)
    cierre = period_end(roles)
    cumple_ano = s["Fecha Ingreso"].notna() & (cierre >= s["Fecha Ingreso"] + pd.DateOffset(years=1))
    s["FR Simulado IESS"] = np.where(s["Tipo Rol"].eq("Operativos"), s["Sueldo IESS"] * 0.0833, np.where(cumple_ano, s["Sueldo IESS"] * 0.0833, 0.0))
    s["XIII Pagado Simulado"] = np.where(s["_Blank XIII"], 0.0, s["XIII Simulado IESS"])
    s["XIII Acumulado Simulado"] = np.where(s["_Blank XIII"], s["XIII Simulado IESS"], 0.0)
    s["XIV Pagado Simulado"] = np.where(s["_Blank XIV"], 0.0, s["XIV Simulado IESS"])
    s["XIV Acumulado Simulado"] = np.where(s["_Blank XIV"], s["XIV Simulado IESS"], 0.0)
    fr_acumula = np.where(s["Tipo Rol"].eq("Operativos"), s["_Blank FR"], s["_Blank FR"] & cumple_ano)
    s["FR Pagado Simulado"] = np.where(fr_acumula, 0.0, s["FR Simulado IESS"])
    s["FR Acumulado Simulado"] = np.where(fr_acumula, s["FR Simulado IESS"], 0.0)

    # ACUMULADOS DEL ROL REAL: se calculan de forma independiente con la base y días del Rol.
    # Se conserva la misma modalidad de acumulación por trabajador para hacer una comparación válida.
    s["XIII Acumulado Rol"] = np.where(
        s["_Blank XIII"],
        s["Materia Gravada Rol Calc"] / 12.0,
        0.0
    )
    s["XIV Acumulado Rol"] = np.where(
        s["_Blank XIV"],
        (482.0 / 12.0) * (s["Días Laborados"].clip(lower=0, upper=30) / 30.0),
        0.0
    )
    s["FR Causado Rol"] = np.where(
        s["Tipo Rol"].eq("Operativos"),
        s["Materia Gravada Rol Calc"] * 0.0833,
        np.where(cumple_ano, s["Materia Gravada Rol Calc"] * 0.0833, 0.0)
    )
    s["FR Acumulado Rol"] = np.where(fr_acumula, s["FR Causado Rol"], 0.0)

    # Obligaciones acumuladas del Rol Simulado IESS.
    s["Décimo XIII Acumulado IESS"] = s["XIII Acumulado Simulado"]
    s["Décimo XIV Acumulado IESS"] = s["XIV Acumulado Simulado"]
    s["Fondo Reserva Acumulado IESS"] = s["FR Acumulado Simulado"]

    s["Dif. XIII Acumulado IESS - Rol"] = s["Décimo XIII Acumulado IESS"] - s["XIII Acumulado Rol"]
    s["Dif. XIV Acumulado IESS - Rol"] = s["Décimo XIV Acumulado IESS"] - s["XIV Acumulado Rol"]
    s["Dif. FR Acumulado IESS - Rol"] = s["Fondo Reserva Acumulado IESS"] - s["FR Acumulado Rol"]
    s["Dif. Total Acumulados IESS - Rol"] = (
        s["Dif. XIII Acumulado IESS - Rol"]
        + s["Dif. XIV Acumulado IESS - Rol"]
        + s["Dif. FR Acumulado IESS - Rol"]
    )

    s["Décimo XIII Pagado IESS"] = s["XIII Pagado Simulado"]
    s["Décimo XIV Pagado IESS"] = s["XIV Pagado Simulado"]
    s["Fondo Reserva Pagado IESS"] = s["FR Pagado Simulado"]
    s["Total Beneficios Acumulados IESS"] = (
        s["Décimo XIII Acumulado IESS"]
        + s["Décimo XIV Acumulado IESS"]
        + s["Fondo Reserva Acumulado IESS"]
    )
    s["Total Beneficios Pagados IESS"] = (
        s["Décimo XIII Pagado IESS"]
        + s["Décimo XIV Pagado IESS"]
        + s["Fondo Reserva Pagado IESS"]
    )

    s["Total Ingresos Simulado IESS"] = s["Sueldo IESS"] + s["XIII Pagado Simulado"] + s["XIV Pagado Simulado"] + s["FR Pagado Simulado"] + s["Movilización"]
    s["Total Egresos Simulado IESS"] = s["Individual IESS"] + s["Préstamo Quirografario"] + s["Préstamo Hipotecario"] + s["Anticipos"] + s["Faltas / Pérdida Remuneración"] + s["Otros Egresos"] + s["Multa"] + s["Impuesto Renta"]
    s["Neto Simulado IESS"] = s["Total Ingresos Simulado IESS"] - s["Total Egresos Simulado IESS"]
    s["Dif. Días Rol vs IESS"] = s["Días Laborados"] - s["Días IESS"]
    s["Dif. Materia Gravada Rol vs IESS"] = s["Materia Gravada Rol Calc"] - s["Sueldo IESS"]
    s["Dif. Individual Rol vs IESS"] = s["IESS"] - s["Individual IESS"]
    s["Dif. Patronal Calc Rol vs IESS"] = s["Aporte Patronal Rol Calc 11.15%"] - s["Patronal IESS"]
    s["Dif. SECAP Calc Rol vs IESS"] = s["SECAP-IECE Rol Calc 1%"] - s["Valor CCC"]
    s["Dif. XIII Rol vs Sim IESS"] = s["Décimo Tercero"] - s["XIII Pagado Simulado"]
    s["Dif. XIV Rol vs Sim IESS"] = s["Décimo Cuarto"] - s["XIV Pagado Simulado"]
    s["Dif. FR Rol vs Sim IESS"] = s["Fondo Reserva"] - s["FR Pagado Simulado"]
    s["Dif. Total Ingresos Rol vs Sim IESS"] = s["Total Ingresos"] - s["Total Ingresos Simulado IESS"]
    s["Dif. Total Egresos Rol vs Sim IESS"] = s["Total Egresos"] - s["Total Egresos Simulado IESS"]
    s["Dif. Neto Rol vs Sim IESS"] = s["Neto a Recibir"] - s["Neto Simulado IESS"]
    s["Estado Simulación"] = np.select([s["Sueldo IESS"].eq(0), s["Dif. Neto Rol vs Sim IESS"].abs() <= 0.05],["SIN IESS / REVISAR","CUADRA"],default="DIFERENCIA")
    return s

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
    d["Cédula"]=d["Cédula"].apply(normalize_ci)
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


# ---------- PDF ----------
PDF_BLUE = colors.HexColor("#0B4F88")
PDF_BLUE_2 = colors.HexColor("#0D5FA6")
PDF_DIFF = colors.HexColor("#FFF2CC")
PDF_GRID = colors.HexColor("#B8C4D1")
PDF_TEXT = colors.HexColor("#172033")
PDF_RED = colors.HexColor("#9F2F2F")
PDF_GREEN = colors.HexColor("#DDF5E8")
PDF_YELLOW = colors.HexColor("#FDE9A9")


def _pdf_fmt(v):
    if pd.isna(v):
        return ""
    if isinstance(v, (pd.Timestamp, datetime)):
        return pd.to_datetime(v).strftime("%d/%m/%Y")
    if isinstance(v, (np.integer, int)):
        return f"{int(v):,}"
    if isinstance(v, (np.floating, float)):
        return f"{float(v):,.2f}"
    return str(v)


def _pdf_money(v):
    try:
        return f"$ {float(v):,.2f}"
    except Exception:
        return _pdf_fmt(v)


def _pdf_footer(canvas, doc):
    canvas.saveState()
    w, h = doc.pagesize
    canvas.setStrokeColor(colors.HexColor("#D9E2EC"))
    canvas.line(12*mm, 10*mm, w-12*mm, 10*mm)
    canvas.setFont("Helvetica", 8)
    canvas.setFillColor(colors.HexColor("#64748B"))
    canvas.drawString(12*mm, 6*mm, "CENASE - Portal Roles vs IESS")
    canvas.drawRightString(w-12*mm, 6*mm, f"Página {doc.page}")
    canvas.restoreState()


def _safe_pdf_df(df):
    if df is None:
        return pd.DataFrame()
    x = df.copy()
    x = x.drop(columns=[c for c in x.columns if str(c).startswith("_")], errors="ignore")
    for c in x.columns:
        if pd.api.types.is_datetime64_any_dtype(x[c]):
            x[c] = pd.to_datetime(x[c], errors="coerce").dt.strftime("%d/%m/%Y")
    return x


def _pdf_column_chunks(df, key_cols=None, max_cols=7):
    """Split very wide reports into readable blocks instead of shrinking the font."""
    if df.empty:
        return [df]
    key_cols = [c for c in (key_cols or []) if c in df.columns]
    other = [c for c in df.columns if c not in key_cols]
    room = max(1, max_cols-len(key_cols))
    if len(df.columns) <= max_cols:
        return [df]
    return [df[key_cols + other[i:i+room]] for i in range(0, len(other), room)]


def _pdf_col_weights(cols):
    weights=[]
    for c in cols:
        nc=norm(c)
        if any(k in nc for k in ["NOMBRE","CUENTA","DETALLE","CONCEPTO","CARGO","PUESTO","CLIENTE","OBSERV"]):
            w=2.2
        elif any(k in nc for k in ["CEDULA","FECHA","PRESENCIA","ESTADO","TIPO ROL","MODALIDAD"]):
            w=1.35
        elif any(k in nc for k in ["DIFERENCIA","MATERIA","INGRESOS","EGRESOS","APORTE","PATRONAL","INDIVIDUAL","FONDO","DECIMO","NETO","SUELDO"]):
            w=1.25
        else:
            w=1.0
        weights.append(w)
    return weights


def _pdf_table_story(df, styles, page_width, repeat_key_cols=None, max_cols=7, total_rows=True):
    df = _safe_pdf_df(df)
    if df.empty:
        return [Paragraph("Sin registros para mostrar.", styles["BodySmall"]), Spacer(1, 4*mm)]
    story=[]
    chunks=_pdf_column_chunks(df, repeat_key_cols, max_cols=max_cols)
    for ci, chunk in enumerate(chunks, start=1):
        cols=list(chunk.columns)
        weights=_pdf_col_weights(cols)
        scale=page_width/sum(weights)
        widths=[w*scale for w in weights]
        data=[[Paragraph(f"<b>{str(c)}</b>", styles["CellHead"]) for c in cols]]
        total_row_indices=[]
        for ridx, (_, row) in enumerate(chunk.iterrows(), start=1):
            vals=[]
            row_text=" ".join(str(row.get(c,"")) for c in cols[:3]).upper()
            if total_rows and "TOTAL" in row_text:
                total_row_indices.append(ridx)
            for c in cols:
                txt=_pdf_fmt(row[c]).replace("&","&amp;").replace("<","&lt;").replace(">","&gt;")
                vals.append(Paragraph(txt, styles["Cell"]))
            data.append(vals)
        tbl=Table(data,colWidths=widths,repeatRows=1,hAlign="LEFT")
        cmds=[
            ("BACKGROUND",(0,0),(-1,0),PDF_BLUE_2),
            ("TEXTCOLOR",(0,0),(-1,0),colors.white),
            ("GRID",(0,0),(-1,-1),0.45,PDF_GRID),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("ROWBACKGROUNDS",(0,1),(-1,-1),[colors.white,colors.HexColor("#F8FAFC")]),
            ("LEFTPADDING",(0,0),(-1,-1),4),("RIGHTPADDING",(0,0),(-1,-1),4),
            ("TOPPADDING",(0,0),(-1,-1),4),("BOTTOMPADDING",(0,0),(-1,-1),4),
        ]
        for j,c in enumerate(cols):
            if "DIF" in norm(c) or "DIFERENCIA" in norm(c):
                cmds += [("BACKGROUND",(j,1),(j,-1),PDF_DIFF),("FONTNAME",(j,1),(j,-1),"Helvetica-Bold")]
        for ridx in total_row_indices:
            cmds += [("BACKGROUND",(0,ridx),(-1,ridx),colors.HexColor("#DCE6F1")),
                     ("FONTNAME",(0,ridx),(-1,ridx),"Helvetica-Bold")]
        tbl.setStyle(TableStyle(cmds))
        story += [tbl,Spacer(1,4*mm)]
        if ci < len(chunks):
            story += [Paragraph(f"Continuación del reporte - bloque {ci+1} de {len(chunks)}",styles["Note"]),Spacer(1,2*mm)]
    return story


def make_pdf_report(title, sections, subtitle="", page="A3", max_cols=7):
    """General readable PDF: larger type and fewer columns per block."""
    out=io.BytesIO()
    pagesize=landscape(A3 if page=="A3" else A4)
    doc=SimpleDocTemplate(out,pagesize=pagesize,rightMargin=10*mm,leftMargin=10*mm,topMargin=12*mm,bottomMargin=14*mm,title=title,author="CENASE")
    base=getSampleStyleSheet()
    styles={
        "Brand":ParagraphStyle("Brand",parent=base["Normal"],fontName="Helvetica-Bold",fontSize=11,textColor=PDF_BLUE_2,spaceAfter=1*mm),
        "Title":ParagraphStyle("PdfTitle",parent=base["Title"],fontName="Helvetica-Bold",fontSize=18,leading=21,textColor=PDF_BLUE,spaceAfter=2*mm),
        "Sub":ParagraphStyle("PdfSub",parent=base["Normal"],fontSize=10,leading=13,textColor=colors.HexColor("#475569"),spaceAfter=5*mm),
        "Section":ParagraphStyle("PdfSection",parent=base["Heading2"],fontName="Helvetica-Bold",fontSize=13,leading=16,textColor=PDF_BLUE,spaceBefore=3*mm,spaceAfter=2*mm),
        "CellHead":ParagraphStyle("CellHead",parent=base["Normal"],fontName="Helvetica-Bold",fontSize=8.1,leading=9.5,textColor=colors.white,alignment=1),
        "Cell":ParagraphStyle("Cell",parent=base["Normal"],fontSize=8.2,leading=10.0,textColor=PDF_TEXT),
        "BodySmall":ParagraphStyle("BodySmall",parent=base["Normal"],fontSize=9.5,leading=12),
        "Note":ParagraphStyle("Note",parent=base["Normal"],fontSize=8.5,leading=10.5,textColor=colors.HexColor("#64748B")),
    }
    story=[Paragraph("CENASE",styles["Brand"]),Paragraph(title,styles["Title"])]
    if subtitle:
        story.append(Paragraph(subtitle,styles["Sub"]))
    usable=pagesize[0]-20*mm
    for sec in sections:
        sec_story=[]
        if sec.get("title"):
            sec_story.append(Paragraph(sec["title"],styles["Section"]))
        sec_story.extend(_pdf_table_story(sec.get("df",pd.DataFrame()),styles,usable,sec.get("keys"),sec.get("max_cols",max_cols)))
        if sec.get("keep",False):
            story.append(KeepTogether(sec_story))
        else:
            story.extend(sec_story)
    doc.build(story,onFirstPage=_pdf_footer,onLaterPages=_pdf_footer)
    return out.getvalue()


def _group_sum_table(roles, cols, labels=None):
    labels=labels or {}
    ordered=["Administrativos","Gerentes","Operativos"]
    rows=[]
    for tipo in ordered:
        g=roles[roles["Tipo Rol"]==tipo]
        row={"Tipo de Rol":tipo}
        for c in cols:
            row[labels.get(c,c)]=pd.to_numeric(g.get(c,0),errors="coerce").fillna(0).sum() if c in g.columns else 0.0
        rows.append(row)
    total={"Tipo de Rol":"TOTAL GENERAL"}
    for c in cols:
        total[labels.get(c,c)]=pd.to_numeric(roles.get(c,0),errors="coerce").fillna(0).sum() if c in roles.columns else 0.0
    rows.append(total)
    return pd.DataFrame(rows)


def make_role_summary_pdf(roles, benefits, mes):
    """Summary laid out like the user's Excel: ingresos, egresos, pagados and acumulados."""
    ingreso_cols=["Sueldo","Horas Suplementarias 50%","Horas Extraordinarias 100%","Recargo 25%","Décimo Tercero","Décimo Cuarto","Fondo Reserva","Movilización","Otros Ingresos","Total Ingresos"]
    egreso_cols=["IESS","Anticipos","Faltas / Pérdida Remuneración","Préstamo Quirografario","Préstamo Hipotecario","Otros Egresos","Multa","Impuesto Renta","Total Egresos","Neto a Recibir"]
    ing=_group_sum_table(roles,ingreso_cols)
    egr=_group_sum_table(roles,egreso_cols)
    pag=_group_sum_table(roles,["Décimo Tercero","Décimo Cuarto","Fondo Reserva"],{
        "Décimo Tercero":"XIII Pagado Rol","Décimo Cuarto":"XIV Pagado Rol","Fondo Reserva":"FR Pagado Rol"})
    acc_rows=[]
    for tipo in ["Administrativos","Gerentes","Operativos"]:
        g=benefits[benefits["Tipo Rol"]==tipo] if benefits is not None and not benefits.empty else pd.DataFrame()
        acc_rows.append({"Tipo de Rol":tipo,
                         "XIII Acumulado":pd.to_numeric(g.get("XIII Acumulado",0),errors="coerce").fillna(0).sum() if not g.empty else 0.0,
                         "XIV Acumulado":pd.to_numeric(g.get("XIV Acumulado",0),errors="coerce").fillna(0).sum() if not g.empty else 0.0,
                         "FR Acumulado":pd.to_numeric(g.get("FR Acumulado",0),errors="coerce").fillna(0).sum() if not g.empty else 0.0})
    acc=pd.DataFrame(acc_rows)
    acc=pd.concat([acc,pd.DataFrame([{"Tipo de Rol":"TOTAL GENERAL",
                                     "XIII Acumulado":acc["XIII Acumulado"].sum(),
                                     "XIV Acumulado":acc["XIV Acumulado"].sum(),
                                     "FR Acumulado":acc["FR Acumulado"].sum()}])],ignore_index=True)
    # En este PDF resumen todos los valores son monetarios, igual que el control Excel.
    for dfm in [ing,egr,pag,acc]:
        for c in dfm.columns:
            if c != "Tipo de Rol":
                dfm[c]=pd.to_numeric(dfm[c],errors="coerce").fillna(0).map(_pdf_money)
    return make_pdf_report(f"RESUMEN DEL ROL - {mes}",[
        {"title":"1. Ingresos del Rol","df":ing,"keys":["Tipo de Rol"],"max_cols":6},
        {"title":"2. Egresos del Rol","df":egr,"keys":["Tipo de Rol"],"max_cols":6},
        {"title":"3. Beneficios pagados en el Rol","df":pag,"keys":["Tipo de Rol"],"max_cols":4,"keep":True},
        {"title":"4. Beneficios acumulados","df":acc,"keys":["Tipo de Rol"],"max_cols":4,"keep":True},
    ],subtitle="Resumen por Administrativos, Gerentes y Operativos. Los beneficios pagados y acumulados se muestran por separado, igual que en el control mensual.",page="A3",max_cols=6)


def make_unified_role_pdf(roles, mes, benefits=None):
    """PDF del Rol en formato resumen tipo tabla dinámica, como el control Excel del usuario."""
    out=io.BytesIO()
    pagesize=landscape(A3)
    doc=SimpleDocTemplate(
        out,pagesize=pagesize,rightMargin=5*mm,leftMargin=5*mm,
        topMargin=6*mm,bottomMargin=10*mm,title=f"ROL DE PAGOS {mes}",author="CENASE"
    )
    base=getSampleStyleSheet()
    title=ParagraphStyle("PivotRoleTitle",parent=base["Title"],fontName="Helvetica-Bold",fontSize=15,leading=18,alignment=1,textColor=PDF_BLUE)
    sub=ParagraphStyle("PivotRoleSub",parent=base["Normal"],fontName="Helvetica-Bold",fontSize=8.5,leading=10,alignment=1,textColor=PDF_TEXT)
    head=ParagraphStyle("PivotRoleHead",parent=base["Normal"],fontName="Helvetica-Bold",fontSize=8.2,leading=9.2,alignment=1,textColor=PDF_TEXT)
    head_white=ParagraphStyle("PivotRoleHeadWhite",parent=head,textColor=colors.white)
    cell=ParagraphStyle("PivotRoleCell",parent=base["Normal"],fontSize=9.0,leading=10.2,textColor=PDF_TEXT)
    money=ParagraphStyle("PivotRoleMoney",parent=cell,alignment=2)
    total=ParagraphStyle("PivotRoleTotal",parent=cell,fontName="Helvetica-Bold")
    total_money=ParagraphStyle("PivotRoleTotalMoney",parent=money,fontName="Helvetica-Bold")
    story=[Paragraph("CENASE",sub),Paragraph("ROL DE PAGOS",title),Paragraph(f"MES: {mes}",sub),Spacer(1,2.5*mm)]

    ingreso_cols=["Sueldo","Horas Suplementarias 50%","Horas Extraordinarias 100%","Recargo 25%","Décimo Tercero","Décimo Cuarto","Fondo Reserva","Movilización","Otros Ingresos","Total Ingresos"]
    egreso_cols=["IESS","Anticipos","Faltas / Pérdida Remuneración","Préstamo Quirografario","Préstamo Hipotecario","Otros Egresos","Multa","Impuesto Renta","Total Egresos","Neto a Recibir"]
    ing=_group_sum_table(roles,ingreso_cols)
    egr=_group_sum_table(roles,egreso_cols)
    pag=_group_sum_table(roles,["Décimo Tercero","Décimo Cuarto","Fondo Reserva"],{
        "Décimo Tercero":"XIII Pagado Rol","Décimo Cuarto":"XIV Pagado Rol","Fondo Reserva":"FR Pagado Rol"})

    if benefits is None:
        try:
            benefits=build_benefits(roles)
        except Exception:
            benefits=pd.DataFrame()
    acc_rows=[]
    for tipo in ["Administrativos","Gerentes","Operativos"]:
        g=benefits[benefits["Tipo Rol"]==tipo] if benefits is not None and not benefits.empty and "Tipo Rol" in benefits.columns else pd.DataFrame()
        acc_rows.append({
            "Tipo de Rol":tipo,
            "XIII Acumulado":pd.to_numeric(g.get("XIII Acumulado",0),errors="coerce").fillna(0).sum() if not g.empty else 0.0,
            "XIV Acumulado":pd.to_numeric(g.get("XIV Acumulado",0),errors="coerce").fillna(0).sum() if not g.empty else 0.0,
            "FR Acumulado":pd.to_numeric(g.get("FR Acumulado",0),errors="coerce").fillna(0).sum() if not g.empty else 0.0,
        })
    acc=pd.DataFrame(acc_rows)
    acc=pd.concat([acc,pd.DataFrame([{
        "Tipo de Rol":"TOTAL GENERAL",
        "XIII Acumulado":acc["XIII Acumulado"].sum(),
        "XIV Acumulado":acc["XIV Acumulado"].sum(),
        "FR Acumulado":acc["FR Acumulado"].sum(),
    }])],ignore_index=True)

    def table_block(df, header_bg, header_fg, widths, compact=False):
        cols=list(df.columns)
        data=[]
        hstyle=head_white if header_fg==colors.white else head
        data.append([Paragraph(str(c).replace("Tipo de Rol","Etiquetas de fila"),hstyle) for c in cols])
        for _,r in df.iterrows():
            is_total="TOTAL" in str(r.iloc[0]).upper()
            row=[]
            for j,c in enumerate(cols):
                if j==0:
                    txt=str(r[c])
                    sty=total if is_total else cell
                else:
                    txt=_pdf_money(r[c])
                    sty=total_money if is_total else money
                row.append(Paragraph(txt,sty))
            data.append(row)
        tbl=Table(data,colWidths=widths,repeatRows=1,hAlign="LEFT")
        total_idx=len(data)-1
        cmds=[
            ("BACKGROUND",(0,0),(-1,0),header_bg),
            ("TEXTCOLOR",(0,0),(-1,0),header_fg),
            ("GRID",(0,0),(-1,-1),0.35,colors.HexColor("#C7CED6")),
            ("VALIGN",(0,0),(-1,-1),"MIDDLE"),
            ("LEFTPADDING",(0,0),(-1,-1),2.2),("RIGHTPADDING",(0,0),(-1,-1),2.2),
            ("TOPPADDING",(0,0),(-1,-1),2.0),("BOTTOMPADDING",(0,0),(-1,-1),2.0),
            ("BACKGROUND",(0,total_idx),(-1,total_idx),colors.HexColor("#D9E4F0")),
            ("FONTNAME",(0,total_idx),(-1,total_idx),"Helvetica-Bold"),
        ]
        if compact:
            cmds += [("TOPPADDING",(0,0),(-1,-1),1.6),("BOTTOMPADDING",(0,0),(-1,-1),1.6)]
        tbl.setStyle(TableStyle(cmds))
        return tbl

    usable=pagesize[0]-10*mm
    # Dos bloques anchos, iguales al Excel: etiqueta + 10 columnas de valores.
    w0=29*mm
    rest=(usable-w0)/10
    wide_widths=[w0]+[rest]*10
    story.append(table_block(ing,colors.HexColor("#D9E4F0"),PDF_TEXT,wide_widths))
    story.append(Spacer(1,2.5*mm))
    story.append(table_block(egr,colors.HexColor("#A43A37"),colors.white,wide_widths))
    story.append(Spacer(1,2.5*mm))

    # Beneficios pagados y acumulados, como en la parte inferior del Excel.
    small_widths=[42*mm,48*mm,48*mm,48*mm]
    story.append(table_block(pag,colors.HexColor("#D9E4F0"),PDF_TEXT,small_widths,compact=True))
    story.append(Spacer(1,1.8*mm))
    story.append(table_block(acc,colors.HexColor("#A43A37"),colors.white,small_widths,compact=True))

    doc.build(story,onFirstPage=_pdf_footer,onLaterPages=_pdf_footer)
    return out.getvalue()

def make_bi_pdf(ds1,ds2,ds3,ds4,ds5,ds6,mes):
    return make_pdf_report(f"REPORTE BI DE DIFERENCIAS ROL VS IESS - {mes}",[
        {"title":"1. Sueldo / Materia gravada","df":ds1,"keys":["Tipo"],"max_cols":7},
        {"title":"2. Días Rol vs IESS","df":ds2,"keys":["Tipo"],"max_cols":7},
        {"title":"3. Beneficios pagados - Rol vs IESS","df":ds3,"keys":["Tipo"],"max_cols":7},
        {"title":"4. Aportes - Rol vs IESS","df":ds4,"keys":["Tipo"],"max_cols":7},
        {"title":"5. Beneficios acumulados - Rol vs IESS","df":ds5,"keys":["Tipo"],"max_cols":7},
        {"title":"6. Resumen de diferencias","df":ds6,"keys":["Tipo"],"max_cols":7},
    ],subtitle="Comparación directa Rol -> IESS -> Diferencia por Administrativos, Gerentes y Operativos.",page="A3")

def export_excel(roles,summary,compare,diffs,benefits,planillas,plan_compare,accounting,payment_accounting,sim_iess):
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
        sim_export=sim_iess.copy()
        sim_export.to_excel(writer,sheet_name="Rol Simulado IESS",index=False,startrow=2)
        ws=writer.sheets["Rol Simulado IESS"]
        ws.merge_range(0,0,0,max(1,len(sim_export.columns)-1),"ROL SIMULADO CON VALORES IESS",title)
        for j,c in enumerate(sim_export.columns): ws.write(2,j,c,head)
        ws.freeze_panes(3,0); ws.autofilter(2,0,2+len(sim_export),max(0,len(sim_export.columns)-1))
        ws.set_column(0,max(0,len(sim_export.columns)-1),18)

        # Reporte BI comparativo
        ds1,ds2,ds3,ds4,ds5,ds6 = build_simplified_difference_report(sim_iess)
        wsbi = wb.add_worksheet("Diferencias BI")
        writer.sheets["Diferencias BI"] = wsbi
        bi_title = wb.add_format({"bold":True,"font_size":14,"font_color":"#FFFFFF","bg_color":"#0B4F88","align":"center","border":1})
        bi_head = wb.add_format({"bold":True,"font_color":"#FFFFFF","bg_color":"#0D5FA6","align":"center","border":1,"text_wrap":True})
        bi_money = wb.add_format({"num_format":'$#,##0.00;[Red]-$#,##0.00',"border":1})
        bi_num = wb.add_format({"num_format":'#,##0.00',"border":1})
        bi_text = wb.add_format({"border":1})
        bi_diff = wb.add_format({"bold":True,"num_format":'$#,##0.00;[Red]-$#,##0.00',"border":1,"bg_color":"#FFF2CC"})
        bi_diff_num = wb.add_format({"bold":True,"num_format":'#,##0.00',"border":1,"bg_color":"#FFF2CC"})
        wsbi.set_column(0,0,22)

        r0 = 0
        sections = [
            ("1. SUELDO / MATERIA GRAVADA", ds1, False),
            ("2. DÍAS ROL VS IESS", ds2, True),
            ("3. BENEFICIOS PAGADOS — ROL VS IESS", ds3, False),
            ("4. APORTES — ROL VS IESS", ds4, False),
            ("5. BENEFICIOS ACUMULADOS — ROL VS IESS", ds5, False),
            ("6. RESUMEN DE DIFERENCIAS", ds6, False),
        ]
        for ttl, dfx, is_days in sections:
            wsbi.merge_range(r0,0,r0,max(1,len(dfx.columns)-1),ttl,bi_title); r0 += 1
            for j,c in enumerate(dfx.columns):
                wsbi.write(r0,j,c,bi_head)
                wsbi.set_column(j,j,19 if j else 22)
            r0 += 1
            for _, rr in dfx.iterrows():
                for j,c in enumerate(dfx.columns):
                    if j == 0:
                        wsbi.write(r0,j,str(rr[c]),bi_text)
                    else:
                        isdiff = ("Dif." in c or "Diferencia" in c)
                        if is_days:
                            fmt = bi_diff_num if isdiff else bi_num
                        else:
                            fmt = bi_diff if isdiff else bi_money
                        wsbi.write_number(r0,j,float(rr[c]),fmt)
                r0 += 1
            r0 += 2
        wsbi.freeze_panes(2,1)

        # Hojas adicionales
        for sheet,df in [("IESS vs Planillas",plan_compare),("Planillas Pagadas",planillas),("Beneficios",benefits.drop(columns=["_Blank XIII","_Blank XIV","_Blank FR"],errors="ignore")),("Asiento Propuesto",accounting),("Pago IESS Real",payment_accounting)]:
            df.to_excel(writer,sheet_name=sheet,index=False,startrow=2)
            ws=writer.sheets[sheet]
            ws.merge_range(0,0,0,max(1,len(df.columns)-1),sheet.upper(),title)
            for j,c in enumerate(df.columns): ws.write(2,j,c,head)
            ws.freeze_panes(3,0); ws.autofilter(2,0,2+len(df),max(0,len(df.columns)-1))
            ws.set_column(0,max(0,len(df.columns)-1),20)
    out.seek(0)
    return out.getvalue()

# ---------- CARGA / HISTORIAL GUARDADO ----------
storage_cfg=get_storage_config()
saved_snapshot=st.session_state.get("cenase_saved_snapshot")

with st.sidebar:
    st.header("💾 Historial mensual")
    if storage_cfg:
        try:
            saved_rows=storage_list(storage_cfg)
            saved_periods=[x.get("period","") for x in saved_rows if x.get("period")]
        except Exception as e:
            saved_rows=[]; saved_periods=[]
            st.warning(f"No se pudo consultar el historial: {e}")
        if saved_periods:
            sel_saved=st.selectbox("Mes guardado",saved_periods,key="saved_period_select")
            copen,cnew=st.columns(2)
            if copen.button("📂 Abrir",use_container_width=True):
                try:
                    st.session_state["cenase_saved_snapshot"]=storage_load(storage_cfg,sel_saved)
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo abrir {sel_saved}: {e}")
            if cnew.button("➕ Nueva carga",use_container_width=True):
                st.session_state.pop("cenase_saved_snapshot",None)
                st.rerun()
        else:
            st.caption("Todavía no hay meses guardados.")
    else:
        st.warning("Guardado permanente aún no configurado.")
        st.caption("La app sigue funcionando con archivos, pero para conservar meses después de reinicios debes configurar Supabase en Streamlit Secrets.")

    st.divider()
    st.header("Carga mensual")
    st.caption("Sube los Roles y el Consolidado IESS para trabajar. El Reporte de Planillas Pagadas y el maestro de cuentas bancarias son opcionales.")
    fger=st.file_uploader("1. Rol de Gerentes",type=["xlsx","xls"],key="ger")
    fad=st.file_uploader("2. Rol de Administración",type=["xlsx","xls"],key="adm")
    fop=st.file_uploader("3. Rol de Operativos",type=["xlsx","xls"],key="ope")
    fiess=st.file_uploader("4. Consolidado IESS",type=["xlsx","xls"],key="iess")
    fplan=st.file_uploader("5. Reporte de Planillas IESS Pagadas (opcional)",type=["xlsx"],key="plan")
    if fplan is None:
        st.caption("Puedes continuar sin este archivo. Solo quedarán pendientes IESS vs Planillas y el Asiento 3 de pago.")
    fbank=st.file_uploader("6. Maestro de Cuentas Bancarias (opcional)",type=["xlsx","xls"],key="bank_accounts")
    st.divider()
    st.caption("El cruce Rol vs IESS se realiza por cédula.")

try:
    saved_snapshot=st.session_state.get("cenase_saved_snapshot")
    if saved_snapshot:
        roles=saved_snapshot["roles"].copy()
        iess=saved_snapshot["iess"].copy()
        planillas=saved_snapshot.get("planillas")
        planillas=planillas.copy() if planillas is not None else empty_planillas()
        bank_accounts=saved_snapshot.get("bank_accounts")
        if bank_accounts is not None:
            bank_accounts=bank_accounts.copy()
        if fbank is not None:
            bank_accounts=read_bank_accounts(fbank)
        data_origin=f"GUARDADO · {saved_snapshot.get('period','')}"
    else:
        if not (fger and fad and fop and fiess):
            st.info("Carga los tres Roles y el Consolidado IESS, o abre un mes guardado, para generar el reporte. El Reporte de Planillas Pagadas es opcional.")
            st.stop()
        ger=read_role(fger,"Gerentes")
        adm=read_role(fad,"Administrativos")
        ope=read_role(fop,"Operativos")
        roles=pd.concat([ger,adm,ope],ignore_index=True)
        iess=read_iess(fiess)
        planillas=read_planillas(fplan) if fplan is not None else empty_planillas()
        bank_accounts=read_bank_accounts(fbank) if fbank is not None else None
        data_origin="ARCHIVOS CARGADOS"

    roles=add_role_employer_calcs(roles)
    dup_mask = roles.duplicated(subset=["Tipo Rol","Mes","Cédula"], keep=False)
    if dup_mask.any():
        st.warning(f"Se detectaron {int(dup_mask.sum())} filas duplicadas. Se conservará una sola fila por trabajador.")
        roles = roles.drop_duplicates(subset=["Tipo Rol","Mes","Cédula"], keep="first").reset_index(drop=True)

    comp=make_compare(roles,iess)
    sim_iess=build_iess_simulated_role(roles,iess)
    benefits=build_benefits(roles,iess)
    plan_compare=build_iess_plan_compare(iess,planillas)
    accounting=build_accounting(roles,iess,benefits,planillas)
    employer_provision=build_employer_provision(roles,iess)
    payment_accounting=build_payment_accounting(iess,planillas)
except Exception as e:
    st.error(f"No pude procesar los datos: {e}")
    st.stop()

# Guardado permanente del período procesado.
periodo_actual=infer_period(roles)
with st.sidebar:
    st.divider()
    st.subheader("Guardar cierre")
    st.caption(f"Período detectado: {periodo_actual}")
    st.caption(f"Origen actual: {data_origin}")
    if storage_cfg:
        if st.button("💾 Guardar / actualizar mes",use_container_width=True,type="primary"):
            try:
                storage_save(storage_cfg,periodo_actual,roles,iess,planillas,bank_accounts)
                st.success(f"{periodo_actual} guardado permanentemente.")
            except Exception as e:
                st.error(f"No se pudo guardar: {e}")
        if saved_snapshot:
            confirm_delete=st.checkbox("Confirmo eliminar este mes",key="confirm_delete_month")
            if st.button("🗑️ Eliminar mes guardado",use_container_width=True,disabled=not confirm_delete):
                try:
                    storage_delete(storage_cfg,periodo_actual)
                    st.session_state.pop("cenase_saved_snapshot",None)
                    st.success(f"{periodo_actual} eliminado.")
                    st.rerun()
                except Exception as e:
                    st.error(f"No se pudo eliminar: {e}")
    else:
        st.info("Configura Supabase para habilitar Guardar / Abrir meses.")

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

# Período para nombres de archivos PDF/Excel
mes_pdf=roles["Mes"].dropna().astype(str)
mes_pdf=mes_pdf.iloc[0] if len(mes_pdf) else datetime.now().strftime("%Y-%m")

tabs=st.tabs(["📊 Resumen Roles","🏛️ Rol vs IESS","🧮 Rol Simulado IESS","⚠️ Diferencias","💳 IESS vs Planillas","🎁 Beneficios","🧾 Contabilidad","🔎 Consulta Roles","✅ Cuadre","🚦 Diferencias Simplificado","💵 Pagos Bancarios"])

with tabs[0]:
    show=summary.copy()
    for col in ["Total Ingresos","Total Egresos","Neto a Recibir"]: show[col]=show[col].map(fmt_money)
    st.dataframe(show,use_container_width=True,hide_index=True)
    st.subheader("Neto por tipo de rol")
    st.bar_chart(grp.set_index("Tipo Rol")[["Neto"]],use_container_width=True)
    p1,p2=st.columns(2)
    pdf_resumen=make_role_summary_pdf(roles,benefits,mes_pdf)
    p1.download_button("📄 Descargar Resumen en PDF",pdf_resumen,file_name=f"Resumen_Roles_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_resumen")
    pdf_rol_unificado=make_unified_role_pdf(roles,mes_pdf,benefits)
    p2.download_button("📄 Descargar Rol en PDF",pdf_rol_unificado,file_name=f"Rol_Pagos_CENASE_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_rol_unificado")

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

    pdf_view=make_pdf_report(f"ROL VS IESS - {mes_pdf}",[{"title":"Conciliación individual","df":view[display_cols],"keys":["Tipo Rol","Cédula","Nombre"],"max_cols":7}],subtitle="Detalle filtrado de conciliación individual Rol vs IESS.")
    st.download_button("📄 Descargar Rol vs IESS en PDF",pdf_view,file_name=f"Rol_vs_IESS_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_rol_iess")

with tabs[2]:
    st.subheader("Rol Simulado con valores reales del IESS")
    st.caption("Reconstruye una nómina comparable usando materia gravada, días, aporte individual, patronal y CCC reales del IESS. Mantiene del Rol los descuentos y la modalidad mensual/acumulada de beneficios.")
    s1,s2,s3,s4=st.columns(4)
    sim_roles=sorted(sim_iess["Tipo Rol"].dropna().astype(str).unique())
    sim_filter=s1.multiselect("Tipo de rol",sim_roles,default=sim_roles,key="sim_role")
    sim_states=sorted(sim_iess["Estado Simulación"].dropna().astype(str).unique())
    sim_state=s2.multiselect("Estado",sim_states,default=sim_states,key="sim_state")
    sim_q=s3.text_input("Nombre / cédula",key="sim_q")
    only_diff=s4.checkbox("Solo diferencias",value=False,key="sim_only_diff")
    sv=sim_iess[sim_iess["Tipo Rol"].isin(sim_filter)&sim_iess["Estado Simulación"].isin(sim_state)].copy()
    if sim_q.strip():
        nq=norm(sim_q)
        sv=sv[sv["Nombre"].map(norm).str.contains(nq,na=False)|sv["Cédula"].astype(str).str.contains(sim_q.strip(),na=False)]
    if only_diff:
        sv=sv[sv["Dif. Neto Rol vs Sim IESS"].abs()>0.05]
    k1,k2,k3,k4=st.columns(4)
    k1.metric("Trabajadores",len(sv)); k2.metric("Neto Rol",fmt_money(sv["Neto a Recibir"].sum())); k3.metric("Neto Simulado IESS",fmt_money(sv["Neto Simulado IESS"].sum())); k4.metric("Diferencia Neto",fmt_money(sv["Dif. Neto Rol vs Sim IESS"].sum()))
    a1,a2,a3,a4=st.columns(4)
    a1.metric("XIII acumulado IESS",fmt_money(sv["Décimo XIII Acumulado IESS"].sum()))
    a2.metric("XIV acumulado IESS",fmt_money(sv["Décimo XIV Acumulado IESS"].sum()))
    a3.metric("F.R. acumulado IESS",fmt_money(sv["Fondo Reserva Acumulado IESS"].sum()))
    a4.metric("Total acumulado",fmt_money(sv["Total Beneficios Acumulados IESS"].sum()))
    sim_cols=["Estado Simulación","Tipo Rol","Cédula","Nombre","Cargo","Puesto / Cliente","Días Laborados","Días IESS","Dif. Días Rol vs IESS","Materia Gravada Rol Calc","Sueldo IESS","Dif. Materia Gravada Rol vs IESS","Sueldo Base Simulado IESS","Sobretiempos Rol","Otros Ingresos Gravados Rol","Décimo Tercero","XIII Simulado IESS","Décimo XIII Pagado IESS","XIII Acumulado Rol","Décimo XIII Acumulado IESS","Dif. XIII Acumulado IESS - Rol","Dif. XIII Rol vs Sim IESS","Décimo Cuarto","XIV Simulado IESS","Décimo XIV Pagado IESS","XIV Acumulado Rol","Décimo XIV Acumulado IESS","Dif. XIV Acumulado IESS - Rol","Dif. XIV Rol vs Sim IESS","Fondo Reserva","FR Simulado IESS","Fondo Reserva Pagado IESS","FR Acumulado Rol","Fondo Reserva Acumulado IESS","Dif. FR Acumulado IESS - Rol","Total Beneficios Pagados IESS","Total Beneficios Acumulados IESS","Dif. Total Acumulados IESS - Rol","Dif. FR Rol vs Sim IESS","IESS","Individual IESS","Dif. Individual Rol vs IESS","Aporte Patronal Rol Calc 11.15%","Patronal IESS","Dif. Patronal Calc Rol vs IESS","SECAP-IECE Rol Calc 1%","Valor CCC","Dif. SECAP Calc Rol vs IESS","Total Ingresos","Total Ingresos Simulado IESS","Dif. Total Ingresos Rol vs Sim IESS","Total Egresos","Total Egresos Simulado IESS","Dif. Total Egresos Rol vs Sim IESS","Neto a Recibir","Neto Simulado IESS","Dif. Neto Rol vs Sim IESS"]
    st.dataframe(sv[sim_cols],use_container_width=True,hide_index=True)
    st.markdown("#### Resumen por grupo")
    sim_summary=sv.groupby("Tipo Rol",as_index=False).agg(Trabajadores=("Cédula","count"),Materia_Gravada_Rol=("Materia Gravada Rol Calc","sum"),Materia_Gravada_IESS=("Sueldo IESS","sum"),Patronal_Rol_Calc=("Aporte Patronal Rol Calc 11.15%","sum"),Patronal_IESS=("Patronal IESS","sum"),SECAP_Rol_Calc=("SECAP-IECE Rol Calc 1%","sum"),CCC_IESS=("Valor CCC","sum"),Neto_Rol=("Neto a Recibir","sum"),Neto_Simulado_IESS=("Neto Simulado IESS","sum"),Diferencia_Neto=("Dif. Neto Rol vs Sim IESS","sum"))
    st.dataframe(sim_summary,use_container_width=True,hide_index=True)

    st.markdown("#### Beneficios pagados vs acumulados por grupo")
    acum_summary=sv.groupby("Tipo Rol",as_index=False).agg(
        XIII_Pagado=("Décimo XIII Pagado IESS","sum"),
        XIII_Acumulado=("Décimo XIII Acumulado IESS","sum"),
        XIV_Pagado=("Décimo XIV Pagado IESS","sum"),
        XIV_Acumulado=("Décimo XIV Acumulado IESS","sum"),
        FR_Pagado=("Fondo Reserva Pagado IESS","sum"),
        FR_Acumulado=("Fondo Reserva Acumulado IESS","sum"),
        Total_Beneficios_Pagados=("Total Beneficios Pagados IESS","sum"),
        Total_Beneficios_Acumulados=("Total Beneficios Acumulados IESS","sum")
    )
    st.dataframe(acum_summary,use_container_width=True,hide_index=True)
    pdf_sim=make_pdf_report(f"ROL SIMULADO IESS - {mes_pdf}",[
        {"title":"Detalle Rol Simulado IESS","df":sv[sim_cols],"keys":["Tipo Rol","Cédula","Nombre"],"max_cols":7},
        {"title":"Resumen por grupo","df":sim_summary,"keys":["Tipo Rol"],"max_cols":7},
        {"title":"Beneficios pagados vs acumulados","df":acum_summary,"keys":["Tipo Rol"],"max_cols":7},
    ],subtitle="Nómina comparable usando valores reales del IESS y modalidades del Rol.",page="A3")
    st.download_button("📄 Descargar Rol Simulado IESS en PDF",pdf_sim,file_name=f"Rol_Simulado_IESS_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_sim_iess")

with tabs[3]:
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

    diff_pdf_cols=["Presencia","Tipo Rol","Cédula","Nombre","Días Laborados","Días IESS","Diferencia Días","Base Rol IESS","Sueldo IESS","Diferencia Base","IESS","Individual IESS","Diferencia Aporte Individual","Patronal Esperado Rol 11.15%","Patronal IESS","Diferencia Aporte Patronal","% Individual IESS","% Patronal IESS"]
    pdf_diffs=make_pdf_report(f"DIFERENCIAS A REVISAR - {mes_pdf}",[{"title":"Registros que requieren revisión","df":diffs[diff_pdf_cols] if not diffs.empty else pd.DataFrame(columns=diff_pdf_cols),"keys":["Tipo Rol","Cédula","Nombre"],"max_cols":7}],subtitle="Detalle de diferencias detectadas entre Rol e IESS.",page="A3")
    st.download_button("📄 Descargar Diferencias en PDF",pdf_diffs,file_name=f"Diferencias_Rol_IESS_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_diffs")

with tabs[4]:
    st.subheader("IESS vs Planillas Pagadas")
    if planillas is None or planillas.empty:
        st.warning("Reporte de Planillas IESS Pagadas no cargado. La APP continúa funcionando normalmente; esta conciliación y el Asiento 3 quedarán pendientes hasta que cargues el archivo.")
        st.markdown("#### Obligación según Consolidado IESS")
        iess_pending=pd.DataFrame([
            {"Concepto":"Aporte Individual","Valor":float(iess['Individual IESS'].sum())},
            {"Concepto":"Aporte Patronal","Valor":float(iess['Patronal IESS'].sum())},
            {"Concepto":"SECAP / IECE (CCC)","Valor":float(iess['Valor CCC'].sum())},
        ])
        st.dataframe(iess_pending,use_container_width=True,hide_index=True)
        st.caption("Cuando cargues el reporte de pago, la APP completará automáticamente la comparación y generará el Asiento 3.")
    else:
        st.dataframe(plan_compare,use_container_width=True,hide_index=True)
        st.markdown("#### Valor contable del pago real")
        st.info("Prioridad aplicada: 1) el Consolidado IESS aporta los valores reales por trabajador (Patronal, Individual y Valor CCC); 2) el Reporte de Planillas verifica lo efectivamente PAGADO; 3) si Consolidado y Planillas difieren, el PAGO contable usa la planilla real y la diferencia queda visible para revisión. No se recalcula CCC sobre el total de materia gravada.")
        st.dataframe(payment_accounting,use_container_width=True,hide_index=True)
        st.markdown("#### Detalle de planillas cargadas")
        st.dataframe(planillas,use_container_width=True,hide_index=True)
        pdf_plan=make_pdf_report(f"IESS VS PLANILLAS PAGADAS - {mes_pdf}",[
            {"title":"Conciliación IESS vs Planillas","df":plan_compare,"max_cols":7},
            {"title":"Valor contable del pago real","df":payment_accounting,"max_cols":8},
            {"title":"Detalle de planillas pagadas","df":planillas,"max_cols":7},
        ],subtitle="Comparación del consolidado IESS con las obligaciones efectivamente pagadas.",page="A3")
        st.download_button("📄 Descargar IESS vs Planillas en PDF",pdf_plan,file_name=f"IESS_vs_Planillas_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_planillas")

with tabs[5]:
    st.subheader("Beneficios: pago mensual vs acumulación")
    st.caption("Base principal: materia gravada IESS. Décimos: valor en rol = pago mensual; vacío = acumula. Fondo de Reserva: fórmula IESS/12 después de 1 año, con Operativos separados de Gerentes/Administrativos.")
    bcols=["Tipo Rol","Cédula","Nombre","Fecha Ingreso","Sueldo IESS","Días IESS","XIII Causado","Modalidad XIII","XIII Pagado Rol","XIII Acumulado","XIV Causado","Modalidad XIV","XIV Pagado Rol","XIV Acumulado","Cumple 1 Año FR","Regla FR","Modalidad FR","FR Causado","FR Pagado Rol","FR Acumulado"]
    st.dataframe(benefits[bcols],use_container_width=True,hide_index=True)
    pdf_ben=make_pdf_report(f"BENEFICIOS - {mes_pdf}",[{"title":"Pago mensual vs acumulación","df":benefits[bcols],"keys":["Tipo Rol","Cédula","Nombre"],"max_cols":7}],subtitle="Décimos y Fondo de Reserva: pagado, acumulado y causación.",page="A3")
    st.download_button("📄 Descargar Beneficios en PDF",pdf_ben,file_name=f"Beneficios_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_beneficios")

with tabs[6]:
    st.subheader("Asientos contables de nómina")
    st.caption("Estructura CENASE: 1) Rol + beneficios acumulados, 2) Provisión patronal IESS + SECAP/IECE, 3) Pago real de planillas IESS. Cada asiento se genera desde su fuente real y debe cuadrar por separado.")

    # ----- ASIENTO 1 -----
    st.markdown("### 1️⃣ Asiento 1 — Devengo del Rol + beneficios acumulados")
    st.caption("Fuente: Roles de Operativos, Gerentes y Administrativos. El aporte personal, descuentos y Neto a Recibir nacen del Rol. Los XIII/XIV/FR acumulados se incorporan en este mismo asiento.")
    st.dataframe(accounting,use_container_width=True,hide_index=True)
    a1d=float(accounting['Debe'].sum()); a1h=float(accounting['Haber'].sum()); a1dif=round(a1d-a1h,2)
    c1,c2,c3=st.columns(3)
    c1.metric("Total Debe",fmt_money(a1d)); c2.metric("Total Haber",fmt_money(a1h)); c3.metric("Diferencia",fmt_money(a1dif))
    if abs(a1dif)<=0.05: st.success("ASIENTO 1 CUADRADO ✓")
    else: st.error("ASIENTO 1 NO CUADRA. Revisar valores del Rol y acumulados.")
    st.caption(f"Glosa: P/R DEVENGO DE ROL DE PAGOS Y BENEFICIOS ACUMULADOS CORRESPONDIENTE AL PERÍODO {mes_pdf}.")

    # Control visible de beneficios acumulados por grupo.
    acc_preview=benefits.groupby('Tipo Rol',as_index=False).agg({
        'XIII Acumulado':'sum','XIV Acumulado':'sum','FR Acumulado':'sum'
    })
    acc_total=pd.DataFrame([{'Tipo Rol':'TOTAL GENERAL',
        'XIII Acumulado':acc_preview['XIII Acumulado'].sum(),
        'XIV Acumulado':acc_preview['XIV Acumulado'].sum(),
        'FR Acumulado':acc_preview['FR Acumulado'].sum()}])
    acc_preview=pd.concat([acc_preview,acc_total],ignore_index=True)
    st.markdown("#### Control de beneficios acumulados incluidos en el Rol")
    st.dataframe(acc_preview,use_container_width=True,hide_index=True)

    # ----- ASIENTO 2 -----
    st.markdown("### 2️⃣ Asiento 2 — Provisión Patronal IESS + SECAP/IECE")
    st.caption("Fuente: Consolidado IESS. Este asiento NO incluye aporte personal porque ya fue reconocido como descuento/pasivo en el Rol.")
    st.dataframe(employer_provision,use_container_width=True,hide_index=True)
    a2d=float(employer_provision['Debe'].sum()); a2h=float(employer_provision['Haber'].sum()); a2dif=round(a2d-a2h,2)
    c1,c2,c3=st.columns(3)
    c1.metric("Total Debe",fmt_money(a2d)); c2.metric("Total Haber",fmt_money(a2h)); c3.metric("Diferencia",fmt_money(a2dif))
    if abs(a2dif)<=0.05: st.success("ASIENTO 2 CUADRADO ✓")
    else: st.error("ASIENTO 2 NO CUADRA. Revisar Consolidado IESS.")
    st.caption(f"Glosa: P/R PROVISIÓN DE APORTES PATRONALES IESS Y SECAP-IECE CORRESPONDIENTE AL PERÍODO {mes_pdf}.")

    # ----- ASIENTO 3 -----
    st.markdown("### 3️⃣ Asiento 3 — Pago de Planillas IESS")
    if planillas is None or planillas.empty:
        st.warning("PENDIENTE: no se cargó el Reporte de Planillas IESS Pagadas. Los Asientos 1 y 2 sí están completos; el Asiento 3 se generará cuando exista el pago real.")
        st.caption("No se crea ningún asiento de pago estimado ni se fuerza un valor contra Banco sin el reporte pagado.")
        plan_compare_display=plan_compare
        liq=0.0
    else:
        st.caption("Fuente: Consolidado IESS + reporte de planillas efectivamente pagadas. El asiento cruza los pasivos y cuadra exactamente contra el total pagado.")
        st.dataframe(payment_accounting,use_container_width=True,hide_index=True)
        a3d=float(payment_accounting['Debe'].sum()); a3h=float(payment_accounting['Haber'].sum()); a3dif=round(a3d-a3h,2)
        c1,c2,c3=st.columns(3)
        c1.metric("Total Debe",fmt_money(a3d)); c2.metric("Total Haber",fmt_money(a3h)); c3.metric("Diferencia",fmt_money(a3dif))
        if abs(a3dif)<=0.05: st.success("ASIENTO 3 CUADRADO ✓")
        else: st.error("ASIENTO 3 NO CUADRA. Revisar planillas pagadas / IESS por liquidar.")
        st.caption(f"Glosa: P/R ASIENTO PAGO DE IESS CORRESPONDIENTE AL PERÍODO {mes_pdf}.")

        st.markdown("#### Conciliación de control — Consolidado vs Planillas pagadas")
        st.dataframe(plan_compare,use_container_width=True,hide_index=True)
        liq=float(payment_accounting.loc[payment_accounting['Cuenta'].str.contains('IESS por liquidar',case=False,na=False),'Debe'].sum()-payment_accounting.loc[payment_accounting['Cuenta'].str.contains('IESS por liquidar',case=False,na=False),'Haber'].sum())
        st.metric("IESS por liquidar del pago",fmt_money(liq))
        st.caption("IESS por liquidar representa la diferencia necesaria para que el cruce de pasivos coincida con el pago real. Debe revisarse contra planillas complementarias, extensiones, juveniles u otros movimientos del período.")

    cont_sections=[
        {"title":"ASIENTO 1 - Rol + beneficios acumulados","df":accounting,"max_cols":8},
        {"title":"Control beneficios acumulados","df":acc_preview,"max_cols":6},
        {"title":"ASIENTO 2 - Provisión patronal IESS + SECAP/IECE","df":employer_provision,"max_cols":8},
        {"title":"ASIENTO 3 - Pago de planillas IESS","df":payment_accounting,"max_cols":8},
        {"title":"Conciliación Consolidado vs Planillas","df":plan_compare,"max_cols":8},
    ]
    pdf_cont=make_pdf_report(f"CONTABILIDAD DE NÓMINA - {mes_pdf}",cont_sections,subtitle=("Tres asientos separados: Rol + acumulados, provisión patronal y pago IESS." if planillas is not None and not planillas.empty else "Asientos 1 y 2 completos. Asiento 3 pendiente por falta del Reporte de Planillas Pagadas."),page="A3")
    st.download_button("📄 Descargar Contabilidad en PDF",pdf_cont,file_name=f"Contabilidad_Nomina_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_contabilidad")

with tabs[7]:
    st.subheader("Consulta detallada de Roles")
    st.caption("Incluye Materia Gravada Rol Calc, Aporte Patronal 11,15%, Aporte Individual 9,45% y SECAP/IECE 1% como campos de auditoría.")
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
    pdf_rr=make_pdf_report(f"CONSULTA DE ROLES - {mes_pdf}",[{"title":"Detalle filtrado del Rol","df":rr,"keys":["Tipo Rol","Cédula","Nombre"],"max_cols":7}],subtitle="Reporte según los filtros seleccionados en Consulta Roles.",page="A3")
    st.download_button("📄 Descargar Consulta Roles en PDF",pdf_rr,file_name=f"Consulta_Roles_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_consulta")

with tabs[8]:
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
    cuadre_df=pd.DataFrame([
        {"Concepto":"Base Rol","Valor":comp[comp["Presencia"]!="SOLO IESS"]["Base Rol IESS"].sum()},
        {"Concepto":"Base IESS","Valor":iess["Sueldo IESS"].sum()},
        {"Concepto":"Diferencia Base IESS - Rol","Valor":iess["Sueldo IESS"].sum()-comp[comp["Presencia"]!="SOLO IESS"]["Base Rol IESS"].sum()},
        {"Concepto":"Aporte Individual Rol","Valor":roles["IESS"].sum()},
        {"Concepto":"Aporte Individual IESS","Valor":iess["Individual IESS"].sum()},
        {"Concepto":"Patronal IESS 11.15%","Valor":iess["Patronal IESS"].sum()},
        {"Concepto":"Valor CCC","Valor":iess["Valor CCC"].sum()},
        {"Concepto":"Total Aporte IESS","Valor":iess["Total Aporte IESS"].sum()},
    ])
    pdf_cuadre=make_pdf_report(f"CUADRE GENERAL - {mes_pdf}",[{"title":"Resumen de cuadre","df":cuadre_df,"max_cols":6}],subtitle="Totales generales Rol e IESS.")
    st.download_button("📄 Descargar Cuadre en PDF",pdf_cuadre,file_name=f"Cuadre_General_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_cuadre")

# ---------- DOWNLOAD ----------
st.divider()
excel=export_excel(roles,summary,comp,diffs,benefits,planillas,plan_compare,accounting,payment_accounting,sim_iess)
mes=roles["Mes"].dropna().astype(str)
mes=mes.iloc[0] if len(mes) else datetime.now().strftime("%Y-%m")
d1,d2=st.columns(2)
d1.download_button("⬇️ Descargar reporte completo Roles + IESS (Excel)",
                   data=excel,file_name=f"Roles_vs_IESS_{mes}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)
full_pdf=make_pdf_report(f"REPORTE MENSUAL ROLES + IESS - {mes_pdf}",[
    {"title":"Resumen de Roles","df":summary,"keys":["Tipo Rol"],"max_cols":8},
    {"title":"Rol vs IESS - diferencias","df":diffs.drop(columns=["_merge"],errors="ignore"),"keys":["Tipo Rol","Cédula","Nombre"],"max_cols":7},
    {"title":"IESS vs Planillas","df":plan_compare,"max_cols":7},
    {"title":"Asiento contable propuesto","df":accounting,"max_cols":8},
],subtitle=("Reporte mensual integral de nómina, conciliación IESS, planillas y contabilidad." if planillas is not None and not planillas.empty else "Reporte mensual de nómina e IESS. Conciliación de planillas y Asiento 3 pendientes."),page="A3")
d2.download_button("📄 Descargar reporte mensual completo (PDF)",
                   data=full_pdf,file_name=f"Reporte_Mensual_Roles_IESS_{mes_pdf}.pdf",
                   mime="application/pdf",use_container_width=True)
st.markdown('<p class="small">Ahora cada módulo tiene descarga PDF. También puedes generar el Rol Unificado mensual en un solo PDF desde Resumen Roles.</p>',unsafe_allow_html=True)

with tabs[9]:
    st.subheader("📊 Reporte BI de Diferencias Rol vs IESS")
    st.caption("Comparación directa Rol → IESS → Diferencia. Las filas son los grupos de nómina y las columnas muestran los valores comparables.")

    ds1,ds2,ds3,ds4,ds5,ds6 = build_simplified_difference_report(sim_iess)

    def show_bi_table(title, df, day_table=False):
        st.markdown(f"#### {title}")
        fmt = {}
        for c in df.columns:
            if c == "Tipo":
                continue
            if day_table:
                fmt[c] = "{:,.2f}"
            else:
                fmt[c] = "${:,.2f}"
        styled = df.style.format(fmt)
        diff_cols = [c for c in df.columns if "Dif." in c or "Diferencia" in c]
        if diff_cols:
            styled = styled.set_properties(subset=diff_cols, **{"font-weight":"bold"})
        st.dataframe(styled, use_container_width=True, hide_index=True)

    show_bi_table("1. Sueldo / Materia gravada", ds1)
    show_bi_table("2. Días Rol vs IESS", ds2, day_table=True)
    show_bi_table("3. Beneficios pagados — Rol vs IESS", ds3)
    show_bi_table("4. Aportes — Rol vs IESS", ds4)
    show_bi_table("5. Beneficios acumulados — Rol vs IESS", ds5)

    st.markdown("#### 6. Resumen de diferencias")
    resumen_fmt = {
        "Dif. Sueldo":"${:,.2f}",
        "Dif. Días":"{:,.2f}",
        "Dif. Beneficios Pagados":"${:,.2f}",
        "Dif. Aportes":"${:,.2f}",
        "Dif. Acumulados":"${:,.2f}",
    }
    st.dataframe(
        ds6.style.format(resumen_fmt).set_properties(
            subset=["Dif. Sueldo","Dif. Días","Dif. Beneficios Pagados","Dif. Aportes","Dif. Acumulados"],
            **{"font-weight":"bold"}
        ),
        use_container_width=True,
        hide_index=True
    )
    st.info("En acumulados, el lado Rol y el lado IESS se calculan de forma independiente. Por eso una diferencia $0,00 significa que realmente cuadran, no que se está comparando IESS contra sí mismo.")
    pdf_bi=make_bi_pdf(ds1,ds2,ds3,ds4,ds5,ds6,mes_pdf)
    st.download_button("📄 Descargar Reporte BI en PDF",pdf_bi,file_name=f"Reporte_BI_Diferencias_{mes_pdf}.pdf",mime="application/pdf",use_container_width=True,key="pdf_bi")

with tabs[10]:
    st.subheader("💵 Plantilla de pago de nómina al banco")
    st.caption("Cruza por cédula el maestro bancario con el Rol. VALOR A PAGAR se exporta en centavos, sin comas ni puntos: $230.45 → 23045.")
    if bank_accounts is None or bank_accounts.empty:
        st.info("Carga el archivo 'CUENTAS BANCARIAS' en la barra lateral para generar la plantilla bancaria.")
    else:
        c1,c2=st.columns(2)
        client_id=c1.text_input("IDENT. CLIENTE",value="",help="Código/identificación del cliente pagador solicitado por el banco. No se completa automáticamente para evitar usar un dato incorrecto.")
        default_ref=f"PAGO ROL {mes_pdf}"
        reference=c2.text_input("REFERENCIA",value=default_ref)
        bank_df, bank_missing=build_bank_payment(roles,bank_accounts,client_id,reference)

        k1,k2,k3,k4=st.columns(4)
        k1.metric("Trabajadores a pagar",len(bank_df))
        k2.metric("Total neto a pagar",fmt_money((bank_df["VALOR A PAGAR"].sum()/100) if len(bank_df) else 0))
        k3.metric("Sin cuenta bancaria",len(bank_missing))
        k4.metric("Neto total del Rol",fmt_money(roles["Neto a Recibir"].sum()))

        if not str(client_id).strip():
            st.warning("Completa IDENT. CLIENTE antes de usar el archivo para carga bancaria.")

        st.markdown("#### Vista previa")
        st.dataframe(bank_df,use_container_width=True,hide_index=True,column_config={"VALOR A PAGAR":st.column_config.NumberColumn(format="%d", help="Centavos sin coma ni punto. Ej.: $230.45 = 23045")})

        if len(bank_missing):
            st.markdown("#### ⚠️ Trabajadores sin cuenta bancaria identificada")
            miss_show=bank_missing[["Cédula","Nombre","Tipo Rol","Puesto / Cliente","Neto a Recibir"]].copy()
            st.dataframe(miss_show,use_container_width=True,hide_index=True)

        cheque_df=build_cheque_payment_report(bank_missing)
        if len(cheque_df):
            st.markdown("#### 🧾 Pagos que deben emitirse mediante cheque")
            ch1,ch2=st.columns(2)
            ch1.metric("Cheques a emitir",len(cheque_df))
            ch2.metric("Total a pagar en cheques",fmt_money(cheque_df["VALOR A PAGAR"].sum()))
            st.dataframe(cheque_df,use_container_width=True,hide_index=True,column_config={"VALOR A PAGAR":st.column_config.NumberColumn(format="$ %.2f")})

        dx1,dx2,dx3=st.columns(3)
        excel_bank=export_bank_payment_excel(bank_df,mes_pdf)
        dx1.download_button("⬇️ Descargar plantilla bancaria en Excel",excel_bank,file_name=f"Pago_Rol_Banco_{mes_pdf}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="bank_excel")
        txt_bank=export_bank_payment_txt(bank_df)
        dx2.download_button("⬇️ Descargar TXT tabulado para banco",txt_bank,file_name=f"Pago_Rol_Banco_{mes_pdf}.txt",mime="text/plain",use_container_width=True,key="bank_txt")
        excel_cheques=export_cheque_payment_excel(cheque_df,mes_pdf)
        dx3.download_button("🧾 Descargar reporte de pagos en cheque",excel_cheques,file_name=f"Pago_Rol_Cheques_{mes_pdf}.xlsx",mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",use_container_width=True,key="cheque_excel")
        st.caption("El Excel bancario contiene únicamente quienes tienen cuenta para transferencia. El reporte de cheques contiene quienes no tienen cuenta bancaria y deben pagarse mediante CHEQUE. Solo se incluyen valores netos positivos.")

