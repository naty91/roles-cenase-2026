
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
<h1>Reporte Consolidado de Roles + Conciliación IESS · v9</h1>
<p>Gerentes · Administración · Operativos · IESS | Consulta, diferencias, cuadre y descarga mensual</p>
</div>
""", unsafe_allow_html=True)

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

def period_end(roles):
    for v in roles["Mes"].dropna().astype(str):
        m=re.match(r"^(\\d{4})-(\\d{1,2})$",v.strip())
        if m:
            y,mo=map(int,m.groups()); return pd.Timestamp(y,mo,1)+pd.offsets.MonthEnd(0)
    return pd.Timestamp.today()+pd.offsets.MonthEnd(0)

def build_benefits(roles, iess):
    """Beneficios sociales con IESS como base principal y rol como forma de pago.

    Reglas CENASE:
    - XIII causado: materia gravada IESS / 12.
    - XIV causado: SBU / 12, proporcional a días IESS.
    - FR causado: materia gravada IESS / 12 cuando ya cumplió un año.
    - La celda del rol define modalidad: valor = pago mensual; vacío = acumulación.
    - Operativos se mantienen separados de Gerentes y Administrativos para contabilidad.
    """
    cierre=period_end(roles)
    # Consolidar IESS por cédula para no duplicar a una persona si el IESS trae más de una línea.
    ii=iess.groupby('Cédula',as_index=False).agg({
        'Sueldo IESS':'sum','Días IESS':'sum','Patronal IESS':'sum','Individual IESS':'sum'
    })
    b=roles.merge(ii,on='Cédula',how='left')
    for c in ['Sueldo IESS','Días IESS','Patronal IESS','Individual IESS']:
        b[c]=pd.to_numeric(b[c],errors='coerce').fillna(0.0)
    b['Fecha Corte']=cierre
    b['Cumple 1 Año FR']=b['Fecha Ingreso'].notna() & (cierre >= b['Fecha Ingreso'] + pd.DateOffset(years=1))

    # Materia gravada real reportada al IESS es la base principal del cálculo.
    b['Base Beneficios IESS']=b['Sueldo IESS']
    b['XIII Causado']=b['Base Beneficios IESS']/12.0
    b['XIV Causado']=(482.0/12.0)*(b['Días IESS'].clip(lower=0,upper=30)/30.0)
    # FONDO DE RESERVA:
    # OPERATIVOS: tratamiento propio CENASE. Se causa sobre la materia gravada IESS
    # del grupo operativo SIN aplicar el filtro general de 1 año usado para Adm/Ger.
    # Adm/Ger: conserva validación de antigüedad.
    b['FR Causado Teórico']=np.where(
        b['Tipo Rol'].eq('Operativos'),
        b['Base Beneficios IESS'] * 0.0833,
        np.where(b['Cumple 1 Año FR'], b['Base Beneficios IESS'] * 0.0833, 0.0)
    )

    # El pago mensual real viene del propio rol.
    # Para Operativos, el gasto contable total del grupo se determina posteriormente
    # sobre la materia gravada IESS completa; el rol sirve para separar pago mensual
    # y acumulación.
    b['FR Causado']=b['FR Causado Teórico']

    b['Modalidad XIII']=np.where(b['_Blank XIII'],'ACUMULA','PAGO MENSUAL')
    b['Modalidad XIV']=np.where(b['_Blank XIV'],'ACUMULA','PAGO MENSUAL')

    # Fondo: misma fórmula legal, pero se reporta y concilia por separado según grupo.
    b['Regla FR']=np.where(
        b['Tipo Rol'].eq('Operativos'),
        'OPERATIVO - IESS + control contra planilla FONDOS',
        'ADM/GER - IESS + modalidad según rol'
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
    return b

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
    """Asiento de ROL / DEVENGO con trazabilidad de fuentes.

    Fuente principal: IESS por trabajador y grupo.
    Pago real: se valida contra planillas y se presenta aparte; préstamos y fondos pagados
    usan el valor real de planilla cuando existe.
    """
    rr=roles.copy(); ii=iess.copy(); bb=benefits.copy()
    tipo=rr[['Cédula','Tipo Rol']].drop_duplicates('Cédula')
    ii=ii.merge(tipo,on='Cédula',how='left'); ii['Tipo Rol']=ii['Tipo Rol'].fillna('Sin Rol')

    rows=[]
    def debit(code,name,val,source='',grupo=''):
        rows.append({'Grupo':grupo,'Cuenta':f'{code} {name}','Debe':round(float(val),2),'Haber':0.0,'Fuente':source})
    def credit(code,name,val,source='',grupo=''):
        rows.append({'Grupo':grupo,'Cuenta':f'{code} {name}','Debe':0.0,'Haber':round(float(val),2),'Fuente':source})
    def overtime(d):
        return float((d['Horas Suplementarias 50%']+d['Horas Extraordinarias 100%']+d['Recargo 25%']).sum())

    # Mantener 3 grupos SIEMPRE separados. Gerentes y Administrativos usan mismas cuentas,
    # pero jamás se mezclan para calcular los valores.
    group_defs=[
        ('Operativos','Vtas.','5.2.1.1'),
        ('Gerentes','Adm.','5.2.1.2'),
        ('Administrativos','Adm.','5.2.1.2')
    ]
    for tipo_rol,suffix,prefix in group_defs:
        r=rr[rr['Tipo Rol'].eq(tipo_rol)]
        i=ii[ii['Tipo Rol'].eq(tipo_rol)]
        b=bb[bb['Tipo Rol'].eq(tipo_rol)]
        base=float(i['Sueldo IESS'].sum())
        ot=overtime(r)
        # Gratificaciones = Otros Ingresos del rol, por instrucción CENASE.
        grat=float(r['Otros Ingresos'].sum())
        sueldo=max(base-ot,0.0)
        patronal=float(i['Patronal IESS'].sum())
        secap=float(i['Valor CCC'].sum())
        # Fondo de Reserva
        if tipo_rol == 'Operativos':
            # Regla contable CENASE para Operativos: 8,33% de toda la materia gravada IESS
            # del grupo. No aplicar filtro general de antigüedad.
            fr=round(base * 0.0833, 2)
        else:
            # Adm/Ger: mantiene elegibilidad individual.
            fr=float(b['FR Causado'].sum())

        # XIII y XIV: el rol ya trae el cálculo de nómina aplicado persona por persona.
        # La fórmula IESS queda como auditoría, pero NO reemplaza el valor del rol.
        xiii=float(r['Décimo Tercero'].sum())
        xiv=float(r['Décimo Cuarto'].sum())

        debit(f'{prefix}.1',f'Sueldos Unificados {suffix}',sueldo,'IESS materia gravada - sobretiempos Rol',tipo_rol)
        debit(f'{prefix}.2',f'Sobretiempos {suffix}',ot,'Rol: suplementarias + extraordinarias + recargo',tipo_rol)
        debit(f'{prefix}.3',f'Gratificaciones {suffix}',grat,'OTROS INGRESOS DEL ROL',tipo_rol)
        debit(f'{prefix}.5',f'Aportes Patronales al IESS {suffix}',patronal,'Consolidado IESS por trabajador',tipo_rol)
        debit(f'{prefix}.6',f'Secap - Iece {suffix}',secap,'SUMA Valor CCC real del Consolidado IESS por trabajador',tipo_rol)
        debit(f'{prefix}.7',f'Fondos de Reserva {suffix}',fr,
              'OPERATIVOS: 8,33% materia gravada IESS completa' if tipo_rol=='Operativos'
              else 'ADM/GER: 8,33% IESS según elegibilidad individual',tipo_rol)
        debit(f'{prefix}.8',f'Décimo Tercer Sueldo {suffix}',xiii,'VALOR REAL CALCULADO EN ROL; IESS / 12 queda como control',tipo_rol)
        debit(f'{prefix}.9',f'Décimo Cuarto Sueldo {suffix}',xiv,'VALOR REAL CALCULADO EN ROL; SBU 482 / 12 queda como control',tipo_rol)

    # Pasivos / descuentos.
    xiii_acc=float(bb['XIII Acumulado'].sum())
    xiv_acc=float(bb['XIV Acumulado'].sum())
    fr_acc_calc=float(bb['FR Acumulado'].sum())
    fr_paid=paid_by_type(planillas,'FONDOS')
    qp_paid=paid_by_type(planillas,'DIVPRE')

    credit('1.1.2.5.11','Anticipos a empleados',rr['Anticipos'].sum(),'Rol')
    credit('2.1.7.1.1','9.45% Aportes Individuales',ii['Individual IESS'].sum(),'Consolidado IESS')
    credit('2.1.7.1.2','Prestamos Quirografarios',qp_paid if qp_paid else rr['Préstamo Quirografario'].sum(),
           'PLANILLA DIVPRE PAGADA' if qp_paid else 'Rol - sin planilla DIVPRE')
    credit('2.1.7.6.1','Décimo Tercer Sueldo',xiii_acc,'Acumulado: celda XIII vacía')
    credit('2.1.7.6.2','Décimo Cuarto Sueldo',xiv_acc,'Acumulado: celda XIV vacía')
    credit('2.1.7.6.4','11.15% Aportes Patronales I.E.S.S.',ii['Patronal IESS'].sum(),'Consolidado IESS')
    secap_total=sum(r['Debe'] for r in rows if 'Secap - Iece' in r['Cuenta'])
    credit('2.1.7.6.5','1% Secap - Iece',secap_total,'SUMA Valor CCC real del Consolidado IESS')
    # Si existe planilla FONDOS, el valor efectivamente pagado queda como fuente prioritaria para el pago.
    # En el asiento de rol se conserva el acumulado calculado; la conciliación muestra la diferencia real.
    credit('2.1.7.6.6','Fondos de Reservas',fr_acc_calc,'Acumulado calculado; validar vs planilla FONDOS real')

    # Otros descuentos reales del rol que no tienen cuenta específica arriba.
    otros=float((rr['Préstamo Hipotecario']+rr['Faltas / Pérdida Remuneración']+rr['Otros Egresos']+rr['Multa']+rr['Impuesto Renta']).sum())
    if abs(otros)>0.005:
        credit('2.1.7.7.2','Otros descuentos de nómina (cta pte)',otros,'Descuentos del rol no clasificados en cuentas anteriores')

    # Resultado natural del asiento de devengo. NO incluye el ajuste Rol vs IESS.
    a0=pd.DataFrame(rows)
    debe=float(a0['Debe'].sum()); haber_identificado=float(a0['Haber'].sum())
    sueldos_pagar=debe-haber_identificado
    credit('2.1.7.7.1','Sueldos por Pagar',sueldos_pagar,'Resultado del devengo IESS luego de pasivos identificados')
    return pd.DataFrame(rows)

def build_payment_accounting(iess,planillas):
    """Asiento/soporte de PAGO REAL IESS. No altera el asiento de rol.
    Si consolidado y planilla difieren, la contabilidad del pago toma la planilla real.
    """
    base=float(iess['Sueldo IESS'].sum())
    cons_plani=float(iess['Individual IESS'].sum()+iess['Patronal IESS'].sum()+iess['Valor CCC'].sum())
    items=[
      ('PLANI','Aportes IESS + CCC/SECAP-IECE según Consolidado',cons_plani),
      ('DIVPRE','Préstamos Quirografarios',np.nan),
      ('FONDOS','Fondos de Reserva',np.nan),
      ('PLTJEM','Juveniles / obligaciones trabajadores',np.nan),
      ('EXTSALCY','Extensión salud / cónyuge',np.nan),
    ]
    rows=[]
    for tipo,concepto,cons in items:
        pag=paid_by_type(planillas,tipo)
        rows.append({
          'Tipo IESS':tipo,'Concepto':concepto,
          'Consolidado / cálculo':cons,
          'Pago Real Planillas':pag,
          'Valor a Contabilizar':pag if pag else (0.0 if pd.isna(cons) else cons),
          'Diferencia vs Consolidado':np.nan if pd.isna(cons) else pag-cons,
          'Fuente Aplicada':'PLANILLA PAGADA' if pag else 'CONSOLIDADO IESS',
          'Estado':'INFORMATIVO' if pd.isna(cons) else ('CUADRA' if abs(pag-cons)<=0.05 else 'REVISAR')
        })
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

    # Obligaciones acumuladas del Rol Simulado IESS.
    s["Décimo XIII Acumulado IESS"] = s["XIII Acumulado Simulado"]
    s["Décimo XIV Acumulado IESS"] = s["XIV Acumulado Simulado"]
    s["Fondo Reserva Acumulado IESS"] = s["FR Acumulado Simulado"]
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

# ---------- UPLOADS ----------
with st.sidebar:
    st.header("Carga mensual")
    st.caption("Sube los 4 archivos del mismo período.")
    fger=st.file_uploader("1. Rol de Gerentes",type=["xlsx","xls"],key="ger")
    fad=st.file_uploader("2. Rol de Administración",type=["xlsx","xls"],key="adm")
    fop=st.file_uploader("3. Rol de Operativos",type=["xlsx","xls"],key="ope")
    fiess=st.file_uploader("4. Consolidado IESS",type=["xlsx","xls"],key="iess")
    fplan=st.file_uploader("5. Reporte de Planillas IESS Pagadas",type=["xlsx"],key="plan")
    st.divider()
    st.caption("El cruce Rol vs IESS se realiza por cédula.")

if not (fger and fad and fop and fiess and fplan):
    st.info("Carga los cinco archivos para generar el reporte y la conciliación Rol vs IESS.")
    st.stop()

try:
    ger=read_role(fger,"Gerentes")
    adm=read_role(fad,"Administrativos")
    ope=read_role(fop,"Operativos")
    roles=pd.concat([ger,adm,ope],ignore_index=True)
    roles=add_role_employer_calcs(roles)
    dup_mask = roles.duplicated(subset=["Tipo Rol","Mes","Cédula"], keep=False)
    if dup_mask.any():
        st.warning(f"Se detectaron {int(dup_mask.sum())} filas duplicadas. Se conservará una sola fila por trabajador.")
        roles = roles.drop_duplicates(subset=["Tipo Rol","Mes","Cédula"], keep="first").reset_index(drop=True)
    iess=read_iess(fiess)
    comp=make_compare(roles,iess)
    sim_iess=build_iess_simulated_role(roles,iess)
    benefits=build_benefits(roles,iess)
    planillas=read_planillas(fplan)
    plan_compare=build_iess_plan_compare(iess,planillas)
    accounting=build_accounting(roles,iess,benefits,planillas)
    payment_accounting=build_payment_accounting(iess,planillas)
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

tabs=st.tabs(["📊 Resumen Roles","🏛️ Rol vs IESS","🧮 Rol Simulado IESS","⚠️ Diferencias","💳 IESS vs Planillas","🎁 Beneficios","🧾 Contabilidad","🔎 Consulta Roles","✅ Cuadre"])

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
    sim_cols=["Estado Simulación","Tipo Rol","Cédula","Nombre","Cargo","Puesto / Cliente","Días Laborados","Días IESS","Dif. Días Rol vs IESS","Materia Gravada Rol Calc","Sueldo IESS","Dif. Materia Gravada Rol vs IESS","Sueldo Base Simulado IESS","Sobretiempos Rol","Otros Ingresos Gravados Rol","Décimo Tercero","XIII Simulado IESS","Décimo XIII Pagado IESS","Décimo XIII Acumulado IESS","Dif. XIII Rol vs Sim IESS","Décimo Cuarto","XIV Simulado IESS","Décimo XIV Pagado IESS","Décimo XIV Acumulado IESS","Dif. XIV Rol vs Sim IESS","Fondo Reserva","FR Simulado IESS","Fondo Reserva Pagado IESS","Fondo Reserva Acumulado IESS","Total Beneficios Pagados IESS","Total Beneficios Acumulados IESS","Dif. FR Rol vs Sim IESS","IESS","Individual IESS","Dif. Individual Rol vs IESS","Aporte Patronal Rol Calc 11.15%","Patronal IESS","Dif. Patronal Calc Rol vs IESS","SECAP-IECE Rol Calc 1%","Valor CCC","Dif. SECAP Calc Rol vs IESS","Total Ingresos","Total Ingresos Simulado IESS","Dif. Total Ingresos Rol vs Sim IESS","Total Egresos","Total Egresos Simulado IESS","Dif. Total Egresos Rol vs Sim IESS","Neto a Recibir","Neto Simulado IESS","Dif. Neto Rol vs Sim IESS"]
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

with tabs[4]:
    st.subheader("IESS vs Planillas Pagadas")
    st.dataframe(plan_compare,use_container_width=True,hide_index=True)
    st.markdown("#### Valor contable del pago real")
    st.info("Prioridad aplicada: 1) el Consolidado IESS aporta los valores reales por trabajador (Patronal, Individual y Valor CCC); 2) el Reporte de Planillas verifica lo efectivamente PAGADO; 3) si Consolidado y Planillas difieren, el PAGO contable usa la planilla real y la diferencia queda visible para revisión. No se recalcula CCC sobre el total de materia gravada.")
    st.dataframe(payment_accounting,use_container_width=True,hide_index=True)
    st.markdown("#### Detalle de planillas cargadas")
    st.dataframe(planillas,use_container_width=True,hide_index=True)

with tabs[5]:
    st.subheader("Beneficios: pago mensual vs acumulación")
    st.caption("Base principal: materia gravada IESS. Décimos: valor en rol = pago mensual; vacío = acumula. Fondo de Reserva: fórmula IESS/12 después de 1 año, con Operativos separados de Gerentes/Administrativos.")
    bcols=["Tipo Rol","Cédula","Nombre","Fecha Ingreso","Sueldo IESS","Días IESS","XIII Causado","Modalidad XIII","XIII Pagado Rol","XIII Acumulado","XIV Causado","Modalidad XIV","XIV Pagado Rol","XIV Acumulado","Cumple 1 Año FR","Regla FR","Modalidad FR","FR Causado","FR Pagado Rol","FR Acumulado"]
    st.dataframe(benefits[bcols],use_container_width=True,hide_index=True)

with tabs[6]:
    st.subheader("Asiento contable propuesto")
    st.caption("Asiento de devengo: IESS es la base principal por trabajador y por los 3 grupos. Gratificaciones = Otros Ingresos del rol. Los pagos reales de IESS se concilian aparte contra las planillas pagadas.")
    st.dataframe(accounting,use_container_width=True,hide_index=True)
    ac1,ac2,ac3=st.columns(3)
    ac1.metric("Total Debe",fmt_money(accounting["Debe"].sum()))
    ac2.metric("Total Haber",fmt_money(accounting["Haber"].sum()))
    ac3.metric("Diferencia",fmt_money(accounting["Debe"].sum()-accounting["Haber"].sum()))
    st.markdown("#### Auditoría contra asiento patrón de enero 2026")
    bench=january_benchmark(accounting,roles)
    st.dataframe(bench,use_container_width=True,hide_index=True)
    st.metric("Líneas patrón que cuadran",f"{int((bench['Estado']=='CUADRA').sum())} / {len(bench)}")

    st.markdown("#### Asiento separado de ajuste Rol vs IESS")
    neto=float(roles["Neto a Recibir"].sum())
    # Para enero, el asiento principal patrón dejó Sueldos por Pagar en 133,650.63.
    # En meses siguientes se toma el valor generado por el asiento principal.
    sp=float(accounting.loc[accounting["Cuenta"].str.startswith("2.1.7.7.1"),"Haber"].sum())
    ajuste=round(sp-neto,2)
    if abs(ajuste)<=0.05:
        st.success("No se requiere asiento de ajuste: Sueldos por Pagar coincide con el neto del rol.")
    else:
        aj=pd.DataFrame([
          {"Cuenta":"2.1.7.7.1 Sueldos por Pagar","Debe":max(ajuste,0),"Haber":max(-ajuste,0)},
          {"Cuenta":"5.2.1.2.72 (-) Ajuste Diferencia Rol vs. IESS","Debe":max(-ajuste,0),"Haber":max(ajuste,0)}
        ])
        st.dataframe(aj,use_container_width=True,hide_index=True)
        st.caption(f"Glosa: P/R AJUSTE POR DIFERENCIA ENTRE VALORES REGISTRADOS SEGÚN IESS Y ROL DE PAGOS CORRESPONDIENTE AL PERÍODO. Valor: {fmt_money(abs(ajuste))}")

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

# ---------- DOWNLOAD ----------
st.divider()
excel=export_excel(roles,summary,comp,diffs,benefits,planillas,plan_compare,accounting,payment_accounting,sim_iess)
mes=roles["Mes"].dropna().astype(str)
mes=mes.iloc[0] if len(mes) else datetime.now().strftime("%Y-%m")
st.download_button("⬇️ Descargar reporte completo Roles + IESS",
                   data=excel,file_name=f"Roles_vs_IESS_{mes}.xlsx",
                   mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                   use_container_width=True)
st.markdown('<p class="small">El archivo descargado incluye además Rol Simulado IESS y los nuevos campos patronales calculados dentro del Rol real.</p>',unsafe_allow_html=True)
