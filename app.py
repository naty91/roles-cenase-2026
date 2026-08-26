
import io
import re
import unicodedata
from datetime import datetime, timedelta

import numpy as np
import pandas as pd
import streamlit as st

st.set_page_config(
    page_title="CENASE | Reporte de Roles",
    page_icon="📊",
    layout="wide",
)

# ---------- ESTILO ----------
st.markdown("""
<style>
    .stApp { background: #f5f7fb; }
    .block-container { padding-top: 1.4rem; padding-bottom: 2rem; max-width: 1500px; }
    h1, h2, h3 { color: #0b2f55; }
    .hero {
        background: linear-gradient(135deg,#073b6f,#0d5fa6);
        padding: 22px 26px;
        border-radius: 18px;
        color: white;
        margin-bottom: 18px;
        box-shadow: 0 8px 24px rgba(13,95,166,.15);
    }
    .hero h1 { color: white; margin: 0; font-size: 2rem; }
    .hero p { margin: 5px 0 0 0; color: #e9f4ff; }
    div[data-testid="stMetric"] {
        background: white;
        border: 1px solid #e5eaf0;
        padding: 14px 16px;
        border-radius: 14px;
        box-shadow: 0 4px 14px rgba(0,0,0,.04);
    }
    div[data-testid="stDataFrame"] {
        background: white;
        border-radius: 12px;
    }
    .ok {
        background:#ecfdf3; color:#166534; border:1px solid #bbf7d0;
        border-radius:12px; padding:10px 13px; margin:8px 0;
    }
    .warn {
        background:#fff7ed; color:#9a3412; border:1px solid #fed7aa;
        border-radius:12px; padding:10px 13px; margin:8px 0;
    }
    .small-note {font-size:.88rem;color:#64748b;}
</style>
""", unsafe_allow_html=True)

st.markdown("""
<div class="hero">
  <h1>Reporte Consolidado de Roles</h1>
  <p>Gerentes · Administración · Operativos | Consulta, cuadre y descarga mensual</p>
</div>
""", unsafe_allow_html=True)

# ---------- UTILIDADES ----------
def norm_text(x):
    if x is None:
        return ""
    s = str(x).strip()
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    s = re.sub(r"\s+", " ", s).upper()
    return s

def excel_date(value):
    if pd.isna(value) or value == "":
        return pd.NaT
    if isinstance(value, (pd.Timestamp, datetime)):
        return pd.to_datetime(value)
    try:
        num = float(value)
        if 20000 < num < 80000:
            return pd.Timestamp("1899-12-30") + pd.to_timedelta(num, unit="D")
    except Exception:
        pass
    return pd.to_datetime(value, errors="coerce")

def money(v):
    try:
        return float(v) if pd.notna(v) and v != "" else 0.0
    except Exception:
        return 0.0

CANONICAL_COLUMNS = [
    "Tipo Rol","Mes","Cédula","Nombre","Fecha Ingreso","Cargo","Puesto / Cliente",
    "Días Laborados","Base","Sueldo","Horas Suplementarias 50%","Horas Extraordinarias 100%",
    "Recargo 25%","Décimo Tercero","Décimo Cuarto","Fondo Reserva","Movilización",
    "Otros Ingresos","Total Ingresos","Préstamo Quirografario","Préstamo Hipotecario",
    "Anticipos","Faltas / Pérdida Remuneración","Otros Egresos","IESS","Multa","Impuesto Renta",
    "Total Egresos","Neto a Recibir","Observaciones","Email"
]

ALIASES = {
    "MAIL":"Email","EMAIL":"Email","MES":"Mes","C.I":"Cédula","CI":"Cédula","CEDULA":"Cédula",
    "NOMBRE":"Nombre","F.INGRESO":"Fecha Ingreso","F INGRESO":"Fecha Ingreso","CARGO":"Cargo",
    "PUESTO":"Puesto / Cliente","D.LAB":"Días Laborados","DIAS":"Días Laborados",
    "DIAS LABORADOS":"Días Laborados","BASE":"Base","SUELDO":"Sueldo",
    "HORAS EXTRAS":"Horas Extraordinarias 100%","H.S. 50%":"Horas Suplementarias 50%",
    "HS 50%":"Horas Suplementarias 50%","H.E. 100":"Horas Extraordinarias 100%",
    "HE 100":"Horas Extraordinarias 100%","RECARGO 25%":"Recargo 25%",
    "DECIMO TERCER":"Décimo Tercero","13 AVO":"Décimo Tercero",
    "DECIMO CUARTO":"Décimo Cuarto","14 AVO":"Décimo Cuarto",
    "FONDO RESER":"Fondo Reserva","F.R.":"Fondo Reserva",
    "MOVILIZACION":"Movilización","OTROS ING.":"Otros Ingresos","OTROS ING":"Otros Ingresos",
    "TOTAL INGRESOS":"Total Ingresos","T. ING":"Total Ingresos",
    "PTMO-QUIROG":"Préstamo Quirografario","P. QUIR":"Préstamo Quirografario",
    "PTMO HIPOT":"Préstamo Hipotecario","ANTICIPO":"Anticipos","ANTICIPOS":"Anticipos",
    "FALTAS":"Faltas / Pérdida Remuneración","OTROS EGRESOS":"Otros Egresos",
    "OTROS EGR":"Otros Egresos","IESS":"IESS","MULTA":"Multa","I.R":"Impuesto Renta",
    "TOTAL EGRESOS":"Total Egresos","T. DESC":"Total Egresos",
    "NETO A RECIBIR":"Neto a Recibir","NETO":"Neto a Recibir",
    "OBSERVACIONES / PENDIENTES":"Observaciones"
}

NUMERIC_COLS = [
    "Días Laborados","Base","Sueldo","Horas Suplementarias 50%","Horas Extraordinarias 100%",
    "Recargo 25%","Décimo Tercero","Décimo Cuarto","Fondo Reserva","Movilización",
    "Otros Ingresos","Total Ingresos","Préstamo Quirografario","Préstamo Hipotecario",
    "Anticipos","Faltas / Pérdida Remuneración","Otros Egresos","IESS","Multa","Impuesto Renta",
    "Total Egresos","Neto a Recibir"
]

def detect_header(raw, tipo):
    # Los archivos actuales tienen encabezados cerca del inicio.
    for i in range(min(12, len(raw))):
        vals = [norm_text(v) for v in raw.iloc[i].tolist()]
        has_name = "NOMBRE" in vals
        has_ci = any(v in ("C.I","CI","CEDULA") for v in vals)
        has_net = any(v in ("NETO","NETO A RECIBIR") for v in vals)
        if has_name and has_ci and has_net:
            return i
    raise ValueError(f"No pude identificar la fila de encabezados del rol {tipo}.")

def find_sheet(file_obj):
    xls = pd.ExcelFile(file_obj)
    for s in xls.sheet_names:
        if norm_text(s) == "LISTA":
            return s
    for s in xls.sheet_names:
        if "LISTA" in norm_text(s):
            return s
    raise ValueError("No se encontró la hoja LISTA/lista.")

def read_role(file_obj, tipo):
    sheet = find_sheet(file_obj)
    raw = pd.read_excel(file_obj, sheet_name=sheet, header=None, dtype=object)
    h = detect_header(raw, tipo)
    headers = [norm_text(x) for x in raw.iloc[h].tolist()]
    data = raw.iloc[h+1:].copy()
    data.columns = headers

    # Renombrar solo columnas conocidas y omitir duplicados vacíos.
    rename = {}
    for c in data.columns:
        key = norm_text(c)
        if key in ALIASES:
            rename[c] = ALIASES[key]
    data = data.rename(columns=rename)

    # Si hubiera columnas duplicadas, tomar la primera no vacía.
    compact = pd.DataFrame(index=data.index)
    for c in list(dict.fromkeys(data.columns)):
        same = data.loc[:, data.columns == c]
        if same.shape[1] == 1:
            compact[c] = same.iloc[:,0]
        else:
            compact[c] = same.bfill(axis=1).iloc[:,0]
    data = compact

    # Filas reales: cédula + nombre. Esto elimina subtotales y firmas.
    if "Cédula" not in data.columns or "Nombre" not in data.columns:
        raise ValueError(f"Faltan Cédula o Nombre en {tipo}.")
    data["Cédula"] = data["Cédula"].astype(str).str.replace(r"\.0$","",regex=True).str.strip()
    data["Nombre"] = data["Nombre"].fillna("").astype(str).str.strip()
    data = data[
        data["Cédula"].str.fullmatch(r"\d{8,13}", na=False)
        & data["Nombre"].ne("")
    ].copy()

    data["Tipo Rol"] = tipo

    for c in CANONICAL_COLUMNS:
        if c not in data.columns:
            data[c] = np.nan

    data["Fecha Ingreso"] = data["Fecha Ingreso"].apply(excel_date)

    # Mes puede venir como número serial de Excel.
    parsed_mes = data["Mes"].apply(excel_date)
    if parsed_mes.notna().any():
        data["Mes"] = parsed_mes.dt.strftime("%Y-%m")
    else:
        data["Mes"] = data["Mes"].astype(str)

    for c in NUMERIC_COLS:
        data[c] = pd.to_numeric(data[c], errors="coerce").fillna(0.0)

    # Para operativos no hay "Cargo"; conservar identificación de Guardia.
    if tipo == "Operativos":
        data["Cargo"] = data["Cargo"].replace("", np.nan).fillna("Guardia")
    else:
        data["Puesto / Cliente"] = data["Puesto / Cliente"].fillna("")

    return data[CANONICAL_COLUMNS].reset_index(drop=True)

def fmt_money(v):
    return f"${v:,.2f}"

def build_excel(filtered, summary):
    out = io.BytesIO()
    with pd.ExcelWriter(out, engine="xlsxwriter", datetime_format="dd/mm/yyyy") as writer:
        wb = writer.book
        fmt_title = wb.add_format({"bold": True, "font_size": 16, "font_color":"#FFFFFF",
                                   "bg_color":"#0B4F88","align":"center","valign":"vcenter"})
        fmt_header = wb.add_format({"bold": True, "font_color":"#FFFFFF","bg_color":"#0D5FA6",
                                    "border":1,"align":"center","valign":"vcenter","text_wrap":True})
        fmt_money_x = wb.add_format({"num_format":"$#,##0.00","border":1})
        fmt_int = wb.add_format({"num_format":"0","border":1})
        fmt_date = wb.add_format({"num_format":"dd/mm/yyyy","border":1})
        fmt_text = wb.add_format({"border":1})
        fmt_sub = wb.add_format({"bold":True,"bg_color":"#EAF2F8","border":1})
        fmt_sub_money = wb.add_format({"bold":True,"bg_color":"#EAF2F8","border":1,"num_format":"$#,##0.00"})

        summary.to_excel(writer, sheet_name="Resumen", index=False, startrow=3)
        ws = writer.sheets["Resumen"]
        ws.merge_range("A1:F1","RESUMEN CONSOLIDADO DE ROLES",fmt_title)
        for j, col in enumerate(summary.columns):
            ws.write(3,j,col,fmt_header)
        ws.set_column("A:A",22)
        ws.set_column("B:B",12)
        ws.set_column("C:F",18,fmt_money_x)
        ws.freeze_panes(4,0)

        # Consolidado y hojas por rol.
        outputs = [("Consolidado", filtered)]
        for role in ["Gerentes","Administrativos","Operativos"]:
            outputs.append((role, filtered[filtered["Tipo Rol"] == role].copy()))

        for sheet_name, df in outputs:
            safe = df.copy()
            safe.to_excel(writer, sheet_name=sheet_name, index=False, startrow=2)
            ws = writer.sheets[sheet_name]
            last_col = max(len(safe.columns)-1, 0)
            ws.merge_range(0,0,0,last_col, f"REPORTE DE ROLES - {sheet_name.upper()}", fmt_title)
            for j, col in enumerate(safe.columns):
                ws.write(2,j,col,fmt_header)
            ws.freeze_panes(3,0)
            ws.autofilter(2,0,2+len(safe),last_col)
            ws.set_column(0,0,18)
            ws.set_column(1,1,10)
            ws.set_column(2,2,14)
            ws.set_column(3,3,37)
            ws.set_column(4,4,14,fmt_date)
            ws.set_column(5,6,27)
            ws.set_column(7,7,14,fmt_int)
            ws.set_column(8,28,17,fmt_money_x)
            ws.set_column(29,30,28)
            # Convertir observaciones/email a texto format no moneda
            if len(safe.columns) > 29:
                ws.set_column(29,30,28,fmt_text)

        writer.close()
    out.seek(0)
    return out.getvalue()

# ---------- CARGA ----------
with st.sidebar:
    st.header("Carga mensual")
    st.caption("Sube los tres archivos correspondientes al mismo mes.")
    f_ger = st.file_uploader("1. Rol de Gerentes", type=["xlsx","xls"], key="ger")
    f_adm = st.file_uploader("2. Rol de Administración", type=["xlsx","xls"], key="adm")
    f_ope = st.file_uploader("3. Rol de Operativos", type=["xlsx","xls"], key="ope")
    st.divider()
    st.caption("Los archivos originales no se modifican. El sistema genera el consolidado para consulta y descarga.")

if not (f_ger and f_adm and f_ope):
    st.info("Sube los tres roles para generar el reporte mensual.")
    st.stop()

try:
    ger = read_role(f_ger, "Gerentes")
    adm = read_role(f_adm, "Administrativos")
    ope = read_role(f_ope, "Operativos")
    all_data = pd.concat([ger, adm, ope], ignore_index=True)
except Exception as e:
    st.error(f"No pude procesar uno de los archivos: {e}")
    st.stop()

# ---------- VALIDACIONES ----------
computed = all_data.groupby("Tipo Rol", as_index=False).agg(
    Empleados=("Cédula","count"),
    Total_Ingresos=("Total Ingresos","sum"),
    Total_Egresos=("Total Egresos","sum"),
    Neto=("Neto a Recibir","sum")
)
total_row = pd.DataFrame([{
    "Tipo Rol":"TOTAL GENERAL",
    "Empleados":int(computed["Empleados"].sum()),
    "Total_Ingresos":computed["Total_Ingresos"].sum(),
    "Total_Egresos":computed["Total_Egresos"].sum(),
    "Neto":computed["Neto"].sum()
}])
summary_all = pd.concat([computed, total_row], ignore_index=True)
summary_all.columns = ["Tipo de Rol","Empleados","Total Ingresos","Total Egresos","Neto a Recibir"]

# Validaciones aritméticas por persona.
all_data["_dif_ing"] = (
    all_data["Sueldo"] + all_data["Horas Suplementarias 50%"] + all_data["Horas Extraordinarias 100%"]
    + all_data["Recargo 25%"] + all_data["Décimo Tercero"] + all_data["Décimo Cuarto"]
    + all_data["Fondo Reserva"] + all_data["Movilización"] + all_data["Otros Ingresos"]
    - all_data["Total Ingresos"]
)
all_data["_dif_neto"] = all_data["Total Ingresos"] - all_data["Total Egresos"] - all_data["Neto a Recibir"]
issues_income = (all_data["_dif_ing"].abs() > 0.05).sum()
issues_net = (all_data["_dif_neto"].abs() > 0.05).sum()

# ---------- KPIs ----------
c1,c2,c3,c4,c5 = st.columns(5)
c1.metric("Empleados", f"{len(all_data):,}")
c2.metric("Total ingresos", fmt_money(all_data["Total Ingresos"].sum()))
c3.metric("Total egresos", fmt_money(all_data["Total Egresos"].sum()))
c4.metric("Neto a pagar", fmt_money(all_data["Neto a Recibir"].sum()))
c5.metric("Puestos operativos", f"{ope['Puesto / Cliente'].replace('',np.nan).nunique():,}")

if issues_income == 0 and issues_net == 0:
    st.markdown('<div class="ok">✓ Cuadre aritmético correcto: ingresos, egresos y neto coinciden dentro de una tolerancia de $0,05.</div>', unsafe_allow_html=True)
else:
    st.markdown(
        f'<div class="warn">⚠ Revisar: {issues_income} registro(s) con diferencia en componentes de ingresos '
        f'y {issues_net} registro(s) con diferencia entre ingresos − egresos y neto.</div>',
        unsafe_allow_html=True
    )

# ---------- TABS ----------
tab_res, tab_cons, tab_cuadre = st.tabs(["📊 Resumen","🔎 Consulta detallada","✅ Cuadre"])

with tab_res:
    st.subheader("Resumen por tipo de rol")
    show_summary = summary_all.copy()
    for c in ["Total Ingresos","Total Egresos","Neto a Recibir"]:
        show_summary[c] = show_summary[c].map(fmt_money)
    st.dataframe(show_summary, use_container_width=True, hide_index=True)

    st.subheader("Neto a pagar por rol")
    chart = computed.set_index("Tipo Rol")[["Neto"]]
    st.bar_chart(chart, use_container_width=True)

    if not ope.empty:
        st.subheader("Operativos por puesto / cliente")
        op_summary = (
            ope.groupby("Puesto / Cliente", dropna=False)
            .agg(Empleados=("Cédula","count"), Neto=("Neto a Recibir","sum"))
            .reset_index()
            .sort_values(["Empleados","Puesto / Cliente"], ascending=[False,True])
        )
        op_summary["Neto"] = op_summary["Neto"].map(fmt_money)
        st.dataframe(op_summary, use_container_width=True, hide_index=True)

with tab_cons:
    st.subheader("Filtros")
    f1,f2,f3,f4 = st.columns(4)
    role_filter = f1.multiselect(
        "Tipo de rol",
        options=sorted(all_data["Tipo Rol"].dropna().unique()),
        default=sorted(all_data["Tipo Rol"].dropna().unique())
    )
    names = f2.text_input("Nombre o cédula", placeholder="Buscar...")
    cargo_opts = sorted([x for x in all_data["Cargo"].dropna().astype(str).unique() if x.strip()])
    cargo_filter = f3.multiselect("Cargo", cargo_opts)
    puesto_opts = sorted([x for x in all_data["Puesto / Cliente"].dropna().astype(str).unique() if x.strip()])
    puesto_filter = f4.multiselect("Puesto / cliente", puesto_opts)

    g1,g2,g3 = st.columns(3)
    min_neto = float(all_data["Neto a Recibir"].min())
    max_neto = float(all_data["Neto a Recibir"].max())
    neto_range = g1.slider(
        "Rango neto a recibir",
        min_value=float(np.floor(min_neto)),
        max_value=float(np.ceil(max_neto)),
        value=(float(np.floor(min_neto)), float(np.ceil(max_neto))),
        step=1.0
    )
    dias_min = float(all_data["Días Laborados"].min())
    dias_max = float(all_data["Días Laborados"].max())
    dias_range = g2.slider(
        "Días laborados",
        min_value=float(np.floor(dias_min)),
        max_value=float(np.ceil(dias_max)),
        value=(float(np.floor(dias_min)), float(np.ceil(dias_max))),
        step=1.0
    )
    cols_default = [
        "Tipo Rol","Cédula","Nombre","Cargo","Puesto / Cliente","Días Laborados",
        "Sueldo","Total Ingresos","Total Egresos","Neto a Recibir"
    ]
    selected_cols = g3.multiselect("Columnas a mostrar", CANONICAL_COLUMNS, default=cols_default)

    filtered = all_data[all_data["Tipo Rol"].isin(role_filter)].copy()
    if names.strip():
        q = norm_text(names)
        mask = filtered["Nombre"].map(norm_text).str.contains(q, na=False) | filtered["Cédula"].astype(str).str.contains(names.strip(), na=False)
        filtered = filtered[mask]
    if cargo_filter:
        filtered = filtered[filtered["Cargo"].isin(cargo_filter)]
    if puesto_filter:
        filtered = filtered[filtered["Puesto / Cliente"].isin(puesto_filter)]
    filtered = filtered[
        filtered["Neto a Recibir"].between(*neto_range)
        & filtered["Días Laborados"].between(*dias_range)
    ]

    k1,k2,k3,k4 = st.columns(4)
    k1.metric("Registros filtrados", len(filtered))
    k2.metric("Ingresos filtrados", fmt_money(filtered["Total Ingresos"].sum()))
    k3.metric("Egresos filtrados", fmt_money(filtered["Total Egresos"].sum()))
    k4.metric("Neto filtrado", fmt_money(filtered["Neto a Recibir"].sum()))

    display = filtered[selected_cols].copy() if selected_cols else filtered[cols_default].copy()
    money_cols = [c for c in display.columns if c in NUMERIC_COLS and c != "Días Laborados"]
    st.dataframe(
        display,
        use_container_width=True,
        hide_index=True,
        column_config={
            **{c: st.column_config.NumberColumn(c, format="$ %.2f") for c in money_cols},
            "Días Laborados": st.column_config.NumberColumn("Días Laborados", format="%.0f"),
            "Fecha Ingreso": st.column_config.DateColumn("Fecha Ingreso", format="DD/MM/YYYY")
        }
    )

    # Resumen filtrado para exportar.
    fsum = filtered.groupby("Tipo Rol", as_index=False).agg(
        Empleados=("Cédula","count"),
        **{"Total Ingresos":("Total Ingresos","sum"),
           "Total Egresos":("Total Egresos","sum"),
           "Neto a Recibir":("Neto a Recibir","sum")}
    )
    total_f = pd.DataFrame([{
        "Tipo Rol":"TOTAL GENERAL","Empleados":int(fsum["Empleados"].sum()),
        "Total Ingresos":fsum["Total Ingresos"].sum(),
        "Total Egresos":fsum["Total Egresos"].sum(),
        "Neto a Recibir":fsum["Neto a Recibir"].sum()
    }])
    fsum = pd.concat([fsum,total_f],ignore_index=True).rename(columns={"Tipo Rol":"Tipo de Rol"})

    export_df = filtered[CANONICAL_COLUMNS].copy()
    excel_bytes = build_excel(export_df, fsum)
    mes_label = all_data["Mes"].dropna().astype(str).replace("NaT","")
    mes_label = mes_label.iloc[0] if len(mes_label) else datetime.now().strftime("%Y-%m")
    st.download_button(
        "⬇️ Descargar Excel consolidado",
        data=excel_bytes,
        file_name=f"Roles_Consolidados_{mes_label}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        use_container_width=True
    )
    st.caption("El Excel incluye: Resumen, Consolidado, Gerentes, Administrativos y Operativos, respetando los filtros aplicados.")

with tab_cuadre:
    st.subheader("Cuadre por rol")
    check = computed.copy()
    check["Diferencia Neto"] = check["Total_Ingresos"] - check["Total_Egresos"] - check["Neto"]
    check = check.rename(columns={
        "Tipo Rol":"Tipo de Rol","Total_Ingresos":"Total Ingresos",
        "Total_Egresos":"Total Egresos","Neto":"Neto a Recibir"
    })
    st.dataframe(
        check,
        use_container_width=True,
        hide_index=True,
        column_config={
            "Total Ingresos": st.column_config.NumberColumn(format="$ %.2f"),
            "Total Egresos": st.column_config.NumberColumn(format="$ %.2f"),
            "Neto a Recibir": st.column_config.NumberColumn(format="$ %.2f"),
            "Diferencia Neto": st.column_config.NumberColumn(format="$ %.2f"),
        }
    )

    st.subheader("Registros con diferencias")
    bad = all_data[(all_data["_dif_ing"].abs() > 0.05) | (all_data["_dif_neto"].abs() > 0.05)].copy()
    if bad.empty:
        st.success("No se detectaron diferencias aritméticas superiores a $0,05.")
    else:
        bad["Dif. componentes ingresos"] = bad["_dif_ing"]
        bad["Dif. neto"] = bad["_dif_neto"]
        st.dataframe(
            bad[["Tipo Rol","Cédula","Nombre","Total Ingresos","Total Egresos","Neto a Recibir",
                 "Dif. componentes ingresos","Dif. neto"]],
            use_container_width=True,
            hide_index=True
        )

st.markdown('<p class="small-note">Sistema preparado para reutilizarse mensualmente con archivos que mantengan la estructura actual de los roles.</p>', unsafe_allow_html=True)
