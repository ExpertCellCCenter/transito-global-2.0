import streamlit as st
import pandas as pd
import numpy as np
from datetime import date, datetime
import pyodbc
import plotly.express as px
from io import BytesIO
from openpyxl.utils import get_column_letter

# -------------------------------------------------
# CONFIG STREAMLIT
# -------------------------------------------------
st.set_page_config(
    page_title="Transito Globlal 2.0",
    page_icon="🚗",
    layout="wide",
)

# ---------- Fancy global styles ----------
st.markdown(
    """
<style>
html, body, [class*="css"]  {
    font-family: "Segoe UI", system-ui, sans-serif;
}

/* Main background */
section.main {
    background: radial-gradient(circle at top left,#020617 0,#0b1120 45%,#020617 100%);
    color: #f9fafb;
}
.block-container {
    padding-top: 1.2rem;
    padding-bottom: 2rem;
}

/* Title */
h1 {
    font-weight: 800 !important;
}

/* Tabs */
.stTabs [role="tablist"] {
    gap: 6px;
}
.stTabs [role="tab"] {
    padding: 6px 14px;
    border-radius: 999px;
    background-color: #020617;
    color: #e5e7eb;
    border: 1px solid rgba(148,163,184,0.4);
}
.stTabs [aria-selected="true"] {
    background: linear-gradient(90deg,#22c55e,#06b6d4);
    color: #0f172a !important;
    border-color: transparent !important;
}

/* Metrics */
[data-testid="stMetric"] {
    background: rgba(15,23,42,0.9);
    border-radius: 16px;
    padding: 10px 14px;
    border: 1px solid rgba(148,163,184,0.55);
    box-shadow: 0 10px 30px rgba(15,23,42,0.7);
}
[data-testid="stMetricValue"] {
    font-size: 1.7rem;
    font-weight: 800;
}

/* Download buttons */
div[data-testid="stDownloadButton"] > button {
    border-radius: 999px;
    background: linear-gradient(90deg,#0ea5e9,#6366f1);
    color: #f9fafb;
    border: none;
    padding: 0.4rem 1.3rem;
    font-weight: 600;
}
div[data-testid="stDownloadButton"] > button:hover {
    filter: brightness(1.1);
}

/* Plotly transparent background */
.js-plotly-plot .plotly .main-svg {
    background-color: rgba(0,0,0,0) !important;
}
</style>
""",
    unsafe_allow_html=True,
)

# -------------------------------------------------
# SMALL HELPER: DF -> EXCEL BYTES (auto-fit + filters)
# -------------------------------------------------
def df_to_excel_bytes(df: pd.DataFrame, sheet_name: str = "Datos") -> bytes:
    """Return an .xlsx file (bytes) with autofilter and auto column width."""
    output = BytesIO()
    with pd.ExcelWriter(output, engine="openpyxl") as writer:
        df.to_excel(writer, index=False, sheet_name=sheet_name)
        wb = writer.book
        ws = writer.sheets[sheet_name]

        max_row = ws.max_row
        max_col = ws.max_column

        # Autofiltro en encabezados
        ws.auto_filter.ref = f"A1:{get_column_letter(max_col)}{max_row}"

        # Auto ancho columnas
        for col_idx in range(1, max_col + 1):
            col_letter = get_column_letter(col_idx)
            max_length = 0
            for cell in ws[col_letter]:
                if cell.value is not None:
                    max_length = max(max_length, len(str(cell.value)))
            ws.column_dimensions[col_letter].width = max_length + 2

    output.seek(0)
    return output.getvalue()


# -------------------------------------------------
# DB CONNECTION
# -------------------------------------------------
@st.cache_resource
def get_connection():
    cfg = st.secrets["sql"]
    driver = cfg["driver"]
    server = cfg["server"]
    database = cfg["database"]
    user = cfg["user"]
    password = cfg["password"]

    conn_str = (
        f"DRIVER={{{driver}}};"
        f"SERVER={server};"
        f"DATABASE={database};"
        f"UID={user};"
        f"PWD={password};"
        "Encrypt=yes;"
        "TrustServerCertificate=yes;"
    )
    return pyodbc.connect(conn_str)


# -------------------------------------------------
# LOAD DATA FROM SQL
# -------------------------------------------------
@st.cache_data(ttl=900)
def load_hoja1():
    """
    Tabla equivalente a Hoja1 de Power BI,
    pero leyendo directamente de reporte_empleado en SQL.

    - Solo ATT, CONTACT CENTER, VIRTUAL
    - Solo puestos teléfono/supervisor
    - Solo empleados ACTIVO
    """

    sql = """
    SELECT DISTINCT
        e.[Nombre Completo] AS NombreCompleto,
        e.[Jefe Inmediato]  AS JefeDirecto,
        e.[Region],
        e.[SubRegion],
        e.[Plaza],
        e.[Tienda],
        e.[Puesto],
        e.[Canal de Venta],
        e.[Tipo Tienda],
        e.[Operacion],
        e.[Estatus]
    FROM reporte_empleado('EMPRESA_MAESTRA',1,'','') AS e
    WHERE
        e.[Canal de Venta] = 'ATT'
        AND e.[Operacion]   = 'CONTACT CENTER'
        AND e.[Tipo Tienda] = 'VIRTUAL'
        AND e.[Puesto] IN (
            'ASESOR TELEFONICO',
            'ASESOR TELEFONICO 7500',
            'EJECUTIVO TELEFONICO 6500 AM',
            'EJECUTIVO TELEFONICO 6500 PM',
            'SUPERVISOR DE CONTACT CENTER'
        )
        AND e.[Estatus] = 'ACTIVO';
    """

    conn = get_connection()
    df = pd.read_sql(sql, conn)

    # Limpieza de textos
    text_cols = [
        "NombreCompleto",
        "JefeDirecto",
        "Region",
        "SubRegion",
        "Plaza",
        "Tienda",
        "Puesto",
        "Canal de Venta",
        "Tipo Tienda",
        "Operacion",
        "Estatus",
    ]
    for col in text_cols:
        df[col] = df[col].astype(str).str.strip()
        df[col] = df[col].replace({"nan": np.nan, "None": np.nan})

    # Como hacías en Power BI: supervisor vacío -> "ENCUBADORA"
    df["JefeDirecto"] = df["JefeDirecto"].fillna("").str.strip()
    df["JefeDirecto"] = df["JefeDirecto"].replace("", "ENCUBADORA")

    # Columna Coordinador igual a JefeDirecto
    df["Coordinador"] = df["JefeDirecto"]

    return df


@st.cache_data(ttl=900)
def load_consulta1(fecha_ini: date, fecha_fin: date) -> pd.DataFrame:
    """
    Replica Consulta1 base (sin las transformaciones).
    """
    fi = fecha_ini.strftime("%Y%m%d")
    ff = fecha_fin.strftime("%Y%m%d")

    sql = f"""
    SELECT
        [Folio],
        [Telefono],
        [Entrega],
        [Tienda solicita] AS Centro,
        [Estatus],
        [Back Office],
        [Fecha creacion],
        [Venta],
        [Vendedor]
    FROM reporte_programacion_entrega('empresa_maestra', 4, '{fi}', '{ff}')
    WHERE
        [Tienda solicita] LIKE 'EXP ATT C CENTER%' AND
        [Estatus] IN ('En entrega','Canc Error','Entregado',
                      'En preparacion','Back Office','Solicitado');
    """
    conn = get_connection()
    df = pd.read_sql(sql, conn)
    return df


# -------------------------------------------------
# TRANSFORMACIONES COMO EN POWER QUERY
# -------------------------------------------------
def transform_consulta1(df_raw: pd.DataFrame, hoja: pd.DataFrame) -> pd.DataFrame:
    df = df_raw.copy()

    # --- Limpieza de textos ---
    for col in ["Centro", "Estatus", "Back Office", "Venta", "Vendedor"]:
        if col in df.columns:
            df[col] = df[col].astype(str).str.strip()
            df[col] = df[col].replace({"nan": np.nan, "None": np.nan})

    # --- Centro Original (sin np.where) ---
    df["Centro Original"] = np.nan  # empezamos todo como NaN (object)

    mask_cc2 = df["Centro"].str.contains("EXP ATT C CENTER 2", na=False)
    mask_juarez = df["Centro"].str.contains("EXP ATT C CENTER JUAREZ", na=False)

    df.loc[mask_cc2, "Centro Original"] = "CC2"
    df.loc[mask_juarez, "Centro Original"] = "Juarez"

    # --- Región ---
    def region_from_centro(c):
        if not isinstance(c, str):
            return np.nan
        if "GDL" in c:
            return "GDL"
        if "MEX" in c:
            return "MEX"
        if "MTY" in c:
            return "MTY"
        if "PUE" in c:
            return "PUE"
        if "TIJ" in c:
            return "TIJ"
        if "VER" in c:
            return "VER"
        return np.nan

    df["Region"] = df["Centro"].apply(region_from_centro)

    # --- Join con Hoja1 (ejecutivos / supervisores) ---
    hoja_join = hoja.rename(
        columns={
            "NombreCompleto": "Nombre Completo",
            "JefeDirecto": "Jefe directo",
        }
    )

    df = df.merge(
        hoja_join[["Nombre Completo", "Jefe directo", "Coordinador"]],
        how="left",
        left_on="Vendedor",
        right_on="Nombre Completo",
    )

    # ya no necesitamos la columna de join
    df.drop(columns=["Nombre Completo"], inplace=True, errors="ignore")

    # --- Status calculado ---
    def status_calc(row):
        est = row.get("Estatus")
        venta = row.get("Venta")

        if est in ("En entrega", "En preparacion", "Solicitado", "Back Office"):
            return "En Transito"

        venta_vacia = pd.isna(venta) or str(venta).strip() == ""

        if est == "Entregado" and venta_vacia:
            return "En Transito"

        if est == "Entregado":
            return "Entregado"

        return est

    df["Status"] = df.apply(status_calc, axis=1)

    # --- Campos de fecha / hora ---
    df["Fecha creacion"] = pd.to_datetime(df["Fecha creacion"])
    df["Fecha"] = df["Fecha creacion"].dt.date
    df["Hora"] = df["Fecha creacion"].dt.hour

    # --- Campos tipo calendario ---
    iso = df["Fecha creacion"].dt.isocalendar()
    df["Año"] = df["Fecha creacion"].dt.year
    df["MesNum"] = df["Fecha creacion"].dt.month
    df["Mes"] = df["Fecha creacion"].dt.strftime("%B")
    df["AñoMes"] = df["Fecha creacion"].dt.strftime("%Y-%m")
    df["Día"] = df["Fecha creacion"].dt.day
    df["Nombre Día"] = df["Fecha creacion"].dt.strftime("%A")
    df["Año Semana"] = (
        iso["year"].astype(str) + "-" + iso["week"].astype(str).str.zfill(2)
    )

    return df


@st.cache_data(ttl=900)
def build_sin_venta(hoja: pd.DataFrame, consulta: pd.DataFrame) -> pd.DataFrame:
    """
    Replica la tabla SinVenta original:
    Hoja1 anti-join Consulta1 por Nombre Completo vs Vendedor.
    """
    tmp = hoja.merge(
        consulta[["Vendedor"]],
        how="left",
        left_on="NombreCompleto",
        right_on="Vendedor",
        indicator=True,
    )
    sinv = tmp[tmp["_merge"] == "left_only"].copy()
    sinv.drop(columns=["Vendedor", "_merge"], inplace=True)
    return sinv


# -------------------------------------------------
# KPI HELPERS (equivalentes a medidas DAX)
# -------------------------------------------------
def kpi_activadas(df: pd.DataFrame) -> int:
    return int((df["Status"] == "Entregado").sum())


def kpi_back(df: pd.DataFrame) -> int:
    return int((df["Estatus"] == "Back Office").sum())


def kpi_en_entrega(df: pd.DataFrame) -> int:
    return int((df["Estatus"] == "En entrega").sum())


def kpi_en_transito(df: pd.DataFrame) -> int:
    return int((df["Status"] == "En Transito").sum())


def kpi_preparacion(df: pd.DataFrame) -> int:
    return int((df["Estatus"] == "En preparacion").sum())


def kpi_solicitados(df: pd.DataFrame) -> int:
    return int((df["Estatus"] == "Solicitado").sum())


def kpi_total_canc_error(df: pd.DataFrame) -> int:
    return int((df["Estatus"] == "Canc Error").sum())


def kpi_total_programadas(df: pd.DataFrame) -> int:
    # igual que Total Programadas = Estatus <> "Canc Error"
    return int((df["Estatus"] != "Canc Error").sum())


def kpi_entregados_sin_venta(df: pd.DataFrame) -> int:
    mask = (df["Estatus"] == "Entregado") & (
        df["Venta"].isna() | (df["Venta"].astype(str).str.strip() == "")
    )
    return int(mask.sum())


def kpi_total_sinventa(df_sinventa: pd.DataFrame) -> int:
    return int(df_sinventa.shape[0])


# -------------------------------------------------
# MAIN APP
# -------------------------------------------------
def main():
    st.title("Dashboard Transito Global  – CC")

    # ----------- Sidebar: rangos de fecha y filtros globales -----------
    st.sidebar.header("Filtros")

    # Rango de fechas equivalente al Calendario de Power BI
    default_start = date(2025, 10, 1)
    default_end = date.today()

    fecha_ini = st.sidebar.date_input("Fecha inicio", default_start)
    fecha_fin = st.sidebar.date_input("Fecha fin", default_end)

    if fecha_ini > fecha_fin:
        st.sidebar.error("La fecha inicio no puede ser mayor que la fecha fin.")
        return

    # Carga datos
    with st.spinner("Cargando datos desde SQL..."):
        hoja = load_hoja1()
        consulta_raw = load_consulta1(fecha_ini, fecha_fin)
        consulta = transform_consulta1(consulta_raw, hoja)
        sinventa = build_sin_venta(hoja, consulta)

    # Filtros de Centro, Supervisor y Mes (como en el panel negro de PBI)
    centros = ["All"] + sorted(
        [c for c in consulta["Centro Original"].dropna().unique().tolist()]
    )
    supervisores = ["All"] + sorted(
        [s for s in consulta["Jefe directo"].dropna().unique().tolist()]
    )
    meses = ["All"] + sorted(consulta["Mes"].unique().tolist())

    centro_sel = st.sidebar.selectbox("Centro", centros, index=0)
    supervisor_sel = st.sidebar.selectbox("Supervisor", supervisores, index=0)
    mes_sel = st.sidebar.selectbox("Mes", meses, index=0)

    df = consulta.copy()
    if centro_sel != "All":
        df = df[df["Centro Original"] == centro_sel]
    if supervisor_sel != "All":
        df = df[df["Jefe directo"] == supervisor_sel]
    if mes_sel != "All":
        df = df[df["Mes"] == mes_sel]

    # SinVenta filtrado por supervisor si aplica
    sinv_fil = sinventa.copy()
    if supervisor_sel != "All":
        sinv_fil = sinv_fil[sinv_fil["JefeDirecto"] == supervisor_sel]

    # ----------- Tabs (equivalentes a páginas de PBI) -----------
    tabs = st.tabs(
        [
            "Resumen",
            "Back Office",
            "Canceladas",
            "Programadas x semana",
            "Programadas (Top Ejecutivos)",
            "Detalle General",
            "Sin Venta",
        ]
    )

    # ==================== TAB 0: RESUMEN ====================
    with tabs[0]:
        st.subheader("Resumen de estatus")

        col1, col2, col3 = st.columns(3)
        with col1:
            st.metric("En preparación", kpi_preparacion(df))
            st.metric("Solicitados", kpi_solicitados(df))
        with col2:
            st.metric("En entrega", kpi_en_entrega(df))
            st.metric("En tránsito", kpi_en_transito(df))
        with col3:
            st.metric("Activadas (Entregado)", kpi_activadas(df))
            st.metric("Back Office", kpi_back(df))

        st.metric("Entregados sin venta (Validación)", kpi_entregados_sin_venta(df))

        st.download_button(
            "Descargar detalle (Excel)",
            data=df_to_excel_bytes(df, "Detalle"),
            file_name=f"detalle_programacion_{fecha_ini}_{fecha_fin}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        )

    # ==================== TAB 1: BACK OFFICE ====================
    with tabs[1]:
        st.subheader("Back Office")

        df_back = df[df["Estatus"] == "Back Office"]
        if df_back.empty:
            st.info("No hay registros con Estatus = 'Back Office' para los filtros actuales.")
        else:
            by_day = df_back.groupby("Fecha", as_index=False).size()
            fig = px.bar(
                by_day,
                x="Fecha",
                y="size",
                title="Total por día (Back Office)",
                labels={"size": "Total Programadas"},
            )
            st.plotly_chart(fig, use_container_width=True)

            day_sel = st.selectbox(
                "Selecciona un día para ver el desglose por hora",
                sorted(by_day["Fecha"].unique()),
            )
            df_day = df_back[df_back["Fecha"] == day_sel]
            by_hour = df_day.groupby("Hora", as_index=False).size()
            fig2 = px.bar(
                by_hour,
                x="Hora",
                y="size",
                title=f"Desglose por hora – {day_sel}",
                labels={"size": "Total Programadas"},
            )
            st.plotly_chart(fig2, use_container_width=True)

            st.download_button(
                "Descargar Back Office (Excel)",
                data=df_to_excel_bytes(df_back, "BackOffice"),
                file_name=f"backoffice_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ==================== TAB 2: CANCELADAS ====================
    with tabs[2]:
        st.subheader("Canceladas (Canc Error)")

        df_canc = df[df["Estatus"] == "Canc Error"]
        if df_canc.empty:
            st.info("No hay registros cancelados para los filtros actuales.")
        else:
            by_day = df_canc.groupby("Fecha", as_index=False).size()
            fig = px.bar(
                by_day,
                x="Fecha",
                y="size",
                title="Canceladas por día",
                labels={"size": "Total Canc Error"},
            )
            st.plotly_chart(fig, use_container_width=True)

            day_sel = st.selectbox(
                "Selecciona un día (Canceladas)",
                sorted(by_day["Fecha"].unique()),
            )
            df_day = df_canc[df_canc["Fecha"] == day_sel]
            by_hour = df_day.groupby("Hora", as_index=False).size()
            fig2 = px.bar(
                by_hour,
                x="Hora",
                y="size",
                title=f"Desglose por hora – {day_sel}",
                labels={"size": "Total Canc Error"},
            )
            st.plotly_chart(fig2, use_container_width=True)

            st.download_button(
                "Descargar Canceladas (Excel)",
                data=df_to_excel_bytes(df_canc, "Canceladas"),
                file_name=f"canceladas_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ==================== TAB 3: PROGRAMADAS X SEMANA ====================
    with tabs[3]:
        st.subheader("Programadas por semana")

        df_prog = df[df["Estatus"] != "Canc Error"]
        if df_prog.empty:
            st.info("No hay programadas para los filtros actuales.")
        else:
            by_week = df_prog.groupby("Año Semana", as_index=False).size()
            fig = px.bar(
                by_week,
                x="Año Semana",
                y="size",
                title="Vista general de programadas por semana",
                labels={"size": "Total Programadas"},
            )
            st.plotly_chart(fig, use_container_width=True)

            st.download_button(
                "Descargar Programadas (Excel)",
                data=df_to_excel_bytes(df_prog, "Programadas"),
                file_name=f"programadas_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ==================== TAB 4: PROGRAMADAS – TOP EJECUTIVOS ====================
    with tabs[4]:
        st.subheader("Top Ejecutivos – Programadas")

        df_prog = df[df["Estatus"] != "Canc Error"]
        if df_prog.empty:
            st.info("No hay programadas para los filtros actuales.")
        else:
            by_exec = (
                df_prog.groupby("Vendedor", as_index=False)
                .size()
                .rename(columns={"size": "Total Programadas"})
                .sort_values("Total Programadas", ascending=False)
                .head(30)
            )
            fig = px.bar(
                by_exec,
                x="Total Programadas",
                y="Vendedor",
                orientation="h",
                title="Top Ejecutivos Global",
                labels={"Vendedor": "Ejecutivo"},
            )
            st.plotly_chart(fig, use_container_width=True)

            st.download_button(
                "Descargar Top Ejecutivos (Excel)",
                data=df_to_excel_bytes(by_exec, "TopEjecutivos"),
                file_name=f"top_ejecutivos_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ==================== TAB 5: DETALLE GENERAL ====================
    with tabs[5]:
        st.subheader("Detalle general programadas")

        if df.empty:
            st.info("Sin datos para los filtros actuales.")
        else:
            grouped = (
                df.groupby(["Jefe directo", "Vendedor"], as_index=False)
                .agg(
                    TotalProgramadas=("Folio", "count"),
                    Activadas=("Status", lambda s: int((s == "Entregado").sum())),
                    EnTransito=("Status", lambda s: int((s == "En Transito").sum())),
                )
                .rename(columns={"Vendedor": "Ejecutivo"})
            )

            st.dataframe(grouped)

            by_sup = (
                grouped.groupby("Jefe directo", as_index=False)["TotalProgramadas"]
                .sum()
                .rename(columns={"Jefe directo": "Supervisor"})
            )
            fig = px.pie(
                by_sup,
                names="Supervisor",
                values="TotalProgramadas",
                title="Programadas por supervisor",
            )
            st.plotly_chart(fig, use_container_width=True)

            st.download_button(
                "Descargar detalle general (Excel)",
                data=df_to_excel_bytes(grouped, "DetalleGeneral"),
                file_name=f"detalle_general_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

    # ==================== TAB 6: SIN VENTA ====================
    with tabs[6]:
        st.subheader("Ejecutivos sin venta")

        total_sinv = kpi_total_sinventa(sinv_fil)
        st.metric("Total ejecutivos sin venta", total_sinv)

        if sinv_fil.empty:
            st.info("No hay ejecutivos sin venta para los filtros actuales.")
        else:
            # Ordena por JefeDirecto y nombre
            df_sinv = sinv_fil.sort_values(["JefeDirecto", "NombreCompleto"])
            st.dataframe(df_sinv[["JefeDirecto", "NombreCompleto"]])

            st.download_button(
                "Descargar Sin Venta (Excel)",
                data=df_to_excel_bytes(df_sinv[["JefeDirecto", "NombreCompleto"]], "SinVenta"),
                file_name=f"sin_venta_{fecha_ini}_{fecha_fin}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )


if __name__ == "__main__":
    main()
