import os
import sys
import glob
import re
import pandas as pd
import plotly.express as px
import streamlit as st


# ============================================================
# 1. CONFIGURACIÓN GENERAL DE LA PÁGINA
# ============================================================

st.set_page_config(
    page_title="Dashboard Autoinspección - La Moderna",
    page_icon="🍪",
    layout="wide"
)

st.title(" Dashboard de Autoinspección")
st.subheader("Fábrica de Galletas La Moderna — Reporte Anual")
st.markdown("---")


# ============================================================
# 2. UBICACIÓN DE LA CARPETA DE DATOS
# ============================================================

if os.environ.get("AUTOINSPECCIONES_DATOS"):
    RUTA_BASE = os.environ["AUTOINSPECCIONES_DATOS"]
else:
    RUTA_BASE = os.path.join(
        os.path.dirname(os.path.abspath(__file__)),
        "datos"
    )


# ============================================================
# 3. CONFIGURACIÓN DE MESES Y CATEGORÍAS
# ============================================================

ORDEN_MESES = [
    "ENERO",
    "FEBRERO",
    "MARZO",
    "ABRIL",
    "MAYO",
    "JUNIO",
    "JULIO",
    "AGOSTO",
    "SEPTIEMBRE",
    "OCTUBRE",
    "NOVIEMBRE",
    "DICIEMBRE"
]

CATEGORIAS = [
    "AP",
    "CP",
    "MO",
    "MS",
    "PL"
]


# ============================================================
# 4. FUNCIÓN PARA OBTENER LOS AÑOS DISPONIBLES
# ============================================================

def obtener_carpetas_anio(ruta_base):
    if not os.path.exists(ruta_base):
        return {}

    carpetas = {}
    for nombre in os.listdir(ruta_base):
        ruta = os.path.join(ruta_base, nombre)

        if not os.path.isdir(ruta):
            continue

        coincidencia = re.search(r"(20\d{2})", nombre)
        if coincidencia:
            anio = int(coincidencia.group(1))
            carpetas[anio] = ruta

    return dict(sorted(carpetas.items()))


# ============================================================
# 5. FUNCIÓN PARA EXTRAER LOS DATOS DE CADA MES
# ============================================================

def extraer_tablas_mes(filepath, mes):
    """
    Lee la hoja 'TABLAS' del Excel del mes y extrae:
    1. Resumen de Categorías por Línea
    2. Totales por Categoría
    3. Hallazgos por Línea de Producción
    4. Hallazgos por Responsable
    5. Hallazgos por Pre Requisito
    """
    try:
        xls = pd.ExcelFile(filepath)

        if "TABLAS" not in xls.sheet_names:
            return None

        df_raw = pd.read_excel(
            xls,
            sheet_name="TABLAS",
            header=None
        )

        # ----------------------------------------------------
        # A) CATEGORÍAS POR LÍNEA
        # ----------------------------------------------------
        df_cat_linea = pd.DataFrame()
        dict_totales_cat = {cat: 0 for cat in CATEGORIAS}

        for idx, row in df_raw.iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if (
                any("TABLA CATEGORÍAS POR LÍNEA" in x for x in row_str)
                or any("LINEAS" in x for x in row_str)
            ):
                sub_df = df_raw.iloc[idx + 1:idx + 13, [8, 9, 10, 11, 12, 13]].copy()
                sub_df.columns = ["LINEA"] + CATEGORIAS
                sub_df = sub_df.dropna(subset=["LINEA"])
                sub_df = sub_df[~sub_df["LINEA"].astype(str).str.upper().str.contains("TOTAL")]

                for col in CATEGORIAS:
                    sub_df[col] = pd.to_numeric(sub_df[col], errors="coerce").fillna(0)
                    dict_totales_cat[col] = sub_df[col].sum()

                sub_df["MES"] = mes
                df_cat_linea = sub_df
                break

        # ----------------------------------------------------
        # B) HALLAZGOS POR LÍNEA
        # ----------------------------------------------------
        df_hallazgos_linea = pd.DataFrame()

        for idx, row in df_raw.iterrows():
            val_0 = str(row.iloc[0]).upper() if pd.notna(row.iloc[0]) else ""
            if "LÍNEAS" in val_0 or "LINEAS" in val_0:
                sub_df = df_raw.iloc[idx + 1:idx + 12, [0, 1]].copy()
                sub_df.columns = ["LINEA", "HALLAZGOS"]
                sub_df = sub_df.dropna(subset=["LINEA"])
                sub_df = sub_df[~sub_df["LINEA"].astype(str).str.upper().str.contains("TOTAL")]
                sub_df["HALLAZGOS"] = pd.to_numeric(sub_df["HALLAZGOS"], errors="coerce").fillna(0)
                sub_df["MES"] = mes
                df_hallazgos_linea = sub_df
                break

        # ----------------------------------------------------
        # C) HALLAZGOS POR RESPONSABLE
        # ----------------------------------------------------
        df_responsables = pd.DataFrame()

        for idx, row in df_raw.iterrows():
            row_str = [str(x).upper() for x in row.values if pd.notna(x)]
            if any("RESPONSABLE" in x for x in row_str):
                sub_df = df_raw.iloc[idx + 2:idx + 12, [5, 7]].copy()
                sub_df.columns = ["RESPONSABLE", "HALLAZGOS"]
                sub_df = sub_df.dropna(subset=["RESPONSABLE"])
                sub_df = sub_df[~sub_df["RESPONSABLE"].astype(str).str.upper().str.contains("TOTAL")]
                sub_df["HALLAZGOS"] = pd.to_numeric(sub_df["HALLAZGOS"], errors="coerce").fillna(0)
                sub_df["MES"] = mes
                df_responsables = sub_df
                break

        # ----------------------------------------------------
        # D) HALLAZGOS POR PRE REQUISITO
        # ----------------------------------------------------
        df_prerequisitos = pd.DataFrame()

        for idx, row in df_raw.iterrows():
            val_a = str(row.iloc[0]).upper() if pd.notna(row.iloc[0]) else ""
            
            if "PRE REQUISITO" in val_a or "PRE-REQUISITO" in val_a:
                # Leer desde la fila siguiente al encabezado (columnas A y B)
                sub_df = df_raw.iloc[idx + 2:idx + 15, [0, 2]].copy()
                sub_df.columns = ["PRE_REQUISITO", "HALLAZGOS"]
                
                sub_df = sub_df.dropna(subset=["PRE_REQUISITO"])
                
                # Convertir a texto, pasar a mayúsculas y ELIMINAR ESPACIOS EXTRA
                sub_df["PRE_REQUISITO"] = (
                    sub_df["PRE_REQUISITO"]
                    .astype(str)
                    .str.strip()
                    .str.upper()
                )
                
                # Filtrar encabezados y la fila TOTAL
                sub_df = sub_df[
                    ~sub_df["PRE_REQUISITO"].str.contains("PRE REQUISITO|PRE-REQUISITO|TOTAL")
                ]
                
                sub_df["HALLAZGOS"] = pd.to_numeric(sub_df["HALLAZGOS"], errors="coerce").fillna(0)
                sub_df["MES"] = mes
                df_prerequisitos = sub_df
                break

        return {
            "cat_linea": df_cat_linea,
            "totales_cat": dict_totales_cat,
            "hallazgos_linea": df_hallazgos_linea,
            "responsables": df_responsables,
            "prerequisitos": df_prerequisitos
        }

    except Exception as e:
        st.error(f"Error al procesar la hoja TABLAS de {os.path.basename(filepath)}: {e}")
        return None


# ============================================================
# 6. CARGAR TODOS LOS MESES DE UN AÑO
# ============================================================

def cargar_todos_los_meses(ruta_folder):
    patron = os.path.join(ruta_folder, "*.xlsx")
    archivos = [f for f in glob.glob(patron) if not os.path.basename(f).startswith("~$")]
    dict_meses = {}

    for archivo in archivos:
        nombre_file = os.path.basename(archivo).upper()
        mes_encontrado = None

        for mes in ORDEN_MESES:
            if mes in nombre_file:
                mes_encontrado = mes
                break

        if mes_encontrado:
            res = extraer_tablas_mes(archivo, mes_encontrado)
            if res:
                dict_meses[mes_encontrado] = res

    return archivos, dict_meses


# ============================================================
# 7. DETECTAR AÑOS DISPONIBLES
# ============================================================

carpetas_anio = obtener_carpetas_anio(RUTA_BASE)

if not carpetas_anio:
    st.error(f"No se encontraron carpetas de años dentro de:\n\n{RUTA_BASE}")
    st.info(
        "Ejemplo de estructura correcta:\n\n"
        "datos/\n"
        "├── Autoinspeccion Líneas 2026/\n"
        "├── Autoinspeccion Líneas 2027/\n"
        "└── Autoinspeccion Líneas 2028/"
    )
    st.stop()


# ============================================================
# 8. MENÚ LATERAL
# ============================================================

if getattr(sys, "frozen", False):
    ruta_imagen = os.path.join(sys._MEIPASS, "image.jpeg")
else:
    ruta_imagen = "image.jpeg"

if os.path.exists(ruta_imagen):
    st.sidebar.image(ruta_imagen, width=80)

st.sidebar.title("Filtros del Dashboard")


# ============================================================
# 9. SELECTOR DE AÑO
# ============================================================

anios_disponibles = list(carpetas_anio.keys())

anio_seleccionado = st.sidebar.selectbox(
    "Seleccionar Año:",
    anios_disponibles,
    index=len(anios_disponibles) - 1
)

RUTA_DATOS = carpetas_anio[anio_seleccionado]


# ============================================================
# 10. BOTÓN PARA ACTUALIZAR Y BOTÓN PARA APAGAR
# ============================================================

if st.sidebar.button("Actualizar Datos de Excel", use_container_width=True):
    st.rerun()

st.sidebar.markdown("---")

if st.sidebar.button("🔴 Detener / Apagar Aplicación", use_container_width=True):
    st.sidebar.warning("Apagando el servidor... Ya puedes cerrar esta pestaña.")
    os._exit(0)

st.sidebar.markdown("---")


# ============================================================
# 11. CARGAR LOS EXCEL DEL AÑO SELECCIONADO
# ============================================================

if not os.path.exists(RUTA_DATOS):
    st.error(f"No existe la carpeta:\n{RUTA_DATOS}")
    st.stop()

archivos_excel, datos_meses = cargar_todos_los_meses(RUTA_DATOS)


# ============================================================
# 12. FILTRO POR MES
# ============================================================

meses_disponibles = [m for m in ORDEN_MESES if m in datos_meses]

filtro_mes = st.sidebar.selectbox(
    "Seleccionar Mes:",
    ["TODOS LOS MESES"] + meses_disponibles
)

st.sidebar.markdown("---")

st.sidebar.info(
    f"**Año seleccionado:** {anio_seleccionado}\n\n"
    f"**Meses procesados:** {len(datos_meses)} / 12"
)


# ============================================================
# 13. CONSOLIDACIÓN DE DATOS
# ============================================================

lista_hallazgos = []
lista_cat = []
lista_resp = []
lista_prereq = []

totales_categorias_global = {cat: 0 for cat in CATEGORIAS}

for m, data in datos_meses.items():
    if filtro_mes == "TODOS LOS MESES" or filtro_mes == m:

        if "hallazgos_linea" in data and not data["hallazgos_linea"].empty:
            lista_hallazgos.append(data["hallazgos_linea"])

        if "cat_linea" in data and not data["cat_linea"].empty:
            lista_cat.append(data["cat_linea"])

        if "responsables" in data and not data["responsables"].empty:
            lista_resp.append(data["responsables"])

        if "prerequisitos" in data and not data["prerequisitos"].empty:
            lista_prereq.append(data["prerequisitos"])

        if "totales_cat" in data:
            for cat, val in data["totales_cat"].items():
                totales_categorias_global[cat] += val


# ============================================================
# 14. CREAR DATAFRAMES CONSOLIDADOS
# ============================================================

if lista_hallazgos:
    df_filtrado_lineas = pd.concat(lista_hallazgos, ignore_index=True)
else:
    df_filtrado_lineas = pd.DataFrame(columns=["LINEA", "HALLAZGOS", "MES"])

if lista_resp:
    df_filtrado_resp = pd.concat(lista_resp, ignore_index=True)
else:
    df_filtrado_resp = pd.DataFrame(columns=["RESPONSABLE", "HALLAZGOS", "MES"])

if lista_prereq:
    df_filtrado_prereq = pd.concat(lista_prereq, ignore_index=True)
else:
    df_filtrado_prereq = pd.DataFrame(columns=["PRE_REQUISITO", "HALLAZGOS", "MES"])


# ============================================================
# 15. TÍTULO DEL AÑO
# ============================================================

st.markdown(f"## Dashboard Anual {anio_seleccionado}")
st.caption(f"Carpeta analizada: {os.path.basename(RUTA_DATOS)}")


# ============================================================
# 16. MÉTRICAS PRINCIPALES
# ============================================================

col1, col2, col3, col4 = st.columns(4)

total_h = df_filtrado_lineas["HALLAZGOS"].sum() if not df_filtrado_lineas.empty else 0

with col1:
    st.metric("TOTAL DE HALLAZGOS", int(total_h))

with col2:
    st.metric("MESES LEÍDOS", f"{len(datos_meses)} / 12")

with col3:
    st.metric("AÑO DE LECTURA", str(anio_seleccionado))

with col4:
    st.metric("PERIODO FILTRADO", filtro_mes if filtro_mes != "TODOS LOS MESES" else "Ene - Dic")

st.markdown("---")


# ============================================================
# 17. VISUALIZACIONES
# ============================================================

st.subheader(" Análisis de Autoinspecciones")

# FILA 1
col_g1, col_g2 = st.columns(2)

with col_g1:
    st.markdown("##### Total Hallazgos por Mes")
    if not df_filtrado_lineas.empty:
        df_por_mes = df_filtrado_lineas.groupby("MES", as_index=False)["HALLAZGOS"].sum()
        df_por_mes["MES"] = pd.Categorical(df_por_mes["MES"], categories=ORDEN_MESES, ordered=True)
        df_por_mes = df_por_mes.sort_values("MES")

        fig_mes = px.bar(
            df_por_mes,
            x="MES",
            y="HALLAZGOS",
            text="HALLAZGOS",
            color_discrete_sequence=["#2b5c8f"]
        )
        fig_mes.update_traces(textposition="outside")
        fig_mes.update_layout(xaxis_title="Mes", yaxis_title="Hallazgos", height=380, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig_mes, use_container_width=True)
    else:
        st.info("No hay datos disponibles.")

with col_g2:
    st.markdown("##### Total Categorías")
    df_pie_cat = pd.DataFrame(list(totales_categorias_global.items()), columns=["Categoría", "Total"])
    fig_cat = px.pie(
        df_pie_cat,
        values="Total",
        names="Categoría",
        hole=0.4,
        color_discrete_sequence=px.colors.qualitative.Pastel
    )
    fig_cat.update_layout(height=380, margin=dict(l=20, r=20, t=30, b=20))
    st.plotly_chart(fig_cat, use_container_width=True)


# FILA 2
col_g3, col_g4 = st.columns(2)

with col_g3:
    st.markdown("##### Hallazgos por Línea de Producción")
    if not df_filtrado_lineas.empty:
        df_por_linea = df_filtrado_lineas.groupby("LINEA", as_index=False)["HALLAZGOS"].sum().sort_values("HALLAZGOS", ascending=False)
        fig_linea = px.bar(
            df_por_linea,
            x="LINEA",
            y="HALLAZGOS",
            text="HALLAZGOS",
            color_discrete_sequence=["#d9534f"]
        )
        fig_linea.update_traces(textposition="outside")
        fig_linea.update_layout(xaxis_title="Línea", yaxis_title="Hallazgos", height=380, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig_linea, use_container_width=True)
    else:
        st.info("No hay datos disponibles.")

with col_g4:
    st.markdown("##### Hallazgos por Responsable")
    if not df_filtrado_resp.empty:
        df_por_resp = df_filtrado_resp.groupby("RESPONSABLE", as_index=False)["HALLAZGOS"].sum().sort_values("HALLAZGOS", ascending=True)
        fig_resp = px.bar(
            df_por_resp,
            x="HALLAZGOS",
            y="RESPONSABLE",
            text="HALLAZGOS",
            orientation="h",
            color_discrete_sequence=["#20b2aa"]
        )
        fig_resp.update_traces(textposition="outside")
        fig_resp.update_layout(xaxis_title="Hallazgos", yaxis_title="Responsable", height=380, margin=dict(l=20, r=20, t=30, b=20))
        st.plotly_chart(fig_resp, use_container_width=True)
    else:
        st.info("No hay datos disponibles.")


# FILA 3 - PRE REQUISITOS
st.markdown("---")
st.markdown("##### Hallazgos por Pre Requisito")

if not df_filtrado_prereq.empty:
    df_por_prereq = df_filtrado_prereq.groupby("PRE_REQUISITO", as_index=False)["HALLAZGOS"].sum().sort_values("HALLAZGOS", ascending=False)
    fig_prereq = px.bar(
        df_por_prereq,
        x="PRE_REQUISITO",
        y="HALLAZGOS",
        text="HALLAZGOS",
        color_discrete_sequence=["#ff9900"]
    )
    fig_prereq.update_traces(textposition="outside")
    fig_prereq.update_layout(xaxis_title="Pre Requisito", yaxis_title="Hallazgos", height=400, margin=dict(l=20, r=20, t=30, b=20))
    st.plotly_chart(fig_prereq, use_container_width=True)
else:
    st.info("No hay datos de Pre Requisitos disponibles.")


# ============================================================
# 18. TABLA DE DATOS
# ============================================================

with st.expander(" Ver datos consolidados"):
    st.dataframe(df_filtrado_lineas, use_container_width=True)


# ============================================================
# 19. INFORMACIÓN DEL SISTEMA
# ============================================================

st.markdown("---")
st.caption(" Dashboard de Autoinspección — Fábrica de Galletas La Moderna")
st.caption("Los datos se obtienen directamente de los archivos Excel mensuales.")