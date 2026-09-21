from pathlib import Path
from io import BytesIO

import matplotlib.pyplot as plt
import pandas as pd
import plotly.express as px
import streamlit as st
from reportlab.lib import colors
from reportlab.lib.enums import TA_CENTER, TA_LEFT
from reportlab.lib.pagesizes import letter
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.units import inch
from reportlab.platypus import (
    SimpleDocTemplate,
    Paragraph,
    Spacer,
    Image,
    PageBreak,
    Table,
    TableStyle,
    KeepTogether,
)


st.set_page_config(
    page_title="Dashboard de Ventas y Rentabilidad",
    page_icon="📊",
    layout="wide",
)

st.markdown(
    """
    <style>
    .stApp {
        background: #FFFFFF;
        color: #1F2937;
    }

    [data-testid="stSidebar"] {
        background: #F5F7FA;
        border-right: 1px solid #D9E1EA;
    }

    [data-testid="stHeader"] {
        background: #FFFFFF;
    }

    [data-testid="stToolbar"] {
        background: #FFFFFF;
    }

    [data-testid="stMarkdownContainer"],
    [data-testid="stMarkdownContainer"] p,
    [data-testid="stMarkdownContainer"] h1,
    [data-testid="stMarkdownContainer"] h2,
    [data-testid="stMarkdownContainer"] h3,
    [data-testid="stSidebar"] label,
    [data-testid="stSidebar"] h2 {
        color: #1F2937;
    }

    [data-testid="stMetric"] {
        background: #F7F9FC;
        border: 1px solid #D9E1EA;
        border-radius: 10px;
        padding: 14px 16px;
        min-height: 105px;
        box-shadow: 0 1px 2px rgba(31, 41, 55, 0.04);
    }

    [data-testid="stMetricLabel"],
    [data-testid="stMetricValue"] {
        color: #1F2937;
    }

    [data-testid="stMetricValue"] {
        font-size: clamp(1.15rem, 2.2vw, 1.8rem);
        white-space: nowrap;
    }

    [data-testid="stDataFrame"] {
        border: 1px solid #D9E1EA;
    }

    [data-testid="stDownloadButton"] button {
        background-color: #FFFFFF !important;
        color: #123B6D !important;
        border: 2px solid #2F80ED !important;
        font-weight: 700 !important;
    }

    [data-testid="stDownloadButton"] button p,
    [data-testid="stDownloadButton"] button span,
    [data-testid="stDownloadButton"] button div {
        color: #123B6D !important;
        font-weight: 700 !important;
    }

    [data-testid="stDownloadButton"] button:hover {
        background-color: #EEF5FF !important;
        border-color: #1F5FAF !important;
    }

    .stButton > button {
        background-color: #163A5F !important;
        color: #FFFFFF !important;
        border: 1px solid #163A5F !important;
        font-weight: 700 !important;
    }

    .stButton > button p,
    .stButton > button span,
    .stButton > button div,
    .stButton > button * {
        color: #FFFFFF !important;
        font-weight: 700 !important;
        opacity: 1 !important;
    }

    [data-testid="stDownloadButton"] button,
    [data-testid="stDownloadButton"] button * {
        color: #123B6D !important;
        opacity: 1 !important;
    }

    .stButton > button:hover {
        background-color: #2F5D8A !important;
        color: #FFFFFF !important;
        border-color: #2F5D8A !important;
    }
    </style>
    """,
    unsafe_allow_html=True,
)



BASE_DIR = Path(__file__).resolve().parent
DATA_FILE = BASE_DIR / "datos_ejemplo.xlsx"


@st.cache_data
def cargar_datos() -> pd.DataFrame:
    """Carga y normaliza el Excel incluido en el repositorio."""
    if not DATA_FILE.exists():
        raise FileNotFoundError(
            f"No se encontró el archivo: {DATA_FILE.name}. "
            "Debe estar en la misma carpeta que app.py."
        )

    datos = pd.read_excel(DATA_FILE)
    datos["Fecha"] = pd.to_datetime(datos["Fecha"])

    numeric_cols = [
        "Ventas", "Unidades", "Clientes", "Costo",
        "Ganancia", "Margen", "Descuentos", "Ticket_Promedio"
    ]
    for col in numeric_cols:
        datos[col] = pd.to_numeric(datos[col], errors="coerce")

    return datos


def moneda(valor: float) -> str:
    return f"${valor:,.2f}"


def calcular_kpis(datos: pd.DataFrame) -> dict[str, float]:
    ventas_totales = datos["Ventas"].sum()
    return {
        "Ventas totales": ventas_totales,
        "Ganancia total": datos["Ganancia"].sum(),
        "Margen ponderado": (
            datos["Ganancia"].sum() / ventas_totales if ventas_totales else 0
        ),
        "Ticket promedio": datos["Ticket_Promedio"].mean(),
        "Unidades vendidas": datos["Unidades"].sum(),
        "Clientes": datos["Clientes"].sum(),
        "Descuentos": datos["Descuentos"].sum(),
    }


def resumen_por(datos: pd.DataFrame, columna: str) -> pd.DataFrame:
    resumen = datos.groupby(columna, as_index=False).agg(
        Ventas=("Ventas", "sum"),
        Ganancia=("Ganancia", "sum"),
        Costo=("Costo", "sum"),
        Descuentos=("Descuentos", "sum"),
        Ticket_Promedio=("Ticket_Promedio", "mean"),
    )
    resumen["Margen"] = resumen["Ganancia"] / resumen["Ventas"]
    return resumen


def configurar_figura(figura):
    figura.update_layout(
        template="plotly_white",
        paper_bgcolor="#FFFFFF",
        plot_bgcolor="#FFFFFF",
        font=dict(color="#1F2937"),
        margin=dict(l=20, r=20, t=55, b=20),
        height=420,
    )
    colores_dashboard = [
        "#1976D2", "#2EAD62", "#F39C12", "#8E5AC7",
        "#E85D75", "#00A6A6", "#5C6BC0", "#6C757D",
    ]
    for i, trace in enumerate(figura.data):
        if hasattr(trace, "marker") and trace.marker is not None:
            trace.marker.color = colores_dashboard[i % len(colores_dashboard)]
        if hasattr(trace, "line") and trace.line is not None:
            trace.line.color = colores_dashboard[i % len(colores_dashboard)]
        if getattr(trace, "text", None) is not None:
            trace.textfont = dict(size=18, color="#1F2937")
            if getattr(trace, "orientation", None) == "h":
                trace.textposition = "outside"
            else:
                trace.textposition = "outside"
    figura.update_xaxes(
        showgrid=True,
        gridcolor="#E5E7EB",
        zeroline=False,
        tickfont=dict(color="#4B5563"),
        title_font=dict(color="#1F2937"),
    )
    figura.update_yaxes(
        showgrid=True,
        gridcolor="#E5E7EB",
        zeroline=False,
        tickfont=dict(color="#4B5563"),
        title_font=dict(color="#1F2937"),
    )
    return figura


def mostrar_grafico(figura):
    st.plotly_chart(configurar_figura(figura), use_container_width=True)


def preparar_analisis(datos: pd.DataFrame):
    resumen_mensual = (
        datos.assign(Mes=datos["Fecha"].dt.to_period("M").astype(str))
        .groupby("Mes", as_index=False)
        .agg(
            Ventas=("Ventas", "sum"),
            Ganancia=("Ganancia", "sum"),
            Costo=("Costo", "sum"),
        )
    )
    resumen_categoria = resumen_por(datos, "Categoria")
    resumen_region = resumen_por(datos, "Region")
    return resumen_mensual, resumen_categoria, resumen_region


def crear_grafica_png(figura, titulo: str = "") -> BytesIO:
    """Convierte un gráfico Plotly en PNG para incorporarlo al PDF."""
    imagen = BytesIO()
    figura.write_image(
        imagen,
        format="png",
        width=1200,
        height=650,
        scale=1,
    )
    imagen.seek(0)
    return imagen


def crear_pdf(datos: pd.DataFrame) -> bytes:
    """
    Genera un informe ejecutivo dinámico con los datos actualmente filtrados.
    Incluye resumen ejecutivo, KPIs, principales hallazgos, explicación de cada
    visualización y recomendaciones accionables.
    """
    kpis = calcular_kpis(datos)
    resumen_mensual, resumen_categoria, resumen_region = preparar_analisis(datos)

    # -----------------------------
    # Cálculos para el análisis narrativo
    # -----------------------------
    ventas_totales = kpis["Ventas totales"]
    ganancia_total = kpis["Ganancia total"]
    margen = kpis["Margen ponderado"]

    mes_mejor_ventas = (
        resumen_mensual.loc[resumen_mensual["Ventas"].idxmax()]
        if not resumen_mensual.empty else None
    )
    mes_menor_ventas = (
        resumen_mensual.loc[resumen_mensual["Ventas"].idxmin()]
        if not resumen_mensual.empty else None
    )

    categoria_top = (
        resumen_categoria.loc[resumen_categoria["Ventas"].idxmax()]
        if not resumen_categoria.empty else None
    )
    categoria_margen = (
        resumen_categoria.loc[resumen_categoria["Margen"].idxmax()]
        if not resumen_categoria.empty else None
    )
    categoria_margen_bajo = (
        resumen_categoria.loc[resumen_categoria["Margen"].idxmin()]
        if not resumen_categoria.empty else None
    )

    region_top = (
        resumen_region.loc[resumen_region["Ventas"].idxmax()]
        if not resumen_region.empty else None
    )
    region_margen = (
        resumen_region.loc[resumen_region["Margen"].idxmax()]
        if not resumen_region.empty else None
    )
    region_margen_bajo = (
        resumen_region.loc[resumen_region["Margen"].idxmin()]
        if not resumen_region.empty else None
    )

    producto_ventas = (
        datos.groupby("Producto", as_index=False)
        .agg(Ventas=("Ventas", "sum"), Ganancia=("Ganancia", "sum"))
        .sort_values("Ventas", ascending=False)
    )
    producto_top = producto_ventas.iloc[0] if not producto_ventas.empty else None

    # Variación entre primer y último mes disponible.
    if len(resumen_mensual) >= 2:
        ventas_inicio = resumen_mensual.iloc[0]["Ventas"]
        ventas_final = resumen_mensual.iloc[-1]["Ventas"]
        variacion_periodo = (
            (ventas_final - ventas_inicio) / ventas_inicio
            if ventas_inicio else 0
        )
    else:
        variacion_periodo = 0

    # Relación entre descuentos y margen.
    if len(datos) >= 2 and datos["Descuentos"].nunique() > 1 and datos["Ventas"].nunique() > 1:
        margen_fila = (datos["Ganancia"] / datos["Ventas"]).replace([float("inf"), -float("inf")], 0).fillna(0)
        corr_descuento_margen = datos["Descuentos"].corr(margen_fila)
    else:
        corr_descuento_margen = 0

    # -----------------------------
    # Gráficas del informe
    # -----------------------------
    colores = {
        "Ventas": "#1F77B4",
        "Ganancia": "#2CA02C",
        "Margen": "#F28E2B",
        "Accesorios": "#2F80ED",
        "Oficina": "#6C757D",
        "Tecnologia": "#163A5F",
        "Movilidad": "#8E5AC7",
        "0–200": "#2F80ED",
        "201–400": "#5B8FF9",
        "401–600": "#7AA6E6",
        "601–800": "#A8BEDB",
        "801+": "#C7D4E3",
    }

    # ---------------------------------------------------------------
    # Visualizaciones ejecutivas.
    # Los colores se fijan directamente sobre los traces para evitar
    # que Kaleido use negro como color por defecto.
    # ---------------------------------------------------------------
    palette = [
        "#1976D2", "#2EAD62", "#F39C12", "#8E5AC7",
        "#E85D75", "#00A6A6", "#5C6BC0", "#6C757D",
    ]

    ticket_pdf = (
        datos.groupby("Categoria", as_index=False)["Ticket_Promedio"]
        .mean()
        .sort_values("Ticket_Promedio", ascending=False)
    )

    descuento_pdf = (
        datos.assign(
            Rango_Descuento=pd.cut(
                datos["Descuentos"],
                bins=[-1, 200, 400, 600, 800, float("inf")],
                labels=["0–200", "201–400", "401–600", "601–800", "801+"],
            ),
            Margen_calculado=(
                datos["Ganancia"] / datos["Ventas"]
            ).replace([float("inf"), -float("inf")], 0).fillna(0),
        )
        .groupby("Rango_Descuento", observed=False, as_index=False)
        .agg(
            Margen_Promedio=("Margen_calculado", "mean"),
            Ventas=("Ventas", "sum"),
            Registros=("Ventas", "size"),
        )
    )

    figs = []

    fig = px.line(
        resumen_mensual,
        x="Mes",
        y=["Ventas", "Ganancia"],
        title="1. Evolución mensual de ventas y ganancia",
        labels={"value": "Monto", "variable": "Indicador", "Mes": "Mes"},
        markers=True,
    )
    fig.update_traces(
        selector={"name": "Ventas"},
        line=dict(color="#1976D2", width=3),
        marker=dict(color="#1976D2", size=8), textfont=dict(size=19),
    )
    fig.update_traces(
        selector={"name": "Ganancia"},
        line=dict(color="#2EAD62", width=3),
        marker=dict(color="#2EAD62", size=8), textfont=dict(size=19),
    )
    figs.append(fig)

    fig = px.bar(
        resumen_categoria,
        x="Categoria",
        y=["Ventas", "Ganancia"],
        barmode="group",
        title="2. Ventas y ganancia por categoría",
        labels={"value": "Monto", "variable": "Indicador"},
    )
    fig.update_traces(
        selector={"name": "Ventas"},
        marker_color="#1976D2",
    )
    fig.update_traces(
        selector={"name": "Ganancia"},
        marker_color="#2EAD62",
    )
    figs.append(fig)

    fig = px.bar(
        resumen_categoria.sort_values("Margen"),
        x="Categoria",
        y="Margen",
        title="3. Margen por categoría",
        labels={"Margen": "Margen", "Categoria": "Categoría"},
    )
    fig.update_traces(
        marker_color=[
            palette[i % len(palette)] for i in range(len(fig.data))
        ]
    )
    # Para este gráfico cada barra es un trace; si Plotly agrupa de otra
    # forma, se vuelve a aplicar el color por índice.
    for i, trace in enumerate(fig.data):
        trace.marker.color = palette[i % len(palette)]
        trace.text = [f"{v:.1%}" for v in resumen_categoria.sort_values("Margen")["Margen"]]
        trace.textposition = "outside"
        trace.textfont = dict(size=19, color="#1F2937")
    figs.append(fig)

    fig = px.bar(
        resumen_region,
        x="Region",
        y=["Ventas", "Ganancia"],
        barmode="group",
        title="4. Ventas y ganancia por región",
        labels={"value": "Monto", "variable": "Indicador"},
    )
    fig.update_traces(
        selector={"name": "Ventas"},
        marker_color="#1976D2",
    )
    fig.update_traces(
        selector={"name": "Ganancia"},
        marker_color="#2EAD62",
    )
    figs.append(fig)

    fig = px.bar(
        producto_ventas.head(10).sort_values("Ventas"),
        x="Ventas",
        y="Producto",
        orientation="h",
        title="5. Productos con mayor facturación",
        labels={"Ventas": "Ventas", "Producto": "Producto"},
        text_auto=".2s",
    )
    for i, trace in enumerate(fig.data):
        trace.marker.color = palette[i % len(palette)]
    figs.append(fig)

    fig = px.bar(
        descuento_pdf,
        x="Rango_Descuento",
        y="Margen_Promedio",
        title="6. Margen promedio por rango de descuento",
        labels={
            "Rango_Descuento": "Rango de descuento",
            "Margen_Promedio": "Margen promedio",
        },
        text_auto=".1%",
    )
    for i, trace in enumerate(fig.data):
        trace.marker.color = palette[i % len(palette)]
    figs.append(fig)

    fig = px.bar(
        ticket_pdf,
        x="Categoria",
        y="Ticket_Promedio",
        title="7. Ticket promedio por categoría",
        labels={"Categoria": "Categoría", "Ticket_Promedio": "Ticket promedio"},
        text_auto=".2f",
    )
    for i, trace in enumerate(fig.data):
        trace.marker.color = palette[i % len(palette)]
    figs.append(fig)

    for fig in figs:
        fig.update_layout(
            template="plotly_white",
            paper_bgcolor="#FFFFFF",
            plot_bgcolor="#FFFFFF",
            font=dict(color="#1F2937", size=18),
            title_font=dict(color="#163A5F", size=24),
            margin=dict(l=80, r=55, t=90, b=70),
            height=640,
            legend=dict(
                bgcolor="rgba(255,255,255,0.95)",
                font=dict(color="#1F2937", size=16),
            ),
        )
        fig.update_xaxes(
            showgrid=True,
            gridcolor="#E5E7EB",
            zeroline=False,
            tickfont=dict(color="#4B5563", size=19),
            title_font=dict(color="#1F2937", size=18),
        )
        fig.update_yaxes(
            showgrid=True,
            gridcolor="#E5E7EB",
            zeroline=False,
            tickfont=dict(color="#4B5563", size=19),
            title_font=dict(color="#1F2937", size=18),
        )
    # -----------------------------
    # Estilos del informe ejecutivo
    # -----------------------------
    styles = getSampleStyleSheet()
    styles.add(
        ParagraphStyle(
            name="ExecutiveTitle",
            parent=styles["Title"],
            fontSize=24,
            leading=29,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#163A5F"),
            spaceAfter=10,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ExecutiveSubtitle",
            parent=styles["Normal"],
            fontSize=11,
            leading=16,
            alignment=TA_CENTER,
            textColor=colors.HexColor("#5B6573"),
            spaceAfter=18,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SectionTitle",
            parent=styles["Heading2"],
            fontSize=16,
            leading=20,
            textColor=colors.HexColor("#163A5F"),
            spaceBefore=8,
            spaceAfter=8,
        )
    )
    styles.add(
        ParagraphStyle(
            name="Subsection",
            parent=styles["Heading3"],
            fontSize=12,
            leading=15,
            textColor=colors.HexColor("#2F5D8A"),
            spaceBefore=8,
            spaceAfter=5,
        )
    )
    styles.add(
        ParagraphStyle(
            name="ExecutiveBody",
            parent=styles["BodyText"],
            fontSize=9.5,
            leading=14,
            textColor=colors.HexColor("#30343B"),
            spaceAfter=7,
        )
    )
    styles.add(
        ParagraphStyle(
            name="SmallNote",
            parent=styles["BodyText"],
            fontSize=8,
            leading=11,
            textColor=colors.HexColor("#68707D"),
        )
    )

    def money(v):
        return f"${v:,.2f}"

    def pct(v):
        return f"{v:.2%}"

    def safe(v):
        return "N/D" if pd.isna(v) else str(v)

    # -----------------------------
    # Documento
    # -----------------------------
    buffer = BytesIO()

    def footer(canvas, doc):
        canvas.saveState()
        canvas.setStrokeColor(colors.HexColor("#D9DEE5"))
        canvas.line(0.65 * inch, 0.48 * inch, 7.85 * inch, 0.48 * inch)
        canvas.setFont("Helvetica", 8)
        canvas.setFillColor(colors.HexColor("#68707D"))
        canvas.drawString(
            0.65 * inch,
            0.30 * inch,
            "Informe ejecutivo de ventas y rentabilidad",
        )
        canvas.drawRightString(
            7.85 * inch,
            0.30 * inch,
            f"Página {doc.page}",
        )
        canvas.restoreState()

    doc = SimpleDocTemplate(
        buffer,
        pagesize=letter,
        rightMargin=0.65 * inch,
        leftMargin=0.65 * inch,
        topMargin=0.65 * inch,
        bottomMargin=0.65 * inch,
        title="Informe Ejecutivo de Ventas y Rentabilidad",
        author="Empresa_de_tegnologia",
    )

    story = []

    # Portada / resumen ejecutivo
    story.extend(
        [
            Spacer(1, 0.35 * inch),
            Paragraph(
                "Informe Ejecutivo de Ventas y Rentabilidad",
                styles["ExecutiveTitle"],
            ),
            Paragraph(
                "Análisis gerencial basado en los datos actualmente seleccionados en el dashboard",
                styles["ExecutiveSubtitle"],
            ),
        ]
    )

    fecha_inicio = datos["Fecha"].min().strftime("%d/%m/%Y")
    fecha_fin = datos["Fecha"].max().strftime("%d/%m/%Y")

    resumen_portada = (
        f"<b>Periodo analizado:</b> {fecha_inicio} al {fecha_fin}<br/>"
        f"<b>Registros:</b> {len(datos):,}<br/>"
        f"<b>Categorías:</b> {datos['Categoria'].nunique()} &nbsp;&nbsp; "
        f"<b>Regiones:</b> {datos['Region'].nunique()} &nbsp;&nbsp; "
        f"<b>Productos:</b> {datos['Producto'].nunique()}"
    )
    story.append(Paragraph(resumen_portada, styles["ExecutiveBody"]))
    story.append(Spacer(1, 10))

    # Tabla ejecutiva de KPIs
    kpi_data = [
        ["Indicador", "Resultado", "Lectura ejecutiva"],
        ["Ventas totales", money(ventas_totales), "Volumen total facturado"],
        ["Ganancia total", money(ganancia_total), "Resultado acumulado"],
        ["Margen ponderado", pct(margen), "Rentabilidad sobre ventas"],
        ["Ticket promedio", money(kpis["Ticket promedio"]), "Valor promedio por operación"],
        ["Unidades vendidas", f"{kpis['Unidades vendidas']:,.0f}", "Volumen de unidades"],
        ["Clientes", f"{kpis['Clientes']:,.0f}", "Base acumulada registrada"],
        ["Descuentos", money(kpis["Descuentos"]), "Importe total descontado"],
    ]

    tabla_kpi = Table(kpi_data, colWidths=[1.45 * inch, 1.35 * inch, 4.0 * inch])
    tabla_kpi.setStyle(
        TableStyle(
            [
                ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor("#163A5F")),
                ("TEXTCOLOR", (0, 0), (-1, 0), colors.white),
                ("FONTNAME", (0, 0), (-1, 0), "Helvetica-Bold"),
                ("FONTNAME", (0, 1), (0, -1), "Helvetica-Bold"),
                ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#F7F9FB")),
                ("TEXTCOLOR", (0, 1), (-1, -1), colors.HexColor("#30343B")),
                ("GRID", (0, 0), (-1, -1), 0.35, colors.HexColor("#D9DEE5")),
                ("VALIGN", (0, 0), (-1, -1), "MIDDLE"),
                ("ALIGN", (1, 1), (1, -1), "RIGHT"),
                ("FONTSIZE", (0, 0), (-1, -1), 8.5),
                ("LEADING", (0, 0), (-1, -1), 11),
                ("TOPPADDING", (0, 0), (-1, -1), 7),
                ("BOTTOMPADDING", (0, 0), (-1, -1), 7),
            ]
        )
    )
    story.append(tabla_kpi)
    story.append(Spacer(1, 14))

    # Resumen ejecutivo
    resumen_general = (
        f"El conjunto analizado registra <b>{money(ventas_totales)}</b> en ventas y "
        f"<b>{money(ganancia_total)}</b> de ganancia, con un margen ponderado de "
        f"<b>{pct(margen)}</b>. El resultado debe interpretarse conjuntamente con "
        f"la composición por categoría, región, producto y nivel de descuentos, "
        f"ya que el volumen de ventas por sí solo no representa toda la rentabilidad."
    )
    story.append(Paragraph("Resumen ejecutivo", styles["SectionTitle"]))
    story.append(Paragraph(resumen_general, styles["ExecutiveBody"]))

    if mes_mejor_ventas is not None:
        story.append(
            Paragraph(
                f"En la evolución temporal, <b>{safe(mes_mejor_ventas['Mes'])}</b> "
                f"presenta el mayor nivel de ventas ({money(mes_mejor_ventas['Ventas'])}), "
                f"mientras que <b>{safe(mes_menor_ventas['Mes'])}</b> registra el menor "
                f"({money(mes_menor_ventas['Ventas'])}). Entre el primer y el último "
                f"mes disponible, las ventas presentan una variación de "
                f"<b>{pct(variacion_periodo)}</b>.",
                styles["ExecutiveBody"],
            )
        )

    if categoria_top is not None:
        story.append(
            Paragraph(
                f"<b>{safe(categoria_top['Categoria'])}</b> es la categoría con mayor "
                f"facturación, con {money(categoria_top['Ventas'])} y un margen de "
                f"{pct(categoria_top['Margen'])}. La categoría con mayor margen es "
                f"<b>{safe(categoria_margen['Categoria'])}</b> ({pct(categoria_margen['Margen'])}), "
                f"lo que permite diferenciar claramente entre liderazgo en volumen y "
                f"liderazgo en rentabilidad.",
                styles["ExecutiveBody"],
            )
        )

    if region_top is not None:
        story.append(
            Paragraph(
                f"Por región, <b>{safe(region_top['Region'])}</b> concentra el mayor "
                f"volumen de ventas ({money(region_top['Ventas'])}), mientras que "
                f"<b>{safe(region_margen['Region'])}</b> presenta el margen regional "
                f"más alto ({pct(region_margen['Margen'])}). Esta diferencia es útil "
                f"para evaluar si el crecimiento comercial está acompañado por una "
                f"rentabilidad equivalente.",
                styles["ExecutiveBody"],
            )
        )

    # Principales hallazgos
    story.append(Paragraph("Principales hallazgos", styles["SectionTitle"]))

    hallazgos = [
        (
            "1. Dinámica temporal",
            (
                f"El comportamiento mensual muestra variaciones relevantes en el "
                f"volumen de ventas. El máximo se observa en {safe(mes_mejor_ventas['Mes'])} "
                f"y el mínimo en {safe(mes_menor_ventas['Mes'])}. La comparación entre "
                f"periodos permite identificar meses que requieren explicación comercial "
                f"y meses que pueden servir como referencia para planificación."
            )
        ),
        (
            "2. Concentración de ventas",
            (
                f"La categoría {safe(categoria_top['Categoria'])} lidera el volumen "
                f"facturado con {money(categoria_top['Ventas'])}. Esto señala dónde se "
                f"concentra actualmente la generación de ingresos y permite evaluar la "
                f"dependencia del resultado respecto de determinadas líneas de negocio."
            )
        ),
        (
            "3. Volumen versus rentabilidad",
            (
                f"La categoría con mayor margen es {safe(categoria_margen['Categoria'])}, "
                f"mientras que {safe(categoria_margen_bajo['Categoria'])} presenta el "
                f"menor margen ({pct(categoria_margen_bajo['Margen'])}). Por tanto, una "
                f"línea con menor volumen puede aportar proporcionalmente más rentabilidad "
                f"que una línea con mayor facturación."
            )
        ),
        (
            "4. Desempeño regional",
            (
                f"{safe(region_top['Region'])} lidera las ventas regionales, mientras "
                f"{safe(region_margen['Region'])} lidera en margen. La diferencia entre "
                f"ambos indicadores sugiere analizar cada mercado no solo por ventas, "
                f"sino también por su capacidad de convertir ventas en ganancia."
            )
        ),
        (
            "5. Productos relevantes",
            (
                f"El producto con mayor facturación acumulada es {safe(producto_top['Producto'])}, "
                f"con {money(producto_top['Ventas'])}. Su contribución debe analizarse "
                f"junto con su ganancia y margen antes de utilizarlo como referencia para "
                f"decisiones de expansión o inventario."
            )
        ),
        (
            "6. Descuentos y margen",
            (
                f"La relación estadística entre descuentos y margen en los registros "
                f"analizados presenta un coeficiente de correlación de "
                f"<b>{corr_descuento_margen:.2f}</b>. Este valor es descriptivo y no "
                f"demuestra causalidad; sirve como señal para profundizar el análisis "
                f"de política de descuentos, precios y rentabilidad."
            )
        ),
    ]

    for titulo, detalle in hallazgos:
        story.append(Paragraph(titulo, styles["Subsection"]))
        story.append(Paragraph(detalle, styles["ExecutiveBody"]))

    # Página de análisis visual
    story.append(PageBreak())
    story.append(Paragraph("Análisis visual y lectura de las gráficas", styles["SectionTitle"]))

    explicaciones = [
        (
            figs[0],
            "Evolución mensual",
            (
                "Esta gráfica permite observar simultáneamente la trayectoria de ventas "
                "y ganancia durante el periodo. La lectura recomendada es identificar "
                "picos y caídas, y posteriormente verificar si los cambios de ventas "
                "se acompañan de movimientos similares en la ganancia. Una separación "
                "creciente entre ambas series puede justificar una revisión de costos "
                "o de la mezcla de productos."
            ),
        ),
        (
            figs[1],
            "Ventas y ganancia por categoría",
            (
                "La comparación por categoría permite distinguir qué líneas aportan "
                "mayor volumen y cuáles convierten ese volumen en ganancia. La categoría "
                "líder en ventas no necesariamente es la líder en ganancia, por lo que "
                "ambos indicadores deben evaluarse conjuntamente."
            ),
        ),
        (
            figs[2],
            "Margen por categoría",
            (
                "El gráfico ordena visualmente la rentabilidad relativa de las categorías. "
                "Las diferencias de margen son relevantes para revisar precios, costos, "
                "descuentos y mezcla comercial. Las categorías de margen bajo requieren "
                "un análisis específico antes de incrementar su volumen."
            ),
        ),
        (
            figs[3],
            "Ventas y ganancia por región",
            (
                "Esta visualización compara el desempeño comercial entre regiones. "
                "Permite identificar mercados que generan mucho volumen y mercados que "
                "destacan por su rentabilidad. La combinación de ambos indicadores ofrece "
                "una lectura más completa del desempeño territorial."
            ),
        ),
        (
            figs[4],
            "Productos con mayor facturación",
            (
                "La gráfica ordena los productos por volumen de ventas y permite "
                "identificar rápidamente los principales generadores de facturación. "
                "Esta lectura facilita priorizar productos para análisis de inventario, "
                "disponibilidad, estrategia comercial y rentabilidad. La facturación "
                "debe complementarse con el margen antes de tomar decisiones."
            ),
        ),
        (
            figs[5],
            "Margen promedio por rango de descuento",
            (
                "En lugar de mostrar cientos de observaciones individuales, esta "
                "visualización agrupa los descuentos en rangos y calcula el margen "
                "promedio de cada grupo. Esto facilita identificar si los niveles de "
                "descuento más elevados están asociados con una menor rentabilidad. "
                "El patrón es descriptivo y no demuestra causalidad."
            ),
        ),
        (
            figs[6],
            "Ticket promedio por categoría",
            (
                "El ticket promedio ayuda a entender el valor medio de las operaciones "
                "en cada categoría. Diferencias importantes pueden orientar estrategias "
                "de venta cruzada, paquetes comerciales, segmentación y priorización "
                "de clientes."
            ),
        ),
    ]

    for i, (fig, titulo, explicacion) in enumerate(explicaciones):
        if i > 0:
            story.append(PageBreak())
        story.append(Paragraph(titulo, styles["Subsection"]))
        story.append(Paragraph(explicacion, styles["ExecutiveBody"]))
        story.append(Image(crear_grafica_png(fig), width=6.9 * inch, height=3.74 * inch))

    story.append(Paragraph("Nota metodológica", styles["SectionTitle"]))
    story.append(
        Paragraph(
            "Este informe es generado automáticamente a partir de los registros "
            "seleccionados en el dashboard. Las conclusiones son descriptivas y se "
            "limitan a los indicadores disponibles en el conjunto de datos. Las "
            "correlaciones no deben interpretarse como relaciones causales. Para "
            "decisiones de negocio se recomienda complementar este análisis con "
            "información de costos, precios, clientes, inventario y contexto comercial "
            "cuando esté disponible.",
            styles["SmallNote"],
        )
    )


    # ---------------------------------------------------------------
    # RECOMENDACIONES EJECUTIVAS
    # Se colocan AL FINAL del informe, después de todas las gráficas.
    # La lógica combina datos observados + escenarios de gestión.
    # Los escenarios son referencias de decisión y no pronósticos.
    # ---------------------------------------------------------------
    margen_portafolio = (
        ganancia_total / ventas_totales if ventas_totales else 0
    )

    # Margen por registro para definir niveles de control.
    margen_registro = (
        (datos["Ganancia"] / datos["Ventas"])
        .replace([float("inf"), -float("inf")], 0)
        .fillna(0)
    )
    margen_p25 = margen_registro.quantile(0.25)

    # Descuentos por bandas.
    datos_rec = datos.copy()
    datos_rec["Margen_calculado"] = margen_registro
    datos_rec["Rango_Descuento"] = pd.cut(
        datos_rec["Descuentos"],
        bins=[-1, 200, 400, 600, 800, float("inf")],
        labels=["0–200", "201–400", "401–600", "601–800", "801+"],
    )

    desc_rec = (
        datos_rec.groupby("Rango_Descuento", observed=False)
        .agg(
            Ventas=("Ventas", "sum"),
            Ganancia=("Ganancia", "sum"),
            Margen=("Margen_calculado", "mean"),
            Registros=("Ventas", "size"),
        )
        .reset_index()
    )

    margen_base_desc = float(
        desc_rec.loc[
            desc_rec["Rango_Descuento"] == "0–200", "Margen"
        ].iloc[0]
    ) if (desc_rec["Rango_Descuento"] == "0–200").any() else margen_portafolio

    # Escenario: si las ventas se mantienen iguales y el margen de los
    # grupos de descuento >200 se recuperara hasta el margen observado
    # en 0–200, cuánto ganancia adicional representaría.
    desc_rec["Ganancia_escenario"] = (
        desc_rec["Ventas"] * margen_base_desc
    )
    desc_rec["Mejora_escenario"] = (
        desc_rec["Ganancia_escenario"] - desc_rec["Ganancia"]
    ).clip(lower=0)

    mejora_desc_total = desc_rec.loc[
        desc_rec["Rango_Descuento"].isin(["201–400", "401–600", "601–800", "801+"]),
        "Mejora_escenario",
    ].sum()

    # Productos.
    producto_det = (
        datos.groupby("Producto", as_index=False)
        .agg(
            Ventas=("Ventas", "sum"),
            Ganancia=("Ganancia", "sum"),
            Descuento=("Descuentos", "mean"),
            Ticket=("Ticket_Promedio", "mean"),
        )
    )
    producto_det["Margen"] = producto_det["Ganancia"] / producto_det["Ventas"]
    producto_det = producto_det.sort_values("Ventas", ascending=False)

    top3 = producto_det.head(3)
    top3_participacion = (
        top3["Ventas"].sum() / ventas_totales if ventas_totales else 0
    )

    # Producto líder de facturación y producto de alto volumen con menor margen.
    producto_lider = producto_det.iloc[0] if not producto_det.empty else None
    producto_menor_margen = (
        producto_det.sort_values("Margen").iloc[0]
        if not producto_det.empty else None
    )

    # Categorías y regiones.
    cat_rec = (
        datos.groupby("Categoria", as_index=False)
        .agg(Ventas=("Ventas", "sum"), Ganancia=("Ganancia", "sum"))
    )
    cat_rec["Margen"] = cat_rec["Ganancia"] / cat_rec["Ventas"]

    cat_menor = cat_rec.sort_values("Margen").iloc[0] if not cat_rec.empty else None
    cat_mayor = cat_rec.sort_values("Margen", ascending=False).iloc[0] if not cat_rec.empty else None

    reg_rec = (
        datos.groupby("Region", as_index=False)
        .agg(Ventas=("Ventas", "sum"), Ganancia=("Ganancia", "sum"))
    )
    reg_rec["Margen"] = reg_rec["Ganancia"] / reg_rec["Ventas"]
    reg_rec["Participacion"] = reg_rec["Ventas"] / ventas_totales if ventas_totales else 0

    region_ventas = reg_rec.sort_values("Ventas", ascending=False).iloc[0] if not reg_rec.empty else None
    region_margen = reg_rec.sort_values("Margen", ascending=False).iloc[0] if not reg_rec.empty else None
    region_menor_margen = reg_rec.sort_values("Margen").iloc[0] if not reg_rec.empty else None

    # Tickets por categoría.
    ticket_cat = (
        datos.groupby("Categoria", as_index=False)["Ticket_Promedio"]
        .mean()
        .sort_values("Ticket_Promedio", ascending=False)
    )
    ticket_alto = ticket_cat.iloc[0] if not ticket_cat.empty else None
    ticket_bajo = ticket_cat.iloc[-1] if not ticket_cat.empty else None

    # Escenario financiero para producto de menor margen:
    # llevarlo al margen total del portafolio, manteniendo ventas constantes.
    mejora_producto_menor = 0
    if producto_menor_margen is not None:
        mejora_producto_menor = max(
            0,
            producto_menor_margen["Ventas"]
            * (margen_portafolio - producto_menor_margen["Margen"])
        )

    # Escenario para la categoría de menor margen:
    mejora_categoria_menor = 0
    if cat_menor is not None:
        mejora_categoria_menor = max(
            0,
            cat_menor["Ventas"]
            * (margen_portafolio - cat_menor["Margen"])
        )

    # ---------------------------------------------------------------
    # Página final de recomendaciones
    # ---------------------------------------------------------------
    story.append(PageBreak())
    story.append(
        Paragraph(
            "Recomendaciones para la toma de decisiones",
            styles["SectionTitle"],
        )
    )
    story.append(
        Paragraph(
            "Esta sección convierte los resultados del análisis en decisiones "
            "concretas. Los importes son calculados con los datos filtrados. "
            "Los escenarios de mejora son referencias para evaluar oportunidades; "
            "no representan una garantía de resultados.",
            styles["ExecutiveBody"],
        )
    )

    rec_style = ParagraphStyle(
        "RecommendationTextFinal",
        parent=styles["ExecutiveBody"],
        fontSize=8.8,
        leading=12,
        spaceAfter=0,
        textColor=colors.HexColor("#30343B"),
    )
    rec_title_style = ParagraphStyle(
        "RecommendationTitleFinal",
        parent=styles["ExecutiveBody"],
        fontSize=9,
        leading=12,
        spaceAfter=0,
        textColor=colors.HexColor("#163A5F"),
    )

    financiera = [
        (
            "1. Poner un límite económico a los descuentos",
            (
                f"Usar $200 como descuento estándar. Los registros de $0–$200 tienen un "
                f"margen promedio de {pct(margen_base_desc)}. Cuando el descuento supera "
                f"$200, el margen baja a {pct(float(desc_rec.loc[desc_rec['Rango_Descuento']=='201–400','Margen'].iloc[0]))} "
                f"en $201–$400 y llega a {pct(float(desc_rec.loc[desc_rec['Rango_Descuento']=='801+','Margen'].iloc[0]))} "
                f"por encima de $800. Autorizar descuentos >$600 solo cuando exista una "
                f"razón comercial documentada."
            ),
        ),
        (
            "2. Medir cuánto dinero se está dejando en la mesa",
            (
                f"Como escenario de gestión, si las ventas se mantuvieran iguales y los "
                f"grupos con descuentos >$200 alcanzaran el margen observado en $0–$200, "
                f"la ganancia adicional teórica sería de aproximadamente "
                f"{money(mejora_desc_total)}. No es un pronóstico: es el tamaño de la "
                f"oportunidad que debe cuantificarse antes de cambiar precios."
            ),
        ),
        (
            "3. Recuperar margen en el producto de menor rentabilidad",
            (
                f"{safe(producto_menor_margen['Producto']) if producto_menor_margen is not None else 'El producto identificado'} "
                f"presenta un margen de {pct(producto_menor_margen['Margen']) if producto_menor_margen is not None else 'N/D'}. "
                f"Su venta acumulada es de {money(producto_menor_margen['Ventas']) if producto_menor_margen is not None else 'N/D'}. "
                f"Como meta financiera, llevarlo al margen total de {pct(margen_portafolio)} "
                f"representaría una mejora potencial de {money(mejora_producto_menor)} "
                f"manteniendo el mismo nivel de ventas. La revisión debe centrarse en precio, "
                f"costo unitario y descuentos."
            ),
        ),
        (
            "4. Atacar primero la categoría con menor margen",
            (
                f"{safe(cat_menor['Categoria']) if cat_menor is not None else 'La categoría identificada'} "
                f"genera {money(cat_menor['Ventas']) if cat_menor is not None else 'N/D'} "
                f"en ventas y tiene un margen de {pct(cat_menor['Margen']) if cat_menor is not None else 'N/D'}. "
                f"Subirla hasta el margen actual del portafolio ({pct(margen_portafolio)}) "
                f"representaría, solo como escenario, cerca de {money(mejora_categoria_menor)} "
                f"de ganancia adicional. Es un objetivo de revisión, no una cifra garantizada."
            ),
        ),
        (
            "5. Usar el margen como condición para aprobar crecimiento",
            (
                f"El portafolio tiene un margen ponderado de {pct(margen_portafolio)}. "
                f"Para nuevos descuentos, campañas o metas comerciales, establecer como "
                f"regla de control que el escenario aprobado muestre ventas, ganancia y "
                f"margen esperado. Operaciones por debajo de {pct(margen_p25)} deben "
                f"revisarse individualmente antes de aumentar su volumen."
            ),
        ),
    ]

    marketing = [
        (
            "1. Construir campañas alrededor de los productos que ya generan demanda",
            (
                f"{safe(producto_lider['Producto']) if producto_lider is not None else 'El producto líder'} "
                f"es el principal generador de ventas con {money(producto_lider['Ventas']) if producto_lider is not None else 'N/D'} "
                f"y un margen de {pct(producto_lider['Margen']) if producto_lider is not None else 'N/D'}. "
                f"Los 3 productos principales concentran {pct(top3_participacion)} de la facturación. "
                f"La primera línea de campañas debe enfocarse en captación y remarketing de estos "
                f"productos, manteniendo como KPI obligatorio el margen, no solo las ventas."
            ),
        ),
        (
            "2. No usar descuentos altos para vender el producto de menor margen",
            (
                f"{safe(producto_menor_margen['Producto']) if producto_menor_margen is not None else 'El producto de menor margen'} "
                f"tiene un margen de {pct(producto_menor_margen['Margen']) if producto_menor_margen is not None else 'N/D'}. "
                f"En la primera prueba comercial, limitar el descuento a $200 y probar en su lugar "
                f"un beneficio de mayor valor percibido: paquete, accesorio, instalación o servicio. "
                f"El objetivo es aumentar conversión sin empujar el margen hacia los rangos de "
                f"descuento de mayor pérdida."
            ),
        ),
        (
            "3. Diseñar una campaña de aumento de ticket",
            (
                f"La categoría con mayor ticket promedio alcanza {money(ticket_alto['Ticket_Promedio']) if ticket_alto is not None else 'N/D'}, "
                f"mientras la menor alcanza {money(ticket_bajo['Ticket_Promedio']) if ticket_bajo is not None else 'N/D'}. "
                f"Para las categorías de menor ticket, probar paquetes y venta cruzada con una meta "
                f"inicial de +5% de ticket. Por ejemplo, sobre un ticket de {money(ticket_bajo['Ticket_Promedio']) if ticket_bajo is not None else 'N/D'}, "
                f"un aumento de 5% equivale a aproximadamente {money(ticket_bajo['Ticket_Promedio']*0.05) if ticket_bajo is not None else 'N/D'} "
                f"adicionales por operación."
            ),
        ),
        (
            "4. Distribuir el presupuesto inicial por región y luego moverlo por rentabilidad",
            (
                f"Como punto de partida para una prueba de medios, se puede distribuir el presupuesto "
                f"según la participación actual de ventas: {safe(region_ventas['Region']) if region_ventas is not None else 'región líder'} "
                f"tiene {pct(region_ventas['Participacion']) if region_ventas is not None else 'N/D'} del negocio. "
                f"Después de 30 días, mover presupuesto hacia las regiones que entreguen mayor "
                f"ganancia por dólar invertido. {safe(region_margen['Region']) if region_margen is not None else 'La región líder en margen'} "
                f"presenta actualmente un margen de {pct(region_margen['Margen']) if region_margen is not None else 'N/D'}."
            ),
        ),
        (
            "5. Convertir cada campaña en una prueba con una regla de decisión",
            (
                f"Usar tres reglas simples: mantener campañas con margen ≥{pct(margen_portafolio)}, "
                f"optimizar las que estén entre {pct(0.30)} y {pct(margen_portafolio)}, y revisar "
                f"las que estén por debajo de 30%. Además, comparar ventas incrementales, ticket "
                f"y descuento promedio. Así una campaña no continúa solo porque aumentó las ventas."
            ),
        ),
    ]

    def crear_tabla_recomendaciones(titulo, recomendaciones, fondo):
        filas = [
            [
                Paragraph(f"<b>{titulo}</b>", rec_title_style),
                Paragraph("<b>Qué hacer y qué medir</b>", rec_title_style),
            ]
        ]
        for nombre, accion in recomendaciones:
            filas.append(
                [
                    Paragraph(f"<b>{nombre}</b>", rec_title_style),
                    Paragraph(accion, rec_style),
                ]
            )

        tabla = Table(
            filas,
            colWidths=[1.85 * inch, 4.65 * inch],
            repeatRows=1,
            hAlign="LEFT",
        )
        tabla.setStyle(
            TableStyle(
                [
                    ("BACKGROUND", (0, 0), (-1, 0), colors.HexColor(fondo)),
                    ("TEXTCOLOR", (0, 0), (-1, 0), colors.HexColor("#163A5F")),
                    ("BACKGROUND", (0, 1), (-1, -1), colors.HexColor("#FAFBFC")),
                    ("GRID", (0, 0), (-1, -1), 0.4, colors.HexColor("#D4DCE5")),
                    ("VALIGN", (0, 0), (-1, -1), "TOP"),
                    ("LEFTPADDING", (0, 0), (-1, -1), 8),
                    ("RIGHTPADDING", (0, 0), (-1, -1), 8),
                    ("TOPPADDING", (0, 0), (-1, -1), 8),
                    ("BOTTOMPADDING", (0, 0), (-1, -1), 8),
                ]
            )
        )
        return tabla

    story.append(
        crear_tabla_recomendaciones(
            "FINANZAS Y CONTABILIDAD",
            financiera,
            "#EAF7EF",
        )
    )
    story.append(Spacer(1, 14))
    story.append(
        crear_tabla_recomendaciones(
            "MARKETING Y VENTAS",
            marketing,
            "#EAF3FF",
        )
    )

    story.append(Spacer(1, 10))

    doc.build(story, onFirstPage=footer, onLaterPages=footer)
    buffer.seek(0)
    return buffer.getvalue()


def main():
    datos = cargar_datos()

    st.title("Dashboard de Ventas y Rentabilidad")
    st.caption(
        "Análisis interactivo basado en datos_ejemplo.xlsx. "
        "El informe se genera con los filtros seleccionados."
    )

    with st.sidebar:
        st.header("Filtros")
        categorias = st.multiselect(
            "Categoría",
            sorted(datos["Categoria"].dropna().unique()),
            default=sorted(datos["Categoria"].dropna().unique()),
        )
        regiones = st.multiselect(
            "Región",
            sorted(datos["Region"].dropna().unique()),
            default=sorted(datos["Region"].dropna().unique()),
        )
        productos = st.multiselect(
            "Producto",
            sorted(datos["Producto"].dropna().unique()),
            default=sorted(datos["Producto"].dropna().unique()),
        )

        fecha_min = datos["Fecha"].min().date()
        fecha_max = datos["Fecha"].max().date()
        periodo = st.date_input(
            "Periodo",
            value=(fecha_min, fecha_max),
            min_value=fecha_min,
            max_value=fecha_max,
        )

    fecha_inicio, fecha_fin = (
        periodo if isinstance(periodo, tuple) else (periodo, periodo)
    )

    filtrados = datos[
        datos["Categoria"].isin(categorias)
        & datos["Region"].isin(regiones)
        & datos["Producto"].isin(productos)
        & datos["Fecha"].dt.date.between(fecha_inicio, fecha_fin)
    ].copy()

    if filtrados.empty:
        st.warning("No hay datos para los filtros seleccionados.")
        return

    kpis = calcular_kpis(filtrados)

    st.subheader("Indicadores clave")
    formatos = [
        moneda,
        moneda,
        lambda x: f"{x:.2%}",
        moneda,
        lambda x: f"{x:,.0f}",
        lambda x: f"{x:,.0f}",
        moneda,
    ]
    items_kpi = list(zip(kpis.items(), formatos))

    for fila in [items_kpi[:4], items_kpi[4:]]:
        tarjetas = st.columns(len(fila))
        for tarjeta, ((nombre, valor), formato) in zip(tarjetas, fila):
            tarjeta.metric(nombre, formato(valor))

    resumen_mensual, resumen_categoria, resumen_region = preparar_analisis(filtrados)

    st.subheader("Evolución y rendimiento")

    izquierda, derecha = st.columns(2)
    with izquierda:
        mostrar_grafico(
            px.line(
                resumen_mensual,
                x="Mes",
                y=["Ventas", "Ganancia"],
                title="Evolución mensual de ventas y ganancia",
            )
        )
    with derecha:
        mostrar_grafico(
            px.bar(
                resumen_categoria,
                x="Categoria",
                y="Ventas",
                title="Ventas por categoría",
                color="Categoria",
            )
        )

    izquierda, derecha = st.columns(2)
    with izquierda:
        mostrar_grafico(
            px.bar(
                resumen_categoria,
                x="Categoria",
                y="Margen",
                title="Margen por categoría",
                color="Categoria",
            )
        )
    with derecha:
        mostrar_grafico(
            px.bar(
                resumen_region,
                x="Region",
                y="Ventas",
                title="Ventas por región",
                color="Region",
            )
        )

    izquierda, derecha = st.columns(2)
    with izquierda:
        productos_dashboard = (
            filtrados.groupby("Producto", as_index=False)
            .agg(Ventas=("Ventas", "sum"))
            .sort_values("Ventas", ascending=False)
            .head(10)
            .sort_values("Ventas")
        )
        mostrar_grafico(
            px.bar(
                productos_dashboard,
                x="Ventas",
                y="Producto",
                orientation="h",
                title="Top 10 productos por facturación",
                text_auto=".2s",
            )
        )
    with derecha:
        filtrados_tmp = filtrados.copy()
        filtrados_tmp["Margen_calculado"] = (
            filtrados_tmp["Ganancia"] / filtrados_tmp["Ventas"]
        ).replace([float("inf"), -float("inf")], 0).fillna(0)
        filtrados_tmp["Rango_Descuento"] = pd.cut(
            filtrados_tmp["Descuentos"],
            bins=[-1, 200, 400, 600, 800, float("inf")],
            labels=["0–200", "201–400", "401–600", "601–800", "801+"],
        )
        descuento_dashboard = (
            filtrados_tmp.groupby("Rango_Descuento", observed=False, as_index=False)
            .agg(Margen=("Margen_calculado", "mean"))
        )
        mostrar_grafico(
            px.bar(
                descuento_dashboard,
                x="Rango_Descuento",
                y="Margen",
                title="Margen promedio por rango de descuento",
                text_auto=".1%",
            )
        )

    st.caption(
        "Ticket promedio = ventas totales ÷ número de operaciones. "
        "Sirve para saber cuánto dinero genera, en promedio, cada compra."
    )

    st.markdown("### Productos destacados")

    col_prod1, col_prod2 = st.columns(2)

    with col_prod1:
        productos_mas_vendidos = (
            filtrados.groupby("Producto", as_index=False)
            .agg(
                Ventas=("Ventas", "sum"),
                Unidades=("Unidades", "sum"),
            )
            .sort_values(["Ventas", "Unidades"], ascending=False)
            .head(10)
            .sort_values("Ventas")
        )

        mostrar_grafico(
            px.bar(
                productos_mas_vendidos,
                x="Ventas",
                y="Producto",
                orientation="h",
                title="Top 10 productos más vendidos",
                text_auto=".2s",
            )
        )

    with col_prod2:
        productos_margen = filtrados.copy()
        productos_margen["Margen_calculado"] = (
            productos_margen["Ganancia"] / productos_margen["Ventas"]
        ).replace([float("inf"), -float("inf")], 0).fillna(0)

        productos_margen = (
            productos_margen.groupby("Producto", as_index=False)
            .agg(
                Ventas=("Ventas", "sum"),
                Ganancia=("Ganancia", "sum"),
                Margen=("Margen_calculado", "mean"),
            )
        )

        # Evitamos productos con ventas insignificantes para que un margen
        # alto de una sola operación no domine la lectura.
        umbral_ventas = productos_margen["Ventas"].sum() * 0.01
        productos_margen = (
            productos_margen[productos_margen["Ventas"] >= umbral_ventas]
            .sort_values(["Margen", "Ganancia"], ascending=False)
            .head(10)
            .sort_values("Margen")
        )

        mostrar_grafico(
            px.bar(
                productos_margen,
                x="Margen",
                y="Producto",
                orientation="h",
                title="Top 10 productos por margen",
                text_auto=".1%",
            )
        )

    ticket_categoria = (
        filtrados.groupby("Categoria", as_index=False)["Ticket_Promedio"].mean()
    )
    mostrar_grafico(
        px.bar(
            ticket_categoria,
            x="Categoria",
            y="Ticket_Promedio",
            title="Ticket promedio por categoría",
            text_auto=".2f",
        )
    )

    st.subheader("Detalle de datos filtrados")
    st.dataframe(
        filtrados.sort_values("Fecha", ascending=False),
        use_container_width=True,
        hide_index=True,
    )

    st.divider()
    st.subheader("📄 Informe gerencial")

    st.write(
        "El botón utiliza los mismos datos y filtros visibles en el dashboard "
        "para construir un informe PDF actualizado."
    )

    if st.button("Generar informe PDF", type="primary", use_container_width=True):
        with st.spinner("Generando informe..."):
            try:
                pdf_bytes = crear_pdf(filtrados)
                st.session_state["pdf_informe"] = pdf_bytes
                st.success("Informe generado correctamente.")
            except Exception as error:
                st.error(
                    "No se pudo generar el PDF. Revisa las dependencias del "
                    f"proyecto. Detalle: {error}"
                )

    if "pdf_informe" in st.session_state:
        st.download_button(
            label="⬇️ Descargar informe PDF",
            data=st.session_state["pdf_informe"],
            file_name="informe_gerencial.pdf",
            mime="application/pdf",
            use_container_width=True,
        )


if __name__ == "__main__":
    main()
