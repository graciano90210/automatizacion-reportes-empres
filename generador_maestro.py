# generador_maestro_hibrido.py
# ======================================================================
# Generador Maestro Híbrido - Crea 1 PDF por ruta (estilo "INFORME JUAN")
# ======================================================================

import os
import glob
import pandas as pd
import pdfplumber
from datetime import datetime
from reportlab.lib.pagesizes import A4
from reportlab.lib.units import mm
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.platypus import SimpleDocTemplate, Paragraph, Spacer, Table, TableStyle
from reportlab.lib import colors

# ------------------ CONFIG ------------------
CARPETA_ENTRADAS = "entradas"
CARPETA_REPORTES = "reportes"
# Rutas tuyas (exactas tal como aparecen en el PDF)
MIS_RUTAS = [
    "01.01. JUAN SP","01.02. FRANCISCO MORATO","01.03. DIADEMA",
    "01.04. RIBEIRÃO PRETO","01.05. PINDA","01.06. RIO",
    "01.07. OSTRAS","01.08. CAMPINAS","02.05. SANTOS",
    "02.06. GUARULHOS","02.07. CABO FRIO CRISTIAN","02.08. SÃO PEDRO",
    "02.09. CANTI"
]

PAGE_SIZE = A4
# --------------------------------------------

os.makedirs(CARPETA_REPORTES, exist_ok=True)

def encontrar_archivo_mas_reciente(patron):
    """Busca el archivo más reciente que contenga `patron` en el nombre dentro de CARPETA_ENTRADAS."""
    ruta_patron = os.path.join(CARPETA_ENTRADAS, f"*{patron}*.xlsx")
    archivos = glob.glob(ruta_patron)
    if not archivos:
        return None
    archivos.sort(key=os.path.getmtime, reverse=True)
    return archivos[0]

def encontrar_pdf_mas_reciente(patron):
    ruta_patron = os.path.join(CARPETA_ENTRADAS, f"*{patron}*.pdf")
    archivos = glob.glob(ruta_patron)
    if not archivos:
        return None
    archivos.sort(key=os.path.getmtime, reverse=True)
    return archivos[0]

def leer_excel_si_existe(ruta):
    """Lee Excel devolviendo DataFrame o None si falla."""
    if not ruta:
        return None
    try:
        df = pd.read_excel(ruta, engine="openpyxl")
        print(f"Leído Excel: {os.path.basename(ruta)} -> {df.shape[0]} filas, {df.shape[1]} cols")
        return df
    except Exception as e:
        print(f"Error leyendo {ruta}: {e}")
        return None

def extraer_clientes_por_ruta_desde_pdf(ruta_pdf):
    """
    Extrae texto del PDF y construye un dict: ruta -> lista de lineas (clientes).
    Se basa en que el PDF contiene secciones que empiezan con "RUTA:" seguido del nombre.
    """
    rutas_map = {}
    if not ruta_pdf:
        return rutas_map
    try:
        with pdfplumber.open(ruta_pdf) as pdf:
            texto = ""
            for p in pdf.pages:
                t = p.extract_text()
                if t:
                    texto += t + "\n"
    except Exception as e:
        print(f"Error abriendo PDF {ruta_pdf}: {e}")
        return rutas_map

    # Normalizamos y separamos por "RUTA:"
    partes = texto.split("RUTA:")
    for parte in partes:
        texto_seccion = parte.strip()
        if not texto_seccion:
            continue
        # Primera línea será el nombre de la ruta (posible)
        lineas = [l.strip() for l in texto_seccion.splitlines() if l.strip()]
        if not lineas:
            continue
        nombre_ruta = lineas[0]
        # Guardamos resto de líneas como datos de clientes (ignorando encabezados)
        datos = []
        for ln in lineas[1:]:
            # descartamos totales
            if ln.upper().startswith("TOTAL") or ln.upper().startswith("PÁGINA") or ln.upper().startswith("FECHA"):
                continue
            # guardamos lineas que parecen contener datos (números + texto)
            if any(char.isdigit() for char in ln):
                datos.append(ln)
        rutas_map[nombre_ruta] = datos
    print(f"Extracción PDF completada. Secciones encontradas: {len(rutas_map)}")
    return rutas_map

def resumen_dataframe_por_ruta(df, ruta_col_name="RUTA", ruta_val=None, max_rows=10):
    """Filtra df por ruta y devuelve un resumen simple (head y totales de columnas numéricas)."""
    if df is None:
        return None, None
    # intentamos encontrar la columna de ruta (insensible a mayúsculas)
    cols_lower = {c.lower(): c for c in df.columns}
    candidate = None
    for key in cols_lower:
        if "ruta" == key or "ruta" in key:
            candidate = cols_lower[key]
            break
    if not candidate:
        # no hay columna ruta
        # devolvemos head global si la ruta_val es None, sino vacío
        if ruta_val is None:
            sample = df.head(max_rows)
            totals = df.select_dtypes('number').sum().to_dict()
            return sample, totals
        else:
            return pd.DataFrame(), {}
    df_r = df[df[candidate].astype(str).str.strip() == str(ruta_val).strip()]
    sample = df_r.head(max_rows)
    totals = df_r.select_dtypes('number').sum().to_dict()
    return sample, totals

def formatear_numero(n):
    try:
        if pd.isna(n):
            return ""
        # si es entero mostrar sin decimales, sino con 2
        if float(n).is_integer():
            return f"{int(n):,}".replace(",", ".")
        return f"{float(n):,.2f}".replace(",", ".")
    except:
        return str(n)

def crear_pdf_reporte_por_ruta(nombre_ruta, fecha_informe, cuadre_df, cuadre_totales,
                               creditos_df, creditos_totales,
                               gastos_df, gastos_totales,
                               clientes_lista, salida_path):
    """Genera PDF con estilo sencillo parecido a 'INFORME JUAN'."""
    doc = SimpleDocTemplate(salida_path, pagesize=PAGE_SIZE,
                            rightMargin=18*mm, leftMargin=18*mm, topMargin=18*mm, bottomMargin=18*mm)
    styles = getSampleStyleSheet()
    story = []

    # Encabezado grande
    h_title = Paragraph("INFORME", ParagraphStyle('Title', parent=styles['Title'], alignment=1, fontSize=18))
    story.append(h_title)
    story.append(Spacer(1, 6))
    subtitle = Paragraph(f"<b>{nombre_ruta}</b>", ParagraphStyle('SubTitle', parent=styles['Normal'], alignment=1, fontSize=12))
    story.append(subtitle)
    story.append(Spacer(1, 8))

    # Fecha
    story.append(Paragraph(f"<b>Fecha:</b> {fecha_informe}", styles['Normal']))
    story.append(Spacer(1, 8))

    # --- CUADRE RUTA ---
    story.append(Paragraph("<b>CUADRE RUTA</b>", styles['Heading2']))
    story.append(Spacer(1, 4))
    if cuadre_df is None or cuadre_df.empty:
        story.append(Paragraph("No hay registro de cuadre para esta ruta.", styles['Normal']))
    else:
        # Si hay algunos totales, muéstralos
        if cuadre_totales:
            linea = " | ".join([f"{k}: {formatear_numero(v)}" for k, v in cuadre_totales.items()])
            story.append(Paragraph(linea, styles['Normal']))
        # tabla con algunas columnas
        tbl_head = list(cuadre_df.columns[:6])  # mostramos hasta 6 columnas para ajuste
        tbl_data = [tbl_head]
        for _, r in cuadre_df.head(15).iterrows():
            row = [str(r.get(c, "")) for c in tbl_head]
            tbl_data.append(row)
        t = Table(tbl_data, hAlign='LEFT')
        t.setStyle(TableStyle([
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
            ('FONTSIZE', (0,0), (-1,-1), 8),
        ]))
        story.append(t)
    story.append(Spacer(1, 10))

    # --- CREDITOS ---
    story.append(Paragraph("<b>CRÉDITOS</b>", styles['Heading2']))
    story.append(Spacer(1, 4))
    if creditos_df is None or creditos_df.empty:
        story.append(Paragraph("No hay créditos registrados para esta ruta.", styles['Normal']))
    else:
        if creditos_totales:
            linea = " | ".join([f"{k}: {formatear_numero(v)}" for k, v in creditos_totales.items()])
            story.append(Paragraph(linea, styles['Normal']))
        cols = list(creditos_df.columns[:6])
        tbl_data = [cols]
        for _, r in creditos_df.head(15).iterrows():
            row = [str(r.get(c, "")) for c in cols]
            tbl_data.append(row)
        t = Table(tbl_data, hAlign='LEFT')
        t.setStyle(TableStyle([
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
            ('FONTSIZE', (0,0), (-1,-1), 8),
        ]))
        story.append(t)
    story.append(Spacer(1, 10))

    # --- GASTOS ---
    story.append(Paragraph("<b>GASTOS Y MOVIMIENTOS</b>", styles['Heading2']))
    story.append(Spacer(1, 4))
    if gastos_df is None or gastos_df.empty:
        story.append(Paragraph("No hay gastos registrados para esta ruta.", styles['Normal']))
    else:
        if gastos_totales:
            linea = " | ".join([f"{k}: {formatear_numero(v)}" for k, v in gastos_totales.items()])
            story.append(Paragraph(linea, styles['Normal']))
        cols = list(gastos_df.columns[:6])
        tbl_data = [cols]
        for _, r in gastos_df.head(15).iterrows():
            row = [str(r.get(c, "")) for c in cols]
            tbl_data.append(row)
        t = Table(tbl_data, hAlign='LEFT')
        t.setStyle(TableStyle([
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold'),
            ('BACKGROUND', (0,0), (-1,0), colors.lightgrey),
            ('GRID', (0,0), (-1,-1), 0.25, colors.grey),
            ('FONTSIZE', (0,0), (-1,-1), 8),
        ]))
        story.append(t)
    story.append(Spacer(1, 10))

    # --- CLIENTES SIN PAGO ---
    story.append(Paragraph("<b>CLIENTES SIN PAGO</b>", styles['Heading2']))
    story.append(Spacer(1, 4))
    if not clientes_lista:
        story.append(Paragraph("No se encontraron clientes sin pago para esta ruta.", styles['Normal']))
    else:
        # mostramos cada cliente en una línea
        for i, cli in enumerate(clientes_lista[:200], start=1):
            story.append(Paragraph(f"{i}. {cli}", ParagraphStyle('cli', parent=styles['Normal'], fontSize=9)))
    story.append(Spacer(1, 6))

    # footer: generación
    story.append(Spacer(1, 12))
    story.append(Paragraph(f"Generado: {datetime.now().strftime('%Y-%m-%d %H:%M:%S')}", ParagraphStyle('footer', parent=styles['Normal'], fontSize=7)))
    doc.build(story)
    print(f"PDF generado: {salida_path}")

def main():
    print("Iniciando generador maestro híbrido...")

    # 1) localizar archivos
    cuadre_path = encontrar_archivo_mas_reciente("ReporteExcelCuadreVariasRutas")
    creditos_path = encontrar_archivo_mas_reciente("ReporteExcelCreditos")
    gastos_path = encontrar_archivo_mas_reciente("ReporteExcelGastos")
    pdf_clientes_path = encontrar_pdf_mas_reciente("ReportePdfClientesSinPagos")

    print("Archivos detectados:")
    print(" - Cuadre:", cuadre_path)
    print(" - Créditos:", creditos_path)
    print(" - Gastos:", gastos_path)
    print(" - PDF Clientes:", pdf_clientes_path)

    # 2) leer excels
    df_cuadre = leer_excel_si_existe(cuadre_path)
    df_creditos = leer_excel_si_existe(creditos_path)
    df_gastos = leer_excel_si_existe(gastos_path)

    # 3) extraer clientes por ruta desde el PDF
    clientes_por_ruta = extraer_clientes_por_ruta_desde_pdf(pdf_clientes_path)

    fecha_informe = datetime.now().strftime("%d/%m/%Y")

    # 4) iterar sobre MIS_RUTAS y generar PDF por cada una
    for ruta in MIS_RUTAS:
        print(f"\nProcesando ruta: {ruta} ...")
        cuadre_subset, cuadre_totales = resumen_dataframe_por_ruta(df_cuadre, ruta_val=ruta)
        creditos_subset, creditos_totales = resumen_dataframe_por_ruta(df_creditos, ruta_val=ruta)
        gastos_subset, gastos_totales = resumen_dataframe_por_ruta(df_gastos, ruta_val=ruta)
        clientes_lista = clientes_por_ruta.get(ruta, [])

        # nombre de salida: limpiar caracteres no deseados
        safe_name = ruta.replace(" ", "_").replace(".", "").replace("/", "_")
        salida = os.path.join(CARPETA_REPORTES, f"Reporte_RUTA_{safe_name}_{datetime.now().strftime('%Y-%m-%d')}.pdf")

        crear_pdf_reporte_por_ruta(
            nombre_ruta=ruta,
            fecha_informe=fecha_informe,
            cuadre_df=cuadre_subset if cuadre_subset is not None else pd.DataFrame(),
            cuadre_totales=cuadre_totales,
            creditos_df=creditos_subset if creditos_subset is not None else pd.DataFrame(),
            creditos_totales=creditos_totales,
            gastos_df=gastos_subset if gastos_subset is not None else pd.DataFrame(),
            gastos_totales=gastos_totales,
            clientes_lista=clientes_lista,
            salida_path=salida
        )

    print("\nProceso completado. Revisa la carpeta 'reportes/'.")

if __name__ == "__main__":
    main()
