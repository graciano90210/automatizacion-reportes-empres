import os
import re
import unicodedata
import pandas as pd
from datetime import datetime, timedelta
from reportlab.lib.pagesizes import letter
from reportlab.pdfgen import canvas
from reportlab.lib import colors
from reportlab.platypus import Table, TableStyle
import PyPDF2
import openpyxl

# ============================================================
# CONFIGURACIÓN
# ============================================================

CARPETA_ENTRADAS = "entradas"
CARPETA_REPORTES = "reportes"

MIS_RUTAS = [
    "01.01. JUAN SP",
    "01.02. FRANCISCO MORATO",
    "01.03. DIADEMA",
    "01.04. RIBEIRAO PRETO",
    "01.05. PINDA",
    "01.06. RIO",
    "01.07. OSTRAS",
    "01.08. CAMPINAS",
    "02.05. SANTOS",
    "02.06. GUARULHOS",
    "02.07. CABO FRIO CRISTIAN",
    "02.08. SÃO PEDRO",
    "02.09. CANTI",
]

# ============================================================
# FUNCIONES AUXILIARES
# ============================================================

def _norm(txt: str) -> str:
    """Normaliza texto removiendo acentos y caracteres especiales."""
    if txt is None:
        return ""
    s = str(txt)
    s = unicodedata.normalize("NFKD", s).encode("ascii", "ignore").decode("ascii")
    return re.sub(r"[^A-Z0-9]", "", s.upper())

def _find_col(df, candidates):
    """Busca columna que coincida con alguno de los candidatos."""
    if df is None or df.empty:
        return None
    norm_map = {_norm(c): c for c in df.columns}
    
    for cand in candidates:
        n = _norm(cand)
        if n in norm_map:
            return norm_map[n]
    
    for key, original in norm_map.items():
        for cand in candidates:
            if _norm(cand) in key:
                return original
    return None

def encontrar_todos_los_archivos(patron_inicio):
    """Encuentra TODOS los archivos que coincidan con el patrón."""
    archivos = [
        os.path.join(CARPETA_ENTRADAS, f)
        for f in os.listdir(CARPETA_ENTRADAS)
        if patron_inicio in f and f.endswith(".xlsx")
    ]
    return archivos

def encontrar_todos_los_pdfs(patron_inicio):
    """Encuentra TODOS los PDFs que coincidan con el patrón."""
    archivos = [
        os.path.join(CARPETA_ENTRADAS, f)
        for f in os.listdir(CARPETA_ENTRADAS)
        if patron_inicio in f and f.endswith(".pdf")
    ]
    return archivos

# ============================================================
# FUNCIONES DE LECTURA DE DATOS
# ============================================================

def leer_excel_cuadre_por_rutas(ruta):
    """Lee el archivo de Cuadre con formato especial."""
    if not ruta:
        return {}
    
    print(f"Leyendo Cuadre: {os.path.basename(ruta)}")
    try:
        workbook = openpyxl.load_workbook(ruta, data_only=True)
        sheet = workbook.active
        
        datos_por_ruta = {}
        
        for row_idx in range(1, sheet.max_row + 1):
            cell_value = sheet.cell(row=row_idx, column=1).value
            if cell_value and isinstance(cell_value, str) and cell_value.startswith("Ruta:"):
                nombre_ruta = cell_value.replace("Ruta:", "").strip()
                
                header_row = sheet[row_idx + 1]
                headers = [cell.value for cell in header_row if cell.value]
                
                data_row = sheet[row_idx + 2]
                
                ruta_data = {}
                for i, header in enumerate(headers):
                    valor_celda = data_row[i].value
                    if isinstance(valor_celda, str) and '$' in valor_celda:
                        valor_limpio = valor_celda.replace('$', '').replace('.', '').replace(',', '.').strip()
                        ruta_data[header] = pd.to_numeric(valor_limpio, errors='coerce')
                    else:
                        ruta_data[header] = valor_celda
                
                datos_por_ruta[nombre_ruta] = ruta_data

        return datos_por_ruta
    except Exception as e:
        print(f"Error leyendo Cuadre: {e}")
        return {}

def leer_excel_simple(ruta):
    """Lee archivo Excel con formato de tabla simple."""
    if not ruta:
        return None
    try:
        df = pd.read_excel(ruta, engine="openpyxl", header=8)
        print(f"Leído: {os.path.basename(ruta)} -> {df.shape[0]} filas")
        
        for col in df.columns:
            if df[col].dtype == "object":
                muestra = df[col].dropna().astype(str).str.strip()
                if muestra.empty:
                    continue
                    
                patron_numero = muestra.str.match(r'^[\s\$R]*-?[\d\.,]+[\s]*$')
                
                if patron_numero.any():
                    limpia = (df[col].astype(str)
                             .str.replace(r'R\$\s*', '', regex=True)
                             .str.replace('$', '', regex=False)
                             .str.replace(' ', '', regex=False)
                             .str.replace('.', '', regex=False)
                             .str.replace(',', '.', regex=False)
                             .str.strip())
                    df[col] = pd.to_numeric(limpia, errors='coerce').fillna(0)
        
        return df
    except Exception as e:
        print(f"Error: {e}")
        return None

def extraer_clientes_desde_pdf(ruta_pdf):
    """Extrae clientes sin pago del PDF basado en el formato de tabla."""
    if not ruta_pdf:
        return {}
    
    try:
        with open(ruta_pdf, "rb") as f:
            pdf = PyPDF2.PdfReader(f)
            texto_completo = ""
            for page in pdf.pages:
                texto_completo += page.extract_text() + "\n"
        
        clientes_por_ruta = {}
        lineas = texto_completo.split("\n")
        ruta_actual = None
        
        for linea in lineas:
            linea_original = linea
            linea = linea.strip()
            
            # Detectar encabezado de ruta (formato: RUTA: 01.02. FRANCISCO MORATO)
            if linea.startswith("RUTA:"):
                ruta_actual = linea.replace("RUTA:", "").strip()
                if ruta_actual not in clientes_por_ruta:
                    clientes_por_ruta[ruta_actual] = []
                print(f"   -> Ruta: '{ruta_actual}'")
                continue
            
            # Si hay ruta actual y la línea tiene el patrón de datos de cliente
            if ruta_actual:
                # Ignorar encabezados de tabla
                if any(x in linea.upper() for x in ["COD. CRÉDITO", "CLIENTE", "DÍAS DE ATRASO", "EXCEPCIÓN", "ATRASO", "PENALIZACIÓN", "TOTAL"]):
                    continue
                
                # Patrón: número (1-2 dígitos), código (6 dígitos), nombre cliente, resto de datos
                # Ejemplo: "10 142519 Maria Polliane Quintan 1 0 1 0"
                patron = re.match(r'^\d{1,2}\s+\d{5,6}\s+(.+?)\s+\d+\s+\d+\s+\d+\s+\d+\s*$', linea)
                
                if patron:
                    # Extraer el código de crédito y nombre del cliente
                    partes = linea.split()
                    if len(partes) >= 6:  # mínimo: # cod nombre dia exc atr pen
                        numero = partes[0]
                        cod_credito = partes[1]
                        # El nombre es todo lo que está entre el código y los últimos 4 números
                        nombre_cliente = ' '.join(partes[2:-4])
                        
                        cliente_info = f"{cod_credito} {nombre_cliente} {' '.join(partes[-4:])}"
                        clientes_por_ruta[ruta_actual].append(cliente_info)
        
        # Resumen
        total = sum(len(v) for v in clientes_por_ruta.values())
        print(f"   -> Total clientes: {total}")
        
        return clientes_por_ruta
    
    except Exception as e:
        print(f"Error extrayendo PDF: {e}")
        import traceback
        traceback.print_exc()
        return {}

def resumen_dataframe_por_ruta(df, ruta_val=None):
    """Filtra DataFrame por valor de ruta."""
    if df is None or df.empty:
        return pd.DataFrame()
    
    col_ruta = _find_col(df, ["RUTA"])
    if not col_ruta:
        return pd.DataFrame()
    
    df_limpio = df.copy()
    df_limpio['_RUTA_NORM'] = df_limpio[col_ruta].astype(str).str.strip()
    ruta_norm = str(ruta_val).strip()
    
    resultado = df_limpio[df_limpio['_RUTA_NORM'] == ruta_norm]
    resultado = resultado.drop('_RUTA_NORM', axis=1, errors='ignore')
    return resultado

# ============================================================
# FUNCIÓN PARA CREAR PDF
# ============================================================

def crear_pdf_reporte_por_ruta(nombre_ruta, fecha_informe, cuadre_totales,
                               creditos_df, gastos_df, clientes_lista, salida_path):
    """Genera el PDF del reporte."""
    ancho, alto = letter
    margen = 40
    
    c = canvas.Canvas(salida_path, pagesize=letter)
    
    # Header
    c.setFillColor(colors.HexColor('#1F4788'))
    c.rect(0, alto - 70, ancho, 70, fill=True, stroke=False)
    c.setFillColor(colors.white)
    c.setFont('Helvetica-Bold', 18)
    c.drawString(margen, alto - 35, f"> {nombre_ruta}")
    c.setFont('Helvetica-Bold', 14)
    c.drawString(ancho - 180, alto - 35, "INFORME")
    
    c.setFillColor(colors.black)
    c.setFont('Helvetica', 10)
    c.drawRightString(ancho - margen, alto - 85, f"Fecha: {fecha_informe}")
    
    y_actual = alto - 120
    
    # CUADRE RUTA
    c.setFont('Helvetica-Bold', 11)
    c.drawString(margen, y_actual, "CUADRE RUTA")
    y_actual -= 5
    c.setStrokeColor(colors.black)
    c.setLineWidth(1)
    c.line(margen, y_actual, ancho - margen, y_actual)
    y_actual -= 20
    
    abonos = cuadre_totales.get('ABONOS', 0)
    desembolsos = cuadre_totales.get('DESEMBOLSOS', 0)
    total_caja = cuadre_totales.get('TOTAL CAJA', 0)
    
    c.setFont('Helvetica-Bold', 9)
    col_width = (ancho - 2 * margen) / 3
    c.drawCentredString(margen + col_width / 2, y_actual, "ABONOS")
    c.drawCentredString(margen + col_width * 1.5, y_actual, "DESEMBOLSOS")
    c.drawCentredString(margen + col_width * 2.5, y_actual, "TOTAL CAJA")
    y_actual -= 15
    
    c.setFont('Helvetica', 10)
    c.drawCentredString(margen + col_width / 2, y_actual, f"{abonos:,.2f}")
    c.drawCentredString(margen + col_width * 1.5, y_actual, f"{desembolsos:,.2f}")
    c.drawCentredString(margen + col_width * 2.5, y_actual, f"{total_caja:,.2f}")
    y_actual -= 30
    
    # GASTOS Y MOVIMIENTOS
    c.setFont('Helvetica-Bold', 11)
    c.drawString(margen, y_actual, "GASTOS Y MOVIMIENTOS")
    y_actual -= 5
    c.line(margen, y_actual, ancho - margen, y_actual)
    y_actual -= 15
    
    if gastos_df is not None and not gastos_df.empty:
        desc_col = _find_col(gastos_df, ["OBSERVACION", "OBSERVACIÓN", "CONCEPTO"])
        val_col = _find_col(gastos_df, ["VALOR", "CAUSANTES"])
        
        if desc_col and val_col:
            tmp = gastos_df[[desc_col, val_col]].head(10).copy()
            tmp.columns = ["DESCRIPCION", "VALOR"]
            tmp["VALOR"] = pd.to_numeric(tmp["VALOR"], errors='coerce').fillna(0)
            tmp["VALOR"] = tmp["VALOR"].apply(lambda x: f"${x:,.2f}")
            
            data = [["DESCRIPCION", "VALOR"]] + tmp.values.tolist()
            tabla = Table(data, colWidths=[(ancho - 2*margen)*0.70, (ancho - 2*margen)*0.25])
            tabla.setStyle(TableStyle([
                ('FONT', (0,0), (-1,0), 'Helvetica-Bold', 9),
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4788')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('FONTSIZE', (0,1), (-1,-1), 8),
                ('ALIGN', (1,0), (1,-1), 'RIGHT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            w, h = tabla.wrapOn(c, ancho, alto)
            tabla.drawOn(c, margen + 5, y_actual - h)
            y_actual -= (h + 10)
        else:
            c.setFont('Helvetica-Oblique', 8)
            c.drawString(margen + 10, y_actual, "Sin datos de gastos")
            y_actual -= 20
    else:
        c.setFont('Helvetica-Oblique', 8)
        c.drawString(margen + 10, y_actual, "Sin datos de gastos")
        y_actual -= 20
    
    y_actual -= 20
    
    # CREDITOS
    c.setFont('Helvetica-Bold', 11)
    c.drawString(margen, y_actual, "CREDITOS")
    y_actual -= 5
    c.line(margen, y_actual, ancho - margen, y_actual)
    y_actual -= 15
    
    if creditos_df is not None and not creditos_df.empty:
        nombre_col = _find_col(creditos_df, ["NOMBRE", "CLIENTE"])
        tel_col = _find_col(creditos_df, ["TELEFONO", "TELÉFONO", "CELULAR"])
        val_col = _find_col(creditos_df, ["VALOR CRÉDITO", "VALOR CREDITO", "VALOR"])
        
        cols = [c for c in [nombre_col, tel_col, val_col] if c]
        
        if nombre_col and val_col:
            tmp = creditos_df[cols].head(8).copy()
            
            nuevos = []
            anchos = []
            for col in cols:
                if col == nombre_col:
                    nuevos.append("CLIENTE")
                    anchos.append((ancho - 2*margen)*0.45)
                elif col == tel_col:
                    nuevos.append("TELEFONO")
                    anchos.append((ancho - 2*margen)*0.30)
                elif col == val_col:
                    nuevos.append("VALOR")
                    anchos.append((ancho - 2*margen)*0.20)
            
            tmp.columns = nuevos
            
            if "VALOR" in tmp.columns:
                tmp["VALOR"] = pd.to_numeric(tmp["VALOR"], errors='coerce').fillna(0)
                tmp["VALOR"] = tmp["VALOR"].apply(lambda x: f"${x:,.2f}")
            
            data = [tmp.columns.tolist()] + tmp.values.tolist()
            tabla = Table(data, colWidths=anchos)
            tabla.setStyle(TableStyle([
                ('FONT', (0,0), (-1,0), 'Helvetica-Bold', 9),
                ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4788')),
                ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
                ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
                ('FONTSIZE', (0,1), (-1,-1), 8),
                ('ALIGN', (len(nuevos)-1,0), (len(nuevos)-1,-1), 'RIGHT'),
                ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
            ]))
            w, h = tabla.wrapOn(c, ancho, alto)
            tabla.drawOn(c, margen + 5, y_actual - h)
            y_actual -= (h + 10)
        else:
            c.setFont('Helvetica-Oblique', 8)
            c.drawString(margen + 10, y_actual, "Sin datos de créditos")
            y_actual -= 20
    else:
        c.setFont('Helvetica-Oblique', 8)
        c.drawString(margen + 10, y_actual, "Sin datos de créditos")
        y_actual -= 20
    
    y_actual -= 20
    
    # CLIENTES SIN PAGO
    c.setFont('Helvetica-Bold', 11)
    c.drawString(margen, y_actual, "CLIENTES SIN PAGO")
    y_actual -= 5
    c.line(margen, y_actual, ancho - margen, y_actual)
    y_actual -= 15
    
    if clientes_lista:
        data = [["#", "N°", "CLIENTE"]]
        for idx, cliente in enumerate(clientes_lista[:15], 1):
            data.append([str(idx), str(idx), cliente])
        
        tabla = Table(data, colWidths=[30, 30, (ancho - 2*margen) - 60])
        tabla.setStyle(TableStyle([
            ('FONT', (0,0), (-1,0), 'Helvetica-Bold', 9),
            ('BACKGROUND', (0,0), (-1,0), colors.HexColor('#1F4788')),
            ('TEXTCOLOR', (0,0), (-1,0), colors.whitesmoke),
            ('GRID', (0,0), (-1,-1), 0.5, colors.grey),
            ('FONTSIZE', (0,1), (-1,-1), 7),
            ('VALIGN', (0,0), (-1,-1), 'MIDDLE'),
        ]))
        w, h = tabla.wrapOn(c, ancho, alto)
        tabla.drawOn(c, margen + 5, y_actual - h)
    else:
        c.setFont('Helvetica-Oblique', 8)
        c.drawString(margen + 10, y_actual, "Sin clientes sin pago")
    
    c.save()
    print(f"✓ PDF: {os.path.basename(salida_path)}")

# ============================================================
# FUNCIÓN PRINCIPAL
# ============================================================

def main():
    print("Iniciando generador de reportes...\n")
    
    if not os.path.exists(CARPETA_REPORTES):
        os.makedirs(CARPETA_REPORTES)
    
    # 1. COMBINAR TODOS LOS ARCHIVOS DE CUADRE
    print("=== CUADRE ===")
    archivos_cuadre = encontrar_todos_los_archivos("ReporteExcelCuadreVariasRutas")
    datos_cuadre_combinados = {}
    for archivo in archivos_cuadre:
        datos = leer_excel_cuadre_por_rutas(archivo)
        datos_cuadre_combinados.update(datos)
    print(f"Total rutas en Cuadre: {len(datos_cuadre_combinados)}\n")
    
    # 2. COMBINAR TODOS LOS ARCHIVOS DE GASTOS
    print("=== GASTOS ===")
    archivos_gastos = encontrar_todos_los_archivos("ReporteExcelGastos")
    lista_gastos = []
    for archivo in archivos_gastos:
        df = leer_excel_simple(archivo)
        if df is not None:
            lista_gastos.append(df)
    df_gastos_completo = pd.concat(lista_gastos, ignore_index=True) if lista_gastos else pd.DataFrame()
    print(f"Total filas de Gastos: {len(df_gastos_completo)}\n")
    
    # 3. COMBINAR TODOS LOS ARCHIVOS DE CRÉDITOS
    print("=== CRÉDITOS ===")
    archivos_creditos = encontrar_todos_los_archivos("ReporteExcelCreditos")
    lista_creditos = []
    for archivo in archivos_creditos:
        df = leer_excel_simple(archivo)
        if df is not None:
            lista_creditos.append(df)
    df_creditos_completo = pd.concat(lista_creditos, ignore_index=True) if lista_creditos else pd.DataFrame()
    print(f"Total filas de Créditos: {len(df_creditos_completo)}\n")
    
    # 4. COMBINAR TODOS LOS PDFs DE CLIENTES
    print("=== CLIENTES SIN PAGO ===")
    archivos_pdfs = encontrar_todos_los_pdfs("ReportePdfClientesSinPagos")
    clientes_combinados = {}
    for archivo in archivos_pdfs:
        print(f"Leyendo: {os.path.basename(archivo)}")
        datos = extraer_clientes_desde_pdf(archivo)
        for ruta, clientes in datos.items():
            if ruta not in clientes_combinados:
                clientes_combinados[ruta] = []
            clientes_combinados[ruta].extend(clientes)
    print(f"Total rutas con clientes: {len(clientes_combinados)}\n")
    
    # Usar siempre la fecha del día anterior en el informe
    fecha_base = datetime.now() - timedelta(days=1)
    fecha_informe = fecha_base.strftime("%d/%m/%Y")
    
    # 5. GENERAR REPORTES POR RUTA
    print("=== GENERANDO REPORTES ===")
    for ruta in MIS_RUTAS:
        print(f"\n📄 {ruta}")
        
        cuadre_totales = datos_cuadre_combinados.get(ruta, {})
        creditos_subset = resumen_dataframe_por_ruta(df_creditos_completo, ruta_val=ruta)
        gastos_subset = resumen_dataframe_por_ruta(df_gastos_completo, ruta_val=ruta)
        clientes_lista = clientes_combinados.get(ruta, [])
        
        print(f"   Cuadre: {'✓' if cuadre_totales else '✗'}")
        print(f"   Gastos: {len(gastos_subset)} filas")
        print(f"   Créditos: {len(creditos_subset)} filas")
        print(f"   Clientes: {len(clientes_lista)} clientes")
        
        nombre_limpio = ruta.replace(".", "").replace(" ", "_")
        salida_path = os.path.join(CARPETA_REPORTES, f"Reporte_{nombre_limpio}.pdf")
        
        crear_pdf_reporte_por_ruta(
            nombre_ruta=ruta,
            fecha_informe=fecha_informe,
            cuadre_totales=cuadre_totales,
            creditos_df=creditos_subset,
            gastos_df=gastos_subset,
            clientes_lista=clientes_lista,
            salida_path=salida_path
        )
    
    print(f"\n✓✓✓ COMPLETADO ✓✓✓")
    print(f"Revisa la carpeta '{CARPETA_REPORTES}/'")

if __name__ == "__main__":
    main()