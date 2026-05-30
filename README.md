# Modulo-contable
# =========================================================
# 1. Importación de Librerías e Interfaz Gráfica (Utilerías)
# =========================================================
import os
import re
import tkinter as tk
from tkinter import filedialog, messagebox
import openpyxl
from openpyxl.utils.cell import get_column_letter, range_boundaries
import pandas as pd
import dbf  # Requisito: pip install dbf
from datetime import datetime

def configurar_ventana_emergente():
    """Crea una ventana base invisible para soportar cuadros de diálogo de forma limpia."""
    root_oculta = tk.Tk()
    root_oculta.withdraw()
    root_oculta.attributes("-topmost", True)
    return root_oculta

def seleccionar_archivo(ventana_padre=None):
    """Abre el explorador de archivos de Windows para seleccionar el libro de Excel."""
    file_path = filedialog.askopenfilename(
        parent=ventana_padre,
        title="Selecciona el libro de Excel con las planillas",
        filetypes=[("Archivos de Excel", "*.xlsx *.xls")],
    )
    return file_path

def seleccionar_archivo_txt(ventana_padre=None):
    """Abre el explorador de archivos de Windows para seleccionar el TXT de origen."""
    file_path = filedialog.askopenfilename(
        parent=ventana_padre,
        title="Selecciona el archivo TXT (Auxiliar de Mayor)",
        filetypes=[("Archivos de Texto", "*.txt")],
    )
    return file_path

def seleccionar_destino_dbf(ventana_padre=None):
    """Abre una ventana flotante para elegir dónde guardar el archivo dBASE (.dbf)"""
    file_path = filedialog.asksaveasfilename(
        parent=ventana_padre,
        title="Guardar archivo contable dBASE",
        defaultextension=".dbf",
        filetypes=[("Archivos dBASE", "*.dbf")],
        initialfile="PARTIDA.DBF"
    )
    return file_path

def seleccionar_destino_excel(ventana_padre=None):
    """Abre una ventana flotante para elegir dónde guardar el archivo de Excel (.xlsx)"""
    file_path = filedialog.asksaveasfilename(
        parent=ventana_padre,
        title="Guardar partida de diario en Excel",
        defaultextension=".xlsx",
        filetypes=[("Archivos de Excel", "*.xlsx")],
        initialfile="PARTIDA_DIARIO.XLSX"
    )
    return file_path

def limpiar_monto(valor):
    """Convierte celdas vacías, guiones o formatos contables a flotantes limpios."""
    if pd.isna(valor):
        return 0.0
    val_str = str(valor).strip()
    if val_str in ["-", "", ".-", ". -"]:
        return 0.0
    try:
        return float(val_str.replace(",", ""))
    except ValueError:
        return 0.0

def buscar_tabla_oficial(wb, nombre_tabla):
    """Localiza una Tabla de Excel por su nombre dentro de cualquier pestaña del libro."""
    for sheet in wb.worksheets:
        if nombre_tabla in sheet.tables:
            return sheet, sheet.tables[nombre_tabla]
    return None, None

# =========================================================
# 2. Motores de Escritura Física (dBASE .DBF y Excel .XLSX)
# =========================================================

def seguro_int(valor):
    """Convierte de forma segura cualquier texto o vacío a número entero."""
    if not valor:
        return 0
    val_str = str(valor).strip()
    if val_str in ["", "-", ".", ".-"]:
        return 0
    try:
        return int(float(val_str))
    except ValueError:
        return 0

def exportar_a_dbf(ruta_dbf, lineas_partida):
    """Mapea de forma estricta los índices de la lista generada por el TXT."""
    ruta_dbf_fisica = os.path.normpath(ruta_dbf)
    try:
        if os.path.exists(ruta_dbf_fisica):
            os.remove(ruta_dbf_fisica)

        estructura_campos = (
            'TIPPART C(10); NUMPART N(10,0); FECHA D; CODCTA C(25); TIPO C(1); '
            'CONCEPTO1 C(40); CONCEPTO2 C(40); CONCEPTO3 C(40); MONTO N(16,2); '
            'MONTO_AUX N(16,2); B_TIPO C(2); B_NUMERO N(10,0); B_FECHACON D; '
            'B_IMPRESIO L; CORPART N(10,0)'
        )
        tabla_dbf = dbf.Table(ruta_dbf_fisica, estructura_campos, codepage='cp850')
        tabla_dbf.open(mode=dbf.READ_WRITE)
        
        for fila in lineas_partida:
            fecha_str = str(fila[2]).strip()
            fecha_pda = datetime.strptime(fecha_str, "%d/%m/%Y").date()
            
            b_impresio_str = str(fila[13]).upper()
            b_impresio_bool = True if b_impresio_str in ["VERDADERO", "TRUE", "T"] else False
            
            fecha_banco = dbf.Date()
            
            registro = (
                str(fila[0]).strip(),   # TIPPART
                seguro_int(fila[1]),    # NUMPART
                fecha_pda,              # FECHA
                str(fila[3]).strip(),   # CODCTA
                str(fila[4]).strip(),   # TIPO
                str(fila[5]).strip(),   # CONCEPTO1
                str(fila[6]).strip(),   # CONCEPTO2
                str(fila[7]).strip(),   # CONCEPTO3
                float(fila[8]),         # MONTO
                float(fila[9]),         # MONTO_AUX
                str(fila[10]).strip(),  # B_TIPO
                seguro_int(fila[11]),   # B_NUMERO
                fecha_banco,            # B_FECHACON
                b_impresio_bool,        # B_IMPRESIO
                seguro_int(fila[14])    # CORPART
            )
            tabla_dbf.append(registro)
        tabla_dbf.close()
        return True
    except Exception as e:
        print(f"Error al escribir el archivo DBF: {str(e)}")
        raise e

def exportar_a_excel(ruta_excel, lineas_partida):
    """Crea un archivo temporal DBF y clona su memoria a Excel limpiando caracteres ilegales."""
    ruta_excel_fisica = os.path.normpath(ruta_excel)
    ruta_temporal_dbf = os.path.normpath(ruta_excel_fisica.replace(".xlsx", "_TEMP.dbf"))
    
    try:
        # 1. Generamos el DBF temporal usando la función de arriba
        exportar_a_dbf(ruta_temporal_dbf, lineas_partida)

        # 2. Leemos la tabla dBASE nativa desde el disco
        tabla_leida = dbf.Table(ruta_temporal_dbf)
        tabla_leida.open(mode=dbf.READ_ONLY)
        
        # 3. Inicializamos openpyxl
        libro_nuevo = openpyxl.Workbook()
        hoja_activa = libro_nuevo.active
        hoja_activa.title = "PDA"
        
        cabeceras = [campo for campo in tabla_leida.field_names]
        hoja_activa.append(cabeceras)
        
        # Filtro para caracteres invisibles o ilegales en celdas de Excel
        re_limpiar_prohibidos = re.compile(r'[\x00-\x08\x0B-\x0C\x0E-\x1F\x7F-\x9F]')
        
        # 4. Volcamos los registros
        for registro_dbf in tabla_leida:
            fila_excel = []
            for campo in cabeceras:
                valor_crudo = registro_dbf[campo]
                
                if isinstance(valor_crudo, datetime) or hasattr(valor_crudo, 'strftime'):
                    fila_excel.append(valor_crudo.strftime("%d/%m/%Y"))
                elif isinstance(valor_crudo, bool):
                    fila_excel.append("FALSO" if not valor_crudo else "VERDADERO")
                else:
                    texto_limpio = str(valor_crudo).strip()
                    texto_limpio = re_limpiar_prohibidos.sub('', texto_limpio)
                    fila_excel.append(texto_limpio)
                    
            hoja_activa.append(fila_excel)
            row_idx = hoja_activa.max_row
            
            # --- CORRECCIÓN CRÍTICA DE SINTAXIS ---
            # Aplicamos formato de texto plano a las columnas de concepto (6, 7 y 8) de manera legal
            for col_idx in [6, 7, 8]:
                hoja_activa.cell(row=row_idx, column=col_idx).data_type = 's'
                
        # 5. Guardamos y limpiamos el temporal
        libro_nuevo.save(ruta_excel_fisica)
        libro_nuevo.close()
        tabla_leida.close()
        
        if os.path.exists(ruta_temporal_dbf):
            os.remove(ruta_temporal_dbf)
        return True
        
    except Exception as e:
        if 'tabla_leida' in locals():
            try:
                tabla_leida.close()
            except:
                pass
        if os.path.exists(ruta_temporal_dbf):
            try:
                os.remove(ruta_temporal_dbf)
            except:
                pass
        print(f"Error en clonación a Excel desde TXT: {str(e)}")
        raise e

# =========================================================
# 3. Procesamiento de Archivo TXT Jerárquico - PARTE A
# =========================================================

def procesar_txt_contable(ruta_txt):
    """Lee el TXT, determina D/H jerárquicamente por el ancho físico de la cuenta de mayor principal y plancha."""
    lineas_partida = []
    
    # Valores dinámicos por defecto extraídos de la cabecera 1
    fecha_defecto = "28/01/2026" 
    num_partida_defecto = 25       
    tipo_partida_defecto = "D01"   
    
    cuenta_actual = ""
    conceptos_acumulados = []

    # Lista secuencial pura para almacenar el Tipo contable ('D' o 'H') en el orden exacto de aparición en disco
    lista_secuencial_naturalezas = []
    
    # Variable estructural de cascada jerárquica de nivel superior
    naturaleza_mayor_activa = "D"

    # 1. ETAPA DE AUDITORÍA EN BRUTO: Evaluamos el ancho exacto de las líneas de disco (Tus coordenadas)
    with open(ruta_txt, 'r', encoding='latin-1', errors='ignore') as f:
        lineas_originales_disco = f.readlines()

    for linea_bruta in lineas_originales_disco:
        # Medimos la longitud física exacta de la línea original en el disco (removiendo solo el salto de línea)
        largo_linea_bruta = len(linea_bruta.replace('\r', '').replace('\n', ''))
        linea_limpia_b = linea_bruta.strip()

        # Captura de estructuras numéricas de cuentas
        match_c_b = re.match(r'^\s*(\d+)\s+(.+)$', linea_bruta)
        if match_c_b:
            cod_potencial = match_c_b.group(1).strip()
            
            # REGLA JERÁRQUICA MAESTRA POR ANCHO DE LÍNEA DE MAYOR (Exactamente 3 dígitos)
            if len(cod_potencial) == 3:
                # TU DESCUBRIMIENTO GEOMÉTRICO: Si la línea de mayor mide 118 caracteres exactos,
                # esa cuenta de mayor principal está en el HABER (H). Si mide 100 caracteres exactos, está en el DEBE (D).
                if largo_linea_bruta == 118:
                    naturaleza_mayor_activa = "H"
                elif largo_linea_bruta == 100:
                    naturaleza_mayor_activa = "D"
                continue

            # ASIGNACIÓN DE SUBCUENTAS OPERATIVAS ANALÍTICAS (Nivel de detalle >= 5 dígitos)
            if len(cod_potencial) >= 5:
                # Si es una línea de cuenta operativa legítima y mide 84 caracteres con un monto válido al final,
                # guardamos en la lista la letra exacta dictaminada por su Cuenta de Mayor de nivel superior activa.
                match_m_b = re.search(r'\s*(\d+[\.,]\d{2})\s*$', linea_limpia_b)
                if match_m_b and largo_linea_bruta == 84:
                    texto_concepto_b = linea_limpia_b[:match_m_b.start()].strip().upper()
                    if (texto_concepto_b and 
                        re.search(r'[A-Z]', texto_concepto_b) and 
                        not any(k in texto_concepto_b for k in ["TOTALES", "PARTIDA", "PASAN", "CUENTAS", "DESCRIPCION"])):
                        lista_secuencial_naturalezas.append(naturaleza_mayor_activa)
                continue

        # Si el renglón es una línea de dinero analítico puro (Parcial que mide 84 caracteres)
        if largo_linea_bruta == 84:
            match_m_b = re.search(r'\s*(\d+[\.,]\d{2})\s*$', linea_limpia_b)
            if match_m_b:
                texto_concepto_b = linea_limpia_b[:match_m_b.start()].strip().upper()
                # BLINDAJE INDUSTRIAL: Validamos que no pertenezca a las líneas basura de quiebres de página
                if (texto_concepto_b and 
                    re.search(r'[A-Z]', texto_concepto_b) and 
                    not any(k in texto_concepto_b for k in ["TOTALES", "PARTIDA", "PASAN", "CUENTAS", "DESCRIPCION"])):
                    lista_secuencial_naturalezas.append(naturaleza_mayor_activa)

    # 2. LEER EL ARCHIVO COMPLETO COMO UN SOLO BLOQUE DE TEXTO EN MEMORIA PARA EL PLANCHADO LÁSER
    contenido_completo = "".join(lineas_originales_disco)

    # EXTRACCIÓN AUTOMÁTICA DE PARÁMETROS CONTABLES DE LA CABECERA 1
    match_fecha_txt = re.search(r'Del\s+(\d+)\s+de\s+([A-Za-zí]+)\s+del\s+(\d{4})', contenido_completo, re.IGNORECASE)
    if match_fecha_txt:
        dia = match_fecha_txt.group(1).zfill(2)
        mes_txt = match_fecha_txt.group(2).lower()
        anio = match_fecha_txt.group(3)
        meses = {"enero": "01", "febrero": "02", "marzo": "03", "abril": "04", "mayo": "05", "junio": "06",
                 "julio": "07", "agosto": "08", "septiembre": "09", "octubre": "10", "noviembre": "11", "diciembre": "12"}
        mes = meses.get(mes_txt, "01")
        fecha_defecto = f"{dia}/{mes}/{anio}"

    match_partida_txt = re.search(r'Partida\s+No\.\s*:\s*([A-Z0-9]+)\s*-\s*(\d+)', contenido_completo, re.IGNORECASE)
    if match_partida_txt:
        tipo_partida_defecto = match_partida_txt.group(1).strip()
        num_partida_defecto = int(match_partida_txt.group(2).strip())

    # 3. PLANCHADO LÁSER BINARIO DE HOJAS INTERMEDIAS (Lógica original de unificación de conceptos)
    patron_quiebre_hoja = re.compile(
        r'[\x97—\-]+\s*\n\s*Pasan\s+[\d\.,]+\s+[\d\.,]+\s*\n\s*[\x97—\-]+.*?'
        r'C\s*U\s*E\s*N\s*T\s*A\s*S\s+D\s*E\s*S\s*C\s*R\s*I\s*P\s*C\s*I\s*O\s*N.*?\n\s*[\x97—\-]+\s*\n', 
        re.DOTALL | re.IGNORECASE
    )
    texto_planchado = patron_quiebre_hoja.sub('', contenido_completo)
    lineas_planchadas = texto_planchado.splitlines()

    # 4. EXTRACCIÓN LINEAL FLUIDA UNIVERSAL CON ACOPLAMIENTO SECUENCIAL ATÓMICO
    for linea_planchada in lineas_planchadas:
        linea_up = linea_planchada.upper()
        contenido_linea = linea_planchada.strip()

        # Filtro de seguridad inicial contra cabeceras remanentes o firmas finales del documento
        if (not contenido_linea or 
            "———" in linea_planchada or "---" in contenido_linea or "___" in contenido_linea or "\x97" in linea_planchada or
            "PÁG" in linea_up or "PAG." in linea_up or "CORPORACION" in linea_up or 
            "AUXILIAR" in linea_up or "PARTIDA" in linea_up or "DEL " in linea_up or
            "CUENTAS" in contenido_linea.replace(" ", "").upper() or
            "AUTORIZADO" in linea_up or "HECHO POR" in linea_up or "REVISADO" in linea_up or
            re.search(r'\d{2}/\d{2}/\d{4}', linea_planchada)):
            continue

        # IDENTIFICACIÓN UNIVERSAL DE ESTRUCTURAS DE CUENTAS
        match_cuenta = re.match(r'^\s*(\d+)\s+(.+)$', linea_planchada)
        if match_cuenta:
            cuenta_potencial = match_cuenta.group(1).strip()
            if len(cuenta_potencial) >= 5:
                cuenta_actual = cuenta_potencial
                conceptos_acumulados = []  # Vaciamos acumulador para evitar títulos de mayor en las glosas
            continue

        # IDENTIFICACIÓN DE TRANSACCIONES REFORZADA: Soporta montos contables con separadores de miles y millones
        match_monto = re.search(r'\s*(\d{1,3}(?:,?\d{3})*[\.,]\d{2})\s*$', contenido_linea)
        
        if match_monto:
            monto_str = match_monto.group(1)
            monto_transaccion = limpiar_monto(monto_str)
            ultimo_concepto = contenido_linea[:match_monto.start()].strip()
            ultimo_concepto_clean = ultimo_concepto.replace(" ", "").upper()
            
            # Blindaje contra subtotales de cierre en la línea del dinero
            if (not ultimo_concepto or 
                "TOTALES" in ultimo_concepto_clean or 
                "PARTIDA" in ultimo_concepto_clean or 
                "PASAN" in ultimo_concepto_clean or 
                "\x97" in ultimo_concepto or
                "---" in ultimo_concepto or 
                "___" in ultimo_concepto):
                continue

            if ultimo_concepto:
                conceptos_acumulados.append(ultimo_concepto)

            # ASIGNACIÓN DE NATURALEZA FIEL JERÁRQUICA:
            # Extraemos de forma atómica el Tipo contable heredado en cascada desde el mayor principal
            tipo_movimiento = lista_secuencial_naturalezas.pop(0) if len(lista_secuencial_naturalezas) > 0 else "D"

            # EXTRACCIÓN PROTEGIDA POR DESEMPAQUETADO INDIVIDUAL (.pop(0)):
            c1_texto = conceptos_acumulados.pop(0) if len(conceptos_acumulados) > 0 else "CONTINUACION DE REGISTRO"
            c2_texto = conceptos_acumulados.pop(0) if len(conceptos_acumulados) > 0 else ""
            c3_texto = conceptos_acumulados.pop(0) if len(conceptos_acumulados) > 0 else ""
            
            # Protección física de 40 bytes para la estructura de campos dBASE
            concepto1 = str(c1_texto).encode('utf-8')[:40].decode('utf-8', errors='ignore').ljust(40)
            concepto2 = str(c2_texto).encode('utf-8')[:40].decode('utf-8', errors='ignore').ljust(40)
            concepto3 = str(c3_texto).encode('utf-8')[:40].decode('utf-8', errors='ignore').ljust(40)

            # Estructura exacta de 15 columnas del DBF contable
            fila_mapeada = [
                tipo_partida_defecto,       # TIPPART  
                num_partida_defecto,        # NUMPART  
                fecha_defecto,              # FECHA    
                cuenta_actual,              # CODCTA   
                tipo_movimiento,            # TIPO (D u H según la cascada jerárquica real de disco)
                concepto1,                  # CONCEPTO1
                concepto2,                  # CONCEPTO2
                concepto3,                  # CONCEPTO3
                monto_transaccion,          # MONTO    
                0.0,                        # MONTO_AUX
                "",                         # B_TIPO
                0,                          # B_NUMERO
                "",                         # B_FECHACON
                "FALSO",                    # B_IMPRESIO
                0                           # CORPART
            ]
            
            lineas_partida.append(fila_mapeada)
            conceptos_acumulados = []  # Dejar limpio para la siguiente transacción del reporte
            
        else:
            if ("PASAN" not in contenido_linea.upper() and 
                "TOTALES" not in contenido_linea.upper() and 
                "---" not in contenido_linea and 
                "___" not in contenido_linea and
                "\x97" not in contenido_linea and
                not contenido_linea.startswith("-") and
                not contenido_linea.startswith("_")):
                conceptos_acumulados.append(contenido_linea)
                
    return lineas_partida

# =========================================================
# 4. Procesamiento de Planilla Excel Original (Indemnizaciones)
# =========================================================
def procesar_partida_diario_excel():
    """Ejecuta el flujo original basado en el libro de Excel (Indemnizaciones y Parámetros)."""
    archivo_excel = seleccionar_archivo()
    if not archivo_excel:
        return

    try:
        print("Cargando estructuras de Tablas de Excel...")
        wb = openpyxl.load_workbook(archivo_excel)

        # 1. LOCALIZAR LAS TRES TABLAS OFICIALES POR SU NOMBRE EN EL LIBRO
        ws_indem, tabla_indem = buscar_tabla_oficial(wb, "INDEMNIZACIONES")
        ws_param, tabla_param = buscar_tabla_oficial(wb, "PARAMETROS")
        ws_pda, tabla_pda = buscar_tabla_oficial(wb, "PDA")

        if not ws_indem or not ws_param or not ws_pda:
            raise ValueError("No se localizó alguna de las Tablas Oficiales (INDEMNIZACIONES, PARAMETROS o PDA).")

        # 2. OBTENER COORDENADAS EXACTAS CON range_boundaries
        min_col_i, min_row_i, max_col_i, max_row_i = range_boundaries(tabla_indem.ref)
        min_col_p, min_row_p, max_col_p, max_row_p = range_boundaries(tabla_param.ref)
        min_col_d, min_row_d, max_col_d, max_row_d = range_boundaries(tabla_pda.ref)

        # Extracción plana de cabeceras y datos
        cabeceras_i = [str(ws_indem.cell(row=min_row_i, column=c).value).strip() for c in range(min_col_i, max_col_i + 1)]
        datos_i = [[ws_indem.cell(row=r, column=c).value for c in range(min_col_i, max_col_i + 1)] for r in range(min_row_i + 1, max_row_i + 1)]
        df_indem = pd.DataFrame(datos_i, columns=cabeceras_i)

        cabeceras_p = [str(ws_param.cell(row=min_row_p, column=c).value).strip() for c in range(min_col_p, max_col_p + 1)]
        datos_p = [[ws_param.cell(row=r, column=c).value for c in range(min_col_p, max_col_p + 1)] for r in range(min_row_p + 1, max_row_p + 1)]
        df_param = pd.DataFrame(datos_p, columns=cabeceras_p)

        df_indem["AREA"] = df_indem["AREA"].astype(str).str.strip()
        df_param["AREA"] = df_param["AREA"].astype(str).str.strip()

        # Convertir a diccionario de parámetros normalizando las llaves
        param_dict = {}
        for idx_p, row_p in df_param.set_index("AREA").iterrows():
            area_key = str(idx_p).strip()
            param_dict[area_key] = {}
            for k, v in row_p.items():
                if pd.isna(v):
                    cuenta_limpia = ""
                else:
                    v_str = str(v).strip()
                    if v_str.endswith(".0"):
                        cuenta_limpia = v_str[:-2]
                    else:
                        cuenta_limpia = v_str
                
                cabecera_normalizada = str(k).strip().upper()
                if "NETO" in cabecera_normalizada:
                    param_dict[area_key]["NETO"] = cuenta_limpia
                else:
                    param_dict[area_key][cabecera_normalizada] = cuenta_limpia

        lineas_partida = []
        correlativo_partida = 1  # Asiento unificado por empleado
        filas_procesadas_indices = []

        rubros_gasto = ["SALARIO", "VACACION", "AGUINALDO", "INDEMNIZACION", "VIATICOS"]
        rubros_retencion = ["ISSS", "AFP", "ISR"]

        # 3. ANALIZAR FILA POR FILA APLICANDO REGLAS DE TRUNCADO Y STATUS
        for idx, row in df_indem.iterrows():
            nombre_original = str(row["NOMBRE"]).strip()

            if nombre_original == "nan" or nombre_original == "" or pd.isna(row["INGRESO"]):
                continue

            status_actual = str(row.get("STATUS", "")).strip().upper()
            if status_actual == "REGISTRADA":
                continue

            # Rellenamos de espacios vacíos fijos para no romper el motor dBASE del Bloque 2
            concepto1 = nombre_original[:40].ljust(40)
            area = str(row["AREA"]).strip()

            if area not in param_dict:
                raise ValueError(f"El área '{area}' de {nombre_original} no está parametrizada.")

            fecha_baja_dt = pd.to_datetime(row["BAJA"])
            f_ingreso = pd.to_datetime(row["INGRESO"]).strftime("%d/%m/%y")
            f_baja_glosa = fecha_baja_dt.strftime("%d/%m/%y")
            
            fecha_partida_dinamica = fecha_baja_dt.strftime("%d/%m/%Y")
            tippart_dinamico = f"D{fecha_baja_dt.strftime('%m')}"

            # Inicializadores para cuadre matemático por empleado
            total_debe_empleado = 0.0
            total_haber_retenciones = 0.0

            # --- Generar Cargos (D) ---
            for rubro in rubros_gasto:
                monto = 0.0
                for col_name in row.index:
                    if str(col_name).strip().upper() == rubro:
                        monto = limpiar_monto(row[col_name])
                        break
                
                if monto > 0:
                    total_debe_empleado += monto
                    codcta = param_dict[area].get(rubro, "")
                    concepto2 = f"{rubro} LIQ {f_ingreso} AL {f_baja_glosa}"[:40].ljust(40)
                    lineas_partida.append(
                        [tippart_dinamico, correlativo_partida, fecha_partida_dinamica, codcta, "D", concepto1, concepto2, "".ljust(40), round(monto, 2), 0.00, "", 0, "", "FALSO", 0]
                    )

            # --- Generar Abonos (H) de Retenciones ---
            for rubro in rubros_retencion:
                monto = 0.0
                for col_name in row.index:
                    if str(col_name).strip().upper() == rubro:
                        monto = limpiar_monto(row[col_name])
                        break
                        
                if monto > 0:
                    total_haber_retenciones += monto
                    codcta = param_dict[area].get(rubro, "")
                    label_glosa = f"RETENCION {rubro}"
                    concepto2 = f"{label_glosa} LIQ {f_ingreso} AL {f_baja_glosa}"[:40].ljust(40)
                    lineas_partida.append(
                        [tippart_dinamico, correlativo_partida, fecha_partida_dinamica, codcta, "H", concepto1, concepto2, "".ljust(40), round(monto, 2), 0.00, "", 0, "", "FALSO", 0]
                    )
            
            # --- GENERAR ABONO (H) DEL NETO DINÁMICO ---
            monto_neto_calculado = total_debe_empleado - total_haber_retenciones
            
            if monto_neto_calculado > 0:
                codcta_neto = param_dict[area].get("NETO", "")
                concepto2_neto = f"NETO A PAGAR LIQ {f_ingreso} AL {f_baja_glosa}"[:40].ljust(40)
                lineas_partida.append(
                    [tippart_dinamico, correlativo_partida, fecha_partida_dinamica, codcta_neto, "H", concepto1, concepto2_neto, "".ljust(40), round(monto_neto_calculado, 2), 0.00, "", 0, "", "FALSO", 0]
                )

            if total_debe_empleado > 0:
                filas_procesadas_indices.append(idx)
                correlativo_partida += 1

        if not lineas_partida:
            root_msg = configurar_ventana_emergente()
            messagebox.showinfo("Sin Cambios", "Todos los registros ya se encuentran marcados como REGISTRADA.")
            root_msg.destroy()
            wb.close()
            return

        ruta_guardar_dbf = seleccionar_destino_dbf()
        if not ruta_guardar_dbf:
            wb.close()
            return

        # 4. LIMPIAR Y ESCRIBIR EN LA TABLA DESTINO "PDA" REDIMENSIONANDO EL OBJETO
        if ws_pda.max_row > min_row_d:
            ws_pda.delete_rows(min_row_d + 1, amount=ws_pda.max_row - min_row_d)

        for r_idx, fila_datos in enumerate(lineas_partida):
            for c_idx, valor in enumerate(fila_datos):
                ws_pda.cell(row=min_row_d + 1 + r_idx, column=min_col_d + c_idx, value=valor)

        nueva_ref_pda = f"{get_column_letter(min_col_d)}{min_row_d}:{get_column_letter(max_col_d)}{min_row_d + len(lineas_partida)}"
        tabla_pda.ref = nueva_ref_pda

        # 5. ACTUALIZAR ESTADO DE STATUS EN LA TABLA ORIGEN "INDEMNIZACIONES"
        col_status_idx = None
        for c in range(min_col_i, max_col_i + 1):
            if str(ws_indem.cell(row=min_row_i, column=c).value).strip().upper() == "STATUS":
                col_status_idx = c
                break

        if col_status_idx is None:
            raise ValueError("No se encontró la columna 'STATUS' en la Tabla Oficial INDEMNIZACIONES.")

        # SOLUCIÓN DE ERROR: Usamos el nombre exacto de la lista 'filas_procesadas_indices'
        for idx_proc in filas_procesadas_indices:
            row_excel_indem = min_row_i + 1 + idx_proc
            ws_indem.cell(row=row_excel_indem, column=col_status_idx, value="REGISTRADA")

        wb.save(archivo_excel)
        wb.close()

        # 6. EXPORTAR A dBASE
        exportar_a_dbf(ruta_guardar_dbf, lineas_partida)

        root_msg = configurar_ventana_emergente()
        messagebox.showinfo("Carga Masiva Exitosa", f"Proceso Completado Exitosamente.\n1. Se actualizó el libro de Excel con las marcas de registro.\n2. Se creó el archivo dBASE en: {ruta_guardar_dbf}")
        root_msg.destroy()


# =========================================================
# 5. Orquestador TXT y Menú Ejecutivo de Arranque Dual
# =========================================================
def ejecutar_conversion_txt(ventana_padre):
    """Orquesta la lectura del TXT y despliega la opción de guardar en DBF o en Excel utilizando la ventana padre."""
    ruta_txt = seleccionar_archivo_txt(ventana_padre)
    if not ruta_txt:
        return

    try:
        # Procesamos el archivo contable usando el motor unificado de disco
        datos = procesar_txt_contable(ruta_txt)
        if not datos:
            messagebox.showwarning("Atención", "No se detectaron transacciones procesables en el TXT.", parent=ventana_padre)
            return
            
        # Ventana intermedia flotante para decidir el formato de la partida de diario
        ventana_opcion = tk.Toplevel(ventana_padre)
        ventana_opcion.title("Selecciona Formato")
        ventana_opcion.geometry("350x130")
        ventana_opcion.attributes("-topmost", True)
        ventana_opcion.grab_set()  # Bloquea interacciones externas en Windows
        
        lbl_opc = tk.Label(ventana_opcion, text="¿En qué formato deseas guardar el resultado?", font=("Arial", 10, "bold"), pady=10)
        lbl_opc.pack()
        
        def guardar_formato_dbf():
            ventana_opcion.destroy()
            ruta_dbf = seleccionar_destino_dbf(ventana_padre)
            if ruta_dbf:
                try:
                    exportar_a_dbf(ruta_dbf, datos)
                    messagebox.showinfo("Éxito", f"¡Partida de Diario Generada!\nArchivo guardado en: {ruta_dbf}", parent=ventana_padre)
                except Exception as e_dbf:
                    messagebox.showerror("Error dBASE", f"No se pudo escribir el archivo DBF:\n{str(e_dbf)}", parent=ventana_padre)
                    
        def guardar_formato_excel():
            ventana_opcion.destroy()
            ruta_xlsx = seleccionar_destino_excel(ventana_padre)
            if ruta_xlsx:
                try:
                    # BLINDAJE DE EXCEL: Ejecuta y captura si falta la librería openpyxl en el sistema
                    exportar_a_excel(ruta_xlsx, datos)
                    messagebox.showinfo("Éxito", f"¡Partida de Diario Generada!\nArchivo guardado en: {ruta_xlsx}", parent=ventana_padre)
                except Exception as e_xlsx:
                    messagebox.showerror("Error de Excel", f"Falta la librería openpyxl o el archivo está abierto.\nPor favor ejecuta en tu consola:\npip install openpyxl\n\nDetalle: {str(e_xlsx)}", parent=ventana_padre)
                    
        btn_dbf = tk.Button(ventana_opcion, text="Exportar a dBASE (.DBF)", command=guardar_formato_dbf, width=28, bg="#fff3cd")
        btn_dbf.pack(pady=3)
        btn_xls = tk.Button(ventana_opcion, text="Exportar a Excel (.XLSX)", command=guardar_formato_excel, width=28, bg="#d1ecf1")
        btn_xls.pack(pady=3)
        
    except Exception as e:
        messagebox.showerror("Error de Conversión", f"Error procesando el TXT:\n{str(e)}", parent=ventana_padre)

def lanzar_procesar_excel(ventana_menu):
    """Oculta temporalmente el menú para procesar Excel y lo restaura al finalizar."""
    ventana_menu.withdraw()
    procesar_partida_diario_excel()
    ventana_menu.deiconify()

if __name__ == "__main__":
    ventana_menu = tk.Tk()
    ventana_menu.title("Módulo de Carga Contable v2.5")
    ventana_menu.geometry("400x180")
    ventana_menu.attributes("-topmost", True)
    
    lbl = tk.Label(ventana_menu, text="¿Qué tipo de archivo deseas procesar hoy?", font=("Arial", 11, "bold"), pady=15)
    lbl.pack()
    
    btn1 = tk.Button(ventana_menu, text="Procesar Planilla (Excel a DBF)", command=lambda: lanzar_procesar_excel(ventana_menu), width=35, bg="#d4edda")
    btn1.pack(pady=5)
    
    btn2 = tk.Button(ventana_menu, text="Procesar Auxiliar de Mayor (TXT a DBF/Excel)", command=lambda: ejecutar_conversion_txt(ventana_menu), width=35, bg="#cce5ff")
    btn2.pack(pady=5)
    
    ventana_menu.mainloop()
