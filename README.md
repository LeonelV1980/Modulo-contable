# Modulo-contable
# =========================================================
# 1. Configuración de Librerías y Entorno de Dependencias
# =========================================================
import os
import re
import sys
from datetime import datetime

# Componentes Maestros de Interfaz Gráfica (Tkinter)
import tkinter as tk
from tkinter import messagebox, filedialog

# Motores de Base de Datos y Estructuración de Reportes
import dbf     # Manejo estricto de tablas dBASE (.dbf)
import pandas as pd  # Generador definitivo de archivos Excel (.xlsx)

# Nota de Control: Si el script arroja un error de importación al iniciar,
# ejecuta en tu terminal de Windows: pip install pandas openpyxl dbf

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
import pandas as pd

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

def seguro_float(valor):
    """Convierte de forma segura cualquier texto o vacío a número decimal (float)."""
    if not valor:
        return 0.0
    val_str = str(valor).strip().replace(',', '')
    try:
        return float(val_str)
    except ValueError:
        return 0.0

def exportar_a_dbf(ruta_dbf, lineas_partida):
    """Mapea de forma estricta por índices posicionales la lista de 15 campos generada por el TXT."""
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
            tippart_val    = str(fila[0]).strip()
            numpart_val    = seguro_int(fila[1])
            fecha_str      = str(fila[2]).strip()
            fecha_pda      = datetime.strptime(fecha_str, "%d/%m/%Y").date()
            codcta_val     = str(fila[3]).strip()
            tipo_val       = str(fila[4]).strip()
            concepto1_val  = str(fila[5]).strip()[:40]
            concepto2_val  = str(fila[6]).strip()[:40]
            concepto3_val  = str(fila[7]).strip()[:40]
            monto_val      = seguro_float(fila[8])
            monto_aux_val  = seguro_float(fila[9])
            b_tipo_val     = str(fila[10]).strip()
            b_numero_val   = seguro_int(fila[11])
            
            fecha_banco_str = str(fila[12]).strip()
            if fecha_banco_str and fecha_banco_str != "None" and "/" in fecha_banco_str:
                fecha_banco = datetime.strptime(fecha_banco_str, "%d/%m/%Y").date()
            else:
                fecha_banco = dbf.Date()
                
            b_impresio_str  = str(fila[13]).upper().strip()
            b_impresio_bool = True if b_impresio_str in ["VERDADERO", "TRUE", "T", "1"] else False
            corpart_val     = seguro_int(fila[14])

            registro = (
                tippart_val, numpart_val, fecha_pda, codcta_val, tipo_val,
                concepto1_val, concepto2_val, concepto3_val, monto_val, monto_aux_val,
                b_tipo_val, b_numero_val, fecha_banco, b_impresio_bool, corpart_val
            )
            tabla_dbf.append(registro)
            
        tabla_dbf.close()
        return True
    except Exception as e:
        print(f"Error al escribir el archivo DBF: {str(e)}")
        raise e

def grabar_mi_excel_pandas(ruta_excel, lineas_partida):
    """
    Escribe la matriz contable completa en las 15 columnas oficiales usando Pandas, 
    forzando que la pestaña de Excel se llame siempre 'PDA' para desactivar el error del UUID.
    """
    try:
        ruta_excel_fisica = os.path.normpath(ruta_excel)
        
        columnas = [
            'TIPPART', 'NUMPART', 'FECHA', 'CODCTA', 'TIPO', 
            'CONCEPTO1', 'CONCEPTO2', 'CONCEPTO3', 'MONTO', 
            'MONTO_AUX', 'B_TIPO', 'B_NUMERO', 'B_FECHACON', 
            'B_IMPRESIO', 'CORPART'
        ]
        
        df = pd.DataFrame(lineas_partida, columns=columnas)
        
        # Saneamiento estricto de textos para erradicar nulos
        df['B_TIPO'] = df['B_TIPO'].fillna("").astype(str)
        df['B_FECHACON'] = df['B_FECHACON'].fillna("").astype(str)
        df['CODCTA'] = df['CODCTA'].astype(str)
        df['CONCEPTO1'] = df['CONCEPTO1'].astype(str)
        df['CONCEPTO2'] = df['CONCEPTO2'].astype(str)
        df['CONCEPTO3'] = df['CONCEPTO3'].astype(str)
        
        # Inyección forzada en caliente con nombre estático
        df.to_excel(ruta_excel_fisica, index=False, sheet_name="PDA", engine='openpyxl')
        return True
        
    except Exception as e:
        print(f"Error fatal en el empaquetador Pandas: {str(e)}")
        raise e





# =====================================================================
# 3. Motor de Procesamiento Jerárquico - PARTE 1 (Auditoría y Cabecera)
# =====================================================================


def limpiar_monto(monto_str):
    """Limpia el texto de montos contables eliminando separadores de miles para su conversión a float."""
    if not monto_str:
        return 0.0
    try:
        return float(monto_str.strip().replace(',', ''))
    except ValueError:
        return 0.0

def procesar_txt_contable(ruta_txt):
    """
    Lee el TXT, determina D/H jerárquicamente por el ancho físico de la cuenta de mayor principal
    y unifica los conceptos respetando los vacíos originales del reporte de disco.
    """
    lineas_partida = []
    
    if not os.path.exists(ruta_txt):
        return lineas_partida
    
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

    # 1. ETAPA DE AUDITORÍA EN BRUTO: Evaluamos el ancho exacto de las líneas de disco
    with open(ruta_txt, 'r', encoding='latin-1', errors='ignore') as f:
        lineas_originales_disco = f.readlines()

    for linea_bruta in lineas_originales_disco:
        largo_linea_bruta = len(linea_bruta.replace('\r', '').replace('\n', ''))
        linea_limpia_b = linea_bruta.strip()

        # Captura de estructuras numéricas de cuentas
        match_c_b = re.match(r'^\s*(\d+)\s+(.+)$', linea_bruta)
        if match_c_b:
            cod_potencial = match_c_b.group(1).strip()
            
            # REGLA JERÁRQUICA MAESTRA POR ANCHO DE LÍNEA DE MAYOR (Exactamente 3 dígitos)
            if len(cod_potencial) == 3:
                if largo_linea_bruta == 118:
                    naturaleza_mayor_activa = "H"
                elif largo_linea_bruta == 100:
                    naturaleza_mayor_activa = "D"
                continue

            # ASIGNACIÓN DE SUBCUENTAS OPERATIVAS ANALÍTICAS (Nivel de detalle >= 5 dígitos)
            if len(cod_potencial) >= 5:
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
        fecha_defecto = f"{dia}/{meses.get(mes_txt, '01')}/{anio}"

    match_partida_txt = re.search(r'Partida\s+No\.\s*:\s*([A-Z0-9]+)\s*-\s*(\d+)', contenido_completo, re.IGNORECASE)
    if match_partida_txt:
        tipo_partida_defecto = match_partida_txt.group(1).strip()
        num_partida_defecto = int(match_partida_txt.group(2).strip())
    # =====================================================================
    # 3. Motor de Procesamiento Jerárquico - PARTE 2 (Alineación Fiel de Glosas)
    # =====================================================================

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

        # IDENTIFICACIÓN DE TRANSACCIONES REFORZADA: Soporta montos contables finales
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
            tipo_movimiento = lista_secuencial_naturalezas.pop(0) if len(lista_secuencial_naturalezas) > 0 else "D"

            # --- DETECTOR ADAPTATIVO ANTI-DESFASE DE CONCEPTOS VACÍOS ---
            total_conceptos = len(conceptos_acumulados)
            
            # Verificamos si el último elemento acumulado tiene la estructura física de un UUID
            es_uuid = False
            if total_conceptos > 0:
                if re.search(r'[A-F0-9]{8}-[A-F0-9]{4}', conceptos_acumulados[-1], re.IGNORECASE):
                    es_uuid = True

            if total_conceptos == 2 and es_uuid:
                # CASO ESPECIAL CORREGIDO: Dejaron el concepto 1 vacío y agregaron concepto 2 y concepto 3 (UUID)
                concepto1 = ""
                concepto2 = str(conceptos_acumulados.pop(0)).strip()
                concepto3 = str(conceptos_acumulados.pop(0)).strip()
            elif total_conceptos >= 3:
                concepto1 = str(conceptos_acumulados.pop(0)).strip()
                concepto2 = str(conceptos_acumulados.pop(0)).strip()
                concepto3 = str(conceptos_acumulados.pop(0)).strip()
            elif total_conceptos == 1:
                concepto1 = str(conceptos_acumulados.pop(0)).strip()
                concepto2 = ""
                concepto3 = ""
            else:
                concepto1 = "CONTINUACION DE REGISTRO"
                concepto2 = ""
                concepto3 = ""

            # --- COMPLETADO SANEADO DE LAS 15 COLUMNAS CONTABLES ---
            monto_aux_val  = 0.0
            b_tipo_val     = ""
            b_numero_val   = 0
            b_fechacon_val = ""
            b_impresio_val = False
            corpart_val    = len(lineas_partida) + 1

            # Estructura de la matriz que alimentará tanto a dBASE como a Pandas de forma simétrica
            fila_mapeada = [
                tipo_partida_defecto,       # 0. TIPPART  
                num_partida_defecto,        # 1. NUMPART  
                fecha_defecto,              # 2. FECHA    
                cuenta_actual,              # 3. CODCTA   
                tipo_movimiento,            # 4. TIPO 
                concepto1,                  # 5. CONCEPTO1
                concepto2,                  # 6. CONCEPTO2
                concepto3,                  # 7. CONCEPTO3
                monto_transaccion,          # 8. MONTO (Aquí cae el valor 2.71, 12.77, etc.)
                monto_aux_val,              # 9. MONTO_AUX
                b_tipo_val,                 # 10. B_TIPO
                b_numero_val,               # 11. B_NUMERO
                b_fechacon_val,             # 12. B_FECHACON
                b_impresio_val,             # 13. B_IMPRESIO
                corpart_val                 # 14. CORPART
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

def seleccionar_archivo_txt(ventana_padre):
    """Muestra el cuadro para abrir el archivo TXT de origen."""
    ruta = filedialog.askopenfilename(
        title="Selecciona el Auxiliar de Mayor (TXT)",
        filetypes=[("Archivos de Texto", "*.txt"), ("Todos los archivos", "*.*")],
        parent=ventana_padre
    )
    return ruta

def seleccionar_destino_dbf(ventana_padre):
    """Muestra el cuadro para guardar el archivo DBF de salida."""
    ruta = filedialog.asksaveasfilename(
        title="Guardar Partida de Diario (dBASE)",
        filetypes=[("Archivos dBASE", "*.dbf")],
        defaultextension=".dbf",
        parent=ventana_padre
    )
    return ruta

def seleccionar_destino_excel(ventana_padre):
    """Muestra el cuadro de diálogo estándar para elegir la ruta de guardado del Excel."""
    ruta = filedialog.asksaveasfilename(
        title="Guardar Partida de Diario en Excel",
        filetypes=[("Archivos de Excel", "*.xlsx")],
        defaultextension=".xlsx",
        parent=ventana_padre
    )
    return ruta







# =========================================================
# 5. Orquestador TXT y Menú Ejecutivo de Arranque Dual
# =========================================================

def ejecutar_conversion_txt(ventana_padre):
    """Orquesta la lectura del TXT y despliega la opción de guardar en DBF o en Excel."""
    ruta_txt = seleccionar_archivo_txt(ventana_padre)
    if not ruta_txt:
        return

    try:
        datos = procesar_txt_contable(ruta_txt)
        if not datos:
            messagebox.showwarning("Atención", "No se detectaron transacciones procesables en el TXT.", parent=ventana_padre)
            return

        ventana_opcion = tk.Toplevel(ventana_padre)
        ventana_opcion.title("Selecciona Formato")
        ventana_opcion.geometry("350x130")
        ventana_opcion.attributes("-topmost", True)
        ventana_opcion.grab_set()

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
                    # --- CONEXIÓN AJUSTADA AL NUEVO MOTOR ---
                    # Obligamos a Python a ejecutar la función renombrada con Pandas
                    grabar_mi_excel_pandas(ruta_xlsx, datos)
                    messagebox.showinfo("Éxito", f"¡Partida de Diario Generada!\nArchivo Excel guardado con éxito en:\n{ruta_xlsx}", parent=ventana_padre)
                except Exception as e_xlsx:
                    messagebox.showerror("Error de Excel", f"No se pudo salvar el archivo.\nDetalle: {str(e_xlsx)}", parent=ventana_padre)

        btn_dbf = tk.Button(ventana_opcion, text="Exportar a dBASE (.DBF)", command=guardar_formato_dbf, width=28, bg="#fff3cd")
        btn_dbf.pack(pady=3)
        btn_xls = tk.Button(ventana_opcion, text="Exportar a Excel (.XLSX)", command=guardar_formato_excel, width=28, bg="#d1ecf1")
        btn_xls.pack(pady=3)

    except Exception as e:
        messagebox.showerror("Error de Conversión", f"Error procesando el TXT:\n{str(e)}", parent=ventana_padre)


def lanzar_procesar_excel(ventana_menu):
    """Oculta la ventana principal, procesa la planilla de Excel y la restaura al terminar."""
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
