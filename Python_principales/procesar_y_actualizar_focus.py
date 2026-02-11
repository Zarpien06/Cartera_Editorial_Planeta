import os
import sys
import json
import shutil
import traceback
import re
import time
import logging
from logging.handlers import RotatingFileHandler
from typing import Optional, Tuple, Dict, Any, List, Union
from pathlib import Path
from datetime import datetime, timedelta

import pandas as pd
import openpyxl
from openpyxl import load_workbook, Workbook
from openpyxl.worksheet.worksheet import Worksheet
from openpyxl.styles import PatternFill
# Importar xlrd y xlwt para manejo de archivos .xls
try:
    import xlrd
    import xlwt
    XLRD_AVAILABLE = True
except ImportError:
    XLRD_AVAILABLE = False

# Nombre del archivo de salida FOCUS procesado
OUTPUT_FILENAME = "FOCUS_ACTUALIZADO.xlsx"

# Directorios base del proyecto
BASE_DIR = Path(__file__).parent
SALIDAS_DIR = BASE_DIR / "salidas"
BACKUP_DIR = SALIDAS_DIR / "backup"

def encontrar_hoja_por_nombre(wb, nombre_buscar: str) -> str:
    nombre_buscar = str(nombre_buscar).strip().lower()
    for nombre_hoja in wb.sheetnames:
        if str(nombre_hoja).strip().lower() == nombre_buscar:
            return nombre_hoja
    raise ValueError(f"No se encontró la hoja: {nombre_buscar}")

# Marcador para identificar archivos generados por este script
GEN_MARKER_KEY = "_FOCUS_GENERATED_BY"
GEN_MARKER_VAL = "procesar_y_actualizar_focus_v1"

def workbook_has_generation_marker(wb: openpyxl.Workbook, ws: Worksheet) -> bool:
    """Detecta si el libro fue generado previamente por este script."""
    try:
        if str(ws["Z1"].value).strip() == GEN_MARKER_VAL:
            return True
    except Exception:
        pass
    try:
        subject = getattr(wb.properties, "subject", "") or ""
        return GEN_MARKER_VAL in str(subject)
    except Exception:
        return False

def set_workbook_generation_marker(wb: openpyxl.Workbook, ws: Worksheet) -> None:
    """Escribe el marcador en el libro y hoja activa para futuras detecciones."""
    try:
        ws["Z1"].value = GEN_MARKER_VAL
        ws.column_dimensions['Z'].hidden = True
    except Exception:
        pass
    try:
        props = wb.properties
        current_subject = props.subject or ""
        if GEN_MARKER_VAL not in current_subject:
            props.subject = (current_subject + " " + GEN_MARKER_VAL).strip()
    except Exception:
        pass

def limpiar_archivos_antiguos(directorio: Path, horas: int = 24) -> None:
    """Elimina archivos con antigüedad mayor a 'horas' en el directorio dado."""
    try:
        limite = datetime.now().timestamp() - horas * 3600
        if directorio.is_dir():
            for p in directorio.iterdir():
                try:
                    if p.is_file() and p.stat().st_mtime < limite:
                        p.unlink(missing_ok=True)
                except Exception:
                    pass
    except Exception:
        pass

# Limpieza inicial al cargar el módulo
limpiar_archivos_antiguos(SALIDAS_DIR, horas=24)
limpiar_archivos_antiguos(BACKUP_DIR, horas=24)

# Configurar el nivel de logging
logger = logging.getLogger(__name__)
logger.setLevel(logging.INFO)
logger.addHandler(logging.StreamHandler())
logger.addHandler(RotatingFileHandler(BASE_DIR / 'focus_processor.log', maxBytes=1048576, backupCount=3, encoding='utf-8'))

from log_bridge import attach_python_logging, log_event
attach_python_logging('PY_FOCUS', logger)
logger.info('Iniciando procesador FOCUS')

MESES = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"]

CUENTAS_BALANCE = [
    # Cuentas explícitas requeridas por FORMATO DEUDA 1.pdf
    "Total cuenta objeto 43001",
    "Total cuenta objeto 43008",
    "Total cuenta objeto 43042",

    # Cuentas detalladas (rangos y cuentas específicas) - BALANCE
    "0080.43002.20", "0080.43002.21", "0080.43002.15",
    "0080.43002.28", "0080.43002.31", "0080.43002.63",

    # Cuentas adicionales de 43002
    "0080.43002.29", "0080.43002.57",
    "0080.43002.60", "0080.43002.62",
    "0080.43002.64", "0080.43002.65", "0080.43002.66", "0080.43002.69",

    # Cuentas 01010 (clientes/grupo planeta y subcuentas)
    "0080.01010.00010",  # CLIENTE GR.ED.PLANETA
    "0080.01010.00600",  # CLIENTE GR.PLANETA CHILE
    "0080.01010.10801", "0080.01010.10802", "0080.01010.10805",
    "0080.01010.10808", "0080.01010.10809", "0080.01010.10817",
    "0080.01010.10818", "0080.01010.10819",

    # Patrones más amplios para captura por prefijo
    "0080.01010",  # todas las cuentas de este rango
    "0080.43002",  # todas las cuentas de este rango
]


def obtener_mes_siguiente(mes_actual: str) -> str:
    """Dado un nombre de mes en español, devuelve el siguiente mes.
    Si no reconoce el mes, devuelve el mismo valor por seguridad."""
    try:
        idx = int(MESES.index(mes_actual))
        return MESES[(idx + 1) % 12]
    except Exception:
        # Intentar normalizar
        lower_map = {m.lower(): i for i, m in enumerate(MESES)}
        key = str(mes_actual).strip().lower()
        if key in lower_map:
            return MESES[(int(lower_map[key]) + 1) % 12]
        return mes_actual


def normalize_text(value: str) -> str:
    """Normaliza texto: minúsculas, preservando ñ y caracteres especiales del español."""
    try:
        import unicodedata
        s = str(value).lower()
        # Normalizar pero preservar ñ y otros caracteres especiales del español
        s = unicodedata.normalize('NFC', s)
        # Reemplazar separadores comunes pero preservar caracteres especiales
        return ''.join(ch if ch.isalnum() or ch.isspace() or ch in 'ñáéíóúü' else ' ' for ch in s)
    except Exception:
        return str(value).lower()


def to_float(valor) -> float:
    """Convierte cualquier valor a float de forma segura"""
    if pd.isna(valor) or valor is None:
        return 0.0
    if isinstance(valor, (int, float)):
        return float(valor)
    try:
        s = str(valor).strip().replace(',', '').replace(' ', '')
        return float(re.sub(r'[^\d.-]', '', s))
    except:
        return 0.0


def buscar_archivo(directorio: Path, patrones: list, extensiones=None, excluir: Optional[List[str]] = None) -> Optional[Path]:
    """
    Busca el archivo más reciente que coincida con alguno de los patrones
    
    Args:
        directorio: Directorio donde buscar los archivos
        patrones: Lista de patrones a buscar en los nombres de archivo
        extensiones: Lista de extensiones permitidas (ej: ['.xlsx', '.xls']). Si es None, acepta todas.
        
    Returns:
        Path al archivo más reciente que coincide con algún patrón, o None si no se encuentra
    """
    archivos = []
    excluir = excluir or ["backup", "old", "prev", "temp"]
    for archivo in directorio.iterdir():
        if archivo.is_file():
            nombre = archivo.name.lower()
            # Verificar si el nombre del archivo contiene alguno de los patrones
            if any(patron.lower() in nombre for patron in patrones):
                # Excluir nombres indeseados
                if any(bad in nombre for bad in excluir):
                    continue
                # Si se especifican extensiones, verificar que el archivo tenga una de ellas
                if extensiones is None or any(nombre.endswith(ext.lower()) for ext in extensiones):
                    archivos.append(archivo)
    
    # Devolver el archivo más reciente si hay coincidencias
    return max(archivos, key=lambda x: x.stat().st_mtime) if archivos else None


def leer_excel(ruta: Path, hoja=0):
    """
    Lee archivo Excel con manejo de errores y soporte para .xls y .xlsx
    
    Args:
        ruta: Ruta al archivo Excel
        hoja: Índice o nombre de la hoja a leer (por defecto: primera hoja)
        
    Returns:
        DataFrame de pandas con los datos de la hoja especificada
        
    Raises:
        ValueError: Si el formato del archivo no es soportado
        Exception: Si ocurre un error al leer el archivo
    """
    try:
        # Verificar que el archivo existe
        if not ruta.exists():
            raise FileNotFoundError(f"El archivo no existe: {ruta}")
            
        # Verificar que el archivo no esté vacío
        if ruta.stat().st_size == 0:
            raise ValueError(f"El archivo está vacío: {ruta}")
            
        # Seleccionar el motor apropiado según la extensión
        file_ext = ruta.suffix.lower()
        
        if file_ext in ['.xlsx', '.xlsm', '.xltx', '.xltm']:
            # Usar openpyxl para archivos .xlsx y formatos modernos
            print(f"  [INFO] Leyendo archivo {ruta.name} con motor openpyxl...")
            return pd.read_excel(ruta, sheet_name=hoja, engine='openpyxl')
        elif file_ext == '.xls':
            # Usar xlrd 2.0.1+ para archivos .xls modernos
            print(f"  [INFO] Leyendo archivo {ruta.name} con motor xlrd...")
            return pd.read_excel(ruta, sheet_name=hoja, engine='xlrd')
        else:
            raise ValueError(f"Formato de archivo no soportado: {file_ext}")
        
    except Exception as e:
        print(f"\n[ERROR] Error al leer el archivo {ruta.name}:")
        print(f"   Tipo de error: {type(e).__name__}")
        print(f"   Mensaje: {str(e)}")
        print(f"   Ruta completa: {ruta.absolute()}")
        print(f"   Tamaño: {ruta.stat().st_size/1024:.2f} KB" if ruta.exists() else "   El archivo no existe")
        raise


def detectar_mes_archivo(archivo: Path) -> str:
    """Detecta el mes del archivo FOCUS"""
    nombre = archivo.name.lower()
    
    print(f"  Analizando archivo: {archivo.name}")
    
    # Buscar por número de mes (01-12)
    for i, mes in enumerate(MESES, 1):
        if f"_{i:02d}_" in nombre or f"_{i:02d}." in nombre:
            print(f"  Mes detectado por número ({i:02d}): {mes}")
            return mes
    
    # Buscar por nombre de mes en el nombre del archivo
    for mes in MESES:
        if mes.lower() in nombre:
            print(f"  Mes detectado por nombre: {mes}")
            return mes
    
    # Si no se encuentra, intentar leer del contenido (todas las hojas)
    try:
        if str(archivo).lower().endswith('.xls'):
            # Para archivos .xls usar xlrd si está disponible
            if XLRD_AVAILABLE:
                wb = xlrd.open_workbook(str(archivo), on_demand=True)
                ws = wb.sheet_by_index(0)
                try:
                    valor_b5 = ws.cell(4, 1).value  # B5 es fila 4, columna 1 (0-based)
                except Exception:
                    valor_b5 = ""
            else:
                valor_b5 = ""
        else:
            # Para archivos .xlsx usar openpyxl y revisar todas las hojas
            wb = load_workbook(archivo, read_only=True, data_only=True)
            try:
                # Revisar B5 y H7 en todas las hojas hasta encontrar un mes válido
                for ws in wb.worksheets:
                    try:
                        valor_b5 = ws['B5'].value
                    except Exception:
                        valor_b5 = None
                    if valor_b5:
                        for mes in MESES:
                            if mes.lower() in str(valor_b5).lower():
                                print(f"  Mes detectado por B5 en hoja '{ws.title}': {mes}")
                                return mes
                    try:
                        valor_h7 = ws['H7'].value
                    except Exception:
                        valor_h7 = None
                    if valor_h7:
                        for mes in MESES:
                            if mes.lower() in str(valor_h7).lower():
                                print(f"  Mes detectado por H7 en hoja '{ws.title}': {mes}")
                                return mes
            finally:
                wb.close()
            
    except Exception as e:
        print(f"  [WARN] No se pudo leer el contenido del archivo para detectar el mes: {e}")
    
    # Default basado en el nombre del archivo
    if 'julio' in nombre:
        return "Julio"
    elif 'junio' in nombre:
        return "Junio"
    elif 'mayo' in nombre:
        return "Mayo"
    
    print(f"  Usando mes por defecto: Junio")
    return "Junio"  # Default


def obtener_trm_detallada(directorio: Path) -> dict:
    """
    Obtiene TRM detallada (usd, eur, fecha) desde trm_config.
    
    Args:
        directorio: Directorio base donde buscar la configuración
        
    Returns:
        dict: Diccionario con las claves: usd, eur, fecha, actualizado_por, origen
    """
    try:
        from trm_config import load_trm
        trm_data = load_trm()
        
        if not trm_data or not isinstance(trm_data, dict):
            raise ValueError("Datos de TRM no válidos")
            
        print(f"  TRM cargada desde trm_config: {trm_data.get('usd')} USD, {trm_data.get('eur')} EUR")
        return trm_data
        
    except Exception as e:
        print(f"  [ERROR] No se pudo cargar la TRM: {str(e)}")
        # Valores por defecto en caso de error
        return {
            'usd': 4000.0,
            'eur': 4500.0,
            'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'actualizado_por': 'sistema',
            'origen': 'error',
            'error': str(e)
        }

def obtener_trm(directorio: Path) -> Optional[float]:
    """Compatibilidad: devuelve sólo EUR si está disponible."""
    data = obtener_trm_detallada(directorio)
    return float(data['eur']) if data and 'eur' in data else None


def actualizar_celdas_trm(ws, base_dir: Path, mes_actual: Optional[str] = None) -> bool:
    """
    Actualiza todas las celdas relacionadas con TRM en la hoja de trabajo.
    
    Args:
        ws: La hoja de trabajo de Excel
        base_dir: Directorio base donde buscar la configuración de TRM
        mes_actual: Mes actual para mostrar en formato "Junio (Mes)"
        
    Returns:
        bool: True si la actualización fue exitosa, False en caso contrario
    """
    print("\nActualizando celdas con valores de TRM...")
    
    try:
        # Obtener la información TRM una sola vez
        trm_info = obtener_trm_detallada(base_dir)
        
        if not trm_info:
            print("  [ERROR] No se pudo obtener la información TRM")
            return False
            
        # Extraer valores TRM
        trm_eur = trm_info.get('eur')
        trm_usd = trm_info.get('usd')
        trm_fecha = trm_info.get('fecha')
        
        if not all([trm_eur, trm_usd, trm_fecha]):
            print(f"  [ERROR] Información TRM incompleta: EUR={trm_eur}, USD={trm_usd}, Fecha={trm_fecha}")
            return False
            
        print(f"  TRM actual: EUR={trm_eur}, USD={trm_usd}, Fecha={trm_fecha}")
        
        # 1. Actualizar tabla de tipo de cambio con nombres de mes
        # Área TIPO DE CAMBIO - CIERRE (generalmente alrededor de E5:G8)
        if mes_actual:
            try:
                # Buscar las celdas que contienen los meses
                # Formato: "Junio (Mes)" con valor 4.777,18
                mes_anterior_idx = MESES.index(mes_actual) - 1 if mes_actual in MESES else 0
                mes_anterior = MESES[mes_anterior_idx % 12]
                
                # Calcular promedio
                promedio_trm = trm_eur
                
                print(f"  [TRM] Mes actual: {mes_actual}, Mes anterior: {mes_anterior}")
                print(f"  [TRM] Promedio TRM EUR: {promedio_trm}")
                
            except Exception as e:
                print(f"  [WARN] Error calculando TRMs históricas: {e}")
        
        # 2. Actualizar tabla de conversión (generalmente en área derecha)
        # FECHA / DÓLAR TRM / EURO/PESO COL
        try:
            # Buscar y actualizar la tabla de conversión
            # Típicamente se encuentra alrededor de las columnas Q-S
            fecha_obj = datetime.strptime(trm_fecha, '%Y-%m-%d %H:%M:%S') if isinstance(trm_fecha, str) else datetime.now()
            fecha_formato = fecha_obj.strftime('%m/%d/%Y')
            
            print(f"  [TRM] Fecha formateada para tabla: {fecha_formato}")
            print(f"  [TRM] USD: {trm_usd}, EUR: {trm_eur}")
            
        except Exception as e:
            print(f"  [WARN] Error formateando fecha para tabla: {e}")
        
        # 3. Actualizar celdas de encabezado (Q3, Q4, P3, P4) 
        # Q3: Fecha actual del sistema
        # P3, P4: No modificar (dejar como estaban)
        # Q4: Valores TRM EUR
        try:
            # Q3: Fecha actual del sistema en formato dd/mm/yyyy
            fecha_actual = datetime.now()
            ws['Q3'].value = fecha_actual
            # Aplicar formato de fecha dd/mm/yyyy
            try:
                ws['Q3'].number_format = 'dd/mm/yyyy'
            except:
                pass
            print(f"  [OK] Actualizada Q3 con fecha actual: {fecha_actual.strftime('%d/%m/%Y')}")
            
            # P3: No modificar - dejar como está
            print(f"  [INFO] P3: Dejando valor original sin modificar")
            
            # P4: No modificar - dejar como está
            print(f"  [INFO] P4: Dejando valor original sin modificar")
            
            # Q4: Valor TRM EUR
            cell_q4 = ws['Q4']
            if cell_q4.data_type != 'f':  # Solo si no es fórmula
                cell_q4.value = trm_eur
                aplicar_formato_numero(ws, "Q4")
                print(f"  [OK] Actualizada Q4 con TRM EUR: {trm_eur}")
            else:
                print(f"  [LOCK] Q4: Se preserva la fórmula existente: {cell_q4.value}")
                
        except Exception as e:
            print(f"  [WARN] Error actualizando celdas de encabezado: {e}")
        
        # 4. Actualizar celdas de pie de página (Q42, R42, S42)
        try:
            # Actualizar fecha (Q42) con la fecha del archivo TRM
            ws['Q42'].value = trm_fecha
            print(f"  [OK] Actualizada Q42 con fecha: {trm_fecha}")
            
            # Actualizar USD (R42)
            ws['R42'].value = trm_usd
            aplicar_formato_numero(ws, "R42")
            print(f"  [OK] Actualizada R42 con TRM USD: {trm_usd}")
            
            # Actualizar EUR (S42)
            ws['S42'].value = trm_eur
            aplicar_formato_numero(ws, "S42")
            print(f"  [OK] Actualizada S42 con TRM EUR: {trm_eur}")
                
            # Asegurar que J7 tenga el valor correcto
            if ws['J7'].value is None or ws['J7'].data_type != 'f':
                ws['J7'].value = trm_eur
                aplicar_formato_numero(ws, "J7")
                print(f"  [OK] Actualizada J7 con TRM EUR: {trm_eur}")
                
            print("  [OK] Actualización de celdas TRM completada")
            return True
            
        except Exception as e:
            print(f"  [ERROR] Error actualizando celdas de pie de página: {e}")
            return False
            
    except Exception as e:
        print(f"  [ERROR] Error en el proceso de actualización de TRM: {e}")
        import traceback
        print(f"  Detalles del error: {traceback.format_exc()}")
        return False


def procesar_balance(df: pd.DataFrame) -> float:
    """
    1. BALANCE: Suma las cuentas especificadas
    Retorna el total en la unidad original (no dividido)
    """
    print("\n=== PROCESANDO BALANCE ===")

    # Intentar detectar columnas exactas usadas en tus archivos
    possible_account_cols = ["Número cuenta", "Numero cuenta", "Cuenta", "Número de cuenta", "Nro cuenta"]
    possible_saldo_cols = ["Saldo AAF variación", "Saldo AAF variacion", "Saldo", "Saldo periodo desviación", "Saldo periodo desviacion", "Libro mayor Saldo", "Saldo AAF"]

    col_cuenta = None
    col_saldo = None

    for c in possible_account_cols:
        if c in df.columns:
            col_cuenta = c
            break
    if col_cuenta is None:
        # fallback: primera columna no numérica
        for c in df.columns:
            # si la columna contiene strings (o tiene dtype object) asumimos que puede ser la cuenta
            is_object_dtype: bool = df[c].dtype == object
            has_strings: bool = df[c].apply(lambda x: isinstance(x, str)).any() == True
            if is_object_dtype or has_strings:
                col_cuenta = c
                break

    for c in possible_saldo_cols:
        if c in df.columns:
            col_saldo = c
            break
    if col_saldo is None:
        # fallback: columna con más valores numéricos
        numeric_counts: List[Tuple[Any, int]] = []
        for c in df.columns:
            def is_valid_number(x):
                if pd.isna(x):
                    return False
                return pd.api.types.is_number(x)
            is_numeric = df[c].apply(is_valid_number)
            count = int(is_numeric.sum())  # type: ignore
            numeric_counts.append((c, count))
        # Sort by count only, using negative count to get descending order
        # This avoids tuple comparison which causes type errors with Hashable column names
        numeric_counts = sorted(numeric_counts, key=lambda x: -x[1])
        if numeric_counts and numeric_counts[0][1] > 0:
            col_saldo = numeric_counts[0][0]
        else:
            # por último, tomar la última columna
            col_saldo = df.columns[-1]

    if col_cuenta is None or col_saldo is None:
        raise ValueError("No se pudieron detectar las columnas de cuenta o saldo en el archivo de balance. Columnas encontradas: " + str(list(df.columns)))

    print(f"  Columna cuenta: {col_cuenta}")
    print(f"  Columna saldo: {col_saldo}")

    # Normalizar columnas: eliminar espacios extra y convertir a str para búsqueda
    df = df.copy()
    df[col_cuenta] = df[col_cuenta].astype(str).str.strip().fillna("")
    # Reemplazar espacios no-break y caracteres invisibles
    df[col_cuenta] = df[col_cuenta].apply(lambda s: re.sub(r'\s+', ' ', s, flags=re.UNICODE).strip())

    total = 0.0
    coincidencias = []
    cuentas_procesadas = set()  # Para evitar duplicados

    # Primero intentar coincidencias exactas con la lista CUENTAS_BALANCE
    for cuenta in CUENTAS_BALANCE:
        # Solo coincidencia exacta, sin búsqueda parcial para evitar falsos positivos
        mask_exact = df[col_cuenta].str.strip().str.lower() == str(cuenta).strip().lower()
        if mask_exact.any():
            # Verificar que el valor sea numérico antes de sumar
            valores = df.loc[mask_exact, col_saldo].apply(to_float)
            # Filtrar valores extremadamente grandes que podrían ser errores
            valores = valores[abs(valores) < 1e12]  # Filtrar valores mayores a 1 billón
            subtotal = valores.sum()
            
            if abs(subtotal) > 0:  # Solo sumar si el valor es significativo
                # Verificar que la cuenta no haya sido procesada antes
                cuenta_key = str(cuenta).strip().lower()
                if cuenta_key not in cuentas_procesadas:
                    total += subtotal
                    cuentas_procesadas.add(cuenta_key)
                    coincidencias.append((cuenta, subtotal, "exacta"))
                    print(f"  [EXACTA] {cuenta}: {subtotal:,.2f}")
                else:
                    print(f"  [DUPLICADA] {cuenta}: {subtotal:,.2f} (ya procesada)")

    # Buscar específicamente por los totales que necesitamos
    print("  Buscando totales específicos en el nombre de cuenta...")
    
    # Definir patrones de cuentas de total que queremos incluir
    patrones_totales = [
        'TOTAL ACTIVO', 'TOTAL DEL ACTIVO', 'TOTAL ACTIVOS',
        'TOTAL PASIVO', 'TOTAL DEL PASIVO', 'TOTAL PASIVOS',
        'TOTAL PATRIMONIO', 'TOTAL DEL PATRIMONIO'
    ]
    
    # Inicializar subtotal para totales
    subtotal_total = 0.0
    
    # Primero, verificar si hay una fila que sea exactamente "TOTAL"
    mask_total = (df[col_cuenta].str.strip().str.upper() == 'TOTAL')
    if mask_total.any():
        subtotal = df.loc[mask_total, col_saldo].apply(to_float).sum()
        if abs(subtotal) > 0 and abs(subtotal) < 1e12:  # Filtro para valores extremos
            subtotal_total += subtotal
            coincidencias.append(("TOTAL", subtotal, "total_general"))
            print(f"  [TOTAL GENERAL] {subtotal:,.2f}")
    
    # Si no encontramos un TOTAL general, buscar por patrones específicos
    if subtotal_total == 0:
        for patron in patrones_totales:
            mask = df[col_cuenta].str.strip().str.upper() == patron.upper()
            if mask.any():
                subtotal = df.loc[mask, col_saldo].apply(to_float).sum()
                if abs(subtotal) > 0 and abs(subtotal) < 1e12:  # Filtro para valores extremos
                    # Verificar que el patrón no haya sido procesado antes
                    patron_key = str(patron).strip().lower()
                    if patron_key not in cuentas_procesadas:
                        subtotal_total += subtotal
                        cuentas_procesadas.add(patron_key)
                        coincidencias.append((patron, subtotal, "total_especifico"))
                        print(f"  [TOTAL] {patron}: {subtotal:,.2f}")
    
    # Solo sumar si encontramos algún total específico
    if abs(subtotal_total) > 0 and abs(subtotal_total) < 1e12:  # Validación adicional
        total += subtotal_total
    else:
        print("  [WARN] Se ignoró un total con valor sospechoso:", subtotal_total)
    
    # Buscar también filas que NO contengan "TOTAL" pero que sean cuentas válidas
    print("\n  === DETALLE DE CÁLCULO DEL BALANCE ===")
    print("  ---------------------------------")
    print("  CUENTAS ENCONTRADAS Y SUS VALORES:")
    
    # Mostrar coincidencias exactas
    if any(t[2] == "exacta" for t in coincidencias):
        print("\n  [COINCIDENCIAS EXACTAS]")
        for cuenta, valor, tipo in coincidencias:
            if tipo == "exacta":
                print(f"  - {cuenta}: {valor:,.2f}")
    
    # Mostrar totales encontrados
    if any(t[2] in ["total_general", "total_especifico"] for t in coincidencias):
        print("\n  [TOTALES ENCONTRADOS]")
        for cuenta, valor, tipo in coincidencias:
            if tipo in ["total_general", "total_especifico"]:
                print(f"  - {cuenta}: {valor:,.2f}")
    
    # Buscar filas sin 'TOTAL' pero con cuentas válidas que comiencen con "0080"
    print("\n  [BUSCANDO CUENTAS 0080]")
    # Buscar cuentas que comienzan con "0080" (sin incluir las que contienen "TOTAL")
    mask_no_total = ~df[col_cuenta].str.contains('total', case=False, na=False)
    mask_cuentas_0080 = df[col_cuenta].str.contains(r'^0080\.', case=False, na=False, regex=True)
    mask_final = mask_no_total & mask_cuentas_0080
    
    if mask_final.any():
        cuentas_encontradas = []
        for idx, row in df[mask_final].iterrows():
            cuenta_str = str(row[col_cuenta]).strip()
            valor = to_float(row[col_saldo])
            if valor != 0:
                # Verificar que la cuenta no haya sido procesada antes
                cuenta_key = cuenta_str.lower()
                if cuenta_key not in cuentas_procesadas:
                    cuentas_encontradas.append((cuenta_str, valor))
                    total += valor
                    cuentas_procesadas.add(cuenta_key)
                    coincidencias.append((cuenta_str, valor, "no_total_0080"))
                else:
                    print(f"  [DUPLICADA] {cuenta_str}: {valor:,.2f} (ya procesada)")
        
        if cuentas_encontradas:
            print("  Se encontraron las siguientes cuentas 0080:")
            for cuenta, valor in cuentas_encontradas:
                print(f"  - {cuenta}: {valor:,.2f}")
    
    # Búsqueda por patrones numéricos (solo si no hubo coincidencias previas)
    if not coincidencias:
        print("\n  [BÚSQUEDA POR PATRONES NUMÉRICOS]")
        print("  No se encontraron coincidencias con CUENTAS_BALANCE. Buscando patrones alternativos...")
        patrones = ["0080.43002", "0080.01010", "43001", "43008", "43042"]
        cuentas_patron = []
        
        for idx, row in df.iterrows():
            cuenta_val = str(row[col_cuenta])
            if any(p in cuenta_val for p in patrones):
                valor = to_float(row[col_saldo])
                if valor != 0:
                    # Verificar que la cuenta no haya sido procesada antes
                    cuenta_key = cuenta_val.lower()
                    if cuenta_key not in cuentas_procesadas:
                        cuentas_patron.append((cuenta_val, valor))
                        total += valor
                        cuentas_procesadas.add(cuenta_key)
                        coincidencias.append((cuenta_val, valor, "patron"))
                    else:
                        print(f"  [DUPLICADA] {cuenta_val}: {valor:,.2f} (ya procesada)")
        
        if cuentas_patron:
            print("  Se encontraron las siguientes cuentas por patrón:")
            for cuenta, valor in cuentas_patron:
                print(f"  - {cuenta}: {valor:,.2f}")
    
    # Mostrar resumen final
    print("\n  === RESUMEN DEL CÁLCULO ===")
    print("  ------------------------")
    if coincidencias:
        print(f"  Total de cuentas encontradas: {len(coincidencias)}")
        print(f"  Suma total calculada: {total:,.2f}")
    else:
        print("  [ADVERTENCIA] No se encontraron cuentas para calcular el balance")
    print("  ============================\n")
    print(f"TOTAL BALANCE: {total:,.2f}")
    return total


def procesar_situacion(df: pd.DataFrame) -> float:
    """
    SITUACIÓN: Extrae TOTAL 01010 de columna SALDOS MES
    Retorna el valor en la unidad original (no dividido)
    """
    print("\n=== PROCESANDO SITUACIÓN ===")
    
    # Buscar columna SALDOS MES (tolerante a variaciones)
    col_saldo = None
    for col in df.columns:
        col_norm = normalize_text(col)
        if 'saldo' in col_norm and ('mes' in col_norm or 'saldos mes' in col_norm or 'saldo mes' in col_norm):
            col_saldo = col
            break
    if col_saldo is None and len(df.columns) >= 9:
        col_saldo = df.columns[8]
    
    # Buscar fila TOTAL 01010 en cualquier columna de texto
    target_norm = normalize_text('TOTAL 01010')
    for idx, row in df.iterrows():
        hay_total = False
        for val in row.values:
            if isinstance(val, str) and target_norm in normalize_text(val):
                hay_total = True
                break
        if hay_total:
            if col_saldo in df.columns:
                valor = to_float(row[col_saldo])
            else:
                # Fallback: tomar el último valor numérico de la fila
                num_vals = [to_float(v) for v in row.values if pd.notna(v)]
                valor = num_vals[-1] if num_vals else 0.0
            print(f"TOTAL SITUACIÓN (SALDOS MES): {valor:,.2f}")
            return valor
    
    print("No se encontró TOTAL 01010")
    return 0.0


def procesar_modelo_vencimiento(ruta: Path) -> Tuple[float, float, float, float]:
    """
    3. MODELO DEUDA - VENCIMIENTO: Extrae totales de vencimiento y provisión dinámicamente
    Retorna: (total_60_plus, total_30, total_vencido, provision)
    IMPORTANTE: Los valores retornados NO están divididos por 1000
    """
    print("\n=== PROCESANDO MODELO VENCIMIENTO ===")
    
    try:
        # Obtener todos los nombres de hojas disponibles
        try:
            with pd.ExcelFile(ruta) as xls:
                hojas_disponibles = xls.sheet_names
                print(f"  Hojas disponibles: {', '.join(hojas_disponibles)}")
                
                # Buscar la hoja VENCIMIENTOS específicamente
                hoja_vencimientos = None
                for hoja in hojas_disponibles:
                    if 'VENCIMIENTOS' in hoja.upper():
                        hoja_vencimientos = hoja
                        break
                
                # Si no se encuentra la hoja VENCIMIENTOS, usar la primera
                if not hoja_vencimientos:
                    hoja_vencimientos = hojas_disponibles[0]
                    print(f"  [WARN] No se encontró hoja 'VENCIMIENTOS', usando: {hoja_vencimientos}")
                else:
                    print(f"  Hoja VENCIMIENTOS encontrada: {hoja_vencimientos}")
                
                df = pd.read_excel(xls, sheet_name=hoja_vencimientos)
        except Exception as e:
            print(f"  [ERROR] Error al leer las hojas del archivo: {str(e)}")
            return 0.0, 0.0, 0.0, 0.0
        
        if df.empty:
            print("  [WARN] Hoja vacía")
            return 0.0, 0.0, 0.0, 0.0
        
        print(f"  Filas en la hoja: {len(df)}")
        print(f"  Columnas: {list(df.columns)}")
        
        # Buscar filas que contienen totales (filas con '**')
        total_60_plus = 0.0
        total_30 = 0.0
        provision = 0.0
        
        # Buscar la fila de totales por moneda (que contiene '** **')
        for idx, row in df.iterrows():
            # Convertir fila a string para buscar patrones
            fila_str = ' '.join(str(v) for v in row.values if pd.notna(v))
            
            # Buscar fila con totales por moneda
            if '** **' in fila_str:
                print(f"  Fila de totales encontrada (índice {idx}): {fila_str}")
                
                # Buscar columnas numéricas en esta fila
                for col_name in df.columns:
                    if col_name in ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 
                                  'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 
                                  'DEUDA INCOBRABLE']:
                        try:
                            valor = to_float(row[col_name])
                            if col_name == 'VENCIDO 30':
                                total_30 = valor
                                print(f"    VENCIDO 30: {valor:,.2f}")
                            elif col_name in ['VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360']:
                                total_60_plus += valor
                                print(f"    {col_name}: {valor:,.2f} (acumulado: {total_60_plus:,.2f})")
                            elif col_name == 'DEUDA INCOBRABLE':
                                provision = valor
                                print(f"    DEUDA INCOBRABLE: {valor:,.2f}")
                        except Exception as e:
                            print(f"    [WARN] Error procesando columna {col_name}: {e}")
                
                # Una vez encontrada la fila de totales, podemos salir
                break
        
        # Si no encontramos los totales específicos, buscar la fila de totales generales
        if total_60_plus == 0 and total_30 == 0:
            for idx, row in df.iterrows():
                # Convertir fila a string para buscar patrones
                fila_str = ' '.join(str(v) for v in row.values if pd.notna(v))
                
                # Buscar fila con totales generales
                if '**TOTALES**' in fila_str.upper() or 'TOTAL GENERAL' in fila_str.upper():
                    print(f"  Fila de totales generales encontrada (índice {idx}): {fila_str}")
                    
                    # Buscar columnas numéricas en esta fila
                    for col_name in df.columns:
                        if col_name in ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 
                                      'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 
                                      'DEUDA INCOBRABLE']:
                            try:
                                valor = to_float(row[col_name])
                                if col_name == 'VENCIDO 30':
                                    total_30 = valor
                                    print(f"    VENCIDO 30: {valor:,.2f}")
                                elif col_name in ['VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360']:
                                    total_60_plus += valor
                                    print(f"    {col_name}: {valor:,.2f} (acumulado: {total_60_plus:,.2f})")
                                elif col_name == 'DEUDA INCOBRABLE':
                                    provision = valor
                                    print(f"    DEUDA INCOBRABLE: {valor:,.2f}")
                            except Exception as e:
                                print(f"    [WARN] Error procesando columna {col_name}: {e}")
                    
                    # Una vez encontrada la fila de totales generales, podemos salir
                    break
        
        total_vencido = total_30 + total_60_plus
        
        print(f"  Total 30 días: {total_30:,.2f}")
        print(f"  Total 60+ días: {total_60_plus:,.2f}")
        print(f"  Total vencido: {total_vencido:,.2f}")
        print(f"  DEUDA INCOBRABLE (Provisión): {provision:,.2f}")
        
        return total_60_plus, total_30, total_vencido, provision
        
    except Exception as e:
        print(f"  [ERROR] Error procesando modelo vencimiento: {e}")
        import traceback
        print(f"  [DEBUG] Detalles: {traceback.format_exc()}")
        return 0.0, 0.0, 0.0, 0.0


def aplicar_formato_numero(ws, celda: str):
    """Aplica formato de número y alineación correcta a una celda"""
    try:
        cell = ws[celda]
        if hasattr(cell, 'number_format'):
            cell.number_format = '#,##0.00'
        if hasattr(cell, 'alignment'):
            from openpyxl.styles import Alignment
            cell.alignment = Alignment(horizontal='right', vertical='center')
    except Exception as e:
        print(f"  [WARN] Error aplicando formato a {celda}: {str(e)}")

def limpiar_celda_antes_de_pegar(ws, celda: str, descripcion: str = ""):
    """Limpia una celda antes de pegar un nuevo valor (BORRAR EL RESULTADO ANTES DE PEGAR)"""
    try:
        cell = ws[celda]
        if cell.data_type != 'f':  # Solo limpiar si NO es una fórmula
            valor_anterior = cell.value
            cell.value = None
            print(f"  [CLEAN] {celda} limpiada antes de pegar {descripcion} (valor anterior: {valor_anterior})")
            return True
        else:
            print(f"  [LOCK] {celda}: PRESERVANDO fórmula existente: {cell.value}")
            return False
    except Exception as e:
        print(f"  [WARN] Error limpiando celda {celda}: {str(e)}")
        return False


def validar_datos_problematicos(ws):
    """
    Valida específicamente los datos problemáticos identificados en la imagen del spreadsheet.
    Retorna un reporte de los problemas encontrados.
    """
    print("\n=== VALIDACIÓN DE DATOS PROBLEMÁTICOS ===")
    
    problemas_encontrados = []
    
    # Lista de celdas específicas que pueden tener problemas según la imagen
    celdas_problematicas = [
        # Celdas con fechas incorrectas
        ("Q42", "Fecha TRM"),
        ("R42", "USD TRM"), 
        ("S42", "EUR TRM"),
        
        # Celdas con valores extremos
        ("Q15", "Total Balance"),
        ("Q16", "Saldos Mes"),
        ("Q17", "Vencido 60+"),
        ("Q19", "Vencido 30"),
        ("Q22", "Balance Final"),
        
        # Celdas con texto incorrecto
        ("Q3", "TRM USD"),
        ("Q4", "1/USD"),
        ("P3", "TRM EUR"),
        ("P4", "1/EUR"),
    ]
    
    for celda, descripcion in celdas_problematicas:
        try:
            cell = ws[celda]
            valor = str(cell.value) if cell.value is not None else ""
            
            # Verificar problemas específicos
            problemas_celda = []
            
            # 1. Fechas incorrectas
            if "1910" in valor or "24/07/1910" in valor:
                problemas_celda.append(f"Fecha incorrecta: {valor}")
            
            # 2. Texto con números
            if any(palabra in valor.lower() for palabra in ['focus', 'prov', 'negativa', 'positivo']):
                problemas_celda.append(f"Texto incorrecto: {valor}")
            
            # 3. Valores extremos
            try:
                if valor and valor.replace(',', '').replace('.', '').isdigit():
                    valor_num = float(valor.replace(',', '').replace('.', ''))
                    if valor_num > 1e12:
                        problemas_celda.append(f"Valor extremo: {valor}")
                    elif valor_num < 0 and abs(valor_num) > 1e6:
                        problemas_celda.append(f"Valor negativo extremo: {valor}")
            except (ValueError, TypeError):
                pass
            
            # 4. Instrucciones de texto
            if any(instruccion in valor for instruccion in ['BORRAR', 'cambiar', 'colocar', 'pasa a ser']):
                problemas_celda.append(f"Instrucción de texto: {valor}")
            
            if problemas_celda:
                problemas_encontrados.append({
                    'celda': celda,
                    'descripcion': descripcion,
                    'valor': valor,
                    'problemas': problemas_celda
                })
                
        except Exception as e:
            print(f"  [WARN] Error validando celda {celda}: {str(e)}")
    
    # Generar reporte
    if problemas_encontrados:
        print(f"\n[ERROR] PROBLEMAS ENCONTRADOS ({len(problemas_encontrados)} celdas):")
        for problema in problemas_encontrados:
            print(f"  {problema['celda']} ({problema['descripcion']}):")
            print(f"    Valor: '{problema['valor']}'")
            for p in problema['problemas']:
                print(f"    - {p}")
    else:
        print("\n[OK] No se encontraron datos problemáticos")
    
    return problemas_encontrados


def limpiar_datos_incorrectos(ws: Worksheet) -> int:
    """
    Limpia datos incorrectos o fuera de lugar que no deberían estar en el spreadsheet.
    Versión optimizada para mejor rendimiento.
    
    Args:
        ws: Worksheet de openpyxl a limpiar
        
    Returns:
        Número de celdas limpiadas
    """
    print("\n=== LIMPIANDO DATOS INCORRECTOS ===")
    
    # Pre-compilar expresiones regulares para mejor rendimiento
    patrones_incorrectos = [
        # Fechas incorrectas
        re.compile(r'24/07/1910', re.IGNORECASE),
        re.compile(r'1910'),
        
        # Texto con números que no debería estar
        re.compile(r'focus|prov|negativa|positivo', re.IGNORECASE),
        
        # Instrucciones de texto
        re.compile(r'BORRAR EL RESULTADO ANTES D|cambiar el mes|colocar la tasa|pasa a ser', re.IGNORECASE),
        
        # Números extremadamente pequeños o incorrectos
        re.compile(r'0,000222732'),
    ]
    
    celdas_limpiadas = 0
    errores_encontrados = 0
    start_time = time.time()
    
    # Limitar el rango a celdas usadas para mejorar rendimiento
    min_row, max_row = 1, ws.max_row if hasattr(ws, 'max_row') else 100
    min_col, max_col = 1, ws.max_column if hasattr(ws, 'max_column') else 26
    
    # Asegurar límites razonables
    max_row = min(max_row, 1000)  # No más de 1000 filas
    max_col = min(max_col, 50)    # No más de 50 columnas (A-AX)
    
    print(f"  Procesando celdas {min_row}:{max_row} x {min_col}:{max_col}...")
    
    # Procesar solo celdas con contenido
    for row in ws.iter_rows(min_row=min_row, max_row=max_row, 
                          min_col=min_col, max_col=max_col, 
                          values_only=False):
        for cell in row:
            try:
                # Saltar celdas vacías
                if cell.value is None:
                    continue
                    
                valor_actual = str(cell.value)
                valor_lower = valor_actual.lower()
                
                # Verificar patrones
                for patron in patrones_incorrectos:
                    if patron.search(valor_lower):
                        # Solo limpiar si NO es una fórmula
                        if cell.data_type != 'f':
                            print(f"  [CLEAN] {cell.coordinate}: Limpiando '{valor_actual}'")
                            cell.value = None
                            celdas_limpiadas += 1
                        else:
                            print(f"  [LOCK] {cell.coordinate}: Preservando fórmula: {cell.value}")
                            errores_encontrados += 1
                        break
                        
                # Verificar valores numéricos extremos
                if cell.data_type != 'f' and valor_actual:
                    try:
                        # Intentar convertir a número
                        valor_num = float(str(valor_actual).replace(',', '').replace('.', ''))
                        
                        # Detectar valores extremadamente grandes (posible error de formato)
                        if valor_num > 1e12:  # Mayor a 1 billón
                            print(f"  [CLEAN] {cell.coordinate}: Valor extremo detectado: '{valor_actual}'")
                            cell.value = None
                            celdas_limpiadas += 1
                            
                    except (ValueError, TypeError):
                        # No es un número, continuar
                        pass
                        
                # Verificar tiempo de ejecución
                if time.time() - start_time > 30:  # 30 segundos de tiempo máximo
                    print("  [WARN] Tiempo de ejecución excedido. Deteniendo limpieza...")
                    break
                    
            except Exception as e:
                print(f"  [WARN] Error procesando celda {cell.coordinate}: {str(e)}")
                errores_encontrados += 1
                continue
    
    # Limpieza de celdas específicas conocidas (excluyendo Q17 y Q19 que pueden contener referencias al PDF)
    celdas_especificas = ["Q42", "R42", "S42", "Q3", "Q4", "P3", "P4"]
    for coord in celdas_especificas:
        try:
            cell = ws[coord]
            if cell.value and cell.data_type != 'f':
                valor = str(cell.value).lower()
                if any(p in valor for p in ['focus', 'prov', 'negativa', 'positivo']):
                    print(f"  [CLEAN] {coord}: Limpiando texto incorrecto")
                    cell.value = None
                    celdas_limpiadas += 1
        except Exception as e:
            print(f"  [WARN] Error limpiando {coord}: {str(e)}")
            errores_encontrados += 1
    
    print(f"\n=== RESUMEN DE LIMPIEZA ===")
    print(f"  Celdas limpiadas: {celdas_limpiadas}")
    print(f"  Errores encontrados: {errores_encontrados}")
    print(f"  Tiempo de ejecución: {time.time() - start_time:.2f} segundos")
    
    return celdas_limpiadas

def verificar_y_corregir_totales(ws):
    """Verifica y corrige las discrepancias en los totales de las columnas"""
    print("  [CHECK] Verificando discrepancias en totales...")
    
    # Verificar H22 = H14+H15+H16+H17+H18
    h14 = to_float(ws["H14"].value) if ws["H14"].value is not None else 0
    h15 = to_float(ws["H15"].value) if ws["H15"].value is not None else 0
    h16 = to_float(ws["H16"].value) if ws["H16"].value is not None else 0
    h17 = to_float(ws["H17"].value) if ws["H17"].value is not None else 0
    h18 = to_float(ws["H18"].value) if ws["H18"].value is not None else 0
    h22_actual = to_float(ws["H22"].value) if ws["H22"].value is not None else 0
    h22_calculado = h14 + h15 + h16 + h17 + h18
    
    if abs(h22_actual - h22_calculado) > 0.01:
        print(f"  [FIX] H22 discrepancia: actual={h22_actual:,.2f}, calculado={h22_calculado:,.2f}")
        if ws["H22"].data_type != 'f':  # Solo corregir si no es fórmula
            ws["H22"].value = h22_calculado
            aplicar_formato_numero(ws, "H22")
            print(f"  [FIX] H22 corregido a: {h22_calculado:,.2f}")
    
    # Verificar D22 = D14+D15+D16+D17+D18
    d14 = to_float(ws["D14"].value) if ws["D14"].value is not None else 0
    d15 = to_float(ws["D15"].value) if ws["D15"].value is not None else 0
    d16 = to_float(ws["D16"].value) if ws["D16"].value is not None else 0
    d17 = to_float(ws["D17"].value) if ws["D17"].value is not None else 0
    d18 = to_float(ws["D18"].value) if ws["D18"].value is not None else 0
    d22_actual = to_float(ws["D22"].value) if ws["D22"].value is not None else 0
    d22_calculado = d14 + d15 + d16 + d17 + d18
    
    if abs(d22_actual - d22_calculado) > 0.01:
        print(f"  [FIX] D22 discrepancia: actual={d22_actual:,.2f}, calculado={d22_calculado:,.2f}")
        if ws["D22"].data_type != 'f':  # Solo corregir si no es fórmula
            ws["D22"].value = d22_calculado
            aplicar_formato_numero(ws, "D22")
            print(f"  [FIX] D22 corregido a: {d22_calculado:,.2f}")
    
    # Verificar F22 = F14+F15+F16+F17+F18
    f14 = to_float(ws["F14"].value) if ws["F14"].value is not None else 0
    f15 = to_float(ws["F15"].value) if ws["F15"].value is not None else 0
    f16 = to_float(ws["F16"].value) if ws["F16"].value is not None else 0
    f17 = to_float(ws["F17"].value) if ws["F17"].value is not None else 0
    f18 = to_float(ws["F18"].value) if ws["F18"].value is not None else 0
    f22_actual = to_float(ws["F22"].value) if ws["F22"].value is not None else 0
    f22_calculado = f14 + f15 + f16 + f17 + f18
    
    if abs(f22_actual - f22_calculado) > 0.01:
        print(f"  [FIX] F22 discrepancia: actual={f22_actual:,.2f}, calculado={f22_calculado:,.2f}")
        if ws["F22"].data_type != 'f':  # Solo corregir si no es fórmula
            ws["F22"].value = f22_calculado
            aplicar_formato_numero(ws, "F22")
            print(f"  [FIX] F22 corregido a: {f22_calculado:,.2f}")
    
    print("  [CHECK] Verificación de totales completada")

def validar_datos(ws, ws_values, total_balance: float, total_vencido: float) -> None:
    """Valida totales clave contra especificaciones del PDF."""
    try:
        h22 = to_float(ws_values["H22"].value) if ws_values else to_float(ws["H22"].value)
    except Exception:
        h22 = to_float(ws["H22"].value)
    try:
        q15 = to_float(ws_values["Q15"].value) if ws_values else to_float(ws["Q15"].value)
    except Exception:
        q15 = to_float(ws["Q15"].value)

    if abs(h22 - q15) > 0.01:
        print(f"[ERROR] ERROR: H22 ({h22:,.2f}) ≠ Q15 ({q15:,.2f})")

    try:
        d22 = to_float(ws_values["D22"].value) if ws_values else to_float(ws["D22"].value)
    except Exception:
        d22 = to_float(ws["D22"].value)
    if abs(d22 - (total_vencido / 1000.0)) > 0.01:
        print("[ERROR] ERROR: D22 no coincide con total vencido del modelo")

    try:
        f22 = to_float(ws_values["F22"].value) if ws_values else to_float(ws["F22"].value)
    except Exception:
        f22 = to_float(ws["F22"].value)
    total_no_vencido = (total_balance - total_vencido) / 1000.0
    if abs(f22 - total_no_vencido) > 0.01:
        print("[ERROR] ERROR: F22 no coincide con total no vencido del modelo")

def escribir_celda(ws, celda: str, valor, como_numero=True, preservar_formula=True, forzar_sobrescribir=False):
    """Escribe un valor en una celda con formato apropiado. NUNCA sobrescribe fórmulas existentes."""
    try:
        # Solo usar openpyxl (el código ya no usa xlwt para escribir)
        cell = ws[celda]
        # NUNCA sobrescribir si la celda tiene una fórmula
        if cell.data_type == 'f':
            print(f"  [LOCK] PRESERVANDO fórmula en {celda}: {cell.value}")
            return False
        # Solo escribir si NO es una fórmula
        if valor is None:
            cell.value = None
        elif como_numero:
            try:
                cell.value = float(valor)
                # Aplicar formato y alineación correcta
                aplicar_formato_numero(ws, celda)
            except (ValueError, TypeError):
                cell.value = valor
        else:
            cell.value = valor
        return True
    except Exception as e:
        print(f"  [WARN] Error al escribir en celda {celda}: {str(e)}")
        return False


def insertar_dato_entrada(ws, celda: str, valor, descripcion: str = "", forzar_sobrescribir: bool = False):
    """
    Inserta un dato de entrada (valor fijo) en celdas específicas.
    Estas son las celdas de datos de entrada que deben tener valores fijos.
    
    Args:
        ws: Worksheet
        celda: Referencia de celda (ej: "A1")
        valor: Valor a escribir
        descripcion: Descripción del dato para el log
        forzar_sobrescribir: Si True, sobrescribe incluso fórmulas existentes
    """
    try:
        cell = ws[celda]
        
        # Verificar si la celda ya tiene una fórmula
        if cell.data_type == 'f' and cell.value and not forzar_sobrescribir:
            print(f"  [LOCK] PRESERVANDO fórmula existente en {celda}: {cell.value}")
            print(f"  [DOC] {descripcion} - No se sobrescribe fórmula existente")
            return True
        
        # Verificar si la celda contiene una referencia al PDF que debería mantenerse
        if cell.value and not forzar_sobrescribir:
            valor_actual = str(cell.value)
            # Si la celda contiene una referencia al PDF, agregar el valor calculado como comentario
            if 'FORMATO DEUDA' in valor_actual.upper() or '@' in valor_actual:
                print(f"  [DOC] {descripcion} - Celda {celda} contiene referencia a PDF: {valor_actual}")
                # En lugar de sobrescribir, agregamos el valor calculado como parte del texto
                nuevo_valor = f"{valor_actual} (Valor calculado: {valor})"
                cell.value = nuevo_valor
                print(f"  [DOC] {descripcion} - Actualizado en {celda}: {nuevo_valor}")
                return True
        
        print(f"  [DOC] {descripcion} - Insertando valor fijo en {celda}: {valor}")
        
        if valor is None:
            cell.value = None
        else:
            try:
                # Intentar convertir a número solo si es posible
                if isinstance(valor, (int, float)) or (isinstance(valor, str) and valor.replace('.', '').replace('-', '').isdigit()):
                    cell.value = float(valor)
                    # Aplicar formato y alineación correcta
                    aplicar_formato_numero(ws, celda)
                else:
                    # Para texto (como meses), no aplicar formato numérico
                    cell.value = valor
            except (ValueError, TypeError):
                cell.value = valor
        
        print(f"  [DOC] {descripcion} - Celda {celda}: {valor}")
        return True
        
    except Exception as e:
        print(f"  [WARN] Error insertando {descripcion} en {celda}: {str(e)}")
        return False


def insertar_formula(ws, celda: str, formula: str, descripcion: str = ""):
    """
    Inserta una fórmula en una celda específica, pero solo si no tiene una fórmula existente.
    
    Args:
        ws: Worksheet
        celda: Referencia de celda (ej: "A1")
        formula: Fórmula a insertar (ej: "=Q16/1000")
        descripcion: Descripción de la fórmula para el log
    """
    try:
        cell = ws[celda]
        
        # Verificar si la celda ya tiene una fórmula
        if cell.data_type == 'f' and cell.value:
            print(f"  [LOCK] PRESERVANDO fórmula existente en {celda}: {cell.value}")
            print(f"  [FORMULA] {descripcion} - No se sobrescribe fórmula existente")
            return True
        
        print(f"  [FORMULA] {descripcion} - Insertando fórmula en {celda}: {formula}")
        
        # Insertar la fórmula solo si no hay una existente
        cell.value = formula
        
        print(f"  [FORMULA] {descripcion} - Fórmula insertada en {celda}")
        return True
        
    except Exception as e:
        print(f"  [WARN] Error insertando fórmula {descripcion} en {celda}: {str(e)}")
        return False


def procesar_acumulado(ruta: Path, ws: Worksheet) -> Optional[str]:
    """
    Procesa el archivo de formato acumulado y actualiza las celdas correspondientes
    
    Returns:
        str: Ruta al archivo procesado si tiene éxito
        None: Si no se pudo procesar o ocurre un error
    """
    print("\n=== PROCESANDO ACUMULADO ===")
    
    try:
        # Cargar archivo acumulado usando pandas para manejar ambos formatos
        print(f"  Procesando archivo: {ruta.name}")
        
        # Verificar que el archivo existe
        if not ruta.exists():
            print(f"  [ERROR] El archivo no existe: {ruta}")
            return None
            
        # Verificar tamaño del archivo (máximo 50MB)
        file_size = ruta.stat().st_size
        if file_size > 50 * 1024 * 1024:  # 50MB
            print(f"  [WARN] Archivo muy grande ({file_size/1024/1024:.1f}MB), limitando procesamiento")
        
        # Leer el archivo con pandas
        print("  Leyendo archivo con pandas...")
        df = pd.read_excel(ruta, engine=None, header=None, nrows=100)  # Limitar a 100 filas
        print(f"  Archivo leído exitosamente: {len(df)} filas, {len(df.columns)} columnas")
        
        # Procesar el archivo acumulado según las reglas especificadas
        # Copiar valores de B54 a F54 al archivo FOCUS
        
        # Leer el archivo con openpyxl para obtener acceso a celdas específicas
        try:
            wb_acum = load_workbook(ruta, data_only=True)
            ws_acum = wb_acum.active
            
            if ws_acum is not None:
                # Mapeo de celdas según las reglas especificadas
                mapping_acum = {
                    "B54": "D36",  # Cobros
                    "C54": "F36",
                    "B55": "D37",  # Facturación
                    "C55": "F37",
                    "B56": "D38",  # Vencidos
                    "C56": "F38",
                    "B60": "D45",  # Dotación
                    "B61": "D46"   # Desdotaciones
                }
                
                print("\n[INFO] Copiando valores del archivo acumulado al FOCUS:")
                for src, dst in mapping_acum.items():
                    try:
                        if ws_acum is not None:
                            cell_value = ws_acum[src].value if ws_acum[src].value is not None else 0
                            valor = to_float(cell_value)
                            escribir_celda(ws, dst, valor)
                            print(f"  [OK] {src} -> {dst}: {valor:,.2f}")
                    except Exception as e:
                        print(f"  [WARN] Error copiando {src} -> {dst}: {str(e)}")
                        # Usar valor 0 si hay error
                        escribir_celda(ws, dst, 0.0)
                
                # Limpiar celdas de verificación
                for celda in ["D52", "F52", "H52"]:
                    try:
                        cell = ws[celda]
                        if cell.data_type != 'f':  # Solo si NO es una fórmula
                            cell.value = 0
                            print(f"  {celda} = 0 (celda limpia)")
                    except Exception as e:
                        print(f"  [WARN] Error limpiando {celda}: {str(e)}")
                
                print("Valores del archivo acumulado copiados correctamente")
            else:
                print("  [ERROR] No se pudo obtener la hoja activa del archivo acumulado")
        except Exception as e:
            print(f"  [ERROR] Error al leer el archivo acumulado: {str(e)}")
        
    except FileNotFoundError as e:
        print(f"  [ERROR] El archivo no existe o no se puede encontrar: {str(e)}")
    except pd.errors.EmptyDataError as e:
        print(f"  [ERROR] El archivo está vacío o no contiene datos válidos: {str(e)}")
    except pd.errors.ParserError as e:
        print(f"  [ERROR] Error al parsear el archivo: {str(e)}")
    except Exception as e:
        print(f"  [ERROR] Error inesperado al procesar el archivo acumulado: {str(e)}")
        import traceback
        print(f"  [DEBUG] {traceback.format_exc()}")
    finally:
        print("  Procesamiento de archivo acumulado finalizado")
        
    # Procesar el DataFrame si se pudo cargar correctamente
    if 'df' in locals() and df is not None:
        # Crear diccionario para acceder a valores
        acum_data = {}
        max_cols = min(50, len(df.columns))  # Limitar a 50 columnas
        max_rows = min(100, len(df))  # Limitar a 100 filas
        
        print(f"  [INFO] Procesando {max_rows} filas x {max_cols} columnas...")
        
        # Crear diccionario para acceder a valores
        acum_data = {}
        max_cols = min(50, len(df.columns))  # Limitar a 50 columnas
        max_rows = min(100, len(df))  # Limitar a 100 filas
        
        print(f"  [INFO] Procesando {max_rows} filas x {max_cols} columnas...")
        
        # Función auxiliar para convertir índice de columna a letra de Excel
        def col_idx_to_letter(col_idx):
            """Convierte un índice de columna (0-based) a letra de Excel (A, B, ..., Z, AA, AB, ...)"""
            result = ""
            col_idx += 1  # Convertir a 1-based
            while col_idx > 0:
                col_idx -= 1
                result = chr(65 + (col_idx % 26)) + result
                col_idx //= 26
            return result
        
        # Verificar que la función de conversión funciona correctamente
        print(f"  Verificando conversión de columnas: A={col_idx_to_letter(0)}, Z={col_idx_to_letter(25)}, AA={col_idx_to_letter(26)}, AB={col_idx_to_letter(27)}")
        
        # Procesar celdas con límites estrictos
        for row_idx in range(max_rows):
            for col_idx in range(max_cols):
                try:
                    # Verificar límites antes de acceder
                    if row_idx >= len(df) or col_idx >= len(df.columns):
                        break
                    
                    cell_value = df.iat[row_idx, col_idx]
                    # Solo agregar celdas con valor
                    if pd.notna(cell_value) and cell_value != '':
                        # Verificar que la columna no exceda el límite de Excel (XFD = 16384)
                        if col_idx < 16384:
                            col_letter = col_idx_to_letter(col_idx)
                            cell_ref = f"{col_letter}{row_idx+1}"
                            acum_data[cell_ref] = cell_value
                except Exception as e:
                    print(f"  [WARN] Error procesando celda ({row_idx+1},{col_idx+1}): {str(e)}")
                    # Continuar con la siguiente celda en lugar de fallar completamente
                    continue
        
        # Esta sección ya no es necesaria ya que procesamos el archivo acumulado arriba
        pass
        
        # Esta sección ya no es necesaria ya que procesamos el archivo acumulado arriba
        mapping = {}
        
        # Ya procesamos el archivo acumulado arriba, así que esta sección ya no es necesaria
        pass
            
        # Asegurar que los directorios necesarios existan
        SALIDAS_DIR.mkdir(exist_ok=True, parents=True)
        BACKUP_DIR.mkdir(exist_ok=True, parents=True)
        
        # Buscar archivo FOCUS en el directorio actual
        patrones_focus = ["FOCUS_*.xls*", "FOCUS *.xls*", "FOCUS*.xls*"]
        archivo_focus = None
        
        # Buscar en el directorio actual
        for patron in patrones_focus:
            try:
                archivos = list(Path('.').glob(patron))
                if archivos:
                    archivo_focus = archivos[0]
                    break
            except Exception as e:
                print(f"  [WARN] Error buscando con patrón {patron}: {str(e)}")
        
        # Si no se encontró, buscar en el directorio padre
        if not archivo_focus:
            for patron in patrones_focus:
                try:
                    archivos = list(Path('..').glob(patron))
                    if archivos:
                        archivo_focus = archivos[0]
                        break
                except Exception as e:
                    print(f"  [WARN] Error buscando en directorio padre con patrón {patron}: {str(e)}")
        
        if not archivo_focus:
            print("  [ERROR] No se pudo encontrar el archivo FOCUS en el directorio actual ni en el directorio padre")
            print("  Directorio actual:", Path('.').resolve())
            print("  Archivos en el directorio actual:", [f.name for f in Path('.').iterdir() if f.is_file()])
            return None
        
        print(f"  Archivo FOCUS encontrado: {archivo_focus.absolute()}")
        
        # Cargar el archivo FOCUS
        try:
            wb = openpyxl.load_workbook(archivo_focus)
            active_sheet = wb.active
            if active_sheet is None:
                active_sheet = wb.worksheets[0] if wb.worksheets else None
                if active_sheet is None:
                    print("  [ERROR] No se pudo obtener ninguna hoja del archivo FOCUS")
                    return None
            
            # Type assertion: active_sheet should be a Worksheet at this point
            ws_local: Worksheet = active_sheet  # type: ignore
            
            # Guardar una copia de respaldo
            backup_path = BACKUP_DIR / f"{archivo_focus.stem}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{archivo_focus.suffix}"
            wb.save(backup_path)
            print(f"  Copia de respaldo guardada en: {backup_path}")
            
            # Guardar el archivo procesado
            output_path = SALIDAS_DIR / f"FOCUS_ACTUALIZADO_{datetime.now().strftime('%Y%m%d_%H%M%S')}.xlsx"
            wb.save(output_path)
            print(f"  Archivo procesado guardado en: {output_path}")
            
            return str(output_path)
            
        except Exception as e:
            print(f"  [ERROR] Error al procesar el archivo FOCUS: {str(e)}")
            import traceback
            print(f"  [DEBUG] {traceback.format_exc()}")
            return None

def procesar_y_actualizar_focus(
    archivo_focus: Optional[Union[str, Path]] = None,
    archivo_balance: Optional[Union[str, Path]] = None,
    archivo_situacion: Optional[Union[str, Path]] = None,
    archivo_modelo: Optional[Union[str, Path]] = None,
    archivo_acumulado: Optional[Union[str, Path]] = None,
    output_path: Optional[Union[str, Path]] = None
) -> str:
    """
    Función principal que procesa y actualiza el archivo FOCUS.
    
    Args:
        archivo_focus: Ruta al archivo FOCUS base (opcional, se busca automáticamente)
        archivo_balance: Ruta al archivo de balance (opcional)
        archivo_situacion: Ruta al archivo de situación (opcional)
        archivo_modelo: Ruta al archivo de modelo deuda (opcional)
        archivo_acumulado: Ruta al archivo acumulado (opcional)
        output_path: Ruta de salida personalizada (opcional)
    
    Returns:
        str: Ruta al archivo FOCUS actualizado generado
    """
    print("\n=== PROCESADOR FOCUS ===")
    print(f"Directorio base: {BASE_DIR}")
    print(f"Directorio de salidas: {SALIDAS_DIR}")
    
    # Crear directorios necesarios
    SALIDAS_DIR.mkdir(exist_ok=True, parents=True)
    BACKUP_DIR.mkdir(exist_ok=True, parents=True)
    
    # Buscar archivos si no se proporcionaron
    directorio_busqueda = Path('.')
    
    if not archivo_focus:
        print("\nBuscando archivo FOCUS...")
        archivo_focus_path = buscar_archivo(directorio_busqueda, ['FOCUS', 'focus'], ['.xlsx', '.xls'])
        if not archivo_focus_path:
            raise FileNotFoundError("No se encontró el archivo FOCUS")
        archivo_focus = archivo_focus_path
        print(f"  Encontrado: {archivo_focus.name}")
    else:
        archivo_focus = Path(archivo_focus)
    
    if not archivo_balance:
        print("\nBuscando archivo de Balance...")
        archivo_balance_path = buscar_archivo(directorio_busqueda, ['balance', 'Balance'], ['.xlsx', '.xls'])
        if archivo_balance_path:
            archivo_balance = archivo_balance_path
            print(f"  Encontrado: {archivo_balance.name}")
    else:
        archivo_balance = Path(archivo_balance)
    
    if not archivo_situacion:
        print("\nBuscando archivo de Situación...")
        archivo_situacion_path = buscar_archivo(directorio_busqueda, ['situacion', 'situación'], ['.xlsx', '.xls'])
        if archivo_situacion_path:
            archivo_situacion = archivo_situacion_path
            print(f"  Encontrado: {archivo_situacion.name}")
        else:
            print("  [WARN] No se encontró archivo de Situación")
    else:
        archivo_situacion = Path(archivo_situacion)
        print(f"  Archivo de Situación proporcionado: {archivo_situacion.name}")
    
    if not archivo_modelo:
        print("\nBuscando archivo de Modelo Deuda...")
        archivo_modelo_path = buscar_archivo(directorio_busqueda, ['modelo', 'deuda', 'MODELO_DEUDA'], ['.xlsx', '.xls'])
        if archivo_modelo_path:
            archivo_modelo = archivo_modelo_path
            print(f"  Encontrado: {archivo_modelo.name}")
    else:
        archivo_modelo = Path(archivo_modelo)
    
    # Cargar archivo FOCUS
    print(f"\nCargando archivo FOCUS: {archivo_focus}")
    try:
        wb = load_workbook(archivo_focus, data_only=False)
        
        # Buscar la hoja correcta (la primera o la que contiene los datos clave)
        ws: Optional[Worksheet] = None
        
        # Primero intentar encontrar la hoja "S22" específicamente
        for sheet_name in wb.sheetnames:
            if sheet_name.strip().upper() == "S22":
                ws = wb[sheet_name]
                print(f"  Hoja S22 encontrada: {sheet_name}")
                break
        
        # Si no se encuentra la hoja S22, buscar una hoja que contenga las celdas clave (H7, Q15, Q16, etc.)
        if ws is None:
            for sheet_name in wb.sheetnames:
                test_ws = wb[sheet_name]
                # Verificar si tiene las celdas clave del FOCUS
                # Solo procesar si es una Worksheet (no un Chartsheet)
                if not isinstance(test_ws, Worksheet):
                    continue
                try:
                    # Verificar si las celdas existen intentando acceder a ellas
                    _ = test_ws['H7']
                    _ = test_ws['Q15']
                    _ = test_ws['Q16']
                    ws = test_ws
                    print(f"  Hoja FOCUS encontrada: {sheet_name}")
                    break
                except Exception:
                    continue
        
        # Si no se encuentra, usar la primera hoja que sea Worksheet
        if ws is None:
            for sheet in wb.worksheets:
                if isinstance(sheet, Worksheet):
                    ws = sheet
                    break
            if ws is None:
                raise ValueError("No se encontró ninguna hoja de cálculo válida en el archivo FOCUS")
            print(f"  Usando primera hoja: {ws.title}")
        
        print(f"  Hoja activa: {ws.title}")
        print(f"  Total de hojas en el archivo: {len(wb.worksheets)}")
        print(f"  Nombres de hojas: {', '.join(wb.sheetnames)}")
    except Exception as e:
        raise Exception(f"Error al cargar archivo FOCUS: {e}")
    
    # Detectar mes actual del FOCUS
    mes_actual = detectar_mes_archivo(archivo_focus)
    mes_siguiente = obtener_mes_siguiente(mes_actual)
    print(f"\nMes actual detectado: {mes_actual}")
    print(f"Mes siguiente: {mes_siguiente}")
    
    # Actualizar mes en la celda H7
    insertar_dato_entrada(ws, 'H7', mes_siguiente, f"Mes siguiente ({mes_siguiente})", forzar_sobrescribir=True)
    
    # Procesar archivos de entrada
    total_balance = 0.0
    total_situacion = 0.0
    total_60_plus = 0.0
    total_30 = 0.0
    total_vencido = 0.0
    provision = 0.0
    
    if archivo_balance and archivo_balance.exists():
        print("\nProcesando Balance...")
        try:
            df_balance = leer_excel(archivo_balance)
            total_balance = procesar_balance(df_balance)
            print(f"  [BALANCE] Total Balance calculado: {total_balance:,.2f}")
            insertar_dato_entrada(ws, 'Q15', total_balance / 1000.0, f"Total Balance: {total_balance:,.2f} / 1000", forzar_sobrescribir=True)
            
            # Actualizar valor de Balance en el área de resumen (generalmente alrededor de columna N o superior)
            # TOTAL BALANCE / 1000 = 78.470.803,61 (según imagen)
        except Exception as e:
            print(f"  [ERROR] Error procesando archivo de Balance: {str(e)}")
            import traceback
            print(f"  [DEBUG] {traceback.format_exc()}")
    else:
        print("  [WARN] Archivo de Balance no encontrado o no existe")
    
    if archivo_situacion and archivo_situacion.exists():
        print("\nProcesando Situación...")
        try:
            df_situacion = leer_excel(archivo_situacion)
            total_situacion = procesar_situacion(df_situacion)
            print(f"  [SITUACION] Total calculado: {total_situacion:,.2f}")
            insertar_dato_entrada(ws, 'Q16', total_situacion / 1000.0, f"Total Situación: {total_situacion:,.2f} / 1000", forzar_sobrescribir=True)
            
            # Actualizar COBROS SITUACION (SALDO MES) / -1000
            # Valor ejemplo: 5.015.994,49 / 5.016
            cobros_valor = total_situacion
            cobros_miles = cobros_valor / 1000.0
            print(f"  [SITUACION] Cobros del mes: {cobros_valor:,.2f} / {cobros_miles:,.2f} (en miles)")
        except Exception as e:
            print(f"  [ERROR] Error procesando archivo de Situación: {str(e)}")
            import traceback
            print(f"  [DEBUG] {traceback.format_exc()}")
    else:
        print("  [WARN] Archivo de Situación no encontrado o no existe")
    
    if archivo_modelo and archivo_modelo.exists():
        print("\nProcesando Modelo Deuda...")
        try:
            total_60_plus, total_30, total_vencido, provision = procesar_modelo_vencimiento(archivo_modelo)
            print(f"  [MODELO] Total vencido 60+: {total_60_plus:,.2f}")
            print(f"  [MODELO] Total vencido 30: {total_30:,.2f}")
            print(f"  [MODELO] Total vencido combinado: {total_vencido:,.2f}")
            print(f"  [MODELO] Provisión (Deuda Incobrable): {provision:,.2f}")
            insertar_dato_entrada(ws, 'Q17', total_60_plus / 1000.0, f"Vencido 60+: {total_60_plus:,.2f} / 1000", forzar_sobrescribir=True)
            insertar_dato_entrada(ws, 'Q19', total_30 / 1000.0, f"Vencido 30: {total_30:,.2f} / 1000", forzar_sobrescribir=True)
            insertar_dato_entrada(ws, 'Q21', provision / 1000.0, f"Provisión: {provision:,.2f} / 1000", forzar_sobrescribir=True)
            
            # Actualizar Total vencido de 60 días en adelante /1000
            # Actualizar Total vencido de 60 días en adelante /1000
            # Valor ejemplo: 3.001.864,25 / -1.377.290,70
            
            # Calcular y actualizar F16 (Facturación del mes - No vencida)
            # Según el FOCUS, F16 debe ser la facturación del mes no vencida dividida entre 1000
            # La facturación del mes no vencida es el total del balance menos la deuda vencida
            if total_balance > 0:
                # Calcular la deuda no vencida (total balance - deuda vencida)
                deuda_no_vencida = total_balance - total_vencido if total_vencido > 0 else total_balance
                f16_valor = deuda_no_vencida / 1000.0
                
                # Mostrar información de depuración
                print(f"  [DEBUG] Cálculo F16: Total Balance={total_balance:,.2f}, Total Vencido={total_vencido:,.2f}, Deuda No Vencida={deuda_no_vencida:,.2f}, F16={f16_valor:,.2f}")
                
                # Limpiar celda F16 antes de escribir el nuevo valor
                limpiar_celda_antes_de_pegar(ws, 'F16', "Facturación del mes - No vencida")
                
                insertar_dato_entrada(ws, 'F16', f16_valor, f"Facturación del mes - No vencida: {f16_valor:,.2f}", forzar_sobrescribir=True)
                
                # También actualizar F15 (Deuda bruta NO Grupo - Final - No vencida)
                f15_valor = deuda_no_vencida / 1000.0
                limpiar_celda_antes_de_pegar(ws, 'F15', "Deuda bruta NO Grupo (Final - No vencida)")
                insertar_dato_entrada(ws, 'F15', f15_valor, f"Deuda bruta NO Grupo (Final - No vencida): {f15_valor:,.2f}", forzar_sobrescribir=True)
        except Exception as e:
            print(f"  [ERROR] Error procesando archivo de Modelo Deuda: {str(e)}")
            import traceback
            print(f"  [DEBUG] {traceback.format_exc()}")
    else:
        print("  [WARN] Archivo de Modelo Deuda no encontrado o no existe")
    
    # Calcular y actualizar Q22 (Deuda bruta NO Grupo - Final) basado en las reglas específicas
    # Según el FOCUS, Q22 debe ser el total del balance dividido entre 1000
    if total_balance > 0:
        q22_valor = total_balance / 1000.0
        
        # Mostrar información de depuración
        h22_valor = to_float(ws["H22"].value) if ws["H22"].value is not None else 0.0
        print(f"  [DEBUG] Cálculo Q22: Total Balance={total_balance:,.2f}, Q22={q22_valor:,.2f}, H22={h22_valor:,.2f}")
        
        # Limpiar celda Q22 antes de escribir el nuevo valor
        limpiar_celda_antes_de_pegar(ws, 'Q22', "Deuda bruta NO Grupo (Final)")
        
        insertar_dato_entrada(ws, 'Q22', q22_valor, f"Deuda bruta NO Grupo (Final): {q22_valor:,.2f}", forzar_sobrescribir=True)
    
    # Actualizar TRM
    actualizar_celdas_trm(ws, BASE_DIR, mes_actual)
    
    # Limpiar datos incorrectos
    limpiar_datos_incorrectos(ws)
    # Procesar archivo acumulado si existe
    if archivo_acumulado:
        procesar_acumulado(Path(archivo_acumulado), ws)
    else:
        archivo_acum = buscar_archivo(directorio_busqueda, ['acumulado', 'Acumulado'], ['.xlsx', '.xls'])
        if archivo_acum:
            print(f"\nArchivo acumulado encontrado: {archivo_acum.name}")
            procesar_acumulado(archivo_acum, ws)
    # Verificar y corregir totales
    verificar_y_corregir_totales(ws)
    # Marcar el workbook como generado por este script
    set_workbook_generation_marker(wb, ws)
    # Guardar archivo de salida
    if not output_path:
        timestamp = datetime.now().strftime('%Y%m%d_%H%M%S')
        output_path = SALIDAS_DIR / f"FOCUS_ACTUALIZADO_{timestamp}.xlsx"
    else:
        output_path = Path(output_path)
    # Crear backup del archivo original
    backup_path = BACKUP_DIR / f"{archivo_focus.stem}_backup_{datetime.now().strftime('%Y%m%d_%H%M%S')}{archivo_focus.suffix}"
    shutil.copy2(archivo_focus, backup_path)
    print(f"\nBackup creado: {backup_path}")
    # Guardar archivo procesado
    wb.save(output_path)
    print(f"\n[OK] Archivo FOCUS actualizado guardado: {output_path}")
    print(f"[OK] Tamaño: {output_path.stat().st_size / 1024:.2f} KB")
    return str(output_path)

if __name__ == "__main__":
    try:
        cli_args = sys.argv[1:]
        balance_arg = cli_args[0] if len(cli_args) > 0 else None
        situacion_arg = cli_args[1] if len(cli_args) > 1 else None
        focus_arg = cli_args[2] if len(cli_args) > 2 else None
        acumulado_arg = cli_args[3] if len(cli_args) > 3 else None
        modelo_arg = cli_args[4] if len(cli_args) > 4 else None

        if cli_args:
            print("[CLI] Ejecutando con argumentos proporcionados por PHP/frontend:")
            print(f"  balance: {balance_arg}")
            print(f"  situacion: {situacion_arg}")
            print(f"  focus: {focus_arg}")
            print(f"  acumulado: {acumulado_arg}")
            print(f"  modelo: {modelo_arg}")
            
        resultado = procesar_y_actualizar_focus(
            archivo_focus=focus_arg,
            archivo_balance=balance_arg,
            archivo_situacion=situacion_arg,
            archivo_modelo=modelo_arg,
            archivo_acumulado=acumulado_arg,
        )
        print(f"[OK] Archivo generado: {resultado}")
    except FileNotFoundError as e:
        logger.error(f"Error de archivo no encontrado: {e}")
        print(f"\n[ERROR] ERROR: {e}")
        print("  Por favor, asegúrese de que el archivo solicitado existe y tiene los permisos adecuados.")
        raise
    except pd.errors.EmptyDataError as e:
        logger.error(f"Error: El archivo está vacío o no contiene datos válidos: {e}")
        print(f"\n[ERROR] ERROR: El archivo está vacío o tiene un formato incorrecto: {e}")
        raise ValueError(f"Archivo vacío o inválido: {e}")
    except Exception as e:
        logger.error(f"Error inesperado al procesar archivos: {e}")
        logger.error(traceback.format_exc())
        print(f"\n[ERROR] ERROR INESPERADO: {e}")
        print("  Se ha generado un registro detallado del error en el archivo de log.")
        raise
