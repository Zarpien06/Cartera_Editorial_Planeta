# -*- coding: utf-8 -*-
"""
Procesador de Cartera PROVCA
Procesa el archivo provca.csv con fecha de cierre 30-11-2025
"""

import os
import sys
import pandas as pd
import logging
from datetime import datetime
import csv

# Configuración de logging unificado
try:
    from config_logging import logger, log_inicio_proceso, log_fin_proceso, log_error_proceso
    USE_UNIFIED_LOGGING = True
except ImportError:
    # Fallback al logging original si config_logging no está disponible
    LOG_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'procesador_cartera.log')
    logging.basicConfig(
        filename=LOG_FILE_PATH,
        level=logging.INFO,
        format='[%(asctime)s] [%(levelname)s] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    USE_UNIFIED_LOGGING = False
    logger = logging.getLogger(__name__)

# Configurar encoding para Windows
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# Directorios
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, 'salidas')
os.makedirs(OUT_DIR, exist_ok=True)

# Mapeo de columnas según especificación PISA
RENOMBRES = {
    'PCCDEM': 'EMPRESA',
    'PCCDAC': 'ACTIVIDAD',
    'PCDEAC': 'EMPRESA_2',
    'PCCDAG': 'CODIGO AGENTE',
    'PCNMAG': 'AGENTE',
    'PCCDCO': 'CODIGO COBRADOR',
    'PCNMCO': 'COBRADOR',
    'PCCDCL': 'CODIGO CLIENTE',
    'PCCDDN': 'IDENTIFICACION',
    'PCNMCL': 'NOMBRE',
    'PCNMCM': 'DENOMINACION COMERCIAL',
    'PCNMDO': 'DIRECCION',
    'PCTLF1': 'TELEFONO',
    'PCNMPO': 'CIUDAD',
    'PCNUFC': 'NUMERO FACTURA',
    'PCORPD': 'TIPO',
    'PCFEFA': 'FECHA',   
    'PCFEVE': 'FECHA VTO',
    'PCVAFA': 'VALOR',
    'PCSALD': 'SALDO',
    ' PCVAFA ': 'VALOR',
    ' PCSALD ': 'SALDO'
}

def convertir_valor_to_float(x):
    """Convierte valores monetarios a float"""
    if pd.isna(x):
        return 0.0
    s = str(x).strip()
    if s == "" or s in ["-", "0"]:
        return 0.0
    s = s.replace("$", "").replace(" ", "").replace("\u200b", "")
    
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    elif "." in s:
        pass  # Mantener el punto como separador decimal
    
    try:
        return float(s)
    except Exception:
        digits = "".join(ch for ch in s if ch.isdigit() or ch == ".")
        return float(digits) if digits else 0.0

def parse_fecha_segura(serie):
    """Parsea fechas con múltiples formatos"""
    def try_parse(val):
        if pd.isna(val) or str(val).strip() == "":
            return pd.NaT
        formatos = ["%Y%m%d", "%d/%m/%Y", "%Y-%m-%d", "%d-%m-%Y", "%Y/%m/%d"]
        for fmt in formatos:
            try:
                return datetime.strptime(str(val).strip(), fmt)
            except ValueError:
                continue
        return pd.to_datetime(val, errors="coerce", dayfirst=True)
    return serie.apply(try_parse)

def procesar_cartera(input_path, output_path=None, fecha_cierre_str="2025-11-30"):
    """
    Procesa el archivo provca.csv
    Fecha de cierre por defecto: 30 de noviembre 2025
    """
    if USE_UNIFIED_LOGGING:
        log_inicio_proceso("CARTERA", input_path)
    else:
        logging.info(f"Iniciando procesamiento: {input_path}")
        logging.info(f"Fecha de cierre: {fecha_cierre_str}")
    print(f"\n=== PROCESADOR DE CARTERA ===")
    print(f"Fecha de cierre: {fecha_cierre_str}")
    
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"No se encontró el archivo: {input_path}")
    
    # Leer archivo CSV con múltiples encodings
    encodings = ['utf-8-sig', 'latin1', 'cp1252', 'iso-8859-1']
    df = None
    
    for encoding in encodings:
        try:
            df = pd.read_csv(
                input_path,
                sep=';',
                encoding=encoding,
                dtype=str,
                keep_default_na=False,
                na_values=[""],
                quoting=csv.QUOTE_MINIMAL
            )
            if df is not None and len(df) > 0:
                if USE_UNIFIED_LOGGING:
                    logger.info(f"[CARTERA] Archivo leído con {encoding}: {len(df)} filas")
                else:
                    logging.info(f"Archivo leído con {encoding}: {len(df)} filas")
                print(f"✓ Archivo leído con {encoding}: {len(df)} registros")
                break
        except Exception as e:
            continue
    
    if df is None or len(df) == 0:
        raise ValueError("No se pudo leer el archivo CSV")
    
    # Renombrar columnas
    df.rename(columns=RENOMBRES, inplace=True)
    if USE_UNIFIED_LOGGING:
        logger.info("[CARTERA] Columnas renombradas")
    else:
        logging.info("Columnas renombradas")
    
    # Eliminar columna PCIMCO si existe
    if "PCIMCO" in df.columns:
        df.drop(columns=["PCIMCO"], inplace=True)
    
    # Eliminar fila empresa PL30
    if 'EMPRESA' in df.columns and 'ACTIVIDAD' in df.columns:
        df['EMPRESA'] = df['EMPRESA'].astype(str).str.strip()
        df['ACTIVIDAD'] = df['ACTIVIDAD'].astype(str).str.strip()
        
        filas_antes = len(df)
        mask = (df['EMPRESA'] == 'PL') & (df['ACTIVIDAD'] == '30')
        df = df[~mask].copy()
        df = df.reset_index(drop=True)
        
        filas_eliminadas = filas_antes - len(df)
        if filas_eliminadas > 0:
            print(f"✓ Eliminadas {filas_eliminadas} filas de PL30")
    
    # Unificar nombres
    if "NOMBRE" in df.columns and "DENOMINACION COMERCIAL" in df.columns:
        df["NOMBRE"] = df["NOMBRE"].astype(str).str.strip()
        df["DENOMINACION COMERCIAL"] = df["DENOMINACION COMERCIAL"].astype(str).str.strip()
        
        mask_vacia = (df["DENOMINACION COMERCIAL"] == "") | (df["DENOMINACION COMERCIAL"] == "nan")
        mask_nombre_ok = (df["NOMBRE"] != "") & (df["NOMBRE"] != "nan")
        
        df.loc[mask_vacia & mask_nombre_ok, "DENOMINACION COMERCIAL"] = df.loc[mask_vacia & mask_nombre_ok, "NOMBRE"]
        print(f"✓ Nombres unificados en DENOMINACION COMERCIAL")
    
    # Convertir valores monetarios
    for col in ["SALDO", "VALOR"]:
        if col in df.columns:
            df[col] = df[col].apply(convertir_valor_to_float)
    
    if "SALDO" in df.columns:
        print(f"✓ Saldo total: ${df['SALDO'].sum():,.0f}")
    else:
        print("⚠ No se encontró la columna SALDO en el archivo")
    
    # Procesar fechas
    for col in ["FECHA", "FECHA VTO"]:
        if col in df.columns:
            df[col] = parse_fecha_segura(df[col])
    
    # Fecha de cierre (último día del mes)
    try:
        fecha_cierre = pd.to_datetime(fecha_cierre_str)
        fecha_cierre = fecha_cierre + pd.offsets.MonthEnd(0)
        print(f"✓ Fecha de cierre procesada: {fecha_cierre.strftime('%Y-%m-%d')}")
    except:
        fecha_cierre = pd.Timestamp.now() + pd.offsets.MonthEnd(0)
        print(f"⚠ Usando fecha actual: {fecha_cierre.strftime('%Y-%m-%d')}")
    
    # Calcular días vencidos
    if "FECHA VTO" in df.columns:
        print(f"✓ Calculando días vencidos usando fecha de cierre: {fecha_cierre.strftime('%Y-%m-%d')}")
        print(f"✓ Primeras 5 fechas de vencimiento: {df['FECHA VTO'].head().tolist()}")
        df["DIAS VENCIDO"] = (fecha_cierre - df["FECHA VTO"]).dt.days
        df["DIAS VENCIDO"] = df["DIAS VENCIDO"].clip(lower=0).fillna(0).astype(int)
        print(f"✓ Primeros 5 días vencidos calculados: {df['DIAS VENCIDO'].head().tolist()}")
        print(f"✓ Estadísticas de días vencidos: min={df['DIAS VENCIDO'].min()}, max={df['DIAS VENCIDO'].max()}, mean={df['DIAS VENCIDO'].mean():.2f}")
    else:
        print("⚠ No se encontró la columna FECHA VTO en el archivo")
    
    # Calcular columnas de vencimiento
    if "SALDO" in df.columns:
        print(f"✓ Calculando columnas de vencimiento para {len(df)} registros")
        print(f"✓ Saldo total antes de vencimientos: ${df['SALDO'].sum():,.0f}")
        print(f"✓ Primeros 5 saldos: {df['SALDO'].head().tolist()}")
        
        df["SALDO NO VENCIDO"] = df.apply(lambda r: r["SALDO"] if 0 <= r["DIAS VENCIDO"] <= 29 else 0.0, axis=1)
        df["VENCIDO 30"] = df.apply(lambda r: r["SALDO"] if 30 <= r["DIAS VENCIDO"] <= 50 else 0.0, axis=1)  # Corrección: 30-50 días vencidos
        df["VENCIDO 60"] = df.apply(lambda r: r["SALDO"] if 60 <= r["DIAS VENCIDO"] <= 89 else 0.0, axis=1)  # Corrección: 60-89 días vencidos
        df["VENCIDO 90"] = df.apply(lambda r: r["SALDO"] if 90 <= r["DIAS VENCIDO"] <= 179 else 0.0, axis=1)
        df["VENCIDO 180"] = df.apply(lambda r: r["SALDO"] if 180 <= r["DIAS VENCIDO"] <= 359 else 0.0, axis=1)
        df["VENCIDO 360"] = df.apply(lambda r: r["SALDO"] if 360 <= r["DIAS VENCIDO"] <= 369 else 0.0, axis=1)
        df["VENCIDO + 360"] = df.apply(lambda r: r["SALDO"] if r["DIAS VENCIDO"] >= 370 else 0.0, axis=1)
        # Deuda incobrable debería ser >= 370 días
        df["DEUDA INCOBRABLE"] = df.apply(lambda r: r["SALDO"] if r["DIAS VENCIDO"] >= 370 else 0.0, axis=1)
        
        # Mostrar totales parciales
        print(f"✓ Saldo no vencido: ${df['SALDO NO VENCIDO'].sum():,.0f}")
        print(f"✓ Vencido 30: ${df['VENCIDO 30'].sum():,.0f}")
        print(f"✓ Vencido 60: ${df['VENCIDO 60'].sum():,.0f}")
        print(f"✓ Vencido 90: ${df['VENCIDO 90'].sum():,.0f}")
        print(f"✓ Vencido 180: ${df['VENCIDO 180'].sum():,.0f}")
        print(f"✓ Vencido 360: ${df['VENCIDO 360'].sum():,.0f}")
        print(f"✓ Vencido + 360: ${df['VENCIDO + 360'].sum():,.0f}")
        print(f"✓ Deuda incobrable: ${df['DEUDA INCOBRABLE'].sum():,.0f}")
    else:
        # Si no hay columna SALDO, inicializar columnas con ceros
        df["SALDO NO VENCIDO"] = 0.0
        df["VENCIDO 30"] = 0.0
        df["VENCIDO 60"] = 0.0
        df["VENCIDO 90"] = 0.0
        df["VENCIDO 180"] = 0.0
        df["VENCIDO 360"] = 0.0
        df["VENCIDO + 360"] = 0.0
        df["DEUDA INCOBRABLE"] = 0.0
        
    # Asegurar que todas las columnas necesarias existan
    columnas_requeridas = [
        'EMPRESA', 'ACTIVIDAD', 'EMPRESA_2', 'CODIGO AGENTE', 'AGENTE', 
        'CODIGO COBRADOR', 'COBRADOR', 'CODIGO CLIENTE', 'IDENTIFICACION', 
        'NOMBRE', 'DENOMINACION COMERCIAL', 'DIRECCION', 'TELEFONO', 'CIUDAD', 
        'NUMERO FACTURA', 'TIPO', 'FECHA', 'FECHA VTO', 'VALOR', 'SALDO', 
        'DIAS VENCIDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 
        'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 
        'DEUDA INCOBRABLE'
    ]
    
    for col in columnas_requeridas:
        if col not in df.columns:
            df[col] = '' if col in ['EMPRESA', 'ACTIVIDAD', 'EMPRESA_2', 'CODIGO AGENTE', 'AGENTE', 
                                   'CODIGO COBRADOR', 'COBRADOR', 'CODIGO CLIENTE', 'IDENTIFICACION', 
                                   'NOMBRE', 'DENOMINACION COMERCIAL', 'DIRECCION', 'CIUDAD', 
                                   'NUMERO FACTURA', 'TIPO', 'FECHA', 'FECHA VTO'] else 0.0
    
    # Resumen de vencimientos
    print(f"\n=== RESUMEN DE VENCIMIENTOS ===")
    if "SALDO NO VENCIDO" in df.columns:
        saldo_no_vencido = df['SALDO NO VENCIDO'].sum()
        vencido_30 = df['VENCIDO 30'].sum()
        vencido_60 = df['VENCIDO 60'].sum()
        vencido_90 = df['VENCIDO 90'].sum()
        vencido_180 = df['VENCIDO 180'].sum()
        vencido_360 = df['VENCIDO 360'].sum()
        vencido_mas_360 = df['VENCIDO + 360'].sum()
        deuda_incobrable = df['DEUDA INCOBRABLE'].sum()
        
        print(f"Saldo no vencido:  ${saldo_no_vencido:,.0f}")
        print(f"Vencido 30 días:   ${vencido_30:,.0f}")
        print(f"Vencido 60 días:   ${vencido_60:,.0f}")
        print(f"Vencido 90 días:   ${vencido_90:,.0f}")
        print(f"Vencido 180 días:  ${vencido_180:,.0f}")
        print(f"Vencido 360 días:  ${vencido_360:,.0f}")
        print(f"Vencido + 360:     ${vencido_mas_360:,.0f}")
        print(f"Deuda incobrable:  ${deuda_incobrable:,.0f}")
        
        # Verificar si todos los valores son cero
        total_vencimientos = saldo_no_vencido + vencido_30 + vencido_60 + vencido_90 + vencido_180 + vencido_360 + vencido_mas_360 + deuda_incobrable
        if total_vencimientos == 0:
            print("⚠ ADVERTENCIA: Todos los valores de vencimiento son cero. Verificar archivo de entrada.")
    else:
        print("No se calcularon vencimientos debido a la ausencia de la columna SALDO")
    
    # Formatear fechas como date
    if "FECHA" in df.columns:
        df["FECHA"] = df["FECHA"].dt.date
    if "FECHA VTO" in df.columns:
        df["FECHA VTO"] = df["FECHA VTO"].dt.date
    
    # Generar nombre de salida
    if not output_path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(OUT_DIR, f"CARTERA_{timestamp}.xlsx")
    elif os.path.isdir(output_path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_path, f"CARTERA_{timestamp}.xlsx")
    
    # Guardar en Excel
    try:
        with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:
            df.to_excel(writer, index=False, sheet_name="Cartera")
            ws = writer.sheets["Cartera"]
            book = writer.book
            
            # Formato de miles sin decimales
            fmt_miles = book.add_format({"num_format": "#,##0;-#,##0;\"-\";@"})
            
            columnas_montos = ["VALOR", "SALDO", "SALDO NO VENCIDO", "VENCIDO 30", 
                              "VENCIDO 60", "VENCIDO 90", "VENCIDO 180", "VENCIDO 360",
                              "VENCIDO + 360", "DEUDA INCOBRABLE"]
            
            for col in columnas_montos:
                if col in df.columns:
                    idx = df.columns.get_loc(col)
                    ws.set_column(idx, idx, 18, fmt_miles)
            
            # Formato de fecha
            fmt_fecha = book.add_format({"num_format": "dd-mm-yyyy"})
            for col in ["FECHA", "FECHA VTO"]:
                if col in df.columns:
                    idx = df.columns.get_loc(col)
                    ws.set_column(idx, idx, 15, fmt_fecha)
        
        print(f"\n✓ Archivo guardado: {output_path}")
        logging.info(f"Procesamiento completado: {output_path}")
        return output_path
    
    except Exception as e:
        logging.error(f"Error al guardar: {e}")
        raise

def main():
    """Punto de entrada principal"""
    try:
        if len(sys.argv) > 1:
            input_path = sys.argv[1]
            output_path = sys.argv[2] if len(sys.argv) > 2 else None
            fecha_cierre = sys.argv[3] if len(sys.argv) > 3 else "2025-11-30"
            
            resultado = procesar_cartera(input_path, output_path, fecha_cierre)
            if USE_UNIFIED_LOGGING:
                log_fin_proceso("CARTERA", resultado or "salida por defecto", 
                              {"registros": "desconocido"})
        else:
            print("=== Procesador de Cartera PROVCA ===")
            input_path = input("Ruta del archivo CSV: ").strip()
            output_path = input("Ruta de salida (Enter para default): ").strip() or None
            
            resultado = procesar_cartera(input_path, output_path, "2025-11-30")
            print(f"\n✓ Procesamiento completado")
            print(f"Archivo: {resultado}")
            if USE_UNIFIED_LOGGING:
                log_fin_proceso("CARTERA", resultado or "salida por defecto", 
                              {"registros": "desconocido"})
    
    except Exception as e:
        if USE_UNIFIED_LOGGING:
            log_error_proceso("CARTERA", str(e))
        else:
            logging.error(f"Error: {str(e)}", exc_info=True)
        print(f"\n✗ ERROR: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()