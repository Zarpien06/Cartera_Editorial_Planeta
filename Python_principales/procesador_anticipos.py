#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Procesador de Anticipos PROVCA - PISA
Transforma archivos CSV de anticipos en formato Excel
"""

import csv
import pandas as pd
import numpy as np
import os
import sys
import logging
import re
from datetime import datetime

# Configuración de logging unificado
try:
    from config_logging import logger, log_inicio_proceso, log_fin_proceso, log_error_proceso
    USE_UNIFIED_LOGGING = True
except ImportError:
    # Fallback al logging original si config_logging no está disponible
    LOG_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'procesador_anticipos.log')
    logging.basicConfig(
        filename=LOG_FILE_PATH,
        level=logging.INFO,
        format='[%(asctime)s] [%(levelname)s] [PROCESADOR_ANTICIPOS] %(message)s',
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
    "NCCDEM": "EMPRESA",
    "NCCDAC": "ACTIVIDAD",
    "NCCDCL": "CODIGO CLIENTE",
    "WWNIT": "NIT/CEDULA",
    "WWNMCL": "NOMBRE COMERCIAL",
    "WWNMDO": "DIRECCION",
    "WWTLF1": "TELEFONO",
    "WWNMPO": "POBLACION",
    "CCCDFB": "CODIGO AGENTE",
    "BDNMNM": "NOMBRE AGENTE",
    "BDNMPA": "APELLIDO AGENTE",
    "NCMOMO": "TIPO ANTICIPO",
    "NCCDR3": "NRO ANTICIPO",
    "NCIMAN": "VALOR ANTICIPO",
    "NCFEGR": "FECHA ANTICIPO",
    '"NCCDEM"': "EMPRESA",
    '"NCCDAC"': "ACTIVIDAD",
    '"NCCDCL"': "CODIGO CLIENTE",
    '"WWNIT"': "NIT/CEDULA",
    '"WWNMCL"': "NOMBRE COMERCIAL",
    '"WWNMDO"': "DIRECCION",
    '"WWTLF1"': "TELEFONO",
    '"WWNMPO"': "POBLACION",
    '"CCCDFB"': "CODIGO AGENTE",
    '"BDNMNM"': "NOMBRE AGENTE",
    '"BDNMPA"': "APELLIDO AGENTE",
    '"NCMOMO"': "TIPO ANTICIPO",
    '"NCCDR3"': "NRO ANTICIPO",
    '"NCIMAN"': "VALOR ANTICIPO",
    '"NCFEGR"': "FECHA ANTICIPO"
}

def convertir_valor_to_float(x):
    """Convierte valores de texto a float"""
    if pd.isna(x) or x is None:
        return 0.0
    s = str(x).strip()
    if s == "" or s in ["-", "0"]:
        return 0.0
    s = s.replace("$", "").replace(" ", "").replace("\u200b", "")
    
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    
    try:
        return float(s)
    except Exception:
        digits = "".join(ch for ch in s if ch.isdigit() or ch == "." or ch == "-")
        return float(digits) if digits else 0.0

def procesar_anticipos(input_path, output_path=None):
    """
    Procesa el archivo de anticipos según las especificaciones PISA
    """
    if USE_UNIFIED_LOGGING:
        log_inicio_proceso("ANTICIPOS", input_path)
    else:
        logging.info(f"Iniciando procesamiento de anticipos: {input_path}")
    print(f"\n=== PROCESADOR DE ANTICIPOS ===")
    
    abs_input_path = os.path.abspath(input_path)
    
    if not os.path.exists(abs_input_path):
        error_msg = f"No se encontró el archivo: {abs_input_path}"
        if USE_UNIFIED_LOGGING:
            log_error_proceso("ANTICIPOS", error_msg)
        else:
            logging.error(error_msg)
        raise FileNotFoundError(error_msg)
    
    # Leer archivo CSV
    try:
        df = pd.read_csv(
            abs_input_path,
            sep=";",
            encoding="latin1",
            dtype=str,
            keep_default_na=False,
            na_values=[""],
            quoting=csv.QUOTE_MINIMAL
        )
        if USE_UNIFIED_LOGGING:
            logger.info(f"[ANTICIPOS] Archivo CSV leído: {len(df)} filas")
        else:
            logging.info(f"Archivo CSV leído: {len(df)} filas")
        print(f"✓ Archivo leído: {len(df)} registros")
    except Exception as e:
        if USE_UNIFIED_LOGGING:
            log_error_proceso("ANTICIPOS", f"Error al leer CSV: {str(e)}")
        else:
            logging.error(f"Error al leer CSV: {str(e)}")
        raise
    
    # Renombrar columnas
    df.rename(columns=RENOMBRES, inplace=True)
    
    # Asegurar columnas mínimas
    columnas_minimas = ["EMPRESA", "ACTIVIDAD", "CODIGO CLIENTE", "NIT/CEDULA", 
                        "NOMBRE COMERCIAL", "VALOR ANTICIPO", "DENOMINACION COMERCIAL"]
    
    for col in columnas_minimas:
        if col not in df.columns:
            df[col] = ""
    
    # Si no hay DENOMINACION COMERCIAL, usar NOMBRE COMERCIAL
    if "NOMBRE COMERCIAL" in df.columns and "DENOMINACION COMERCIAL" in df.columns:
        df["DENOMINACION COMERCIAL"] = df["DENOMINACION COMERCIAL"].where(df["DENOMINACION COMERCIAL"] != "", df["NOMBRE COMERCIAL"])
    
    # Limpiar caracteres no imprimibles
    invalid_chars = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: invalid_chars.sub("", x) if isinstance(x, str) else x
        )
    
    # Convertir VALOR ANTICIPO a numérico y multiplicar por -1
    if "VALOR ANTICIPO" in df.columns:
        df["VALOR ANTICIPO"] = df["VALOR ANTICIPO"].apply(convertir_valor_to_float)
        df["VALOR ANTICIPO"] = df["VALOR ANTICIPO"] * -1
        df["VALOR ANTICIPO"] = pd.to_numeric(df["VALOR ANTICIPO"], errors='coerce').fillna(0)
        
        total_anticipos = df["VALOR ANTICIPO"].sum()
        print(f"✓ Total anticipos: ${abs(total_anticipos):,.0f} (multiplicado por -1)")
    
    # Crear columna SALDO (copia de VALOR ANTICIPO)
    df["SALDO"] = df["VALOR ANTICIPO"]
    
    # Formatear fechas
    if "FECHA ANTICIPO" in df.columns:
        df["FECHA ANTICIPO"] = df["FECHA ANTICIPO"].astype(str).str.strip()
        df["FECHA ANTICIPO"] = df["FECHA ANTICIPO"].replace(['', 'nan', 'None', 'NaT', 'NaN', '<NA>'], np.nan)
        
        def parse_fecha(fecha_str):
            if pd.isna(fecha_str) or fecha_str in ['', 'NaT', 'nan', 'None', 'NaN', '<NA>']:
                return np.nan
            
            fecha_str = str(fecha_str).strip()
            if not fecha_str or fecha_str.lower() in ['nan', 'none', 'nat', 'null', '', '<na>']:
                return np.nan
            
            # Lista ampliada de formatos de fecha
            formatos = [
                '%Y%m%d', '%Y-%m-%d', '%d-%m-%Y', '%Y/%m/%d', '%d/%m/%Y',
                '%Y-%m-%d %H:%M:%S', '%d/%m/%Y %H:%M:%S', '%Y%m%d%H%M%S'
            ]
            
            for fmt in formatos:
                try:
                    return pd.to_datetime(fecha_str, format=fmt, errors='raise')
                except:
                    continue
            
            # Intento final con inferencia automática
            try:
                return pd.to_datetime(fecha_str, infer_datetime_format=True, errors='coerce')
            except:
                return np.nan
        
        # Aplicar parseo de fecha con manejo de errores mejorado
        try:
            df["FECHA ANTICIPO"] = df["FECHA ANTICIPO"].apply(parse_fecha)
            df["FECHA ANTICIPO"] = df["FECHA ANTICIPO"].apply(
                lambda x: x.date() if pd.notna(x) and hasattr(x, 'date') else x
            )
        except Exception as e:
            print(f"[WARN] Error al parsear fechas: {e}")
            # En caso de error, dejar las fechas como están
            pass
    
    # Crear LINEA_VENTA combinando EMPRESA y ACTIVIDAD
    if 'EMPRESA' in df.columns and 'ACTIVIDAD' in df.columns:
        df['ACTIVIDAD'] = df['ACTIVIDAD'].astype(str).str.strip()
        df['LINEA_VENTA'] = df['EMPRESA'].astype(str).str.strip() + df['ACTIVIDAD'].str.zfill(2)
        print(f"✓ Línea de venta creada")
    else:
        df['LINEA_VENTA'] = ''
    
    # Eliminar LINEA DE NEGOCIO ya que no está en la lista de columnas requeridas
    if 'LINEA DE NEGOCIO' in df.columns:
        df = df.drop(columns=['LINEA DE NEGOCIO'])
    
    # Asignar moneda según la línea de venta
    df['MONEDA'] = 'PESOS COL'  # Por defecto es pesos colombianos
    
    # Líneas en dólares
    lineas_usd = ['PL11', 'PL18', 'PL57', 'PL72', 'PL17']
    # Líneas en euros
    lineas_eur = ['PL16', 'PL41', 'PL68']
    
    # Actualizar moneda según la LINEA_VENTA
    df.loc[df['LINEA_VENTA'].isin(lineas_usd), 'MONEDA'] = 'USD'
    df.loc[df['LINEA_VENTA'].isin(lineas_eur), 'MONEDA'] = 'EUR'
        
    # Agregar columnas requeridas para el archivo de anticipos
    df['SALDO'] = df['VALOR ANTICIPO']  # Los anticipos son valores negativos
    df['SALDO NO VENCIDO'] = df['VALOR ANTICIPO']  # Los anticipos son valores negativos
    df['DEUDA INCOBRABLE'] = 0.0
    
    # Eliminar columnas EMPRESA y ACTIVIDAD ya que su información está representada en LINEA_VENTA
    if 'EMPRESA' in df.columns:
        df = df.drop(columns=['EMPRESA'])
    if 'ACTIVIDAD' in df.columns:
        df = df.drop(columns=['ACTIVIDAD'])
    
    # Definir columnas de salida según especificación PISA para anticipos
    COLUMNAS_SALIDA = [
        'LINEA_VENTA', 
        'MONEDA', 
        'SALDO', 
        'SALDO NO VENCIDO', 
        'DEUDA INCOBRABLE',
        'CODIGO CLIENTE', 
        'NIT/CEDULA',
        'NOMBRE COMERCIAL',
        'DIRECCION', 
        'TELEFONO', 
        'POBLACION',
        'CODIGO AGENTE', 
        'NOMBRE AGENTE',
        'APELLIDO AGENTE', 
        'TIPO ANTICIPO',
        'NRO ANTICIPO', 
        'VALOR ANTICIPO',
        'FECHA ANTICIPO'
    ]
    
    # Asegurar que todas las columnas requeridas existan
    for col in COLUMNAS_SALIDA:
        if col not in df.columns:
            df[col] = ""
    
    # Seleccionar y ordenar columnas finales
    df = df[COLUMNAS_SALIDA]
    
    # Generar nombre de archivo de salida
    if not output_path:
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(OUT_DIR, f"ANTICIPOS_{timestamp}.xlsx")
    elif os.path.isdir(output_path):
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(output_path, f"ANTICIPOS_{timestamp}.xlsx")
    
    abs_output_path = os.path.abspath(output_path)
    
    # Asegurar directorio de salida
    output_dir = os.path.dirname(abs_output_path)
    if not os.path.exists(output_dir):
        os.makedirs(output_dir, exist_ok=True)
    
    # Guardar en Excel
    try:
        writer = pd.ExcelWriter(abs_output_path, engine="xlsxwriter")
        df.to_excel(writer, index=False, sheet_name="Anticipos")
        
        ws = writer.sheets["Anticipos"]
        book = writer.book
        
        # Formatos
        fmt_text = book.add_format({"num_format": "@"})
        fmt_numero = book.add_format({"num_format": "#,##0;-#,##0;\"-\";@"})
        fmt_fecha = book.add_format({"num_format": "dd-mm-yyyy"})
        
        # Aplicar formatos
        columnas_montos = ["VALOR ANTICIPO"]
        
        for idx, col_name in enumerate(df.columns):
            if col_name in columnas_montos:
                ws.set_column(idx, idx, 20, fmt_numero)
            elif col_name in ["FECHA ANTICIPO"]:
                ws.set_column(idx, idx, 15, fmt_fecha)
            else:
                ws.set_column(idx, idx, 20, fmt_text)
        
        writer.close()
        
        print(f"\n✓ Archivo guardado: {abs_output_path}")
        logging.info(f"Archivo generado: {abs_output_path}")
        return abs_output_path
    
    except Exception as e:
        logging.error(f"Error al guardar Excel: {str(e)}")
        raise

def main():
    """Punto de entrada principal"""
    try:
        if len(sys.argv) > 1:
            input_path = sys.argv[1]
            output_path = sys.argv[2] if len(sys.argv) > 2 else None
            resultado = procesar_anticipos(input_path, output_path)
            if USE_UNIFIED_LOGGING:
                log_fin_proceso("ANTICIPOS", resultado or "salida por defecto", 
                              {"registros": "desconocido"})
        else:
            print("=== Procesador de Anticipos ===")
            input_path = input("Ruta del archivo CSV: ").strip()
            output_path = input("Ruta de salida (Enter para default): ").strip() or None
            
            resultado = procesar_anticipos(input_path, output_path)
            print(f"\n✓ Procesamiento completado")
            if USE_UNIFIED_LOGGING:
                log_fin_proceso("ANTICIPOS", resultado or "salida por defecto", 
                              {"registros": "desconocido"})
    
    except Exception as e:
        if USE_UNIFIED_LOGGING:
            log_error_proceso("ANTICIPOS", str(e))
        else:
            logging.error(f"Error: {str(e)}", exc_info=True)
        print(f"\n✗ ERROR: {str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()