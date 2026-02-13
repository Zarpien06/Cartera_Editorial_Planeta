#!/usr/bin/env python3
# -*- coding: utf-8 -*-
"""
Procesador de Anticipos PROVCA - PISA
Transforma archivos CSV de anticipos con la misma estructura que cartera
para consolidación en Modelo Deuda
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

# Mapeo de columnas según especificación PISA para anticipos
RENOMBRES = {
         "NCCDEM": "EMPRESA",
         "NCCDAC": "ACTIVIDAD",
         "NCCDCL": "CODIGO CLIENTE",
         "WWNIT": "NRO DOCUMENTO",
         "WWNMCL": "NOMBRE COMERCIAL",
         "WWNMDO": "DIRECCION",
         "WWTLF1": "TELEFONO",
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
         '"WWNIT"': "NRO DOCUMENTO",
         '"WWNMCL"': "NOMBRE COMERCIAL",
         '"WWNMDO"': "DIRECCION",
         '"WWTLF1"': "TELEFONO",
         '"CCCDFB"': "CODIGO AGENTE",
         '"BDNMNM"': "NOMBRE AGENTE",
         '"BDNMPA"': "APELLIDO AGENTE",
         '"NCMOMO"': "TIPO ANTICIPO",
         '"NCCDR3"': "NRO ANTICIPO",
         '"NCIMAN"': "VALOR ANTICIPO",
         '"NCFEGR"': "FECHA ANTICIPO"
}

def info(msg):
    print(msg)
    if USE_UNIFIED_LOGGING:
        logger.info(msg)
    else:
        logging.info(msg)

def error(msg):
    print(f"\nERROR: {msg}")
    if USE_UNIFIED_LOGGING:
        logger.error(msg)
    else:
        logging.error(msg)

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

def parse_fecha_segura(serie):
    """Parsea fechas manejando múltiples formatos"""
    serie = serie.astype(str).str.strip()
    fechas = pd.Series(pd.NaT, index=serie.index)

    # Detectar formato YYYYMMDD exacto (8 dígitos)
    mask_ymd = serie.str.match(r"^\d{8}$")

    # Parsear YYYYMMDD explícitamente
    if mask_ymd.any():
        fechas.loc[mask_ymd] = pd.to_datetime(
            serie.loc[mask_ymd],
            format="%Y%m%d",
            errors="coerce"
        )

    # Parsear resto como formato día primero (DD/MM/YYYY)
    if (~mask_ymd).any():
        fechas.loc[~mask_ymd] = pd.to_datetime(
            serie.loc[~mask_ymd],
            dayfirst=True,
            errors="coerce"
        )

    return fechas

def procesar_anticipos(input_path, output_path=None, fecha_cierre_str="2025-11-30"):
    """
    Procesa el archivo de anticipos según procedimiento.
    Los anticipos van en las columnas: SALDO, SALDO NO VENCIDO (por vencer)
    y deben tener la misma estructura que cartera para consolidación.
    """
    if USE_UNIFIED_LOGGING:
        log_inicio_proceso("ANTICIPOS", input_path)
    
    info("\n=== PROCESADOR DE ANTICIPOS ===")
    info(f"Fecha de cierre: {fecha_cierre_str}")
    
    if not os.path.exists(input_path):
        raise FileNotFoundError(f"No se encontró el archivo: {input_path}")
    
    # -------------------------
    # 1. LEER ARCHIVO CSV
    # -------------------------
    encodings = ['utf-8-sig', 'latin1', 'cp1252']
    df = None

    for enc in encodings:
        try:
            df = pd.read_csv(
                input_path,
                sep=";",
                encoding=enc,
                dtype=str,
                keep_default_na=False,
                na_values=[""]
            )
            if len(df) > 0:
                info(f"✓ Archivo leído con encoding {enc}")
                break
        except:
            continue

    if df is None:
        raise ValueError("No se pudo leer el archivo")

    info(f"✓ Total registros iniciales: {len(df)}")
    
    # -------------------------
    # 2. RENOMBRAR COLUMNAS
    # -------------------------
    df.rename(columns=RENOMBRES, inplace=True)
    info("✓ Columnas renombradas")
    
    # -------------------------
    # 3. LIMPIAR CARACTERES NO IMPRIMIBLES
    # -------------------------
    invalid_chars = re.compile(r"[\x00-\x08\x0B\x0C\x0E-\x1F]")
    for col in df.select_dtypes(include=['object']).columns:
        df[col] = df[col].apply(
            lambda x: invalid_chars.sub("", x) if isinstance(x, str) else x
        )
    
    # -------------------------
    # 4. CONVERTIR VALOR ANTICIPO Y MULTIPLICAR POR -1
    # -------------------------
    if "VALOR ANTICIPO" in df.columns:
        df["VALOR ANTICIPO"] = df["VALOR ANTICIPO"].apply(convertir_valor_to_float)
        df["VALOR ANTICIPO"] = df["VALOR ANTICIPO"] * -1
        total_anticipos = df["VALOR ANTICIPO"].sum()
        info(f"✓ Total anticipos: ${abs(total_anticipos):,.2f} (multiplicado por -1)")
    else:
        df["VALOR ANTICIPO"] = 0.0
    
    # -------------------------
    # 5. MAPEAR A ESTRUCTURA DE CARTERA
    # Según procedimiento: "estos datos van en las columnas de 
    # SALDOS, SALDO POR VENCER y SALDO NO VENCIDO"
    # -------------------------
    
    # Crear columnas compatibles con estructura de cartera
    df["SALDO"] = df["VALOR ANTICIPO"]
    df["NO VENCIDO"] = df["VALOR ANTICIPO"]  # Los anticipos son saldos por vencer
    df["TOTAL POR VENCER"] = df["VALOR ANTICIPO"]
    
    # Columnas de vencimiento en cero (anticipos no tienen vencimiento)
    df["VENCIDO 30"] = 0.0
    df["VENCIDO 60"] = 0.0
    df["VENCIDO 90"] = 0.0
    df["VENCIDO 180"] = 0.0
    df["VENCIDO 360"] = 0.0
    df["VENCIDO +360"] = 0.0
    df["DEUDA INCOBRABLE"] = 0.0
    df["MORA TOTAL"] = 0.0
    df["SALDO VENCIDO TOTAL"] = 0.0
    df["DIAS VENCIDO"] = 0
    df["DIAS POR VENCER"] = 0
    df["% DOTACION"] = 0
    df["VALOR DOTACION"] = 0.0
    
    # -------------------------
    # 6. MAPEAR CAMPOS DE ANTICIPOS A CARTERA
    # -------------------------
    # Crear mapeo de campos
    df["IDENTIFICACION"] = df.get("NIT/CEDULA", "")
    df["NOMBRE"] = df.get("NOMBRE COMERCIAL", "")
    df["DENOMINACION COMERCIAL"] = df.get("NOMBRE COMERCIAL", "")
    df["DIRECCION"] = df.get("DIRECCION", "")
    df["TELEFONO"] = df.get("TELEFONO", "")
    df["CIUDAD"] = df.get("POBLACION", "")
    df["NUMERO FACTURA"] = df.get("NRO ANTICIPO", "")
    df["TIPO"] = df.get("TIPO ANTICIPO", "ANTICIPO")
    df["VALOR"] = df["VALOR ANTICIPO"]
    
    # Agente
    if "NOMBRE AGENTE" in df.columns and "APELLIDO AGENTE" in df.columns:
        df["AGENTE"] = (df["NOMBRE AGENTE"].astype(str).str.strip() + " " + 
                       df["APELLIDO AGENTE"].astype(str).str.strip()).str.strip()
    else:
        df["AGENTE"] = ""
    
    # -------------------------
    # 7. PROCESAR FECHAS
    # -------------------------
    if "FECHA ANTICIPO" in df.columns:
        df["FECHA"] = parse_fecha_segura(df["FECHA ANTICIPO"])
        df["FECHA VTO"] = df["FECHA"]  # Misma fecha
    else:
        df["FECHA"] = pd.NaT
        df["FECHA VTO"] = pd.NaT
    
    fecha_cierre = pd.to_datetime(fecha_cierre_str)
    
    # Separar fecha de vencimiento
    df["DIA VTO"] = df["FECHA VTO"].dt.day
    df["MES VTO"] = df["FECHA VTO"].dt.month
    df["AÑO VTO"] = df["FECHA VTO"].dt.year
    
    # -------------------------
    # 8. ASEGURAR TODAS LAS COLUMNAS DE CARTERA
    # -------------------------
    columnas_cartera = [
        "EMPRESA",
        "ACTIVIDAD",
        "CODIGO CLIENTE",
        "NRO DOCUMENTO",
        "DENOMINACION COMERCIAL",
        "DIRECCION",
        "TELEFONO",
        "CIUDAD",
        "CODIGO AGENTE",
        "NOMBRE AGENTE",
        "APELLIDO AGENTE",
        "TIPO ANTICIPO",
        "NRO ANTICIPO",
        "VALOR ANTICIPO",
        "FECHA ANTICIPO"
    ]
    
    for col in columnas_cartera:
        if col not in df.columns:
            if col in ["SALDO", "VALOR", "NO VENCIDO", "VENCIDO 30", "VENCIDO 60", 
                      "VENCIDO 90", "VENCIDO 180", "VENCIDO 360", "VENCIDO +360",
                      "DEUDA INCOBRABLE", "MORA TOTAL", "SALDO VENCIDO TOTAL",
                      "VALOR DOTACION", "TOTAL POR VENCER"]:
                df[col] = 0.0
            elif col in ["FECHA", "FECHA VTO"]:
                df[col] = pd.NaT
            elif col in ["DIAS VENCIDO", "DIAS POR VENCER", "% DOTACION"]:
                df[col] = 0
            else:
                df[col] = ""
    
    # -------------------------
    # 9. GENERAR NOMBRE DE SALIDA
    # -------------------------
    nombre_mes = fecha_cierre.strftime("%B").upper()
    anio = fecha_cierre.strftime("%Y")

    MESES_ES = {
        "JANUARY": "ENERO", "FEBRUARY": "FEBRERO", "MARCH": "MARZO",
        "APRIL": "ABRIL", "MAY": "MAYO", "JUNE": "JUNIO",
        "JULY": "JULIO", "AUGUST": "AGOSTO", "SEPTEMBER": "SEPTIEMBRE",
        "OCTOBER": "OCTUBRE", "NOVEMBER": "NOVIEMBRE", "DECEMBER": "DICIEMBRE"
    }
    nombre_mes = MESES_ES.get(nombre_mes, nombre_mes)

    if not output_path:
        output_path = os.path.join(
            OUT_DIR,
            f"ANTICIPOS_{nombre_mes}_{anio}.xlsx"
        )
    
    # -------------------------
    # 10. RESUMEN
    # -------------------------
    resumen = pd.DataFrame({
        "CONCEPTO": [
            "SALDO TOTAL ANTICIPOS",
            "NO VENCIDO (POR VENCER)",
            "TOTAL REGISTROS"
        ],
        "VALOR": [
            df["SALDO"].sum(),
            df["NO VENCIDO"].sum(),
            len(df)
        ]
    })
    
    # -------------------------
    # 11. EXPORTAR A EXCEL
    # -------------------------
    info("\n=== GENERANDO ARCHIVO EXCEL ===")
    
    with pd.ExcelWriter(
        output_path,
        engine="xlsxwriter",
        date_format="dd/mm/yyyy",
        datetime_format="dd/mm/yyyy"
    ) as writer:

        # Seleccionar solo columnas de cartera en orden
        df_salida = df[columnas_cartera]
        df_salida.to_excel(writer, index=False, sheet_name="ANTICIPOS")
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN")

        workbook = writer.book
        worksheet = writer.sheets["ANTICIPOS"]
        worksheet_resumen = writer.sheets["RESUMEN"]
        
        # Formatos
        date_format = workbook.add_format({'num_format': 'dd/mm/yyyy'})
        number_format = workbook.add_format({'num_format': '#,##0.00'})
        percent_format = workbook.add_format({'num_format': '0%'})
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': "#020066",
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1
        })
        
        # Aplicar formato a encabezados
        for col_num, value in enumerate(df_salida.columns.values):
            worksheet.write(0, col_num, value, header_format)
        
        # Identificar columnas por tipo
        fecha_cols = []
        valor_cols = []
        percent_cols = []
        
        for i, col in enumerate(df_salida.columns):
            if df_salida[col].dtype == 'datetime64[ns]':
                fecha_cols.append(i)
            elif col == "% DOTACION":
                percent_cols.append(i)
            elif col in ["SALDO", "VALOR", "NO VENCIDO", "VENCIDO 30", "VENCIDO 60", 
                        "VENCIDO 90", "VENCIDO 180", "VENCIDO 360", "VENCIDO +360",
                        "DEUDA INCOBRABLE", "MORA TOTAL", "SALDO VENCIDO TOTAL",
                        "VALOR DOTACION", "TOTAL POR VENCER", "VALOR ANTICIPO"]:
                valor_cols.append(i)
        
        # Autoajustar columnas
        for i, col in enumerate(df_salida.columns):
            if df_salida[col].dtype == 'datetime64[ns]':
                max_len = max(len(col), 12) + 2
            else:
                max_len = max(
                    df_salida[col].astype(str).map(len).max(),
                    len(col)
                ) + 2
            
            if i in fecha_cols:
                worksheet.set_column(i, i, max_len, date_format)
            elif i in percent_cols:
                worksheet.set_column(i, i, max_len, percent_format)
            elif i in valor_cols:
                worksheet.set_column(i, i, max_len, number_format)
            else:
                worksheet.set_column(i, i, max_len)
        
        # Formato para resumen
        worksheet_resumen.set_column(0, 0, 30)
        worksheet_resumen.set_column(1, 1, 20, number_format)
        
        for col_num, value in enumerate(resumen.columns.values):
            worksheet_resumen.write(0, col_num, value, header_format)

    info(f"\n✓ Archivo generado: {output_path}")
    info(f"✓ Total registros: {len(df)}")
    info(f"✓ Total anticipos: ${abs(df['SALDO'].sum()):,.2f}")
    
    if USE_UNIFIED_LOGGING:
        log_fin_proceso("ANTICIPOS", output_path, {"registros": len(df)})
    
    return output_path

def main():
    """Punto de entrada principal"""
    try:
        input_path = None
        output_path = None
        # Fecha cierre automática: último día del mes actual
        hoy = datetime.today()
        ultimo_dia_mes = pd.Period(hoy.strftime("%Y-%m")).end_time.date()
        fecha_cierre = str(ultimo_dia_mes)

        if len(sys.argv) > 1:
            input_path = sys.argv[1]

        if len(sys.argv) > 2:
            arg2 = sys.argv[2]
            try:
                pd.to_datetime(arg2, format="%Y-%m-%d")
                fecha_cierre = arg2
            except:
                output_path = arg2

        if len(sys.argv) > 3:
            fecha_cierre = sys.argv[3]

        if not input_path:
            raise ValueError("No se recibió archivo de entrada")

        resultado = procesar_anticipos(input_path, output_path, fecha_cierre)

        info(f"\n{'='*60}")
        info(f"PROCESO COMPLETADO EXITOSAMENTE")
        info(f"{'='*60}")
        info(f"Archivo: {resultado}")

    except Exception as e:
        if USE_UNIFIED_LOGGING:
            log_error_proceso("ANTICIPOS", str(e))
        else:
            logging.error(f"Error: {str(e)}", exc_info=True)
        print(f"\n{'='*60}")
        print(f"ERROR EN EL PROCESO")
        print(f"{'='*60}")
        print(f"{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()
