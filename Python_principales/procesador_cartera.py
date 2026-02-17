# -*- coding: utf-8 -*-
"""
Procesador de Cartera PROVCA 
Procesa archivos CSV de cartera cumpliendo con el formato establecido.

"""

import os
import sys
import pandas as pd
import logging
from datetime import datetime
import numpy as np

# ---------------------
# Configurar encoding para Windows
# ---------------------
if sys.platform == 'win32':
    try:
        sys.stdout.reconfigure(encoding='utf-8')
        sys.stderr.reconfigure(encoding='utf-8')
    except AttributeError:
        import codecs
        sys.stdout = codecs.getwriter('utf-8')(sys.stdout.buffer, 'strict')
        sys.stderr = codecs.getwriter('utf-8')(sys.stderr.buffer, 'strict')

# ---------------------
# Logging
# ---------------------
LOG_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'procesador_cartera.log')
logging.basicConfig(
    filename=LOG_FILE_PATH,
    level=logging.INFO,
    format='[%(asctime)s] [%(levelname)s] %(message)s',
    datefmt='%Y-%m-%d %H:%M:%S'
)
logger = logging.getLogger(__name__)

def info(msg):
    print(msg)
    logging.info(msg)

def error(msg):
    print(f"\nERROR: {msg}")
    logging.error(msg)

def warning(msg):
    print(f"\n⚠️ ADVERTENCIA: {msg}")
    logging.warning(msg)

# ---------------------
# Directorios
# ---------------------
BASE_DIR = os.path.dirname(os.path.abspath(__file__))
OUT_DIR = os.path.join(BASE_DIR, 'salidas')
os.makedirs(OUT_DIR, exist_ok=True)

# ---------------------
# MAPEO DE COLUMNAS SEGÚN PROCEDIMIENTO
# ---------------------
RENOMBRES = {
    'PCCDEM': 'EMPRESA',
    'PCCDAC': 'ACTIVIDAD',
    'PCDEAC': 'EMPRESA CODIGO AGENTE',
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
    'PCSALD': 'SALDO'
}

# ---------------------
# Funciones auxiliares
# ---------------------
def convertir_valor_to_float(x):
    """Convierte valores monetarios a float"""
    if pd.isna(x):
        return 0.0
    s = str(x).strip().replace("$", "").replace(" ", "").replace("\u200b", "")
    if s in ["", "-", "0"]:
        return 0.0
    if "," in s and "." in s:
        s = s.replace(".", "").replace(",", ".")
    elif "," in s:
        s = s.replace(",", ".")
    try:
        return float(s)
    except:
        digits = "".join(ch for ch in s if ch.isdigit() or ch == "." or ch == "-")
        return float(digits) if digits and digits != "-" else 0.0

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

# ---------------------
# Función principal
# ---------------------
def procesar_cartera(input_path, output_path=None, fecha_cierre_str=None):
    
    # Calcular fecha de cierre automática si no se proporciona
    if fecha_cierre_str is None:
        hoy = datetime.today()
        ultimo_dia_mes = pd.Period(hoy.strftime("%Y-%m")).end_time.date()
        fecha_cierre_str = str(ultimo_dia_mes)
    
    info("\n=== PROCESADOR DE CARTERA PROVCA ===")
    info(f"Fecha de cierre: {fecha_cierre_str}")

    if not os.path.exists(input_path):
        raise FileNotFoundError(f"No se encontró el archivo: {input_path}")

    # -------------------------
    # 1. LEER ARCHIVO
    # -------------------------
    encodings = ['latin1', 'utf-8-sig', 'cp1252']
    df = None

    for enc in encodings:
        try:
            df = pd.read_csv(input_path, sep=';', encoding=enc, dtype=str)
            if len(df) > 0:
                info(f"✓ Archivo leído con encoding {enc}")
                break
        except:
            continue

    if df is None:
        raise ValueError("No se pudo leer el archivo")

    info(f"✓ Total registros iniciales: {len(df)}")
    info(f"✓ Columnas originales: {list(df.columns)}")

    # -------------------------
    # 2. ELIMINAR COLUMNA PCIMCO (columna U)
    # -------------------------
    if 'PCIMCO' in df.columns:
        df.drop(columns=['PCIMCO'], inplace=True)
        info("✓ Columna PCIMCO eliminada")

    # -------------------------
    # 3. RENOMBRAR COLUMNAS según procedimiento
    # -------------------------
    df.rename(columns=RENOMBRES, inplace=True)
    info("✓ Columnas renombradas")

    # -------------------------
    # 4. ELIMINAR FILA PL30 (EMPRESA=PL y ACTIVIDAD=30)
    # -------------------------
    registros_antes = len(df)
    
    # Convertir a string y limpiar espacios para comparación exacta
    df["EMPRESA"] = df["EMPRESA"].astype(str).str.strip()
    df["ACTIVIDAD"] = df["ACTIVIDAD"].astype(str).str.strip()
    
    # Debug: ver cuántos registros PL existen
    registros_pl = len(df[df["EMPRESA"] == "PL"])
    if registros_pl > 0:
        info(f"ℹ️  Encontrados {registros_pl} registros de empresa 'PL'")
        # Ver actividades de PL
        actividades_pl = df[df["EMPRESA"] == "PL"]["ACTIVIDAD"].unique()
        info(f"ℹ️  Actividades de PL: {list(actividades_pl)}")
    
    # Contar registros PL30 antes de eliminar
    registros_pl30 = len(df[(df["EMPRESA"] == "PL") & (df["ACTIVIDAD"] == "30")])
    if registros_pl30 > 0:
        info(f"ℹ️  Encontrados {registros_pl30} registros de PL30 (EMPRESA='PL' y ACTIVIDAD='30')")
        # Mostrar saldos de estos registros
        saldos_pl30 = df[(df["EMPRESA"] == "PL") & (df["ACTIVIDAD"] == "30")]["SALDO"].tolist()
        info(f"ℹ️  Saldos PL30: {saldos_pl30}")
    
    # Eliminar registros donde EMPRESA='PL' Y ACTIVIDAD='30'
    df = df[~((df["EMPRESA"] == "PL") & (df["ACTIVIDAD"] == "30"))]
    registros_eliminados = registros_antes - len(df)
    
    if registros_eliminados > 0:
        info(f"✓ Eliminados {registros_eliminados} registros de PL30")
    else:
        warning("⚠️  No se encontraron registros PL30 para eliminar")
        warning("    Verificar que el CSV tenga registros con EMPRESA='PL' y ACTIVIDAD='30'")

    # -------------------------
    # 5. UNIFICAR NOMBRES EN DENOMINACION COMERCIAL
    # Copiar NOMBRE a DENOMINACION COMERCIAL cuando esté vacía
    # -------------------------
    mask_vacio = df["DENOMINACION COMERCIAL"].isna() | (df["DENOMINACION COMERCIAL"].str.strip() == "")
    df.loc[mask_vacio, "DENOMINACION COMERCIAL"] = df.loc[mask_vacio, "NOMBRE"]
    info("✓ Nombres unificados en DENOMINACION COMERCIAL")

    # -------------------------
    # 6. CONVERSIÓN MONETARIA
    # -------------------------
    df["VALOR"] = df["VALOR"].apply(convertir_valor_to_float)
    df["SALDO"] = df["SALDO"].apply(convertir_valor_to_float)
    info("✓ Valores monetarios convertidos")

    # -------------------------
    # 7. CONVERTIR FECHAS (mantener como datetime para cálculos)
    # -------------------------
    df["FECHA_TEMP"] = parse_fecha_segura(df["FECHA"])
    df["FECHA VTO_TEMP"] = parse_fecha_segura(df["FECHA VTO"])
    
    # Reportar fechas inválidas
    fechas_invalidas_fecha = df["FECHA_TEMP"].isna().sum()
    fechas_invalidas_vto = df["FECHA VTO_TEMP"].isna().sum()
    
    if fechas_invalidas_fecha > 0:
        warning(f"{fechas_invalidas_fecha} registros con FECHA inválida")
    if fechas_invalidas_vto > 0:
        warning(f"{fechas_invalidas_vto} registros con FECHA VTO inválida")
    
    info("✓ Fechas parseadas")

    # -------------------------
    # 8. CREAR TRES COLUMNAS PARA ABRIR FECHAS DE VENCIMIENTO
    # -------------------------
    df["DIA VTO"] = df["FECHA VTO_TEMP"].dt.day
    df["MES VTO"] = df["FECHA VTO_TEMP"].dt.month
    df["AÑO VTO"] = df["FECHA VTO_TEMP"].dt.year
    info("✓ Fechas de vencimiento separadas en día, mes y año")

    # -------------------------
    # 9. CALCULAR DÍAS VENCIDOS
    # -------------------------
    fecha_cierre = pd.to_datetime(fecha_cierre_str)
    df["DIAS VENCIDO"] = (fecha_cierre - df["FECHA VTO_TEMP"]).dt.days
    # Los días negativos significan que aún no vencen
    df["DIAS VENCIDO"] = df["DIAS VENCIDO"].apply(lambda x: x if x > 0 else 0)
    info("✓ Días vencidos calculados")

    # -------------------------
    # 10. CALCULAR DÍAS POR VENCER
    # -------------------------
    df["DIAS POR VENCER"] = (df["FECHA VTO_TEMP"] - fecha_cierre).dt.days
    # Los días negativos significan que ya vencieron
    df["DIAS POR VENCER"] = df["DIAS POR VENCER"].apply(lambda x: x if x > 0 else 0)
    info("✓ Días por vencer calculados")
     
    # -------------------------
    # 11. CALCULAR SALDO VENCIDO
    # -------------------------
    df["SALDO VENCIDO"] = df["SALDO"].where(df["DIAS VENCIDO"] > 0, 0)
    info("✓ Saldo vencido calculado")

     # -------------------------
     # 12. REORDENAR COLUMNAS
     # -------------------------
    columnas = list(df.columns)
     
    if "SALDO VENCIDO" in columnas:
         columnas.remove("VALOR")
         columnas.remove("SALDO")
     
         idx = columnas.index("SALDO VENCIDO")
     
         columnas.insert(idx, "SALDO")
         columnas.insert(idx, "VALOR")
     
    df = df[columnas]
     
    info("✓ Columnas reordenadas correctamente")
     
    # -------------------------
    # 13. CALCULAR % DOTACIÓN (100% si días vencidos >= 180)
    # -------------------------
    df["% DOTACION"] = df["DIAS VENCIDO"].apply(lambda x: 1.0 if x >= 180 else 0.0)
    
    # -------------------------
    # 14. CALCULAR VALOR DOTACIÓN (saldo si días >= 180)
    # -------------------------
    df["VALOR DOTACION"] = df["SALDO"].where(df["DIAS VENCIDO"] >= 180, 0)
    info("✓ % Dotación y Valor Dotación calculados")

    # -------------------------
    # 15. COLUMNAS DE LOS ÚLTIMOS 6 MESES VENCIDOS
    # -------------------------
    mes_cierre = fecha_cierre.month
    anio_cierre = fecha_cierre.year
    
    meses_atras = []
    for i in range(6, 0, -1):
        mes_calc = mes_cierre - i
        anio_calc = anio_cierre
        if mes_calc <= 0:
            mes_calc += 12
            anio_calc -= 1
        meses_atras.append((anio_calc, mes_calc))
    
    # Crear columnas para cada mes
    for i, (anio, mes) in enumerate(meses_atras, 1):
        nombre_col = f"VTO MES {i}"
        # Facturas vencidas en ese mes específico
        df[nombre_col] = df.apply(
            lambda row: row["SALDO"] if (
                pd.notna(row["FECHA VTO_TEMP"]) and 
                row["FECHA VTO_TEMP"].year == anio and 
                row["FECHA VTO_TEMP"].month == mes and
                row["DIAS VENCIDO"] > 0
            ) else 0,
            axis=1
        )
    
    info("✓ Columnas de últimos 6 meses vencidos creadas")

    # -------------------------
    # 16. VALOR >= 180 DÍAS VENCIDOS
    # -------------------------
    df["VALOR >= 180 DIAS"] = df["SALDO"].where(df["DIAS VENCIDO"] >= 180, 0)

    # -------------------------
    # 17. CALCULAR MORA TOTAL
    # -------------------------
    df["MORA TOTAL"] = df["SALDO VENCIDO"]

    # -------------------------
    # 18. COLUMNAS POR VENCER - PRÓXIMOS 3 MESES
    # -------------------------
    meses_adelante = []
    for i in range(1, 4):
        mes_calc = mes_cierre + i
        anio_calc = anio_cierre
        if mes_calc > 12:
            mes_calc -= 12
            anio_calc += 1
        meses_adelante.append((anio_calc, mes_calc))
    
    # Crear columnas para cada mes por vencer
    for i, (anio, mes) in enumerate(meses_adelante, 1):
        nombre_col = f"POR VENCER MES {i}"
        # Facturas que vencen en ese mes específico
        df[nombre_col] = df.apply(
            lambda row: row["SALDO"] if (
                pd.notna(row["FECHA VTO_TEMP"]) and 
                row["FECHA VTO_TEMP"].year == anio and 
                row["FECHA VTO_TEMP"].month == mes and
                row["DIAS POR VENCER"] > 0
            ) else 0,
            axis=1
        )
    
    info("✓ Columnas de próximos 3 meses por vencer creadas")

    # -------------------------
    # 19. CALCULAR VALOR MAYOR A 90 DÍAS POR VENCER
    # -------------------------
    df["MAYOR 90 DIAS POR VENCER"] = df["SALDO"].where(df["DIAS POR VENCER"] >= 90, 0)

    # -------------------------
    # 20. CALCULAR TOTAL POR VENCER
    # -------------------------
    df["TOTAL POR VENCER"] = df["SALDO"].where(df["DIAS POR VENCER"] > 0, 0)
    
    # VALIDACIÓN 1: Mora Total + Total Por Vencer = Saldo
    df["VALIDACION_MORA_VENCER"] = (
        (df["MORA TOTAL"] + df["TOTAL POR VENCER"]).round(2) == df["SALDO"].round(2)
    )
    info("✓ Validación Mora + Por Vencer realizada")

    # -------------------------
    # 21. RANGOS DE VENCIMIENTO (columnas principales del reporte)
    # -------------------------
    # Saldo no vencido: días vencidos de 0 a 29
    df["SALDO NO VENCIDO"] = df["SALDO"].where(df["DIAS VENCIDO"].between(0, 29), 0)
    
    # Vencido 30: días vencidos de 30 a 59
    df["VENCIDO 30"] = df["SALDO"].where(df["DIAS VENCIDO"].between(30, 59), 0)
    
    # Vencido 60: días vencidos de 60 a 89
    df["VENCIDO 60"] = df["SALDO"].where(df["DIAS VENCIDO"].between(60, 89), 0)
    
    # Vencido 90: días vencidos de 90 a 179
    df["VENCIDO 90"] = df["SALDO"].where(df["DIAS VENCIDO"].between(90, 179), 0)
    
    # Vencido 180: días vencidos de 180 a 359
    df["VENCIDO 180"] = df["SALDO"].where(df["DIAS VENCIDO"].between(180, 359), 0)
    
    # Vencido 360: días vencidos de 360 a 369
    df["VENCIDO 360"] = df["SALDO"].where(df["DIAS VENCIDO"].between(360, 369), 0)
    
    # Vencido +360: días vencidos de +370
    df["VENCIDO +360"] = df["SALDO"].where(df["DIAS VENCIDO"] >= 370, 0)
    
    info("✓ Rangos de vencimiento calculados")

    # -------------------------
    # VALIDACIÓN 2: Suma de rangos = Saldo
    # -------------------------
    df["SUMA_RANGOS"] = (
        df["SALDO NO VENCIDO"] +
        df["VENCIDO 30"] +
        df["VENCIDO 60"] +
        df["VENCIDO 90"] +
        df["VENCIDO 180"] +
        df["VENCIDO 360"] +
        df["VENCIDO +360"]
    )

    df["VALIDACION_RANGOS"] = (
        df["SUMA_RANGOS"].round(2) == df["SALDO"].round(2)
    )
    info("✓ Validación de rangos realizada")

    # -------------------------
    # 22. DEUDA INCOBRABLE = VALOR DOTACIÓN
    # -------------------------
    df["DEUDA INCOBRABLE"] = df["VALOR DOTACION"]

    # -------------------------
    # CONVERTIR FECHAS A STRING FORMATO DD/MM/YYYY PARA EXCEL
    # (Evitar problema de ##### en columnas)
    # -------------------------
    df["FECHA"] = df["FECHA_TEMP"].dt.strftime("%d/%m/%Y")
    df["FECHA VTO"] = df["FECHA VTO_TEMP"].dt.strftime("%d/%m/%Y")
    
    # Rellenar NaT con cadena vacía
    df["FECHA"] = df["FECHA"].fillna("")
    df["FECHA VTO"] = df["FECHA VTO"].fillna("")
    
    # Eliminar columnas temporales
    df.drop(columns=["FECHA_TEMP", "FECHA VTO_TEMP"], inplace=True)
    
    info("✓ Fechas formateadas como texto DD/MM/YYYY")

    # -------------------------
    # IDENTIFICAR REGISTROS CON PROBLEMAS
    # -------------------------
    registros_mora_vencer_invalidos = df[~df["VALIDACION_MORA_VENCER"]].copy()
    registros_rangos_invalidos = df[~df["VALIDACION_RANGOS"]].copy()
    
    if len(registros_mora_vencer_invalidos) > 0:
        warning(f"{len(registros_mora_vencer_invalidos)} registros NO cumplen: Mora + Por Vencer = Saldo")
    
    if len(registros_rangos_invalidos) > 0:
        warning(f"{len(registros_rangos_invalidos)} registros NO cumplen: Suma Rangos = Saldo")

    # -------------------------
    # RESUMEN GENERAL
    # -------------------------
    resumen = pd.DataFrame({
        "CONCEPTO": [
            "SALDO TOTAL",
            "SALDO NO VENCIDO",
            "MORA TOTAL",
            "TOTAL POR VENCER",
            "DEUDA INCOBRABLE",
            "VALOR DOTACION",
            "VALOR >= 180 DIAS",
            "MAYOR 90 DIAS POR VENCER",
            "VENCIDO 30",
            "VENCIDO 60",
            "VENCIDO 90",
            "VENCIDO 180",
            "VENCIDO 360",
            "VENCIDO +360"
        ],
        "VALOR": [
            df["SALDO"].sum(),
            df["SALDO NO VENCIDO"].sum(),
            df["MORA TOTAL"].sum(),
            df["TOTAL POR VENCER"].sum(),
            df["DEUDA INCOBRABLE"].sum(),
            df["VALOR DOTACION"].sum(),
            df["VALOR >= 180 DIAS"].sum(),
            df["MAYOR 90 DIAS POR VENCER"].sum(),
            df["VENCIDO 30"].sum(),
            df["VENCIDO 60"].sum(),
            df["VENCIDO 90"].sum(),
            df["VENCIDO 180"].sum(),
            df["VENCIDO 360"].sum(),
            df["VENCIDO +360"].sum()
        ]
    })

    # -------------------------
    # VALIDACIONES FINALES
    # -------------------------
    validaciones = pd.DataFrame({
        "VALIDACION": [
            "Registros con Mora + Vencer = Saldo",
            "Registros con Rangos = Saldo",
            "Total registros procesados",
            "% Validación Mora+Vencer",
            "% Validación Rangos",
            "Registros con FECHA inválida",
            "Registros con FECHA VTO inválida"
        ],
        "RESULTADO": [
            df["VALIDACION_MORA_VENCER"].sum(),
            df["VALIDACION_RANGOS"].sum(),
            len(df),
            f"{(df['VALIDACION_MORA_VENCER'].sum() / len(df) * 100):.2f}%",
            f"{(df['VALIDACION_RANGOS'].sum() / len(df) * 100):.2f}%",
            fechas_invalidas_fecha,
            fechas_invalidas_vto
        ]
    })

    info("\n=== VALIDACIONES ===")
    info(f"Registros válidos (Mora+Vencer): {df['VALIDACION_MORA_VENCER'].sum()}/{len(df)}")
    info(f"Registros válidos (Rangos): {df['VALIDACION_RANGOS'].sum()}/{len(df)}")
    info(f"Fechas FECHA inválidas: {fechas_invalidas_fecha}")
    info(f"Fechas VTO inválidas: {fechas_invalidas_vto}")

    # -------------------------
    # GENERAR NOMBRE DE SALIDA
    # -------------------------
    if output_path is None:
        fecha_str = datetime.now().strftime("%Y%m%d_%H%M%S")
        output_path = os.path.join(OUT_DIR, f"CARTERA_{fecha_str}.xlsx")

    # -------------------------
    # EXPORTAR EXCEL CON FORMATO MEJORADO
    # -------------------------
    info("\n=== GENERANDO ARCHIVO EXCEL ===")
    
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:

        # Hoja principal con detalle
        df.to_excel(writer, index=False, sheet_name="DETALLE")
        
        # Hoja de resumen
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN")
        
        # Hoja de validaciones
        validaciones.to_excel(writer, index=False, sheet_name="VALIDACIONES")
        
        # Hoja de registros con problemas en Mora+Vencer (si existen)
        if len(registros_mora_vencer_invalidos) > 0:
            registros_mora_vencer_invalidos.to_excel(writer, index=False, sheet_name="ERROR_MORA_VENCER")
        
        # Hoja de registros con problemas en Rangos (si existen)
        if len(registros_rangos_invalidos) > 0:
            registros_rangos_invalidos.to_excel(writer, index=False, sheet_name="ERROR_RANGOS")

        workbook  = writer.book
        worksheet_detalle = writer.sheets["DETALLE"]
        worksheet_resumen = writer.sheets["RESUMEN"]
        worksheet_validaciones = writer.sheets["VALIDACIONES"]
        
        # Formatos mejorados
        date_format = workbook.add_format({
            'num_format': 'dd/mm/yyyy',
            'align': 'center'
        })
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'align': 'right'
        })
        percent_format = workbook.add_format({
            'num_format': '0%',
            'align': 'center'
        })
        header_format = workbook.add_format({
            'bold': True,
            'bg_color': "#123269",
            'font_color': 'white',
            'align': 'center',
            'valign': 'vcenter',
            'border': 1,
            'text_wrap': True
        })
        text_format = workbook.add_format({
            'align': 'left'
        })
        
        # Aplicar formato a encabezados del detalle
        for col_num, value in enumerate(df.columns.values):
            worksheet_detalle.write(0, col_num, value, header_format)
            worksheet_detalle.set_row(0, 30)  # Altura de fila de encabezado
        
        # Identificar columnas por tipo
        fecha_cols = []
        valor_cols = []
        percent_cols = []
        integer_cols = []
        
        for i, col in enumerate(df.columns):
            if col in ["FECHA", "FECHA VTO"]:
                fecha_cols.append(i)
            elif col in ["DIA VTO", "MES VTO", "AÑO VTO", "DIAS VENCIDO", "DIAS POR VENCER"]:
                integer_cols.append(i)
            elif col == "% DOTACION":
                percent_cols.append(i)
            elif col in ["SALDO", "VALOR", "SALDO NO VENCIDO", "VENCIDO 30", "VENCIDO 60", 
                        "VENCIDO 90", "VENCIDO 180", "VENCIDO 360", "VENCIDO +360",
                        "DEUDA INCOBRABLE", "MORA TOTAL", "VALOR DOTACION",
                        "VTO MES 1", "VTO MES 2", "VTO MES 3", "VTO MES 4", "VTO MES 5", "VTO MES 6",
                        "SALDO VENCIDO", "VALOR >= 180 DIAS", 
                        "POR VENCER MES 1", "POR VENCER MES 2", "POR VENCER MES 3",
                        "MAYOR 90 DIAS POR VENCER", "TOTAL POR VENCER", "SUMA_RANGOS"]:
                valor_cols.append(i)
        
        # Autoajustar columnas con anchos mínimos garantizados
        for i, col in enumerate(df.columns):
            # Calcular ancho base
            if col in ["FECHA", "FECHA VTO"]:
                # Ancho mínimo garantizado para fechas (evitar ####)
                max_len = 15
            elif col in ["DIA VTO", "MES VTO", "AÑO VTO"]:
                max_len = 10
            elif col in ["DIAS VENCIDO", "DIAS POR VENCER"]:
                max_len = 16
            elif i in valor_cols:
                # Ancho mínimo garantizado para valores monetarios
                max_len = max(20, len(col) + 2)
            elif i in percent_cols:
                max_len = max(12, len(col) + 2)
            else:
                # Para textos, calcular dinámicamente
                try:
                    max_len_data = df[col].astype(str).map(len).max()
                    max_len = max(max_len_data, len(col)) + 2
                except:
                    max_len = len(col) + 2
            
            # Limitar ancho máximo para evitar columnas muy anchas
            max_len = min(max_len, 50)
            
            # Aplicar formato según tipo
            if i in fecha_cols:
                worksheet_detalle.set_column(i, i, max_len, text_format)
            elif i in percent_cols:
                worksheet_detalle.set_column(i, i, max_len, percent_format)
            elif i in valor_cols:
                worksheet_detalle.set_column(i, i, max_len, number_format)
            elif i in integer_cols:
                worksheet_detalle.set_column(i, i, max_len, workbook.add_format({'align': 'center'}))
            else:
                worksheet_detalle.set_column(i, i, max_len, text_format)
        
        # Formato para hoja resumen
        worksheet_resumen.set_column(0, 0, 35, text_format)
        worksheet_resumen.set_column(1, 1, 25, number_format)
        
        # Aplicar formato a encabezados del resumen
        for col_num, value in enumerate(resumen.columns.values):
            worksheet_resumen.write(0, col_num, value, header_format)
        worksheet_resumen.set_row(0, 30)
        
        # Formato para hoja validaciones
        worksheet_validaciones.set_column(0, 0, 40, text_format)
        worksheet_validaciones.set_column(1, 1, 25, text_format)
        
        # Aplicar formato a encabezados de validaciones
        for col_num, value in enumerate(validaciones.columns.values):
            worksheet_validaciones.write(0, col_num, value, header_format)
        worksheet_validaciones.set_row(0, 30)
        
        # Formatear hojas de errores si existen
        if len(registros_mora_vencer_invalidos) > 0:
            worksheet_error_mora = writer.sheets["ERROR_MORA_VENCER"]
            for col_num, value in enumerate(registros_mora_vencer_invalidos.columns.values):
                worksheet_error_mora.write(0, col_num, value, header_format)
            worksheet_error_mora.set_row(0, 30)
            
        if len(registros_rangos_invalidos) > 0:
            worksheet_error_rangos = writer.sheets["ERROR_RANGOS"]
            for col_num, value in enumerate(registros_rangos_invalidos.columns.values):
                worksheet_error_rangos.write(0, col_num, value, header_format)
            worksheet_error_rangos.set_row(0, 30)

    info(f"\n✓ Archivo generado correctamente: {output_path}")
    info(f"✓ Total registros procesados: {len(df)}")
    info(f"✓ Saldo total: ${df['SALDO'].sum():,.2f}")
    info(f"✓ Mora total: ${df['MORA TOTAL'].sum():,.2f}")
    info(f"✓ Deuda incobrable: ${df['DEUDA INCOBRABLE'].sum():,.2f}")
    
    return output_path


# ---------------------
# Main
# ---------------------
def main():
    try:
        input_path = None
        output_path = None
        fecha_cierre = None

        if len(sys.argv) > 1:
            input_path = sys.argv[1]

        if len(sys.argv) > 2:
            arg2 = sys.argv[2]
            # Si el segundo argumento es una fecha (YYYY-MM-DD)
            try:
                pd.to_datetime(arg2, format="%Y-%m-%d")
                fecha_cierre = arg2
            except:
                output_path = arg2

        if len(sys.argv) > 3:
            fecha_cierre = sys.argv[3]

        if not input_path:
            raise ValueError("No se recibió archivo de entrada")

        resultado = procesar_cartera(input_path, output_path, fecha_cierre)

        info(f"\n{'='*60}")
        info(f"PROCESO COMPLETADO EXITOSAMENTE")
        info(f"{'='*60}")
        info(f"Archivo: {resultado}")

    except Exception as e:
        logging.error(f"Error: {str(e)}", exc_info=True)
        print(f"\n{'='*60}")
        print(f"ERROR EN EL PROCESO")
        print(f"{'='*60}")
        print(f"{str(e)}")
        sys.exit(1)

if __name__ == "__main__":
    main()