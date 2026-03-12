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
    # LIMPIAR ESPACIOS EN CAMPOS DE TEXTO
    # -------------------------
    columnas_texto = [
        "NOMBRE",
        "DENOMINACION COMERCIAL",
        "AGENTE",
        "COBRADOR",
        "CIUDAD",
        "DIRECCION"
    ]
    
    for col in columnas_texto:
        if col in df.columns:
            df[col] = (
                df[col]
                .astype(str)
                .str.strip()                       # Quitar espacios inicio/fin
                .str.replace(r'\s+', ' ', regex=True)  # Quitar dobles espacios internos
            )
    
    info("✓ Espacios limpiados en nombres y textos")
    
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
    # 7.1 FILTRAR REGISTROS MAYORES A FECHA DE CIERRE
    # -------------------------
    fecha_cierre = pd.to_datetime(fecha_cierre_str)
    info(f"Fecha de cierre REAL usada en cálculos: {fecha_cierre}")

    registros_antes_filtro = len(df)
    
    # Solo eliminar facturas emitidas después del cierre
    df = df[df["FECHA_TEMP"] <= fecha_cierre]
    
    registros_despues_filtro = len(df)
    registros_filtrados = registros_antes_filtro - registros_despues_filtro
    
    if registros_filtrados > 0:
        info(f"✓ Eliminados {registros_filtrados} registros con FECHA posterior al cierre")

    # Conteo de registros por mes antes de filtrar
    df["MES_FECHA"] = df["FECHA_TEMP"].dt.to_period("M")
    resumen_mes = df.groupby("MES_FECHA").size()
    info(f"📊 Registros por mes antes de filtrar:\n{resumen_mes}")

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
    df["SALDO VENCIDO"] = df["SALDO"].where(df["FECHA VTO_TEMP"] <= fecha_cierre, 0)
    info("✓ Saldo vencido calculado")

    print("\n===== DEBUG FECHAS ACTIVIDAD 18 =====")

    df_18 = df[df["ACTIVIDAD"] == "18"]
    
    print("Facturas con VTO = fecha cierre:")
    print(
        df_18[df_18["FECHA VTO_TEMP"] == fecha_cierre][
            ["NUMERO FACTURA", "FECHA VTO", "SALDO"]
        ]
    )
    
    print("\nTotal saldo con VTO igual a cierre:")
    print(
        df_18[df_18["FECHA VTO_TEMP"] == fecha_cierre]["SALDO"].sum()
    )
    
     
    # -------------------------
    # 12. REORDENAR COLUMNAS
    # -------------------------
    columnas = list(df.columns)
    
    if "SALDO VENCIDO" in columnas:
    
        # Quitar columnas para reinsertarlas
        columnas.remove("VALOR")
        columnas.remove("SALDO")
    
        if "DIAS VENCIDO" in columnas:
            columnas.remove("DIAS VENCIDO")
        
        if "DIAS POR VENCER" in columnas:
            columnas.remove("DIAS POR VENCER")
    
        idx = columnas.index("SALDO VENCIDO")
    
        # Insertar antes de SALDO VENCIDO
        columnas.insert(idx, "SALDO")
        columnas.insert(idx, "VALOR")
    
        # Insertar después de SALDO VENCIDO
        columnas.insert(idx + 3, "DIAS VENCIDO")
        columnas.insert(idx + 4, "DIAS POR VENCER")
    
    df = df[columnas]
    
    info("✓ Columnas reordenadas correctamente (DIAS VENCIDO y DIAS POR VENCER después de SALDO VENCIDO)")
    
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
    # 15. COLUMNAS DE LOS ÚLTIMOS 6 MESES VENCIDOS (MES 1 = MES CIERRE)
    # -------------------------
    
    mes_cierre = fecha_cierre.month
    anio_cierre = fecha_cierre.year
    
    for i in range(0, 6):  # 0 = mes cierre, 1 = mes anterior, etc
        mes_calc = mes_cierre - i
        anio_calc = anio_cierre
    
        if mes_calc <= 0:
            mes_calc += 12
            anio_calc -= 1
    
        nombre_col = f"VTO MES {i+1}"
    
        df[nombre_col] = df.apply(
            lambda row: row["SALDO"] if (
                pd.notna(row["FECHA VTO_TEMP"]) and
                row["FECHA VTO_TEMP"].year == anio_calc and
                row["FECHA VTO_TEMP"].month == mes_calc and
                row["FECHA VTO_TEMP"] <= fecha_cierre
            ) else 0,
            axis=1
        )
    
    info("✓ Columnas VTO MES 1-6 creadas correctamente (MES 1 = MES CIERRE)")

    # -------------------------
    # 16. VALOR >= 180 DÍAS VENCIDOS
    # -------------------------
    df["VALOR >= 180 DIAS"] = df["SALDO"].where(df["DIAS VENCIDO"] >= 180, 0)

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
    
    print(df.columns)
    
    # -------------------------
    # 20. RANGOS DE VENCIMIENTO (se calculan PRIMERO para usarlos en MORA TOTAL)
    # -------------------------
    df["SALDO"] = df["SALDO"].round(2)

    df["SALDO NO VENCIDO"] = df["SALDO"].where(df["DIAS VENCIDO"].between(0, 29), 0)
    df["VENCIDO 30"]       = df["SALDO"].where(df["DIAS VENCIDO"].between(30, 59), 0)
    df["VENCIDO 60"]       = df["SALDO"].where(df["DIAS VENCIDO"].between(60, 89),  0)
    df["VENCIDO 90"]       = df["SALDO"].where(df["DIAS VENCIDO"].between(90, 179), 0)
    df["VENCIDO 180"]      = df["SALDO"].where(df["DIAS VENCIDO"].between(180, 359), 0)
    df["VENCIDO 360"]      = df["SALDO"].where(df["DIAS VENCIDO"].between(360, 369), 0)
    df["VENCIDO +360"]     = df["SALDO"].where(df["DIAS VENCIDO"] >= 370, 0)

    # MORA TOTAL = VTO MES 1+2+3+4+5+6 + VALOR >= 180 DIAS
    df["MORA TOTAL"] = (
        df["VTO MES 1"] + df["VTO MES 2"] + df["VTO MES 3"] +
        df["VTO MES 4"] + df["VTO MES 5"] + df["VTO MES 6"] +
        df["VALOR >= 180 DIAS"]
    ).round(2)

    # TOTAL POR VENCER = POR VENCER MES 1+2+3 + MAYOR 90 DIAS POR VENCER
    df["TOTAL POR VENCER"] = (df["SALDO"] - df["MORA TOTAL"]).round(2)
    
    # Diferencia técnica por redondeo
    df["DIFERENCIA_REAL"] = (
        df["SALDO"] - (df["MORA TOTAL"] + df["TOTAL POR VENCER"])
    ).round(4)
    
    # Validación tolerante a centavos
    df["VALIDACION_MORA_VENCER"] = (
        df["DIFERENCIA_REAL"].abs() < 0.01
    )
    
    # -------------------------
    # REPORTE DEBUG GLOBAL
    # -------------------------
    
    total_saldo = df["SALDO"].sum().round(2)
    total_mora = df["MORA TOTAL"].sum().round(2)
    total_vencer = df["TOTAL POR VENCER"].sum().round(2)
    
    info("===== VALIDACIÓN CONTABLE =====")
    info(f"SALDO TOTAL: {total_saldo}")
    info(f"MORA TOTAL: {total_mora}")
    info(f"POR VENCER TOTAL: {total_vencer}")
    info(f"SUMA MORA+VENCER: {(total_mora + total_vencer).round(2)}")
    info(f"DIFERENCIA GLOBAL: {(total_saldo - (total_mora + total_vencer)).round(4)}")

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
    df["FECHA"] = df["FECHA_TEMP"]
    df["FECHA VTO"] = df["FECHA VTO_TEMP"]
    
    df.drop(columns=["FECHA_TEMP", "FECHA VTO_TEMP"], inplace=True)
    
    info("✓ Fechas mantenidas como datetime real")

    # =========================
    # ORDENAR CRONOLÓGICAMENTE POR FECHA
    # =========================
    
    df = df.sort_values(by=["FECHA", "FECHA VTO"], ascending=[True, True])
    
    info("✓ Registros ordenados cronológicamente por FECHA y FECHA VTO")
    
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
    # TABLA TIPO DINÁMICA POR ACTIVIDAD
    # -------------------------
    
    tabla_dinamica = (
        df.groupby("ACTIVIDAD", dropna=False)
          .agg({
              "SALDO": "sum",
              "TOTAL POR VENCER": "sum",
              "MORA TOTAL": "sum",
              "VALOR DOTACION": "sum"
          })
          .reset_index()
    )
    
    # Renombrar columnas para que se vean como en tu imagen
    tabla_dinamica.columns = [
        "Etiquetas de fila",
        "Suma de SALDO",
        "Suma de TOTAL POR VENCER",
        "Suma de MORA TOTAL",
        "Suma de VALOR DOTACION"
    ]
    
    # Agregar fila Total general
    total_general = pd.DataFrame({
        "Etiquetas de fila": ["Total general"],
        "Suma de SALDO": [tabla_dinamica["Suma de SALDO"].sum()],
        "Suma de TOTAL POR VENCER": [tabla_dinamica["Suma de TOTAL POR VENCER"].sum()],
        "Suma de MORA TOTAL": [tabla_dinamica["Suma de MORA TOTAL"].sum()],
        "Suma de VALOR DOTACION": [tabla_dinamica["Suma de VALOR DOTACION"].sum()],
    })
    
    tabla_dinamica = pd.concat([tabla_dinamica, total_general], ignore_index=True)
    
    info("✓ Tabla tipo dinámica creada por ACTIVIDAD")
    
    # -------------------------
    # ELIMINAR COLUMNAS INTERNAS
    # -------------------------
    columnas_internas = [
        "MES_FECHA",
        "DIFERENCIA_REAL",
        "VALIDACION_MORA_VENCER",
        "VALIDACION_RANGOS",
        "SUMA_RANGOS" 
    ]
    
    df.drop(columns=[c for c in columnas_internas if c in df.columns], inplace=True)
    
    info("✓ Columnas internas eliminadas del archivo final")

    # -------------------------
    # ORDEN FINAL DE COLUMNAS
    # -------------------------
    
    orden_columnas = [
    "EMPRESA",
    "ACTIVIDAD",
    "EMPRESA CODIGO AGENTE",
    "CODIGO AGENTE",
    "AGENTE",
    "CODIGO COBRADOR",
    "COBRADOR",
    "CODIGO CLIENTE",
    "IDENTIFICACION",
    "NOMBRE",
    "DENOMINACION COMERCIAL",
    "DIRECCION",
    "TELEFONO",
    "CIUDAD",
    "NUMERO FACTURA",
    "TIPO",
    "FECHA",
    "FECHA VTO",
    "DIA VTO",
    "MES VTO",
    "AÑO VTO",
    "VALOR",
    "SALDO",
    "SALDO VENCIDO",
    "DIAS VENCIDO",
    "DIAS POR VENCER",
    "% DOTACION",
    "VALOR DOTACION",
    "VTO MES 1",
    "VTO MES 2",
    "VTO MES 3",
    "VTO MES 4",
    "VTO MES 5",
    "VTO MES 6",
    "VALOR >= 180 DIAS",
    "MORA TOTAL",
    "POR VENCER MES 1",
    "POR VENCER MES 2",
    "POR VENCER MES 3",
    "MAYOR 90 DIAS POR VENCER",
    "TOTAL POR VENCER",
    "SALDO NO VENCIDO",
    "VENCIDO 30",
    "VENCIDO 60",
    "VENCIDO 90",
    "VENCIDO 180",
    "VENCIDO 360",
    "VENCIDO +360",
    "TOTAL POR VENCER",
    "DEUDA INCOBRABLE"
    ]
    
    df = df[[c for c in orden_columnas if c in df.columns]]
    
    info("✓ Columnas ordenadas según formato final PROVCA")
    
    # -------------------------
    # EXPORTAR EXCEL CON FORMATO USANDO XlsxWriter
    # -------------------------
    info("\n=== GENERANDO ARCHIVO EXCEL ===")
    
    with pd.ExcelWriter(output_path, engine="xlsxwriter") as writer:

        # =========================
        # Exportar hojas
        # =========================
        df.to_excel(writer, index=False, sheet_name="DETALLE")
        resumen.to_excel(writer, index=False, sheet_name="RESUMEN")
        validaciones.to_excel(writer, index=False, sheet_name="VALIDACIONES")
        
        if len(registros_mora_vencer_invalidos) > 0:
            registros_mora_vencer_invalidos.to_excel(writer, index=False, sheet_name="ERROR_MORA_VENCER")
    
        if len(registros_rangos_invalidos) > 0:
            registros_rangos_invalidos.to_excel(writer, index=False, sheet_name="ERROR_RANGOS")
    
        tabla_dinamica.to_excel(writer, index=False, sheet_name="TABLA_DINAMICA")
    
        # =========================
        # Obtener workbook
        # =========================
        workbook = writer.book
    
        # =========================
        # Crear formatos (ANTES DE USARLOS)
        # =========================
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
    
        number_format = workbook.add_format({
            'num_format': '#,##0.00',
            'align': 'right'
        })
    
        # =========================
        # Obtener worksheets
        # =========================
        worksheet_detalle = writer.sheets["DETALLE"]
        worksheet_resumen = writer.sheets["RESUMEN"]
        worksheet_validaciones = writer.sheets["VALIDACIONES"]
        worksheet_dinamica = writer.sheets["TABLA_DINAMICA"]
    
        # =========================
        # Formato hoja dinámica
        # =========================
        worksheet_dinamica.set_column(0, 0, 20, text_format)
        worksheet_dinamica.set_column(1, 4, 25, number_format)
    
        for col_num, value in enumerate(tabla_dinamica.columns.values):
            worksheet_dinamica.write(0, col_num, value, header_format)
    
        worksheet_dinamica.set_row(0, 30)
          
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
        
        
        # Reescribir fechas como datetime real (blindaje total)
        for row_num, row_data in enumerate(df.values):
            for col_num, cell_data in enumerate(row_data):
        
                if col_num in fecha_cols and pd.notna(cell_data):
                    worksheet_detalle.write_datetime(
                        row_num + 1,
                        col_num,
                        cell_data,
                        date_format
                    )
        
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
                worksheet_detalle.set_column(i, i, max_len, date_format)
            
            elif i in percent_cols:
                worksheet_detalle.set_column(i, i, max_len, percent_format)
            
            elif i in valor_cols:
                worksheet_detalle.set_column(i, i, max_len, number_format)
            
            elif i in integer_cols:
                worksheet_detalle.set_column(i, i, max_len, workbook.add_format({
                    'align': 'center'
                }))
            
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
        if len(sys.argv) < 2:
            raise ValueError("Debe indicar el archivo de entrada")

        input_path = sys.argv[1]
        fecha_cierre = None
        output_path = None

        # Si envían fecha
        if len(sys.argv) >= 3:
            try:
                pd.to_datetime(sys.argv[2], format="%Y-%m-%d")
                fecha_cierre = sys.argv[2]
            except:
                raise ValueError("La fecha debe tener formato YYYY-MM-DD")

        # Si envían nombre de archivo de salida
        if len(sys.argv) >= 4:
            output_path = sys.argv[3]

        resultado = procesar_cartera(input_path, output_path, fecha_cierre)

        info(f"\n{'='*60}")
        info("PROCESO COMPLETADO EXITOSAMENTE")
        info(f"{'='*60}")
        info(f"Archivo: {resultado}")

    except Exception as e:
        logging.error(f"Error: {str(e)}", exc_info=True)
        print(f"\n{'='*60}")
        print("ERROR EN EL PROCESO")
        print(f"{'='*60}")
        print(str(e))
        sys.exit(1)

if __name__ == "__main__":
    main()
