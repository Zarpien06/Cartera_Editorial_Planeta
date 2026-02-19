# ==========================================================
# MODELO DE DEUDA - Cumplimiento completo procedimiento
# Departamento de Cartera
# Versión corregida con todos los puntos del procedimiento:
#   - Columna % DOTACION
#   - Columna SALDO VENCIDO
#   - Columna MAYOR 90 DIAS
#   - PL10 corregido a LIBRERIAS 2 según procedimiento (pág. 7)
#   - Anticipos como registros (NO compensación)
#   - TRM último día mes anterior desde trm.json
#   - Totales PESOS, DIVISAS (moneda + COP) y VENCIMIENTO
# ==========================================================

import argparse
import sys
import pandas as pd
import numpy as np
import os
import logging
from datetime import datetime, timedelta
from typing import Any, Optional
import json

# ---------------- Logging unificado ----------------
try:
    from config_logging import logger, log_inicio_proceso, log_fin_proceso, log_error_proceso
    USE_UNIFIED_LOGGING = True
except ImportError:
    LOG_FILE_PATH = os.path.join(os.path.dirname(os.path.abspath(__file__)), 'modelo_deuda.log')
    logging.basicConfig(
        filename=LOG_FILE_PATH,
        level=logging.INFO,
        format='[%(asctime)s] [%(levelname)s] [MODELO_DEUDA] %(message)s',
        datefmt='%Y-%m-%d %H:%M:%S'
    )
    USE_UNIFIED_LOGGING = False
    logger = logging.getLogger(__name__)

try:
    from log_bridge import attach_python_logging, log_event
    attach_python_logging('PY_MODELO_DEUDA')
    log_event('PY_MODELO_DEUDA', 'INFO', 'Iniciando modelo de deuda')
except ImportError:
    pass

if not USE_UNIFIED_LOGGING:
    logging.info("Iniciando modelo de deuda")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SALIDAS_DIR = os.path.join(BASE_DIR, 'salidas')
os.makedirs(SALIDAS_DIR, exist_ok=True)

# ---------------- Parámetros del negocio ----------------
LINEAS_PESOS = [
    ('CT', '80'), ('ED', '41'), ('ED', '44'), ('ED', '47'),
    ('PL', '10'), ('PL', '15'), ('PL', '20'), ('PL', '21'),
    ('PL', '23'), ('PL', '25'), ('PL', '28'), ('PL', '29'),
    ('PL', '31'), ('PL', '32'), ('PL', '53'), ('PL', '56'),
    ('PL', '60'), ('PL', '62'), ('PL', '63'), ('PL', '64'),
    ('PL', '65'), ('PL', '66'), ('PL', '69')
]

LINEAS_DIVISAS = [
    ('PL', '11'), ('PL', '18'), ('PL', '57'), ('PL', '41')
]

# *** Tabla Negocio-Canal corregida según procedimiento pág. 7 ***
# CORRECCIÓN OBS-4: PL10 -> LIBRERIAS 2 (no LIBRERIAS 3 como estaba antes)
TABLA_NEGOCIO_CANAL = {
    'PL10': {'NEGOCIO': 'LIBRERIAS 2', 'CANAL': 'LIBRERIAS 2'},   # CORREGIDO: era LIBRERIAS 3
    'PL15': {'NEGOCIO': 'E-COMMERCE',  'CANAL': 'E-COMMERCE'},
    'PL20': {'NEGOCIO': 'LIBRERIAS 1', 'CANAL': 'LIBRERIAS 1'},
    'PL21': {'NEGOCIO': 'LIBRERIAS 2', 'CANAL': 'LIBRERIAS 2'},
    'PL23': {'NEGOCIO': 'LIBRERIAS 1', 'CANAL': 'LIBRERIAS 1'},
    'PL25': {'NEGOCIO': 'LIBRERIAS 1', 'CANAL': 'LIBRERIAS 1'},
    'PL28': {'NEGOCIO': 'SALDOS',      'CANAL': 'SALDOS'},
    'PL29': {'NEGOCIO': 'SALDOS',      'CANAL': 'SALDOS'},
    'PL31': {'NEGOCIO': 'SALDOS',      'CANAL': 'SALDOS'},
    'PL32': {'NEGOCIO': 'DISTRIBUIDORES', 'CANAL': 'DISTRIBUIDORES'},
    'PL53': {'NEGOCIO': 'LIBRERIAS 3', 'CANAL': 'LIBRERIAS 3'},
    'PL56': {'NEGOCIO': 'OTROS DIGITAL','CANAL': 'OTROS DIGITAL'},
    'PL57': {'NEGOCIO': 'PRENSA USD',  'CANAL': 'PRENSA USD'},
    'PL60': {'NEGOCIO': 'OTROS',       'CANAL': 'OTROS'},
    'PL62': {'NEGOCIO': 'PRENSA',      'CANAL': 'PRENSA'},
    'PL63': {'NEGOCIO': 'LIBRERIAS 3', 'CANAL': 'LIBRERIAS 3'},
    'PL64': {'NEGOCIO': 'OTROS',       'CANAL': 'OTROS'},
    'PL65': {'NEGOCIO': 'OTROS',       'CANAL': 'OTROS'},
    'PL66': {'NEGOCIO': 'OTROS DIGITAL','CANAL': 'OTROS DIGITAL'},
    'PL69': {'NEGOCIO': 'LIBRERIAS 1', 'CANAL': 'LIBRERIAS 1'},
    'PL11': {'NEGOCIO': 'EXPORTACION USD', 'CANAL': 'EXPORTACION USD'},
    'PL18': {'NEGOCIO': 'EXPORTACION USD', 'CANAL': 'EXPORTACION USD'},
    'PL41': {'NEGOCIO': 'EXPORTACION EURO','CANAL': 'EXPORTACION EURO'},
    'CT80': {'NEGOCIO': 'TINTA CLUB DEL LIBRO', 'CANAL': 'TINTA'},
    'ED41': {'NEGOCIO': 'EDUCACION',   'CANAL': 'EDUCACION'},
    'ED44': {'NEGOCIO': 'OTROS DIGITAL','CANAL': 'OTROS DIGITAL'},
    'ED47': {'NEGOCIO': 'EDUCACION',   'CANAL': 'EDUCACION'},
}

# ---------------- Utilidades ----------------
def safe_float_conversion(val: Any) -> float:
    if val is None:
        return 0.0
    if isinstance(val, (int, float, np.number)):
        return float(val)
    try:
        return float(str(val).replace(",", "").replace(" ", ""))
    except (ValueError, TypeError):
        return 0.0

def safe_isna_check(val: Any) -> bool:
    try:
        return pd.isna(val)
    except Exception:
        return val is None

def obtener_ultimo_dia_mes_anterior(fecha_cierre_str: str) -> str:
    fecha_cierre = datetime.strptime(fecha_cierre_str, "%Y-%m-%d")
    primer_dia_mes = fecha_cierre.replace(day=1)
    ultimo_dia_mes_anterior = primer_dia_mes - timedelta(days=1)
    return ultimo_dia_mes_anterior.strftime("%Y-%m-%d")

def cargar_tasas_cambio(fecha_cierre: str,
                        usd_override: Optional[float] = None,
                        eur_override: Optional[float] = None):
    TRM_FILE = os.path.join(BASE_DIR, "trm.json")
    fecha_trm = obtener_ultimo_dia_mes_anterior(fecha_cierre)

    if not os.path.exists(TRM_FILE):
        raise FileNotFoundError(
            f"No existe archivo trm.json en {BASE_DIR}. "
            "Créelo con el formato: {{\"YYYY-MM-DD\": {{\"USD\": 0.0, \"EUR\": 0.0}}}}"
        )

    with open(TRM_FILE, "r", encoding="utf-8") as f:
        data = json.load(f)

    if fecha_trm not in data:
        raise ValueError(
            f"No existe TRM registrada para la fecha {fecha_trm} "
            "(último día del mes anterior al cierre)"
        )

    trm_dolar = float(data[fecha_trm].get("USD", 0))
    trm_euro  = float(data[fecha_trm].get("EUR", 0))

    if usd_override and usd_override > 0:
        trm_dolar = usd_override
        print(f"  [INFO] TRM USD sobrescrita por CLI: {trm_dolar}")

    if eur_override and eur_override > 0:
        trm_euro = eur_override
        print(f"  [INFO] TRM EUR sobrescrita por CLI: {trm_euro}")

    print(f"  [OK] TRM utilizada ({fecha_trm})  USD: {trm_dolar:,.4f}  |  EUR: {trm_euro:,.4f}")
    return trm_dolar, trm_euro, fecha_trm

def _last_day_of_month(dt: pd.Timestamp) -> pd.Timestamp:
    return dt.to_period('M').to_timestamp('M')

def _moneda_por_linea(linea_key: str) -> str:
    k = str(linea_key).strip().upper()
    if k in ('PL11', 'PL18', 'PL57'):
        return 'DÓLAR'
    if k == 'PL41':
        return 'EURO'
    return 'PESOS COL'

def _build_linea_key(emp: str, act: str) -> str:
    emp = str(emp).strip().upper()
    act = str(act).strip().upper().replace('.0', '')
    return f"{emp}{act}"

def _ensure_datetime(series):
    try:
        return pd.to_datetime(series, dayfirst=True, errors='coerce')
    except Exception:
        return pd.to_datetime(series, errors='coerce')

# ---------------- Lectura flexible de archivos ----------------
def leer_archivo(archivo: str) -> pd.DataFrame:
    if not os.path.exists(archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {archivo}")
    if archivo.lower().endswith('.csv'):
        for enc in ['utf-8-sig', 'latin1', 'cp1252', 'utf-8']:
            for sep in [';', ',', '\t', '|']:
                try:
                    df = pd.read_csv(archivo, encoding=enc, sep=sep)
                    if len(df.columns) > 1 or len(df) > 0:
                        return df
                except Exception:
                    continue
        return pd.read_csv(archivo, encoding='latin1', sep=';')
    if archivo.lower().endswith(('.xlsx', '.xls')):
        return pd.read_excel(archivo, engine='openpyxl')
    return pd.read_csv(archivo, encoding='latin1', sep=';')

# ============================================================
# CÁLCULO COMPLETO DE CAMPOS (procedimiento completo pág. 3)
# ============================================================
def calcular_campos_provision(df: pd.DataFrame) -> pd.DataFrame:
    """
    Aplica todas las transformaciones del procedimiento al archivo de provisión:
    - Elimina columna PCIMCO (columna U)
    - Elimina fila PL30 con saldo -614.000
    - Unifica NOMBRE en DENOMINACION COMERCIAL
    - Convierte fechas a datetime con dayfirst=True
    - Abre FECHA VTO en columnas DIA / MES / AÑO  (procedimiento pág. 3)
    - Calcula DIAS VENCIDOS y DIAS POR VENCER
    - Calcula SALDO VENCIDO                        (OBS-2 corregida)
    - Calcula % DOTACION                           (OBS-1 corregida)
    - Calcula DEUDA INCOBRABLE (= valor dotación)
    - Crea 7 buckets de vencimiento
    - Valida suma buckets == SALDO
    - Calcula MORA TOTAL y TOTAL POR VENCER
    - Valida MORA TOTAL + TOTAL POR VENCER == SALDO
    - Calcula POR VENCER M1, M2, M3
    - Calcula MAYOR 90 DIAS                        (OBS-3 corregida)
    """

    # -- 1. Eliminar columna PCIMCO (columna U del procedimiento) --
    for col in list(df.columns):
        if col.strip().upper() == 'PCIMCO':
            df = df.drop(columns=[col])
            print("  [OK] Columna PCIMCO eliminada")

    # -- 2. Unificar DENOMINACION COMERCIAL sin borrar datos existentes --
    if 'PCNMCL' in df.columns:
        if 'DENOMINACION COMERCIAL' not in df.columns:
            df['DENOMINACION COMERCIAL'] = df['PCNMCL']
        else:
            df['DENOMINACION COMERCIAL'] = df['DENOMINACION COMERCIAL'].fillna(df['PCNMCL'])
    if 'NOMBRE' in df.columns and 'DENOMINACION COMERCIAL' in df.columns:
        df['DENOMINACION COMERCIAL'] = df['DENOMINACION COMERCIAL'].fillna(df['NOMBRE'])

    # -- 3. Convertir fechas a datetime formato día-mes-año --
    if 'FECHA' in df.columns:
        df['FECHA'] = _ensure_datetime(df['FECHA'])
    if 'FECHA VTO' in df.columns:
        df['FECHA VTO'] = _ensure_datetime(df['FECHA VTO'])

    # -- 4. Abrir FECHA VTO en día, mes, año (procedimiento pág. 3) --
    if 'FECHA VTO' in df.columns:
        df['VTO_DIA']  = df['FECHA VTO'].dt.day
        df['VTO_MES']  = df['FECHA VTO'].dt.month
        df['VTO_ANIO'] = df['FECHA VTO'].dt.year

    # -- 5. Línea de negocio --
    if 'EMPRESA' in df.columns and 'ACTIVIDAD' in df.columns:
        df['LINEA DE NEGOCIO'] = df.apply(
            lambda r: _build_linea_key(r['EMPRESA'], r['ACTIVIDAD']), axis=1
        )

    # -- 6. Eliminar fila específica PL30 - saldo -614.000 --
    try:
        if 'ACTIVIDAD' in df.columns and 'SALDO' in df.columns:
            df['__SAL_NUM__'] = pd.to_numeric(df['SALDO'], errors='coerce')
            cond = (
                df['ACTIVIDAD'].astype(str).str.strip() == '30'
            ) & (
                df['__SAL_NUM__'].round(0) == -614000
            )
            removed = cond.sum()
            if removed:
                print(f"  [OK] Fila PL30 (-614.000) eliminada: {removed} registro(s)")
            df = df.loc[~cond].drop(columns=['__SAL_NUM__'])
    except Exception:
        pass

    # -- 7. Fecha de corte: último día del mes del documento más reciente --
    if 'FECHA' in df.columns and df['FECHA'].notna().any():
        fecha_corte = _last_day_of_month(df['FECHA'].max())
    else:
        fecha_corte = _last_day_of_month(pd.Timestamp.today())

    # -- 8. SALDO a numérico --
    if 'SALDO' in df.columns:
        df['SALDO'] = pd.to_numeric(df['SALDO'], errors='coerce').fillna(0.0)
    else:
        df['SALDO'] = 0.0

    # -- 9. Días vencidos y días por vencer --
    if 'FECHA VTO' in df.columns:
        df['DIAS VENCIDOS']  = (fecha_corte - df['FECHA VTO']).dt.days
        df['DIAS POR VENCER'] = (df['FECHA VTO'] - fecha_corte).dt.days.clip(lower=0)
    else:
        df['DIAS VENCIDOS']   = 0
        df['DIAS POR VENCER'] = 0

    # -- 10. SALDO VENCIDO (procedimiento pág. 3 - OBS-2 corregida) --
    # Monto vencido de cada factura (saldo de facturas con días vencidos > 0)
    df['SALDO VENCIDO'] = np.where(df['DIAS VENCIDOS'] > 0, df['SALDO'], 0.0)

    # -- 11. % DOTACION (procedimiento pág. 3 - OBS-1 corregida) --
    # 100% si DIAS VENCIDOS >= 180, 0% en caso contrario
    df['% DOTACION'] = np.where(df['DIAS VENCIDOS'] >= 180, 1.0, 0.0)

    # -- 12. DEUDA INCOBRABLE = SALDO x % DOTACION (= valor dotación) --
    df['DEUDA INCOBRABLE'] = np.where(df['DIAS VENCIDOS'] >= 180, df['SALDO'], 0.0)

    # -- 13. Siete buckets de vencimiento (procedimiento pág. 3) --
    df['SALDO NO VENCIDO'] = np.where(df['DIAS VENCIDOS'] <= 0,                                         df['SALDO'], 0.0)
    df['VENCIDO 30']       = np.where((df['DIAS VENCIDOS'] >= 30)  & (df['DIAS VENCIDOS'] <= 59),       df['SALDO'], 0.0)
    df['VENCIDO 60']       = np.where((df['DIAS VENCIDOS'] >= 60)  & (df['DIAS VENCIDOS'] <= 89),       df['SALDO'], 0.0)
    df['VENCIDO 90']       = np.where((df['DIAS VENCIDOS'] >= 90)  & (df['DIAS VENCIDOS'] <= 179),      df['SALDO'], 0.0)
    df['VENCIDO 180']      = np.where((df['DIAS VENCIDOS'] >= 180) & (df['DIAS VENCIDOS'] <= 359),      df['SALDO'], 0.0)
    df['VENCIDO 360']      = np.where((df['DIAS VENCIDOS'] >= 360) & (df['DIAS VENCIDOS'] <= 369),      df['SALDO'], 0.0)
    df['VENCIDO + 360']    = np.where(df['DIAS VENCIDOS'] >= 370,                                       df['SALDO'], 0.0)

    # -- 14. Validación: suma de buckets == SALDO --
    suma_buckets = (
        df['SALDO NO VENCIDO'] + df['VENCIDO 30'] + df['VENCIDO 60'] +
        df['VENCIDO 90'] + df['VENCIDO 180'] + df['VENCIDO 360'] + df['VENCIDO + 360']
    ).round(2)
    mismatches_buckets = ((suma_buckets - df['SALDO'].round(2)).abs() > 0.01).sum()
    if mismatches_buckets:
        print(f"  [WARN] {mismatches_buckets} factura(s): suma(buckets) != SALDO")
    else:
        print("  [OK] Validacion buckets: suma(buckets) = SALDO en todas las facturas")

    # -- 15. MORA TOTAL y TOTAL POR VENCER --
    df['MORA TOTAL'] = df[['VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90',
                            'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360']].sum(axis=1)
    df['TOTAL POR VENCER'] = df['SALDO NO VENCIDO'].fillna(0.0)

    # -- 16. Validación: MORA TOTAL + TOTAL POR VENCER == SALDO --
    mismatches_mora = (
        ((df['MORA TOTAL'] + df['TOTAL POR VENCER']).round(2) - df['SALDO'].round(2)).abs() > 0.01
    ).sum()
    if mismatches_mora:
        print(f"  [WARN] {mismatches_mora} factura(s): MORA TOTAL + POR VENCER != SALDO")
    else:
        print("  [OK] Validacion mora: MORA TOTAL + TOTAL POR VENCER = SALDO en todas las facturas")

    # -- 17. Por vencer en los 3 meses siguientes al cierre --
    if 'FECHA VTO' in df.columns:
        meses_diff = (
            df['FECHA VTO'].dt.to_period('M') - fecha_corte.to_period('M')
        ).apply(lambda p: p.n)
        df['POR VENCER M1'] = np.where(meses_diff == 1, df['SALDO'], 0.0)
        df['POR VENCER M2'] = np.where(meses_diff == 2, df['SALDO'], 0.0)
        df['POR VENCER M3'] = np.where(meses_diff == 3, df['SALDO'], 0.0)
    else:
        df['POR VENCER M1'] = 0.0
        df['POR VENCER M2'] = 0.0
        df['POR VENCER M3'] = 0.0

    # -- 18. MAYOR 90 DIAS (procedimiento pág. 3 - OBS-3 corregida) --
    # Suma de todos los buckets con vencimiento >= 90 días
    df['MAYOR 90 DIAS'] = df[['VENCIDO 90', 'VENCIDO 180',
                               'VENCIDO 360', 'VENCIDO + 360']].sum(axis=1)

    return df

# ============================================================
# HOJA VENCIMIENTO
# ============================================================
def crear_hoja_vencimientos(df_pesos_final: pd.DataFrame,
                            df_divisas_final: pd.DataFrame) -> pd.DataFrame:

    columnas_sumables = [
        'SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60',
        'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360',
        'VENCIDO + 360', 'DEUDA INCOBRABLE'
    ]

    df_all = pd.concat([df_pesos_final, df_divisas_final], ignore_index=True)

    if 'LINEA DE NEGOCIO' in df_all.columns:
        df_all['NEGOCIO'] = df_all['LINEA DE NEGOCIO'].apply(
            lambda k: TABLA_NEGOCIO_CANAL.get(
                str(k).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'}
            )['NEGOCIO']
        )
        df_all['CANAL'] = df_all['LINEA DE NEGOCIO'].apply(
            lambda k: TABLA_NEGOCIO_CANAL.get(
                str(k).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'}
            )['CANAL']
        )
        df_all['MONEDA'] = df_all['LINEA DE NEGOCIO'].apply(
            lambda k: _moneda_por_linea(str(k).strip().upper())
        )

    df_all['MONEDA']  = df_all.get('MONEDA', 'PESOS COL')
    df_all['CLIENTE'] = df_all.get('DENOMINACION COMERCIAL', '')

    grp_cols = ['NEGOCIO', 'CANAL', 'MONEDA', 'CLIENTE']
    for c in grp_cols:
        if c not in df_all.columns:
            df_all[c] = ''

    df_sum = (
        df_all
        .groupby(grp_cols, as_index=False)[columnas_sumables]
        .sum()
    )

    df_sum.insert(0, 'PAIS',       'COLOMBIA')
    df_sum.insert(3, 'COBRO/PAGO', 'CLIENTE')
    df_sum = df_sum.rename(columns={'SALDO': 'SALDO TOTAL'})

    # Totales generales por moneda (al final, según procedimiento pág. 6/8)
    totales_moneda = (
        df_sum
        .groupby('MONEDA', as_index=False)[[
            'SALDO TOTAL', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60',
            'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360',
            'VENCIDO + 360', 'DEUDA INCOBRABLE'
        ]]
        .sum()
    )
    totales_moneda['PAIS']       = 'COLOMBIA'
    totales_moneda['NEGOCIO']    = 'TOTAL'
    totales_moneda['CANAL']      = 'TOTAL'
    totales_moneda['COBRO/PAGO'] = ''
    totales_moneda['CLIENTE']    = 'TOTAL GENERAL POR MONEDA'
    totales_moneda = totales_moneda[df_sum.columns]

    df_final = pd.concat([df_sum, totales_moneda], ignore_index=True)
    return df_final

# ============================================================
# FUNCIÓN PRINCIPAL: CREAR MODELO DE DEUDA
# ============================================================
def crear_modelo_deuda(archivo_provision: str,
                       archivo_anticipos: str,
                       output_file: str = '1_Modelo_Deuda.xlsx',
                       usd_override: Optional[float] = None,
                       eur_override: Optional[float] = None) -> str:

    if USE_UNIFIED_LOGGING:
        log_inicio_proceso("MODELO_DEUDA", f"{archivo_provision} + {archivo_anticipos}")
    else:
        logging.info("Iniciando modelo de deuda")

    print("\n" + "=" * 62)
    print("  MODELO DE DEUDA -- Procedimiento Departamento de Cartera")
    print("=" * 62)

    # -- PASO 1: TRM --
    print("\n[1/7] Cargando tasas de cambio...")

    # Intentar cargar desde trm_config si existe, fallback a trm.json directo
    try:
        from trm_config import load_trm
        trm_data   = load_trm()
        trm_dolar  = trm_data["usd"]
        trm_euro   = trm_data["eur"]
        fecha_trm  = trm_data["fecha"]
        if usd_override and usd_override > 0:
            trm_dolar = usd_override
        if eur_override and eur_override > 0:
            trm_euro = eur_override
        print(f"  [OK] TRM ({fecha_trm})  USD: {trm_dolar:,.4f}  |  EUR: {trm_euro:,.4f}")
    except ImportError:
        # Leer provisión para obtener fecha de cierre y calcular último día mes anterior
        _df_temp = leer_archivo(archivo_provision)
        _mapeo_temp = {'PCCDEM': 'EMPRESA', 'PCCDAC': 'ACTIVIDAD', 'PCFEFA': 'FECHA'}
        _df_temp = _df_temp.rename(columns={k: v for k, v in _mapeo_temp.items() if k in _df_temp.columns})
        if 'FECHA' in _df_temp.columns:
            _df_temp['FECHA'] = _ensure_datetime(_df_temp['FECHA'])
            _fecha_cierre = _last_day_of_month(_df_temp['FECHA'].max()).strftime("%Y-%m-%d")
        else:
            _fecha_cierre = datetime.today().strftime("%Y-%m-%d")
        trm_dolar, trm_euro, fecha_trm = cargar_tasas_cambio(
            _fecha_cierre, usd_override, eur_override
        )

    # -- PASO 2: Leer archivos --
    print("\n[2/7] Leyendo archivos de entrada...")
    df_provision_raw = leer_archivo(archivo_provision)
    df_anticipos_raw = leer_archivo(archivo_anticipos)
    print(f"  [OK] Provisión: {len(df_provision_raw):,} registros")
    print(f"  [OK] Anticipos: {len(df_anticipos_raw):,} registros")

    # -- PASO 3: Renombrar columnas provisión (procedimiento pág. 2) --
    print("\n[3/7] Normalizando provisión y calculando vencimientos...")
    mapeo_cartera = {
        'PCCDEM': 'EMPRESA',
        'PCCDAC': 'ACTIVIDAD',
        'PCDEAC': 'EMPRESA',
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
    }
    df_provision = df_provision_raw.rename(
        columns={k: v for k, v in mapeo_cartera.items() if k in df_provision_raw.columns}
    )

    # Aplicar todas las transformaciones del procedimiento
    df_provision = calcular_campos_provision(df_provision)

    # -- PASO 4: Procesar anticipos (procedimiento pág. 4) --
    print("\n[4/7] Procesando anticipos (registros negativos, no compensación)...")
    mapeo_anticipos = {
        'NCCDEM': 'EMPRESA',
        'NCCDAC': 'ACTIVIDAD',
        'NCCDCL': 'CODIGO CLIENTE',
        'WWNIT':  'IDENTIFICACION',
        'WWNMCL': 'DENOMINACION COMERCIAL',
        'WWNMDO': 'DIRECCION',
        'WWTLF1': 'TELEFONO',
        'WWNMPO': 'CIUDAD',
        'CCCDFB': 'CODIGO AGENTE',
        'BDNMNM': 'NOMBRE AGENTE',
        'BDNMPA': 'APELLIDO AGENTE',
        'NCMOMO': 'TIPO ANTICIPO',
        'NCCDR3': 'NUMERO ANTICIPO',
        'NCIMAN': 'VALOR ANTICIPO',
        'NCFEGR': 'FECHA',
    }
    df_anticipos = df_anticipos_raw.rename(
        columns={k: v for k, v in mapeo_anticipos.items() if k in df_anticipos_raw.columns}
    )

    if 'EMPRESA' in df_anticipos.columns and 'ACTIVIDAD' in df_anticipos.columns:
        df_anticipos['LINEA DE NEGOCIO'] = df_anticipos.apply(
            lambda r: _build_linea_key(r['EMPRESA'], r['ACTIVIDAD']), axis=1
        )

    # Valor anticipo * -1 (debe ser negativo, procedimiento pág. 4)
    if 'VALOR ANTICIPO' in df_anticipos.columns:
        df_anticipos['SALDO'] = (
            pd.to_numeric(df_anticipos['VALOR ANTICIPO'], errors='coerce').fillna(0.0) * -1
        )
    else:
        df_anticipos['SALDO'] = 0.0

    # Normalizar columna DENOMINACION COMERCIAL en anticipos
    if 'DENOMINACION COMERCIAL' not in df_anticipos.columns:
        df_anticipos['DENOMINACION COMERCIAL'] = ''

    df_anticipos['FECHA'] = _ensure_datetime(df_anticipos.get('FECHA', pd.NaT))

    # Fecha de corte desde provisión
    if 'FECHA' in df_provision.columns and df_provision['FECHA'].notna().any():
        fecha_corte = _last_day_of_month(df_provision['FECHA'].max())
    else:
        fecha_corte = _last_day_of_month(pd.Timestamp.today())

    df_anticipos['FECHA VTO']      = fecha_corte
    df_anticipos['VTO_DIA']        = fecha_corte.day
    df_anticipos['VTO_MES']        = fecha_corte.month
    df_anticipos['VTO_ANIO']       = fecha_corte.year
    df_anticipos['DIAS VENCIDOS']  = 0
    df_anticipos['DIAS POR VENCER']= 0

    # Anticipos: van en SALDO NO VENCIDO (procedimiento pág. 5)
    df_anticipos['SALDO VENCIDO']   = 0.0
    df_anticipos['% DOTACION']      = 0.0
    df_anticipos['DEUDA INCOBRABLE']= 0.0
    df_anticipos['SALDO NO VENCIDO']= df_anticipos['SALDO']
    for col in ['VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90',
                'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360']:
        df_anticipos[col] = 0.0
    df_anticipos['MORA TOTAL']      = 0.0
    df_anticipos['TOTAL POR VENCER']= df_anticipos['SALDO']
    df_anticipos['POR VENCER M1']   = 0.0
    df_anticipos['POR VENCER M2']   = 0.0
    df_anticipos['POR VENCER M3']   = 0.0
    df_anticipos['MAYOR 90 DIAS']   = 0.0

    print(f"  [OK] {len(df_anticipos):,} anticipos procesados")

    # -- PASO 5: Separar PESOS y DIVISAS (procedimientos pág. 5-6) --
    print("\n[5/7] Separando hojas PESOS y DIVISAS...")
    lineas_pesos_keys   = {f"{cod}{act}" for cod, act in LINEAS_PESOS}
    lineas_divisas_keys = {f"{cod}{act}" for cod, act in LINEAS_DIVISAS}

    df_pesos   = df_provision[df_provision['LINEA DE NEGOCIO'].isin(lineas_pesos_keys)].copy()
    df_divisas = df_provision[df_provision['LINEA DE NEGOCIO'].isin(lineas_divisas_keys)].copy()

    ant_pesos = df_anticipos[df_anticipos['LINEA DE NEGOCIO'].isin(lineas_pesos_keys)].copy()
    ant_div   = df_anticipos[df_anticipos['LINEA DE NEGOCIO'].isin(lineas_divisas_keys)].copy()

    df_pesos['MONEDA']   = 'PESOS COL'
    ant_pesos['MONEDA']  = 'PESOS COL'
    df_divisas['MONEDA'] = df_divisas['LINEA DE NEGOCIO'].apply(_moneda_por_linea)
    ant_div['MONEDA']    = ant_div['LINEA DE NEGOCIO'].apply(_moneda_por_linea)

    df_pesos_final   = pd.concat([df_pesos,   ant_pesos], ignore_index=True)
    df_divisas_final = pd.concat([df_divisas, ant_div],   ignore_index=True)

    # Asignar NEGOCIO y CANAL
    for dfX in (df_pesos_final, df_divisas_final):
        dfX['NEGOCIO'] = dfX['LINEA DE NEGOCIO'].apply(
            lambda k: TABLA_NEGOCIO_CANAL.get(
                str(k).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'}
            )['NEGOCIO']
        )
        dfX['CANAL'] = dfX['LINEA DE NEGOCIO'].apply(
            lambda k: TABLA_NEGOCIO_CANAL.get(
                str(k).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'}
            )['CANAL']
        )

    total_pesos = df_pesos_final['SALDO'].sum() if 'SALDO' in df_pesos_final.columns else 0.0
    total_div   = df_divisas_final['SALDO'].sum() if 'SALDO' in df_divisas_final.columns else 0.0
    print(f"  [OK] Registros PESOS:   {len(df_pesos_final):,}  |  Saldo: ${total_pesos:,.0f}")
    print(f"  [OK] Registros DIVISAS: {len(df_divisas_final):,}  |  Saldo moneda original: {total_div:,.2f}")

    # -- PASO 6: Hoja VENCIMIENTO --
    print("\n[6/7] Generando hoja VENCIMIENTO...")
    df_vencimientos = crear_hoja_vencimientos(df_pesos_final, df_divisas_final)
    print(f"  [OK] {len(df_vencimientos):,} filas en hoja VENCIMIENTO")

    # -- PASO 7: Exportar a Excel --
    print("\n[7/7] Guardando archivo Excel...")
    output_path = os.path.join(SALIDAS_DIR, output_file)

    # Columnas numéricas a formatear
    cols_num = [
        'SALDO', 'VALOR', 'SALDO NO VENCIDO',
        'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90',
        'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360',
        'DEUDA INCOBRABLE', 'SALDO TOTAL', 'MORA TOTAL',
        'TOTAL POR VENCER', 'POR VENCER M1', 'POR VENCER M2', 'POR VENCER M3',
        'SALDO VENCIDO', 'MAYOR 90 DIAS',
    ]

    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        wb = writer.book
        fmt_miles    = wb.add_format({'num_format': '#,##0.00;-#,##0.00;"-";@'})
        fmt_pct      = wb.add_format({'num_format': '0%'})
        fmt_texto    = wb.add_format({'num_format': '@'})
        fmt_header   = wb.add_format({
            'bold': True, 'bg_color': '#1F3864', 'font_color': '#FFFFFF',
            'border': 1, 'align': 'center', 'valign': 'vcenter', 'text_wrap': True
        })
        fmt_total    = wb.add_format({
            'bold': True, 'bg_color': '#D6E4F0', 'num_format': '#,##0.00;-#,##0.00;"-";@',
            'border': 1
        })
        fmt_total_txt= wb.add_format({
            'bold': True, 'bg_color': '#D6E4F0', 'border': 1
        })
        fmt_alt      = wb.add_format({'bg_color': '#F2F7FB'})

        def _limpiar_df(df_data: pd.DataFrame) -> pd.DataFrame:
            """Reemplaza NaN e Inf en columnas numericas por 0.0 para evitar error xlsxwriter."""
            df_out = df_data.copy()
            for col in df_out.columns:
                if col in cols_num or col == '% DOTACION':
                    df_out[col] = pd.to_numeric(df_out[col], errors='coerce').fillna(0.0)
                    df_out[col] = df_out[col].replace([float('inf'), float('-inf')], 0.0)
            return df_out

        def _safe_num(val):
            """Devuelve float limpio o 0.0 si es NaN/Inf/None."""
            try:
                v = float(val)
                if v != v or v == float('inf') or v == float('-inf'):
                    return 0.0
                return v
            except (TypeError, ValueError):
                return 0.0

        def _escribir_hoja(df_data: pd.DataFrame, nombre_hoja: str,
                           fila_total_marker: Optional[str] = None,
                           col_total_marker: Optional[str] = None):
            """Escribe una hoja con encabezados formateados y fila de total destacada."""
            # Limpiar NaN/Inf antes de escribir
            df_data = _limpiar_df(df_data)

            df_data.to_excel(writer, sheet_name=nombre_hoja, index=False, startrow=1)
            ws = writer.sheets[nombre_hoja]

            # Encabezados formateados
            for col_idx, col_name in enumerate(df_data.columns):
                ws.write(0, col_idx, col_name, fmt_header)

            # Formato por columna
            for col_idx, col_name in enumerate(df_data.columns):
                width = 20
                if col_name in cols_num:
                    ws.set_column(col_idx, col_idx, width, fmt_miles)
                elif col_name == '% DOTACION':
                    ws.set_column(col_idx, col_idx, 12, fmt_pct)
                else:
                    ws.set_column(col_idx, col_idx, width, fmt_texto)

            # Resaltar filas de total
            if fila_total_marker and col_total_marker and col_total_marker in df_data.columns:
                for row_idx, val in enumerate(df_data[col_total_marker]):
                    if str(val).startswith(fila_total_marker):
                        # Fila Excel = row_idx + 2 (1 encabezado manual + 1 offset startrow)
                        for c_idx, c_name in enumerate(df_data.columns):
                            cell_val = df_data.iloc[row_idx][c_name]
                            if c_name in cols_num:
                                ws.write(row_idx + 2, c_idx, _safe_num(cell_val), fmt_total)
                            elif c_name == '% DOTACION':
                                ws.write(row_idx + 2, c_idx, _safe_num(cell_val), fmt_total)
                            else:
                                ws.write(row_idx + 2, c_idx, str(cell_val) if cell_val else '', fmt_total_txt)

            ws.freeze_panes(1, 0)

        # ===================================
        # HOJA PESOS
        # ===================================
        if not df_pesos_final.empty:
            totales_pesos = df_pesos_final[[c for c in cols_num if c in df_pesos_final.columns]].sum().to_frame().T
            totales_pesos['LINEA DE NEGOCIO']    = 'TOTAL GENERAL'
            totales_pesos['DENOMINACION COMERCIAL'] = ''
            totales_pesos['MONEDA']              = 'PESOS COL'
            for col in df_pesos_final.columns:
                if col not in totales_pesos.columns:
                    totales_pesos[col] = ''
            totales_pesos = totales_pesos[df_pesos_final.columns]

            df_pesos_salida = pd.concat([df_pesos_final, totales_pesos], ignore_index=True)
            _escribir_hoja(df_pesos_salida, 'PESOS', 'TOTAL GENERAL', 'LINEA DE NEGOCIO')
            print("  [OK] Hoja PESOS escrita")

        # ===================================
        # HOJA DIVISAS
        # ===================================
        if not df_divisas_final.empty:

            # Subtotales por moneda (valor original)
            tot_moneda = (
                df_divisas_final
                .groupby('MONEDA', as_index=False)[[c for c in cols_num if c in df_divisas_final.columns]]
                .sum()
            )
            tot_moneda['__TIPO__'] = 'TOTAL MONEDA'

            # Subtotales convertidos a COP x TRM
            def _convertir_cop(row):
                m = row.get('MONEDA', '')
                factor = trm_dolar if m == 'DÓLAR' else (trm_euro if m == 'EURO' else 1.0)
                r = row.copy()
                for c in cols_num:
                    if c in r.index:
                        r[c] = safe_float_conversion(r[c]) * factor
                return r

            tot_cop = tot_moneda.copy().apply(_convertir_cop, axis=1)
            tot_cop['__TIPO__'] = f'TOTAL MONEDA (COP) -- TRM USD:{trm_dolar:,.2f} | EUR:{trm_euro:,.2f}'

            df_divisas_salida = pd.concat(
                [df_divisas_final, tot_moneda, tot_cop], ignore_index=True
            )
            _escribir_hoja(df_divisas_salida, 'DIVISAS', 'TOTAL MONEDA', '__TIPO__')
            print("  [OK] Hoja DIVISAS escrita")

        # ===================================
        # HOJA VENCIMIENTO
        # ===================================
        if not df_vencimientos.empty:
            _escribir_hoja(df_vencimientos, 'VENCIMIENTO', 'TOTAL GENERAL', 'CLIENTE')
            print("  [OK] Hoja VENCIMIENTO escrita")

    print("\n" + "=" * 62)
    print("  MODELO DE DEUDA GENERADO EXITOSAMENTE")
    print(f"  Archivo : {output_path}")
    print(f"  TRM USD : ${trm_dolar:,.4f}  |  TRM EUR: ${trm_euro:,.4f}  ({fecha_trm})")
    print("=" * 62 + "\n")

    logging.info(f"Modelo de deuda generado: {output_path}")
    if USE_UNIFIED_LOGGING:
        log_fin_proceso("MODELO_DEUDA", output_file)

    return output_path

# ============================================================
# CLI
# ============================================================
def main():
    parser = argparse.ArgumentParser(
        description="Genera Modelo de Deuda -- Procedimiento Departamento de Cartera"
    )
    parser.add_argument("archivo_provision",
                        help="Archivo de provisión PISA (provca.csv o .xlsx)")
    parser.add_argument("archivo_anticipos",
                        help="Archivo de anticipos PISA (clanti.csv o .xlsx)")
    parser.add_argument("-o", "--output-file",
                        help="Nombre del archivo de salida (*.xlsx)",
                        default=None)
    parser.add_argument("--usd", type=float,
                        help="TRM USD override (ej: 4350.50). Ignora trm.json para USD.")
    parser.add_argument("--eur", type=float,
                        help="TRM EUR override (ej: 4712.80). Ignora trm.json para EUR.")
    args = parser.parse_args()

    if not args.output_file:
        ts = datetime.now().strftime("%Y%m%d_%H%M%S")
        args.output_file = f"MODELO_DEUDA_{ts}.xlsx"

    try:
        crear_modelo_deuda(
            args.archivo_provision,
            args.archivo_anticipos,
            args.output_file,
            usd_override=args.usd,
            eur_override=args.eur,
        )
    except Exception as e:
        if USE_UNIFIED_LOGGING:
            log_error_proceso("MODELO_DEUDA", str(e))
        print(f"\n[ERROR] {e}")
        logging.error(f"Error en modelo de deuda: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()