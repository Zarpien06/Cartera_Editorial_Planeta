# ==========================================================
# MODELO DE DEUDA - Cumplimiento completo procedimiento
# Departamento de Cartera - Editorial Planeta Colombia
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

# --------------------------------------------------
# ORDEN DE COLUMNAS PARA MODELO DE DEUDA
# --------------------------------------------------
ORDEN_COLUMNAS_MODELO = [
    'LINEA DE NEGOCIO',
    'CODIGO CLIENTE',
    'IDENTIFICACION',
    'DENOMINACION COMERCIAL',
    'DIRECCION',
    'TELEFONO',
    'CIUDAD',
    'NUMERO FACTURA',
    'TIPO',
    'FECHA',
    'FECHA VTO',

    'VALOR',
    'SALDO',
    'SALDO NO VENCIDO',
    'SALDO VENCIDO',
    'DIAS VENCIDOS',
    'DIAS POR VENCER',
    '% DOTACION',
    'DEUDA INCOBRABLE',

    'VENCIDO 30',
    'VENCIDO 60',
    'VENCIDO 90',
    'VENCIDO 180',
    'VENCIDO 360',
    'VENCIDO + 360',
    'PROVISION'
]
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

def _build_linea_key(emp, act):
    emp = str(emp).strip().upper()
    try:
        act = str(int(float(act)))
    except:
        act = str(act).strip()
    act = act.lstrip('0') or '0'  
    return f"{emp}{act}"

def _ensure_datetime(series):
    try:
        return pd.to_datetime(series, dayfirst=True, errors='coerce')
    except Exception:
        return pd.to_datetime(series, errors='coerce')

def excluir_pl16_pl68(df: pd.DataFrame, nombre_df: str) -> pd.DataFrame:
    """Excluye PL16 y PL68 con validación inmediata"""
    
    before = len(df)
    pl16_before = len(df[df['LINEA DE NEGOCIO'] == 'PL16'])
    pl68_before = len(df[df['LINEA DE NEGOCIO'] == 'PL68'])
    
    # Excluir
    df = df[~df['LINEA DE NEGOCIO'].isin(['PL16', 'PL68'])].copy()
    after = len(df)
    excluidos = before - after
    
    # Validar después
    pl16_after = len(df[df['LINEA DE NEGOCIO'] == 'PL16'])
    pl68_after = len(df[df['LINEA DE NEGOCIO'] == 'PL68'])
    
    print(f"\n  [{nombre_df}] Exclusión de PL16/PL68")
    print(f"    Antes: {before:,} registros (PL16={pl16_before}, PL68={pl68_before})")
    print(f"    Excluidos: {excluidos}")
    print(f"    Después: {after:,} registros")
    
    # VALIDACIÓN CRÍTICA
    if pl16_after > 0 or pl68_after > 0:
        raise ValueError(
            f"[ERROR] CRÍTICO en {nombre_df}: "
            f"PL16={pl16_after}, PL68={pl68_after} TODAVÍA PRESENTES"
        )
    
    print(f" [OK] - PL16 y PL68 eliminados correctamente")
    return df 
    
# --------------------------------------------------
# ORDENAR COLUMNAS MODELO DE DEUDA
# --------------------------------------------------
def ordenar_columnas_modelo(df):
    columnas_existentes = [c for c in ORDEN_COLUMNAS_MODELO if c in df.columns]
    otras_columnas = [c for c in df.columns if c not in columnas_existentes]
    return df[columnas_existentes + otras_columnas]

# ---------------- Lectura flexible de archivos ----------------
def leer_archivo(archivo: str) -> pd.DataFrame:
    if not os.path.exists(archivo):
        raise FileNotFoundError(f"No se encontró el archivo: {archivo}")
    ext = archivo.lower().split('.')[-1]
    if ext == 'csv':
        for enc in ['latin1', 'cp1252', 'utf-8-sig', 'utf-8', 'iso-8859-1']:
            for sep in [';','|',',','\t']:
                try:
                    df = pd.read_csv(archivo, encoding=enc, sep=sep)
                    if df.shape[1] > 1:
                        return df
                except:
                    pass
        raise ValueError(f"No se pudo leer el CSV: {archivo}")
    elif ext in ['xlsx','xls']:
      # Para Excel, buscar la hoja con datos
      xls = pd.ExcelFile(archivo)
      hojas = xls.sheet_names
      
      print(f"  [INFO] Hojas encontradas: {hojas}")
      
      # Intentar hojas en orden de prioridad
      hojas_prioridad = ['DETALLE', 'CARTERA', 'DATA', 'Sheet1', 'Hoja1']
      
      for hoja_preferida in hojas_prioridad:
          if hoja_preferida in hojas:
              print(f"  [OK] Usando hoja: {hoja_preferida}")
              return pd.read_excel(archivo, sheet_name=hoja_preferida, engine='openpyxl')
      
      # Si no encuentra, usa la primera hoja CON DATOS
      for hoja in hojas:
          try:
              df = pd.read_excel(archivo, sheet_name=hoja, engine='openpyxl')
              if len(df) > 0 and len(df.columns) > 1:
                  print(f"  [OK] Usando hoja: {hoja}")
                  return df
          except:
              continue
      
      # Fallbac
      return pd.read_excel(archivo, sheet_name=0, engine='openpyxl')
  
    else:
        raise ValueError(f"Formato no soportado: {archivo}")

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
    # Limpiar nombres de columnas
    df.columns = (
       df.columns
        .str.replace('\n', ' ', regex=True)
        .str.replace('\r', ' ', regex=True)
        .str.strip()
    )
        
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
        df['FECHA'] = pd.to_datetime(df['FECHA'], dayfirst=True, errors='coerce')
    
    if 'FECHA VTO' in df.columns:
        df['FECHA VTO'] = pd.to_datetime(df['FECHA VTO'], dayfirst=True, errors='coerce')
    
    # -- 4. Abrir FECHA VTO en día, mes, año --
    # ELIMINADO: No se necesitan columnas separadas VTO_DIA, VTO_MES, VTO_ANIO
    
    # -- 5. Línea de negocio --
    if 'EMPRESA' in df.columns and 'ACTIVIDAD' in df.columns:
        df['LINEA DE NEGOCIO'] = df.apply(
            lambda r: _build_linea_key(r['EMPRESA'], r['ACTIVIDAD']), axis=1
        )
    elif 'PCCDEM' in df.columns and 'PCCDAC' in df.columns:
        # Si vienen con nombres PISA, renombra primero
        df['EMPRESA'] = df['PCCDEM']
        df['ACTIVIDAD'] = df['PCCDAC']
        df['LINEA DE NEGOCIO'] = df.apply(
            lambda r: _build_linea_key(r['EMPRESA'], r['ACTIVIDAD']), axis=1
        )
    else:
        # Fallback: crear columna vacía
        print("  [WARN] No encontró EMPRESA/ACTIVIDAD o PCCDEM/PCCDAC")
        df['LINEA DE NEGOCIO'] = 'SIN_CLASIFICAR'
        
    # -- 6. Eliminar fila específica PL30 --
    try:
        if 'ACTIVIDAD' in df.columns and 'SALDO' in df.columns:
            df['__SAL_NUM__'] = pd.to_numeric(df['SALDO'], errors='coerce')
            cond = (
                df['ACTIVIDAD'].astype(str).str.strip() == '30'
            ) & (
                df['__SAL_NUM__'].round(0) == -614000
            )
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
    # Convertir días vencidos a número
    df['DIAS VENCIDOS'] = pd.to_numeric(df['DIAS VENCIDOS'], errors='coerce').fillna(0)
    
    # Detectar anticipos
    if 'TIPO' in df.columns:
     es_anticipo = df['TIPO'].astype(str).str.upper().str.contains('ANTICIPO', na=False)
    else:
     es_anticipo = pd.Series(False, index=df.index)
    
    # Calcular saldo vencido
    df['SALDO VENCIDO'] = np.where(
        (~es_anticipo) & (df['DIAS VENCIDOS'] > 0),
        df['SALDO'],
        0.0
    )
     
    # -- 11. % DOTACION --
    if '% DOTACION' not in df.columns:
        df['% DOTACION'] = 0.0
    else:
        df['% DOTACION'] = pd.to_numeric(df['% DOTACION'], errors='coerce').fillna(0.0)
    
    # -- 12. DEUDA INCOBRABLE = SALDO * % DOTACION --
    df['DEUDA INCOBRABLE'] = (df['SALDO'] * df['% DOTACION']).round(2)
    
    # PROVISION
    df['PROVISION'] = df['DEUDA INCOBRABLE']
              
    # Detectar anticipos si existen en provisión
    if 'TIPO' in df.columns:
        tipo_upper = df['TIPO'].astype(str).str.upper()
    
        es_anticipo = (
            tipo_upper.str.contains('ANT', na=False) |
            tipo_upper.str.contains('ANTIC', na=False)
        )
    else:
        es_anticipo = pd.Series(False, index=df.index)
            
    # -- 13. Siete buckets de vencimiento (procedimiento pág. 3) --
    df['SALDO NO VENCIDO'] = np.where(
        es_anticipo,
        df['SALDO'],  # ← MANTENER NEGATIVO (sin .abs())
        np.where(df['DIAS VENCIDOS'] <= 0, df['SALDO'], 0.0)
    )

    # VENCIDO 30: Entre 1 y 29 días vencidos
    df['VENCIDO 30'] = np.where(
        (~es_anticipo) & (df['DIAS VENCIDOS'].between(1, 29)),  # 1-29
        df['SALDO'],
        0.0
    )
    
    # VENCIDO 60: Entre 30 y 59 días vencidos
    df['VENCIDO 60'] = np.where(
        (~es_anticipo) & (df['DIAS VENCIDOS'].between(30, 59)),  # 30-59
        df['SALDO'],
        0.0
    )
    
    # VENCIDO 90: Entre 60 y 89 días vencidos
    df['VENCIDO 90'] = np.where(
        (~es_anticipo) & (df['DIAS VENCIDOS'].between(60, 89)),  # 60-89
        df['SALDO'],
        0.0
    )
    
    # VENCIDO 180: Entre 90 y 179 días vencidos
    df['VENCIDO 180'] = np.where(
        (~es_anticipo) & (df['DIAS VENCIDOS'].between(90, 179)),  # 90-179
        df['SALDO'],
        0.0
    )
    
    # VENCIDO 360: Entre 180 y 359 días vencidos
    df['VENCIDO 360'] = np.where(
        (~es_anticipo) & (df['DIAS VENCIDOS'].between(180, 359)),  # 180-359
        df['SALDO'],
        0.0
    )
    
    # VENCIDO + 360: 360 o más días vencidos
    df['VENCIDO + 360'] = np.where(
        (~es_anticipo) & (df['DIAS VENCIDOS'] >= 360),  # >= 360
        df['SALDO'],
        0.0
    )
        
    # Anticipos nunca pueden ir a vencimientos
    df.loc[es_anticipo, [
        'VENCIDO 30',
        'VENCIDO 60',
        'VENCIDO 90',
        'VENCIDO 180',
        'VENCIDO 360',
        'VENCIDO + 360'
    ]] = 0
        
    # -- 14. Validación: suma de buckets == SALDO --
    # Validación 1: Suma de buckets == SALDO
    suma_buckets = (
        df['SALDO NO VENCIDO'] + df['VENCIDO 30'] + df['VENCIDO 60'] +
        df['VENCIDO 90'] + df['VENCIDO 180'] + df['VENCIDO 360'] + df['VENCIDO + 360']
    ).round(2)
    
    mismatches_buckets = (
        ((suma_buckets - df['SALDO'].round(2)).abs() > 0.01) &
        (df['SALDO'] >= 0)
    ).sum()
    
    if mismatches_buckets:
        print(f"  [WARN] {mismatches_buckets} factura(s): suma(buckets) != SALDO")
    else:
        print("  [OK] Validacion buckets: suma(buckets) = SALDO en todas las facturas")
    
    # Validación 2: MORA TOTAL + TOTAL POR VENCER == SALDO
    mismatches_mora = (
        ((df['MORA TOTAL'] + df['TOTAL POR VENCER']).round(2) - df['SALDO'].round(2)).abs() > 0.01
    ).sum()
    
    if mismatches_mora:
        print(f"  [WARN] {mismatches_mora} factura(s): MORA TOTAL + POR VENCER != SALDO")
    else:
        print("  [OK] Validacion mora: MORA TOTAL + TOTAL POR VENCER = SALDO en todas las facturas")
    
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

    # -- 17 y 18: ELIMINADAS
    # No se crean columnas POR VENCER M1, M2, M3 ni MAYOR 90 DIAS
    
    return df

# ============================================================
# HOJA VENCIMIENTO
# ============================================================
def crear_hoja_vencimientos(df_pesos_final: pd.DataFrame,
                            df_divisas_final: pd.DataFrame) -> pd.DataFrame:
    
    # =====================================================
    # ASEGURAR COLUMNA SALDO VENCIDO
    # =====================================================
    columnas_venc = [
        'VENCIDO 30',
        'VENCIDO 60',
        'VENCIDO 90',
        'VENCIDO 180',
        'VENCIDO 360',
        'VENCIDO + 360'
    ]
    
    for df_tmp in (df_pesos_final, df_divisas_final):
    
        for col in columnas_venc:
            if col not in df_tmp.columns:
                df_tmp[col] = 0
    
        if 'SALDO VENCIDO' not in df_tmp.columns:
            df_tmp['SALDO VENCIDO'] = (
                df_tmp['VENCIDO 30'] +
                df_tmp['VENCIDO 60'] +
                df_tmp['VENCIDO 90'] +
                df_tmp['VENCIDO 180'] +
                df_tmp['VENCIDO 360'] +
                df_tmp['VENCIDO + 360']
            )
            
    columnas_sumables = [
        'SALDO',
        'SALDO NO VENCIDO',
        'SALDO VENCIDO',
        'VENCIDO 30',
        'VENCIDO 60',
        'VENCIDO 90',
        'VENCIDO 180',
        'VENCIDO 360',
        'VENCIDO + 360',
        'DEUDA INCOBRABLE'
    ]

    df_all = pd.concat([df_pesos_final, df_divisas_final], ignore_index=True)

    if 'LINEA DE NEGOCIO' in df_all.columns:
        df_all['NEGOCIO'] = df_all['LINEA DE NEGOCIO'].apply(
            lambda k: TABLA_NEGOCIO_CANAL.get(
                str(k).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'}
            )['NEGOCIO']
        )
        # CANAL = solo número de actividad
        df_all['CANAL'] = df_all['LINEA DE NEGOCIO'].apply(
        lambda k: str(k).strip().upper()[2:] if isinstance(k, str) and len(str(k)) > 2 else ''
        )
        df_all['MONEDA'] = df_all['LINEA DE NEGOCIO'].apply(
            lambda k: _moneda_por_linea(str(k).strip().upper())
        )

    df_all['MONEDA'] = df_all.get('MONEDA', 'PESOS COL').astype(str).str.strip()
    df_all['CLIENTE'] = df_all.get('DENOMINACION COMERCIAL', '').astype(str).str.strip()
    df_all['CLIENTE'] = df_all['CLIENTE'].replace('', 'ANTICIPO SIN CLIENTE')
    df_all['CLIENTE'] = df_all['CLIENTE'].fillna('ANTICIPO SIN CLIENTE')
    
    # Limpiar espacios en TODAS las columnas de agrupamiento
    df_all['NEGOCIO'] = df_all['NEGOCIO'].astype(str).str.strip()
    df_all['CANAL'] = df_all['CANAL'].astype(str).str.strip()
    df_all['MONEDA'] = df_all['MONEDA'].astype(str).str.strip()
    df_all['CLIENTE'] = df_all['CLIENTE'].astype(str).str.strip()
    df_all['CLIENTE'] = (
    df_all['CLIENTE']
    .astype(str)
    .str.strip()
    .str.upper()
    )
     
    grp_cols = ['NEGOCIO', 'CANAL', 'MONEDA', 'CLIENTE']
    for c in grp_cols:
        if c not in df_all.columns:
            df_all[c] = ''
    # Ajustar anticipos para que no resten saldo total
    if 'TIPO' in df_all.columns:
        es_anticipo = df_all['TIPO'].astype(str).str.upper().str.contains('ANT', na=False)
    else:
        es_anticipo = pd.Series(False, index=df_all.index)
        
    # SALDO NO VENCIDO siempre positivo
    df_all.loc[es_anticipo, 'SALDO NO VENCIDO'] = df_all.loc[es_anticipo, 'SALDO']  # mantener negativo
    df_all.loc[es_anticipo, 'SALDO VENCIDO']    = 0.0
    df_all.loc[es_anticipo, ['VENCIDO 30','VENCIDO 60','VENCIDO 90',
                              'VENCIDO 180','VENCIDO 360','VENCIDO + 360']] = 0.0
    
    # SALDO VENCIDO siempre cero
    df_all.loc[es_anticipo, 'SALDO VENCIDO'] = 0
        
    df_sum = (
        df_all
        .groupby(grp_cols, as_index=False)[columnas_sumables]
        .sum()
    )

    df_sum.insert(0, 'Pais',       'COLOMBIA')  # ← 'PAIS' → 'Pais'
    df_sum.insert(3, 'COBRO/PAGO', 'CLIENTE')
    
    # RENOMBRAR CON ESPACIOS AL INICIO
    df_sum = df_sum.rename(columns={
        'SALDO': 'SALDO TOTAL',          
        'SALDO NO VENCIDO': 'Saldo No vencido',
        'SALDO VENCIDO': 'Saldo Vencido',
        'VENCIDO 30': 'Vencido 30',
        'VENCIDO 60': 'Vencido 60',
        'VENCIDO 90': 'Vencido 90',
        'VENCIDO 180': 'Vencido 180',
        'VENCIDO 360': 'Vencido 360',
        'VENCIDO + 360': 'Vencido + 360',
        'DEUDA INCOBRABLE': 'DEUDA INCOBRABLE'
    })
    
    df_sum['SALDO TOTAL'] = df_sum['Saldo No vencido'] + df_sum['Saldo Vencido']
    
    # Totales generales por moneda (al final, según procedimiento pág. 6/8)
    totales_moneda = (
        df_sum
        .groupby('MONEDA', as_index=False)[[
            'SALDO TOTAL',
            'Saldo No vencido',
            'Saldo Vencido',
            'Vencido 30',
            'Vencido 60',
            'Vencido 90',
            'Vencido 180',
            'Vencido 360',
            'Vencido + 360',
            'DEUDA INCOBRABLE'
        ]]
        .sum()
    )
 
    totales_moneda['SALDO TOTAL'] = (
        totales_moneda['Saldo No vencido'] + totales_moneda['Saldo Vencido']
    )
        
    totales_moneda['Pais'] = 'COLOMBIA'  # ← 'PAIS' → 'Pais' también
    totales_moneda['NEGOCIO']    = 'TOTAL'
    totales_moneda['CANAL']      = 'TOTAL'
    totales_moneda['COBRO/PAGO'] = ''
    totales_moneda['CLIENTE']    = 'TOTAL GENERAL POR MONEDA'
    totales_moneda = totales_moneda[df_sum.columns]

    df_final = pd.concat([df_sum, totales_moneda], ignore_index=True)
    return df_final

def crear_hoja_usd_euro_vencimientos(df_pesos_final: pd.DataFrame,
                                         df_divisas_cop: pd.DataFrame) -> pd.DataFrame:
        """Crea vencimientos SOLO con divisas (USD y EUR) - valores originales"""
        
        columnas_sumables = [
            'SALDO',
            'SALDO NO VENCIDO',
            'SALDO VENCIDO',
            'VENCIDO 30',
            'VENCIDO 60',
            'VENCIDO 90',
            'VENCIDO 180',
            'VENCIDO 360',
            'VENCIDO + 360',
            'DEUDA INCOBRABLE'
        ]
    
        df_all = df_divisas_cop.copy() # ← SOLO DIVISAS, no pesos
    
        if 'LINEA DE NEGOCIO' in df_all.columns:
            df_all['NEGOCIO'] = df_all['LINEA DE NEGOCIO'].apply(
                lambda k: TABLA_NEGOCIO_CANAL.get(
                    str(k).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'}
                )['NEGOCIO']
            )
            df_all['CANAL'] = df_all['LINEA DE NEGOCIO'].apply(
                lambda k: str(k).strip().upper()[2:] if isinstance(k, str) and len(str(k)) > 2 else ''
            )
            df_all['MONEDA'] = df_all['LINEA DE NEGOCIO'].apply(
                lambda k: _moneda_por_linea(str(k).strip().upper())
            )
    
        df_all['MONEDA'] = df_all.get('MONEDA', 'PESOS COL').astype(str).str.strip()
        df_all['CLIENTE'] = df_all.get('DENOMINACION COMERCIAL', '').astype(str).str.strip()
        
        df_all['NEGOCIO'] = df_all['NEGOCIO'].astype(str).str.strip()
        df_all['CANAL'] = df_all['CANAL'].astype(str).str.strip()
        df_all['MONEDA'] = df_all['MONEDA'].astype(str).str.strip()
        df_all['CLIENTE'] = df_all['CLIENTE'].astype(str).str.strip()
         
        grp_cols = ['NEGOCIO', 'CANAL', 'MONEDA', 'CLIENTE']
        for c in grp_cols:
            if c not in df_all.columns:
                df_all[c] = ''
    
        df_sum = (
            df_all
            .groupby(grp_cols, as_index=False)[columnas_sumables]
            .sum()
        )
    
        df_sum.insert(0, 'Pais',       'COLOMBIA')
        df_sum.insert(3, 'COBRO/PAGO', 'CLIENTE')
        
        df_sum = df_sum.rename(columns={
            'SALDO': 'SALDO TOTAL',            
            'SALDO NO VENCIDO': 'Saldo No vencido',
            'SALDO VENCIDO': 'Saldo Vencido',
            'VENCIDO 30': 'Vencido 30',
            'VENCIDO 60': 'Vencido 60',
            'VENCIDO 90': 'Vencido 90',
            'VENCIDO 180': 'Vencido 180',
            'VENCIDO 360': 'Vencido 360',
            'VENCIDO + 360': 'Vencido + 360',
            'DEUDA INCOBRABLE': 'DEUDA INCOBRABLE'
        })
        
        df_sum['SALDO TOTAL'] = df_sum['Saldo No vencido'] + df_sum['Saldo Vencido']
        
        totales_moneda = (
            df_sum.groupby('MONEDA', as_index=False)[[
                'SALDO TOTAL',        
                'Saldo No vencido',
                'Saldo Vencido',
                'Vencido 30',
                'Vencido 60',
                'Vencido 90',
                'Vencido 180',
                'Vencido 360',
                'Vencido + 360',
                'DEUDA INCOBRABLE'
            ]].sum()
        )
        
        totales_moneda['SALDO TOTAL'] = (
            totales_moneda['Saldo No vencido'] + totales_moneda['Saldo Vencido']
        )
        
        totales_moneda['Pais'] = 'COLOMBIA'
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
    
    # --- Limpiar filas que sean header duplicado ---
    def limpiar_headers_duplicados(df):
        columnas = list(df.columns)
        return df[
            df.apply(
                lambda row: list(row.astype(str).str.strip()) != columnas,
                axis=1
            )
        ].reset_index(drop=True)
    
    df_provision_raw = limpiar_headers_duplicados(df_provision_raw)
    df_anticipos_raw = limpiar_headers_duplicados(df_anticipos_raw)
    
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
    
    # EXCLUSIÓN 1: PROVISIÓN
    df_provision = excluir_pl16_pl68(df_provision, "PROVISIÓN")
   
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
    # Asegurar identificación en anticipos
    if 'IDENTIFICACION' not in df_anticipos.columns:
        df_anticipos['IDENTIFICACION'] = ''
    
    df_anticipos['IDENTIFICACION'] = (
        df_anticipos['IDENTIFICACION']
        .astype(str)
        .str.strip()
    )
    # Si viene vacío usar código cliente
    df_anticipos['IDENTIFICACION'] = df_anticipos['IDENTIFICACION'].replace('', np.nan)
    df_anticipos['IDENTIFICACION'] = df_anticipos['IDENTIFICACION'].fillna(df_anticipos['CODIGO CLIENTE'])
    
    if 'EMPRESA' in df_anticipos.columns and 'ACTIVIDAD' in df_anticipos.columns:
        df_anticipos['LINEA DE NEGOCIO'] = df_anticipos.apply(
            lambda r: _build_linea_key(r['EMPRESA'], r['ACTIVIDAD']), axis=1
        )
    # EXCLUSIÓN 2: ANTICIPOS (CRÍTICA)
    df_anticipos = excluir_pl16_pl68(df_anticipos, "ANTICIPOS")

    # Valor anticipo * -1 (debe ser negativo, procedimiento pág. 4)
    if 'VALOR ANTICIPO' in df_anticipos.columns:
        valor_ant = pd.to_numeric(df_anticipos['VALOR ANTICIPO'], errors='coerce').fillna(0.0)
    
        # Forzar negativo siempre (aunque venga negativo o positivo)
        valor_ant = -abs(valor_ant)
    
        df_anticipos['VALOR'] = valor_ant
        df_anticipos['SALDO'] = valor_ant
    else:
        df_anticipos['VALOR'] = 0.0
        df_anticipos['SALDO'] = 0.0

    # Normalizar columna DENOMINACION COMERCIAL en anticipos
    if 'DENOMINACION COMERCIAL' not in df_anticipos.columns:
        df_anticipos['DENOMINACION COMERCIAL'] = 'ANTICIPO'
    else:
        df_anticipos['DENOMINACION COMERCIAL'] = df_anticipos['DENOMINACION COMERCIAL'].fillna('ANTICIPO')

    df_anticipos['FECHA'] = _ensure_datetime(df_anticipos.get('FECHA', pd.NaT))

    # Fecha de corte desde provisión
    if 'FECHA' in df_provision.columns and df_provision['FECHA'].notna().any():
        fecha_corte = _last_day_of_month(df_provision['FECHA'].max())
    else:
        fecha_corte = _last_day_of_month(pd.Timestamp.today())

    df_anticipos['FECHA VTO'] = pd.NaT
    
    # =====================================================
    # ANTICIPOS NO PUEDEN TENER VENCIMIENTO
    # =====================================================
    df_anticipos['DIAS VENCIDOS']   = 0
    df_anticipos['DIAS POR VENCER'] = 0
    
    # Forzar tipo entero (evita 0.0 en Excel)
    df_anticipos['DIAS VENCIDOS']   = df_anticipos['DIAS VENCIDOS'].astype(int)
    df_anticipos['DIAS POR VENCER'] = df_anticipos['DIAS POR VENCER'].astype(int)

    # Anticipos: van en SALDO NO VENCIDO (procedimiento pág. 5)
    df_anticipos['SALDO VENCIDO']   = 0.0
    df_anticipos['SALDO NO VENCIDO'] = pd.to_numeric(
      df_anticipos['SALDO'], errors='coerce'
    ).fillna(0)
    df_anticipos['% DOTACION']      = 0.0
    df_anticipos['DEUDA INCOBRABLE']= 0.0
    for col in ['VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90',
                'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360']:
        df_anticipos[col] = 0.0
    df_anticipos['MORA TOTAL']      = 0.0
    df_anticipos['TOTAL POR VENCER'] = df_anticipos['SALDO']
    
    # Columnas numéricas esperadas en el modelo
    columnas_numericas_modelo = [
        'SALDO', 'VALOR', 'SALDO VENCIDO', '% DOTACION',
        'DEUDA INCOBRABLE', 'SALDO NO VENCIDO',
        'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90',
        'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360',
        'MORA TOTAL', 'TOTAL POR VENCER'
    ]

    for col in columnas_numericas_modelo:
        if col not in df_anticipos.columns:
            df_anticipos[col] = 0.0
        else:
            df_anticipos[col] = pd.to_numeric(
                df_anticipos[col], errors='coerce'
            ).fillna(0.0)

    # Asegurar enteros donde corresponde
    df_anticipos['DIAS VENCIDOS']   = df_anticipos['DIAS VENCIDOS'].fillna(0).astype(int)
    df_anticipos['DIAS POR VENCER'] = df_anticipos['DIAS POR VENCER'].fillna(0).astype(int)
    
    print(f"  [OK] {len(df_anticipos):,} anticipos procesados")

    # -- PASO 5: Separar PESOS y DIVISAS (procedimientos pág. 5-6) --
    
    lineas_pesos_keys   = {f"{cod}{act}" for cod, act in LINEAS_PESOS}
    lineas_divisas_keys = {f"{cod}{act}" for cod, act in LINEAS_DIVISAS}
    
    # 1️⃣ primero crear los dataframes
    df_pesos   = df_provision[df_provision['LINEA DE NEGOCIO'].isin(lineas_pesos_keys)].copy()
    df_divisas = df_provision[df_provision['LINEA DE NEGOCIO'].isin(lineas_divisas_keys)].copy()
    
    print("\nLineas en PESOS:")
    print(df_pesos['LINEA DE NEGOCIO'].value_counts())
    print("\nLineas en DIVISAS:")
    print(df_divisas['LINEA DE NEGOCIO'].value_counts())
    
    # 2️⃣ moneda
    df_pesos['MONEDA']   = 'PESOS COL'
    df_divisas['MONEDA'] = df_divisas['LINEA DE NEGOCIO'].apply(_moneda_por_linea)
    
    # 3️⃣ ordenar columnas
    df_pesos = ordenar_columnas_modelo(df_pesos)
    df_divisas = ordenar_columnas_modelo(df_divisas)
    

    # -------------------------------------------------------
    # NORMALIZAR ESTRUCTURA DE ANTICIPOS (CLAVE PARA CONCAT FINAL )
    # -------------------------------------------------------
    
    # Unificar nombre agente
    if 'NOMBRE AGENTE' in df_anticipos.columns:
        df_anticipos['AGENTE'] = (
            df_anticipos.get('NOMBRE AGENTE', '').astype(str) + ' ' +
            df_anticipos.get('APELLIDO AGENTE', '').astype(str)
        ).str.strip()
    
    # Unificar tipo
    df_anticipos['TIPO'] = 'ANTICIPO'
    
    # Si no existe NUMERO FACTURA, usar NUMERO ANTICIPO
    if 'NUMERO ANTICIPO' in df_anticipos.columns:
        df_anticipos['NUMERO FACTURA'] = df_anticipos['NUMERO ANTICIPO']
     
    # -------------------------------------------------------
    # SEPARAR ANTICIPOS POR MONEDA (DESPUÉS DE NORMALIZAR)
    # -------------------------------------------------------
    
    df_anticipos['MONEDA'] = df_anticipos['LINEA DE NEGOCIO'].apply(_moneda_por_linea)

    # SOLUCIÓN
    df_anticipos['MONEDA'] = df_anticipos['MONEDA'].fillna('PESOS COL')
    ant_div   = df_anticipos[df_anticipos['MONEDA'] != 'PESOS COL'].copy() 
    ant_pesos = df_anticipos[df_anticipos['MONEDA'] == 'PESOS COL'].copy()
          
    # -------------------------------------------------------
    # COLUMNAS OFICIALES DEL MODELO (ESTRUCTURA CONTROLADA)
    # -------------------------------------------------------
    
    columnas_modelo = [
        'LINEA DE NEGOCIO',
        'CODIGO AGENTE', 'AGENTE',
        'CODIGO CLIENTE', 'IDENTIFICACION',
        'DENOMINACION COMERCIAL', 'DIRECCION',
        'TELEFONO', 'CIUDAD',
        'NUMERO FACTURA', 'TIPO',
        'FECHA', 'FECHA VTO',
        'VALOR',  'SALDO',
        'SALDO VENCIDO',  'DIAS VENCIDOS',
        'DIAS POR VENCER',  '% DOTACION',
        'SALDO NO VENCIDO',
        'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90',
        'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360',
        'MONEDA', 'DEUDA INCOBRABLE'
    ]
        
   # ==========================================================
    # BLINDAJE ESTRUCTURAL CONTRA COLUMNAS FALTANTES
    # ==========================================================
    
    def blindar_columnas(df, columnas_modelo, nombre_df):
     """
     Garantiza estructura exacta SIN ERRORES
     - Crea columnas faltantes
     - Elimina columnas extras
     - Respeta el orden exacto
     """
     
     # 1️⃣ Crear columnas faltantes
     for col in columnas_modelo:
         if col not in df.columns:
             df[col] = None
     
     # 2️⃣ Mantener estructura sin perder datos
     df = df.reindex(columns=columnas_modelo)
     
     # 3️⃣ Reordenar EXACTAMENTE como el modelo
     df = df[columnas_modelo]
     
     # 4️⃣ Rellenar NaN solo en numéricas
     for col in columnas_modelo:
         if col in df.columns:
             try:
                 if df[col].dtype in ['float64', 'int64', 'Int64', 'Float64']:
                     df[col] = df[col].fillna(0)
             except:
                 pass
     
     return df
    
    # ==========================================================
    # APLICACIÓN DEL BLINDAJE
    # ==========================================================
    
    columnas_modelo_sin_moneda = [c for c in columnas_modelo if c != 'MONEDA']
    
    # Luego en blindar_columnas para PESOS
    df_pesos   = blindar_columnas(df_pesos, columnas_modelo_sin_moneda, "df_pesos")
    ant_pesos  = blindar_columnas(ant_pesos, columnas_modelo_sin_moneda, "ant_pesos")
    
    # Y para DIVISAS sí mantener MONEDA
    df_divisas = blindar_columnas(df_divisas, columnas_modelo, "df_divisas")
    ant_div    = blindar_columnas(ant_div, columnas_modelo, "ant_div")
    
    print("ANTICIPOS PESOS:", ant_pesos['SALDO'].sum())
    print("ANTICIPOS DIVISAS:", ant_div['SALDO'].sum())
    
    # ==========================================================
    # CONCAT FINAL SEGURO
    # ==========================================================
    
    df_pesos_final   = pd.concat([df_pesos, ant_pesos], ignore_index=True)
    df_divisas_final = pd.concat([df_divisas, ant_div], ignore_index=True)
    
    # =====================================================
    # REPARAR ANTICIPOS DESPUES DEL CONCAT
    # =====================================================
    
    for dfX in (df_pesos_final, df_divisas_final):

      es_anticipo = dfX['TIPO'].astype(str).str.upper().str.contains('ANT', na=False)
  
      # anticipos
      dfX.loc[es_anticipo, 'SALDO NO VENCIDO'] = dfX.loc[es_anticipo, 'SALDO']
      dfX.loc[es_anticipo, 'SALDO VENCIDO'] = 0
  
      # facturas normales
      dfX.loc[~es_anticipo, 'SALDO VENCIDO'] = (
          dfX.loc[~es_anticipo, [
              'VENCIDO 30',
              'VENCIDO 60',
              'VENCIDO 90',
              'VENCIDO 180',
              'VENCIDO 360',
              'VENCIDO + 360'
          ]].sum(axis=1)
      )
      
    # ==========================================================
    # VALIDACIÓN ESTRUCTURAL CRÍTICA
    # ==========================================================
    
    # Validación flexible: validar que TENGAN las columnas necesarias
    columnas_pesos_reales = set(df_pesos_final.columns)
    columnas_pesos_esperadas = set(columnas_modelo_sin_moneda)
    
    columnas_divisas_reales = set(df_divisas_final.columns)
    columnas_divisas_esperadas = set(columnas_modelo)
    
    # PESOS
    if columnas_pesos_reales != columnas_pesos_esperadas:
        faltantes = columnas_pesos_esperadas - columnas_pesos_reales
        extras = columnas_pesos_reales - columnas_pesos_esperadas
        if faltantes:
            print(f"[WARN] PESOS: Columnas FALTANTES -> {faltantes}")
            for col in faltantes:
                df_pesos_final[col] = None
        if extras:
            print(f"[WARN] PESOS: Columnas EXTRA (eliminadas) -> {extras}")
            df_pesos_final = df_pesos_final.drop(columns=extras)
        # Reordenar
        df_pesos_final = df_pesos_final[columnas_modelo_sin_moneda]
    
    # DIVISAS
    if columnas_divisas_reales != columnas_divisas_esperadas:
        faltantes = columnas_divisas_esperadas - columnas_divisas_reales
        extras = columnas_divisas_reales - columnas_divisas_esperadas
        if faltantes:
            print(f"[WARN] DIVISAS: Columnas FALTANTES -> {faltantes}")
            for col in faltantes:
                df_divisas_final[col] = None
        if extras:
            print(f"[WARN] DIVISAS: Columnas EXTRA (eliminadas) -> {extras}")
            df_divisas_final = df_divisas_final.drop(columns=extras)
        # Reordenar
        df_divisas_final = df_divisas_final[columnas_modelo]
    
    print("[OK] Blindaje estructural completado correctamente")

    for dfX in (df_pesos_final, df_divisas_final):

     # Si es anticipo (saldo negativo), no debe tener días
     dfX.loc[dfX['SALDO'] < 0, 'DIAS VENCIDOS'] = 0
     dfX.loc[dfX['SALDO'] < 0, 'DIAS POR VENCER'] = 0
 
     # Asegurar que no haya NaN y que sean enteros
     dfX['DIAS VENCIDOS'] = (
         pd.to_numeric(dfX['DIAS VENCIDOS'], errors='coerce')
         .fillna(0)
         .astype(int)
     )
 
     dfX['DIAS POR VENCER'] = (
         pd.to_numeric(dfX['DIAS POR VENCER'], errors='coerce')
         .fillna(0)
         .astype(int)
     )

    total_pesos = df_pesos_final['SALDO'].sum() if 'SALDO' in df_pesos_final.columns else 0.0
    total_div   = df_divisas_final['SALDO'].sum() if 'SALDO' in df_divisas_final.columns else 0.0
    print(f"  [OK] Registros PESOS:   {len(df_pesos_final):,}  |  Saldo: ${total_pesos:,.0f}")
    print(f"  [OK] Registros DIVISAS: {len(df_divisas_final):,}  |  Saldo moneda original: {total_div:,.2f}")

    # =====================================================
    # CONVERTIR DIVISAS A COP PARA CONSOLIDADO
    # =====================================================
    
    df_divisas_cop = df_divisas_final.copy()
    
    # -------------------------------
    # Validación defensiva
    # -------------------------------
    if 'MONEDA' not in df_divisas_cop.columns:
        raise ValueError("La columna MONEDA no existe en df_divisas_cop")
    
    # -------------------------------
    # Normalizar moneda
    # -------------------------------
    df_divisas_cop['MONEDA'] = (
        df_divisas_cop['MONEDA']
        .astype(str)
        .str.strip()
        .str.upper()
    )
    
    # ---------------------------------
    # Aplicar TRM según moneda
    # ---------------------------------
    
    df_divisas_cop['SALDO_COP'] = np.where(
        df_divisas_cop['MONEDA'] == 'DÓLAR',
        df_divisas_cop['SALDO'] * trm_dolar,
        np.where(
            df_divisas_cop['MONEDA'] == 'EURO',
            df_divisas_cop['SALDO'] * trm_euro,
            0.0
        )
    )
    
    # Redondear a 2 decimales
    df_divisas_cop['SALDO_COP'] = df_divisas_cop['SALDO_COP'].round(2)
    
    total_divisas_cop = df_divisas_cop['SALDO_COP'].sum()
    
    print(f"  [OK] Total DIVISAS convertido a COP: ${total_divisas_cop:,.0f}")
    
    # -------------------------------
    # Crear factor de conversión por fila
    # -------------------------------
    if 'MONEDA' not in df_divisas_cop.columns:
     df_divisas_cop['MONEDA'] = df_divisas_cop['LINEA DE NEGOCIO'].apply(_moneda_por_linea)
    
    # VERSIÓN SEGURA - con reset_index para evitar problemas de alineación
    df_divisas_cop = df_divisas_cop.reset_index(drop=True)
    
    factor = df_divisas_cop['MONEDA'].str.upper().map({
    'DÓLAR': trm_dolar,
    'DOLAR': trm_dolar,
    'USD': trm_dolar,
    'EURO': trm_euro,
    'EUR': trm_euro,
    'PESOS COL': 1.0
     }).fillna(1.0)
    
    # Columnas a convertir
    columnas_a_convertir = [
        'SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60',
        'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360',
        'VENCIDO + 360', 'DEUDA INCOBRABLE', 'MORA TOTAL', 'TOTAL POR VENCER'
    ]
    
    for col in columnas_a_convertir:
        if col in df_divisas_cop.columns:
            df_divisas_cop[col] = (
                pd.to_numeric(df_divisas_cop[col], errors='coerce').fillna(0) * factor
            ).round(2)
            
    # -------------------------------
    # Agregar columna de trazabilidad
    # -------------------------------
    df_divisas_cop['TRM_USADA'] = factor
    
    # -------------------------------
    # Logs de control financiero
    # -------------------------------
    total_original = pd.to_numeric(
        df_divisas_final['SALDO'], errors='coerce'
    ).fillna(0).sum()
    
    total_convertido = pd.to_numeric(
        df_divisas_cop['SALDO'], errors='coerce'
    ).fillna(0).sum()
    
    print(f"  [OK] Total divisas original: {total_original:,.2f}")
    print(f"  [OK] Total divisas en COP: {total_convertido:,.2f}")
    
    # -------------------------------
    # Validaciones adicionales
    # -------------------------------
    
    # 1️⃣ Validar que sí hubo conversión
    if round(total_convertido, 2) == round(total_original, 2):
        print("[WARN] Advertencia: Revisar TRM, no hubo variación en conversión.")
    
    # 2️⃣ Validar que no quedó en cero por error
    if abs(total_convertido) < 1:
        print("[WARN] Advertencia: Total convertido es cercano a cero. Revisar TRM.")
    
    print("  [OK] Conversión de divisas finalizada correctamente.")

    # -- PASO 6: Hoja VENCIMIENTO --
    print("\n[6/7] Generando hoja VENCIMIENTO...")
    df_vencimientos = crear_hoja_vencimientos(df_pesos_final, df_divisas_cop)
    
    # Hoja USD/EURO sin conversión
    df_usd_euro_vencimientos = crear_hoja_usd_euro_vencimientos(df_pesos_final, df_divisas_cop)
    print(f"  [OK] {len(df_usd_euro_vencimientos):,} filas USD_EURO_VENCIMIENTOS")
    print(f"  [OK] {len(df_vencimientos):,} filas en hoja VENCIMIENTO")
    
    # --- Quitar hora de fechas antes de exportar ---
    for dfX in (df_pesos_final, df_divisas_final, df_divisas_cop, df_vencimientos, df_usd_euro_vencimientos):
        if 'FECHA' in dfX.columns:
            dfX['FECHA'] = pd.to_datetime(dfX['FECHA'], errors='coerce').dt.date
        if 'FECHA VTO' in dfX.columns:
            dfX['FECHA VTO'] = pd.to_datetime(dfX['FECHA VTO'], errors='coerce').dt.date
           
    # -- PASO 7: Exportar a Excel --
    print("\n[7/7] Guardando archivo Excel...")
    output_path = os.path.join(SALIDAS_DIR, output_file)

    # Columnas numéricas a formatear
    cols_num = [
        'SALDO',
        'VALOR',
        'SALDO NO VENCIDO',
        'SALDO VENCIDO',
        'VENCIDO 30',
        'VENCIDO 60',
        'VENCIDO 90',
        'VENCIDO 180',
        'VENCIDO 360',
        'VENCIDO + 360',
        'DEUDA INCOBRABLE',
        'SALDO TOTAL',
        'MORA TOTAL',
        'TOTAL POR VENCER',
        'Saldo No vencido',
        'Saldo Vencido',
        'Vencido 30',
        'Vencido 60',
        'Vencido 90',
        'Vencido 180',
        'Vencido 360',
        'Vencido + 360',
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

            df_data.to_excel(writer, sheet_name=nombre_hoja, index=False, startrow=0)
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
        # LIMPIEZA GLOBAL (UNA SOLA VEZ)
        # ===================================
        df_pesos_final = df_pesos_final.loc[:, ~df_pesos_final.columns.duplicated(keep='first')]
        df_divisas_final = df_divisas_final.loc[:, ~df_divisas_final.columns.duplicated(keep='first')]
        df_vencimientos = df_vencimientos.loc[:, ~df_vencimientos.columns.duplicated(keep='first')]
        
        for dfX in (df_pesos_final, df_divisas_final, df_vencimientos):
            if not dfX.empty:
                columnas_numericas = dfX.select_dtypes(include=['number']).columns
                dfX[columnas_numericas] = dfX[columnas_numericas].fillna(0)
          
        # ===================================
        # HOJA PESOS
        # ===================================
        if not df_pesos_final.empty:
        
            totales_pesos = (
                df_pesos_final[[c for c in cols_num if c in df_pesos_final.columns]]
                .sum()
                .to_frame()
                .T
            )
        
            totales_pesos['LINEA DE NEGOCIO'] = 'TOTAL GENERAL'
            totales_pesos['DENOMINACION COMERCIAL'] = ''
            totales_pesos['MONEDA'] = 'PESOS COL'
        
            # Asegurar mismas columnas
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

            # Totales ya convertidos correctamente desde df_divisas_cop
            tot_cop = (
                df_divisas_cop
                .groupby('MONEDA', as_index=False)[[c for c in cols_num if c in df_divisas_cop.columns]]
                .sum()
            )
            
            tot_cop['__TIPO__'] = f'TOTAL MONEDA (COP) -- TRM USD:{trm_dolar:,.2f} | EUR:{trm_euro:,.2f}'

            df_divisas_salida = pd.concat(
                [df_divisas_final, tot_moneda, tot_cop], ignore_index=True
            )
            
            # Antes de escribir VENCIMIENTOS_EXTRANJERO (línea ~1295)

            # Filtrar NAT
            df_divisas_final = df_divisas_final[df_divisas_final['TIPO'].astype(str).str.strip() != 'NAT']
            
            # En los totales por moneda, agregar la moneda en el nombre:
            tot_cop['__TIPO__'] = (
                f"TOTAL MONEDA: {tot_cop['MONEDA'].values[0]} (COP) "
                f"-- TRM USD:{trm_dolar:,.2f} | EUR:{trm_euro:,.2f}"
            )
            
            _escribir_hoja(df_divisas_salida, 'VENCIMIENTOS_EXTRANJERO', 'TOTAL MONEDA', '__TIPO__')
            print("  [OK] Hoja VENCIMIENTOS_EXTRANJERO escrita")

        # ===================================
        # HOJA VENCIMIENTO
        # ===================================
        if not df_vencimientos.empty:
            _escribir_hoja(df_vencimientos, 'VENCIMIENTO', 'TOTAL GENERAL', 'CLIENTE')
            print("  [OK] Hoja VENCIMIENTO escrita")
            
        # Hoja USD_EURO_VENCIMIENTOS
        if not df_usd_euro_vencimientos.empty:
            _escribir_hoja(df_usd_euro_vencimientos, 'USD_EURO_VENCIMIENTOS', 'TOTAL GENERAL', 'CLIENTE')
            print("  [OK] Hoja USD_EURO_VENCIMIENTOS escrita")
            
        # ===================================
        # HOJA TASAS TRM
        # ===================================
        df_tasas = pd.DataFrame([
            {'Concepto': 'Tasa de Cambio USD (Dólar)', 'Valor': trm_dolar, 'Moneda': 'COP por USD'},
            {'Concepto': 'Tasa de Cambio EUR (Euro)', 'Valor': trm_euro, 'Moneda': 'COP por EUR'},
            {'Concepto': 'Fecha TRM', 'Valor': fecha_trm, 'Moneda': ''},
            {'Concepto': 'Fecha de Generación', 'Valor': datetime.now().strftime('%Y-%m-%d %H:%M:%S'), 'Moneda': ''},
        ])
        
        df_tasas.to_excel(writer, sheet_name='TASAS_TRM', index=False, startrow=1)
        ws_tasas = writer.sheets['TASAS_TRM']
        
        # Encabezados formateados
        ws_tasas.write(0, 0, 'Concepto', fmt_header)
        ws_tasas.write(0, 1, 'Valor', fmt_header)
        ws_tasas.write(0, 2, 'Unidad', fmt_header)
        
        # Ancho de columnas
        ws_tasas.set_column(0, 0, 30)
        ws_tasas.set_column(1, 1, 15)
        ws_tasas.set_column(2, 2, 15)
        
        print("  [OK] Hoja TASAS_TRM escrita")

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