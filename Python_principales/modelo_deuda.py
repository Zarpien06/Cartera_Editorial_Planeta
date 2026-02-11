import argparse
import sys
import pandas as pd
import numpy as np
import os
import logging
import re
from datetime import datetime
from typing import Union, Any

# Configuración de logging unificado
try:
    from config_logging import logger, log_inicio_proceso, log_fin_proceso, log_error_proceso
    USE_UNIFIED_LOGGING = True
except ImportError:
    # Fallback al logging original si config_logging no está disponible
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

if USE_UNIFIED_LOGGING:
    # Este mensaje se mostrará cuando se llame a la función principal
    pass
else:
    logging.info("Iniciando modelo de deuda")

BASE_DIR = os.path.dirname(os.path.abspath(__file__))
SALIDAS_DIR = os.path.join(BASE_DIR, 'salidas')
os.makedirs(SALIDAS_DIR, exist_ok=True)

# Líneas de venta según especificación
LINEAS_PESOS = [
    ('CT', '80'), ('ED', '41'), ('ED', '44'), ('ED', '47'),
    ('PL', '10'), ('PL', '15'), ('PL', '20'), ('PL', '21'),
    ('PL', '23'), ('PL', '25'), ('PL', '28'), ('PL', '29'),
    ('PL', '31'), ('PL', '32'), ('PL', '53'), ('PL', '56'),
    ('PL', '60'), ('PL', '62'), ('PL', '63'), ('PL', '64'),
    ('PL', '65'), ('PL', '66'), ('PL', '69')
]

LINEAS_DIVISAS = [
    ('PL', '11'), ('PL', '18'), ('PL', '57'), ('PL', '72'),
    ('PL', '17'), ('PL', '16'), ('PL', '41'), ('PL', '68')
]

# Tabla Negocio-Canal según especificación exacta
TABLA_NEGOCIO_CANAL = {
    'PL10': {'NEGOCIO': 'LIBRERIAS 2', 'CANAL': 'LIBRERIAS 2'},
    'PL15': {'NEGOCIO': 'E-COMMERCE', 'CANAL': 'E-COMMERCE'},
    'PL20': {'NEGOCIO': 'LIBRERIAS 1', 'CANAL': 'LIBRERIAS 1'},
    'PL21': {'NEGOCIO': 'LIBRERIAS 2', 'CANAL': 'LIBRERIAS 2'},
    'PL23': {'NEGOCIO': 'LIBRERIAS 3', 'CANAL': 'LIBRERIAS 3'},
    'PL25': {'NEGOCIO': 'LIBRERIAS 1', 'CANAL': 'LIBRERIAS 1'},
    'PL28': {'NEGOCIO': 'SALDOS', 'CANAL': 'SALDOS'},
    'PL29': {'NEGOCIO': 'SALDOS', 'CANAL': 'SALDOS'},
    'PL31': {'NEGOCIO': 'SALDOS', 'CANAL': 'SALDOS'},
    'PL32': {'NEGOCIO': 'DISTRIBUIDORES', 'CANAL': 'DISTRIBUIDORES'},
    'PL53': {'NEGOCIO': 'LIBRERIAS 3', 'CANAL': 'LIBRERIAS 3'},
    'PL56': {'NEGOCIO': 'OTROS DIGITAL', 'CANAL': 'OTROS DIGITAL'},
    'PL57': {'NEGOCIO': 'PRENSA', 'CANAL': 'PRENSA'},
    'PL60': {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'},
    'PL62': {'NEGOCIO': 'PRENSA', 'CANAL': 'PRENSA'},
    'PL63': {'NEGOCIO': 'LIBRERIAS 3', 'CANAL': 'LIBRERIAS 3'},
    'PL64': {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'},
    'PL65': {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'},
    'PL66': {'NEGOCIO': 'OTROS DIGITAL', 'CANAL': 'OTROS DIGITAL'},
    'PL69': {'NEGOCIO': 'LIBRERIAS 1', 'CANAL': 'LIBRERIAS 1'},
    'PL11': {'NEGOCIO': 'EXPORTACION', 'CANAL': 'EXPORTACION'},
    'PL16': {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'},
    'PL17': {'NEGOCIO': 'EXPORTACION', 'CANAL': 'EXPORTACION'},
    'PL18': {'NEGOCIO': 'EXPORTACION', 'CANAL': 'EXPORTACION'},
    'PL41': {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'},
    'PL68': {'NEGOCIO': 'OTROS', 'CANAL': 'OTROS'},
    'PL72': {'NEGOCIO': 'EXPORTACION', 'CANAL': 'EXPORTACION'},
    'CT80': {'NEGOCIO': 'TINTA', 'CANAL': 'TINTA'},
    'ED41': {'NEGOCIO': 'EDUCACION', 'CANAL': 'EDUCACION'},
    'ED44': {'NEGOCIO': 'EDUCACION', 'CANAL': 'EDUCACION'},
    'ED47': {'NEGOCIO': 'EDUCACION', 'CANAL': 'EDUCACION'}
}

def safe_float_conversion(val: Any) -> float:
    """Convierte de forma segura cualquier valor a float, manejando objetos pandas."""
    if val is None:
        return 0.0
    if isinstance(val, (int, float, np.number)):
        return float(val)
    if hasattr(val, 'item'):
        # Para valores de numpy/pandas que pueden extraer un escalar
        try:
            return float(val.item())
        except (ValueError, AttributeError, TypeError):
            pass
    if hasattr(val, '__iter__') and not isinstance(val, (str, bytes)):
        # Si es iterable (como Series o array), tomar el primer elemento si es escalar
        try:
            if len(val) == 1:
                return float(list(val)[0])
        except (ValueError, TypeError):
            pass
    try:
        return float(val)
    except (ValueError, TypeError):
        return 0.0

def safe_isna_check(val: Any) -> bool:
    """Verifica de forma segura si un valor es NA/NaN, manejando objetos pandas."""
    try:
        # Si es un pandas Series o similar, usar .any()
        if hasattr(val, 'any') and not pd.api.types.is_scalar(val):
            return bool(val.isna().any() if hasattr(val, 'isna') else val.any())
        else:
            # Es un valor escalar, usar pd.isna normalmente
            return pd.isna(val)
    except:
        # En caso de error, intentar con pd.isna directamente
        return pd.isna(val)

def cargar_tasas_cambio():
    """Carga las tasas de cambio desde el módulo trm_config"""
    try:
        from trm_config import load_trm
        trm_data = load_trm()
        
        if not trm_data or not isinstance(trm_data, dict):
            print("⚠ Datos de TRM inválidos. Usando valores por defecto")
            return 4000.0, 4500.0, datetime.now().strftime('%Y-%m-%d')
        
        # Obtener tasas con valores por defecto si no existen
        tasa_usd = float(trm_data.get('usd', 4000.0))
        tasa_eur = float(trm_data.get('eur', 4500.0))
        fecha = str(trm_data.get('fecha', datetime.now().strftime('%Y-%m-%d')))
        
        # Validar que las tasas sean positivas
        if tasa_usd <= 0:
            tasa_usd = 4000.0
        if tasa_eur <= 0:
            tasa_eur = 4500.0
        
        print(f"[OK] TRM cargada - USD: ${tasa_usd:,.2f}, EUR: ${tasa_eur:,.2f}")
        return tasa_usd, tasa_eur, fecha
        
    except ImportError:
        print(f"[WARN] Módulo trm_config no encontrado. Usando TRM por defecto - USD: $4000, EUR: $4500")
        return 4000.0, 4500.0, datetime.now().strftime('%Y-%m-%d')
    except Exception as e:
        print(f"[ERROR] Error al cargar TRM: {e}. Usando valores por defecto")
        return 4000.0, 4500.0, datetime.now().strftime('%Y-%m-%d')

def crear_modelo_deuda(archivo_provision, archivo_anticipos, output_file='1_Modelo_Deuda.xlsx'):
    """
    Crea el archivo Modelo Deuda combinando cartera y anticipos
    Los archivos de entrada deben venir de procesador_cartera.py y procesador_anticipos.py
    """
    
    if USE_UNIFIED_LOGGING:
        log_inicio_proceso("MODELO_DEUDA", f"{archivo_provision} + {archivo_anticipos}")
    else:
        logging.info("Iniciando modelo de deuda")
    
    print("\n" + "="*60)
    print("  GENERADOR DE MODELO DE DEUDA")
    print("="*60)
    
    # Cargar TRM
    print("\n[1/6] Cargando tasas de cambio...")
    trm_dolar, trm_euro, fecha_trm = cargar_tasas_cambio()
    
    # Leer archivos procesados
    print("\n[2/6] Leyendo archivos procesados...")
    
    def leer_archivo(archivo):
        """Lee archivo Excel o CSV"""
        if not os.path.exists(archivo):
            raise FileNotFoundError(f"No se encontró: {archivo}")
        
        if archivo.lower().endswith('.csv'):
            encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']
            separators = [',', ';', '\t', '|']  # Probar diferentes separadores
            
            for encoding in encodings:
                for sep in separators:
                    try:
                        df = pd.read_csv(archivo, encoding=encoding, sep=sep)
                        # Verificar si la lectura fue exitosa (si hay más de una columna)
                        if len(df.columns) > 1 or len(df) > 0:
                            return df
                    except:
                        continue
            
            # Si falla con todos los separadores y codificaciones, usar valores predeterminados
            return pd.read_csv(archivo, encoding='latin1', sep=';')
        elif archivo.lower().endswith(('.xlsx', '.xls')):
            return pd.read_excel(archivo, engine='openpyxl')
        else:
            # Si no es CSV ni Excel, intentar como CSV por defecto
            encodings = ['utf-8-sig', 'utf-8', 'latin1', 'cp1252']
            separators = [',', ';', '\t', '|']
            
            for encoding in encodings:
                for sep in separators:
                    try:
                        df = pd.read_csv(archivo, encoding=encoding, sep=sep)
                        if len(df.columns) > 1 or len(df) > 0:
                            return df
                    except:
                        continue
            
            return pd.read_csv(archivo, encoding='latin1', sep=';')
    
    df_provision = leer_archivo(archivo_provision)
    df_anticipos = leer_archivo(archivo_anticipos)
    
    print(f"  [OK] Cartera: {len(df_provision)} registros")
    print(f"  [OK] Anticipos: {len(df_anticipos)} registros")
    
    # Crear columna LINEA DE NEGOCIO
    print("\n[3/6] Creando líneas de negocio...")
    
    # Mapeo de columnas de código a nombres
    mapeo_cartera = {
        'PCCDEM': 'EMPRESA',
        'PCCDAC': 'ACTIVIDAD',
        'PCNMCL': 'NOMBRE CLIENTE',
        'PCNMCM': 'DENOMINACION COMERCIAL',
        'PCVAFA': 'VALOR FACTURA',
        'PCSALD': 'SALDO',
        'PCIMCO': 'DEUDA INCOBRABLE',
        'PCFEVE': 'FECHA VTO',
        'PCFEFA': 'FECHA',
        'PCORPD': 'TIPO',
        'PCTLF1': 'TELEFONO',
        'PCNMDO': 'DIRECCION',
        'PCNMPO': 'CIUDAD',
        'PCNUFC': 'NUMERO FACTURA',
        'PCCDCL': 'CODIGO CLIENTE',
        'PCCDDN': 'IDENTIFICACION'
    }
    
    mapeo_anticipos = {
        'NCCDEM': 'EMPRESA',
        'NCCDAC': 'ACTIVIDAD',
        'WWNMCL': 'NOMBRE CLIENTE',
        'BDNMNM': 'DENOMINACION COMERCIAL',
        'NCCDR3': 'VALOR ANTICIPO',
        'NCFEGR': 'FECHA',
        'NCCDCL': 'CODIGO CLIENTE',
        'WWNIT': 'IDENTIFICACION',
        'WWNMDO': 'DIRECCION',
        'WWTLF1': 'TELEFONO',
        'WWNMPO': 'CIUDAD',
        'NCMOMO': 'SALDO',
        'NCIMAN': 'DEUDA INCOBRABLE'
    }
    
    # Aplicar mapeo a cartera
    for codigo, nombre in mapeo_cartera.items():
        if codigo in df_provision.columns:
            df_provision = df_provision.rename(columns={codigo: nombre})
    
    # Aplicar mapeo a anticipos
    for codigo, nombre in mapeo_anticipos.items():
        if codigo in df_anticipos.columns:
            df_anticipos = df_anticipos.rename(columns={codigo: nombre})
    
    # Para cartera (EMPRESA + ACTIVIDAD)
    if 'EMPRESA' in df_provision.columns and 'ACTIVIDAD' in df_provision.columns:
        # Limpiar valores de EMPRESA y ACTIVIDAD
        emp = df_provision['EMPRESA'].astype(str).str.strip().str.upper()
        act = df_provision['ACTIVIDAD'].astype(str).str.strip().str.upper()
        # Eliminar decimales si existen
        act = act.str.replace(r'\.0+$', '', regex=True)
        df_provision['LINEA DE NEGOCIO'] = emp + act
        print(f"  [DEBUG] Líneas de negocio en cartera: {df_provision['LINEA DE NEGOCIO'].nunique()} líneas únicas")
    
    # Para anticipos (usar LINEA_VENTA o crear)
    if 'LINEA_VENTA' in df_anticipos.columns:
        df_anticipos['LINEA DE NEGOCIO'] = df_anticipos['LINEA_VENTA'].astype(str).str.strip().str.upper()
        print(f"  [DEBUG] Líneas de negocio en anticipos (LINEA_VENTA): {df_anticipos['LINEA DE NEGOCIO'].nunique()} líneas únicas")
    elif 'EMPRESA' in df_anticipos.columns and 'ACTIVIDAD' in df_anticipos.columns:
        # Limpiar valores de EMPRESA y ACTIVIDAD
        emp = df_anticipos['EMPRESA'].astype(str).str.strip().str.upper()
        act = df_anticipos['ACTIVIDAD'].astype(str).str.strip().str.upper()
        # Eliminar decimales si existen
        act = act.str.replace(r'\.0+$', '', regex=True)
        df_anticipos['LINEA DE NEGOCIO'] = emp + act
        print(f"  [DEBUG] Líneas de negocio en anticipos (EMPRESA+ACTIVIDAD): {df_anticipos['LINEA DE NEGOCIO'].nunique()} líneas únicas")
    
    # Mapear VALOR ANTICIPO a SALDO si no existe
    if 'VALOR ANTICIPO' in df_anticipos.columns and 'SALDO' not in df_anticipos.columns:
        df_anticipos['SALDO'] = df_anticipos['VALOR ANTICIPO']
    
    # Mapear NOMBRE COMERCIAL a DENOMINACION COMERCIAL si no existe
    if 'NOMBRE COMERCIAL' in df_anticipos.columns and 'DENOMINACION COMERCIAL' not in df_anticipos.columns:
        df_anticipos['DENOMINACION COMERCIAL'] = df_anticipos['NOMBRE COMERCIAL']
    
    print(f"  [OK] Líneas de negocio creadas")
    
    # HOJA 1: PESOS
    print("\n[4/6] Procesando hoja PESOS...")
    lineas_pesos_keys = [f"{cod}{act}" for cod, act in LINEAS_PESOS]
    
    df_pesos = df_provision[df_provision['LINEA DE NEGOCIO'].isin(lineas_pesos_keys)].copy()
    print(f"  [OK] Cartera en pesos: {len(df_pesos)} registros")
    
    # Agregar anticipos en pesos
    if 'MONEDA' in df_anticipos.columns:
        # Crear una máscara más precisa para identificar monedas en pesos
        mask_pesos = ~df_anticipos['MONEDA'].str.contains('USD|EUR|DOLAR|EURO', case=False, na=True)
        # También considerar NaN o valores vacíos como pesos
        mask_pesos = mask_pesos | df_anticipos['MONEDA'].isna() | (df_anticipos['MONEDA'].str.strip() == '')
        df_anticipos_pesos = df_anticipos[mask_pesos].copy()
        
        if not df_anticipos_pesos.empty:
            print(f"  [DEBUG] Anticipos en pesos encontrados: {len(df_anticipos_pesos)} registros")
            
            # Asegurar que los anticipos tengan la misma estructura que cartera
            columnas_requeridas = ['LINEA DE NEGOCIO', 'DENOMINACION COMERCIAL', 'MONEDA', 'SALDO', 
                                 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90', 
                                 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE']
            
            # Asegurar que todas las columnas requeridas existan
            for col in columnas_requeridas:
                if col not in df_anticipos_pesos.columns:
                    if col in ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 
                               'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 
                               'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE']:
                        df_anticipos_pesos[col] = 0.0
                    else:
                        df_anticipos_pesos[col] = ''
            
            # Asegurar que los saldos de anticipos se distribuyan correctamente en las columnas de vencimiento
            if 'SALDO' in df_anticipos_pesos.columns:
                # Para anticipos, todo el saldo se considera no vencido
                df_anticipos_pesos['SALDO NO VENCIDO'] = df_anticipos_pesos['SALDO'].astype('float64')
                
                # Inicializar otras columnas de vencimiento a cero, pero si existen en los datos de entrada, mantener sus valores
                for col_venc in ['VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360']:
                    if col_venc not in df_anticipos_pesos.columns:
                        df_anticipos_pesos[col_venc] = 0.0
                
                # Para DEUDA INCOBRABLE, si existe en los datos de entrada, mantener su valor, de lo contrario inicializar en 0
                if 'DEUDA INCOBRABLE' not in df_anticipos_pesos.columns:
                    df_anticipos_pesos['DEUDA INCOBRABLE'] = 0.0
            
            # Filtrar anticipos por líneas de negocio en pesos
            if 'LINEA DE NEGOCIO' in df_anticipos_pesos.columns:
                df_anticipos_pesos = df_anticipos_pesos[df_anticipos_pesos['LINEA DE NEGOCIO'].isin(lineas_pesos_keys)]
            
            # Asegurar que los anticipos tengan las mismas columnas que cartera
            # Primero, obtener las columnas que existen en cartera
            columnas_cartera = df_pesos.columns.tolist() if not df_pesos.empty else []
            
            # Asegurar que anticipos tenga todas las columnas de cartera
            for col in columnas_cartera:
                if col not in df_anticipos_pesos.columns:
                    if col in ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 
                               'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 
                               'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE']:
                        df_anticipos_pesos[col] = 0.0
                    else:
                        df_anticipos_pesos[col] = ''
            
            # Reordenar columnas de anticipos para que coincidan con cartera
            if columnas_cartera:  # Solo reordenar si hay columnas
                # Convertir a lista explícitamente para evitar problemas de tipo con Pyright
                # Usar reindex en lugar de reindex(columns=) para evitar errores de tipo en Pyright
                df_anticipos_pesos = df_anticipos_pesos.reindex(columns=columnas_cartera)
            
            # Aplicar lógica de compensación de anticipos con cartera
            # Los anticipos tienen valores negativos (en la hoja de anticipos ya están multiplicados por -1), por lo tanto se suman a la cartera (lo que reduce el saldo)
            print(f"  [DEBUG] Aplicando compensación de anticipos a cartera...")
            
            # Agrupar anticipos por DENOMINACION COMERCIAL para aplicarlos a todos los registros del cliente
            if 'DENOMINACION COMERCIAL' in df_anticipos_pesos.columns:
                print(f"  [DEBUG] Procesando anticipos para compensación")
                print(f"  [DEBUG] Total anticipos antes de agrupar: {df_anticipos_pesos['SALDO'].sum()}")
                
                # Agrupar anticipos por cliente para sumar todos los anticipos del cliente
                anticipos_agrupados_cliente = df_anticipos_pesos.groupby('DENOMINACION COMERCIAL').agg({
                    'SALDO': 'sum',
                    'SALDO NO VENCIDO': 'sum',
                    'VENCIDO 30': 'sum',
                    'VENCIDO 60': 'sum',
                    'VENCIDO 90': 'sum',
                    'VENCIDO 180': 'sum',
                    'VENCIDO 360': 'sum',
                    'VENCIDO + 360': 'sum',
                    'DEUDA INCOBRABLE': 'sum'
                }).reset_index()
                
                print(f"  [DEBUG] Anticipos agrupados por cliente: {len(anticipos_agrupados_cliente)} clientes")
                
                # Aplicar compensación a los registros de cartera
                for idx, row in anticipos_agrupados_cliente.iterrows():
                    mask = df_pesos['DENOMINACION COMERCIAL'] == row['DENOMINACION COMERCIAL']
                    
                    if mask.any():  # type: ignore[truthy-bool]
                        # Calcular el saldo total antes de la compensación para este cliente
                        saldo_antes = safe_float_conversion(df_pesos.loc[mask, 'SALDO'].sum())
                        
                        print(f"    [DEBUG] Aplicando compensación a: {row['DENOMINACION COMERCIAL']}")
                        print(f"    [DEBUG] Saldo antes: {saldo_antes}, Anticipo a aplicar: {row['SALDO']}")
                        
                        # Aplicar compensación a cada columna de saldo a todos los registros del cliente
                        # IMPORTANTE: Los anticipos tienen valores negativos (en la hoja de anticipos ya están multiplicados por -1), por lo tanto se suman a la cartera (lo que reduce el saldo)
                        df_pesos.loc[mask, 'SALDO'] += row['SALDO']
                        df_pesos.loc[mask, 'SALDO NO VENCIDO'] += row['SALDO NO VENCIDO']
                        df_pesos.loc[mask, 'VENCIDO 30'] += row['VENCIDO 30']
                        df_pesos.loc[mask, 'VENCIDO 60'] += row['VENCIDO 60']
                        df_pesos.loc[mask, 'VENCIDO 90'] += row['VENCIDO 90']
                        df_pesos.loc[mask, 'VENCIDO 180'] += row['VENCIDO 180']
                        df_pesos.loc[mask, 'VENCIDO 360'] += row['VENCIDO 360']
                        df_pesos.loc[mask, 'VENCIDO + 360'] += row['VENCIDO + 360']
                        df_pesos.loc[mask, 'DEUDA INCOBRABLE'] += row['DEUDA INCOBRABLE']
                        
                        saldo_despues = safe_float_conversion(df_pesos.loc[mask, 'SALDO'].sum())
                        print(f"    [DEBUG] Compensación aplicada a: {row['DENOMINACION COMERCIAL']}")
                        print(f"    [DEBUG] Antes: {saldo_antes}, Despues: {saldo_despues}, Diferencia: {saldo_despues - saldo_antes}")
                        print(f"    [DEBUG] Total registros afectados: {mask.sum()}")
                        print(f"    [DEBUG] Verificando si hay anticipos negativos: {(row['SALDO'] < 0) if pd.api.types.is_scalar(row['SALDO']) else (row['SALDO'] < 0).any()}")
                    else:
                        # Si no hay cartera para este cliente, agregar el anticipo como nuevo registro
                        # Tomar un registro cualquiera de los anticipos para este cliente
                        cliente_anticipos = df_anticipos_pesos[
                            df_anticipos_pesos['DENOMINACION COMERCIAL'] == row['DENOMINACION COMERCIAL']
                        ]
                        nuevo_registro = cliente_anticipos.iloc[0].copy()
                        df_pesos = pd.concat([df_pesos, pd.DataFrame([nuevo_registro])], ignore_index=True)
                        print(f"    [DEBUG] Agregando anticipo como nuevo registro para: {row['DENOMINACION COMERCIAL']}")
            
            print(f"  [OK] Anticipos compensados: {len(df_anticipos_pesos)} registros procesados")
            print(f"  [OK] Total PESOS después de compensación: {len(df_pesos)} registros")
    
    # Calcular totales para PESOS
    columnas_suma = ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60',
                     'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360',
                     'DEUDA INCOBRABLE']
    
    # Agregar columnas de NEGOCIO y CANAL a PESOS basadas en LINEA DE NEGOCIO
    if 'LINEA DE NEGOCIO' in df_pesos.columns:
        df_pesos['NEGOCIO'] = df_pesos['LINEA DE NEGOCIO'].apply(
            lambda x: TABLA_NEGOCIO_CANAL.get(str(x).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': str(x)})['NEGOCIO']
        )
        df_pesos['CANAL'] = df_pesos['LINEA DE NEGOCIO'].apply(
            lambda x: TABLA_NEGOCIO_CANAL.get(str(x).strip().upper(), {'NEGOCIO': str(x), 'CANAL': str(x)})['CANAL']
        )
    
    # No filtrar registros con saldos cero o negativos para mantener visibles los anticipos que superan la cartera
    # Los valores negativos deben mostrarse explícitamente en los informes financieros
    print(f"  [DEBUG] Registros PESOS antes de filtrar: {len(df_pesos)}")
    print(f"  [DEBUG] Se mantienen todos los registros incluyendo saldos negativos después de compensación")
    
    # No agrupar por cliente si ya se aplicó compensación, para mantener la compensación por línea de negocio
    # La agrupación se hará en la hoja de vencimientos manteniendo las líneas de negocio
    
    # Agregar columna de NEGOCIO basada en LINEA DE NEGOCIO
    if 'LINEA DE NEGOCIO' in df_pesos.columns:
        df_pesos['NEGOCIO'] = df_pesos['LINEA DE NEGOCIO'].apply(
            lambda x: TABLA_NEGOCIO_CANAL.get(str(x).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': str(x)})['NEGOCIO']
        )
        
        # Eliminar columnas origen que ya no son necesarias
        columnas_a_eliminar = ['LINEA DE NEGOCIO', 'EMPRESA', 'ACTIVIDAD']
        for col in columnas_a_eliminar:
            if col in df_pesos.columns:
                df_pesos = df_pesos.drop(columns=[col])
    
    total_pesos = safe_float_conversion(df_pesos['SALDO'].sum()) if 'SALDO' in df_pesos.columns else 0.0
    print(f"  [OK] Saldo total PESOS: ${total_pesos:,.0f}")
    
    # HOJA 2: DIVISAS
    print("\n[5/6] Procesando hoja DIVISAS...")
    lineas_divisas_keys = [f"{cod}{act}" for cod, act in LINEAS_DIVISAS]
    
    df_divisas = df_provision[df_provision['LINEA DE NEGOCIO'].isin(lineas_divisas_keys)].copy()
    print(f"  [OK] Cartera en divisas: {len(df_divisas)} registros")
    
    # Agregar anticipos en divisas
    if 'MONEDA' in df_anticipos.columns:
        df_anticipos_divisas = df_anticipos[
            df_anticipos['MONEDA'].str.contains('USD|EUR|DOLAR|EURO', case=False, na=False)
        ].copy()
        
        if not df_anticipos_divisas.empty:
            print(f"  [DEBUG] Anticipos en divisas encontrados: {len(df_anticipos_divisas)} registros")
            
            # Asegurar que los anticipos tengan la misma estructura que cartera
            columnas_requeridas = ['LINEA DE NEGOCIO', 'DENOMINACION COMERCIAL', 'MONEDA', 'SALDO', 
                                 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90', 
                                 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE']
            
            # Asegurar que todas las columnas requeridas existan
            for col in columnas_requeridas:
                if col not in df_anticipos_divisas.columns:
                    if col in ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 
                               'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 
                               'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE']:
                        df_anticipos_divisas[col] = 0.0
                    else:
                        df_anticipos_divisas[col] = ''
            
            # Asegurar que los saldos de anticipos se distribuyan correctamente en las columnas de vencimiento
            if 'SALDO' in df_anticipos_divisas.columns:
                # Para anticipos, todo el saldo se considera no vencido
                df_anticipos_divisas['SALDO NO VENCIDO'] = df_anticipos_divisas['SALDO'].astype('float64')
                
                # Inicializar otras columnas de vencimiento a cero, pero si existen en los datos de entrada, mantener sus valores
                for col_venc in ['VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360']:
                    if col_venc not in df_anticipos_divisas.columns:
                        df_anticipos_divisas[col_venc] = 0.0
                
                # Para DEUDA INCOBRABLE, si existe en los datos de entrada, mantener su valor, de lo contrario inicializar en 0
                if 'DEUDA INCOBRABLE' not in df_anticipos_divisas.columns:
                    df_anticipos_divisas['DEUDA INCOBRABLE'] = 0.0
            
            # Filtrar anticipos por líneas de negocio en divisas
            if 'LINEA DE NEGOCIO' in df_anticipos_divisas.columns:
                df_anticipos_divisas = df_anticipos_divisas[df_anticipos_divisas['LINEA DE NEGOCIO'].isin(lineas_divisas_keys)]
            
            # Asegurar que los anticipos tengan las mismas columnas que cartera
            # Primero, obtener las columnas que existen en cartera
            columnas_cartera = df_divisas.columns.tolist() if not df_divisas.empty else []
            
            # Asegurar que anticipos tenga todas las columnas de cartera
            for col in columnas_cartera:
                if col not in df_anticipos_divisas.columns:
                    if col in ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 
                               'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 
                               'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE']:
                        df_anticipos_divisas[col] = 0.0
                    else:
                        df_anticipos_divisas[col] = ''
            
            # Reordenar columnas de anticipos para que coincidan con cartera
            if columnas_cartera:  # Solo reindexar si hay columnas
                # Convertir a lista explícitamente para evitar problemas de tipo con Pyright
                # Usar reindex en lugar de reindex(columns=) para evitar errores de tipo en Pyright
                df_anticipos_divisas = df_anticipos_divisas.reindex(columns=columnas_cartera)
            
            # Aplicar lógica de compensación de anticipos con cartera
            # Los anticipos tienen valores negativos (en la hoja de anticipos ya están multiplicados por -1), por lo tanto se suman a la cartera (lo que reduce el saldo)
            print(f"  [DEBUG] Aplicando compensación de anticipos divisas a cartera...")
            
            # Agrupar anticipos por DENOMINACION COMERCIAL para aplicarlos a todos los registros del cliente
            if 'DENOMINACION COMERCIAL' in df_anticipos_divisas.columns:
                print(f"  [DEBUG] Procesando anticipos divisas para compensación")
                print(f"  [DEBUG] Total anticipos divisas antes de agrupar: {df_anticipos_divisas['SALDO'].sum()}")
                
                # Agrupar anticipos por cliente para sumar todos los anticipos del cliente
                anticipos_agrupados_cliente = df_anticipos_divisas.groupby('DENOMINACION COMERCIAL').agg({
                    'SALDO': 'sum',
                    'SALDO NO VENCIDO': 'sum',
                    'VENCIDO 30': 'sum',
                    'VENCIDO 60': 'sum',
                    'VENCIDO 90': 'sum',
                    'VENCIDO 180': 'sum',
                    'VENCIDO 360': 'sum',
                    'VENCIDO + 360': 'sum',
                    'DEUDA INCOBRABLE': 'sum'
                }).reset_index()
                
                print(f"  [DEBUG] Anticipos divisas agrupados por cliente: {len(anticipos_agrupados_cliente)} clientes")
                
                # Aplicar compensación a los registros de cartera
                for idx, row in anticipos_agrupados_cliente.iterrows():
                    mask = df_divisas['DENOMINACION COMERCIAL'] == row['DENOMINACION COMERCIAL']
                    
                    if mask.any():  # type: ignore[truthy-bool]
                        # Calcular el saldo total antes de la compensación para este cliente
                        saldo_antes = safe_float_conversion(df_divisas.loc[mask, 'SALDO'].sum())
                        
                        print(f"    [DEBUG] Aplicando compensación divisas a: {row['DENOMINACION COMERCIAL']}")
                        print(f"    [DEBUG] Saldo antes: {saldo_antes}, Anticipo a aplicar: {row['SALDO']}")
                        
                        # Aplicar compensación a cada columna de saldo a todos los registros del cliente
                        df_divisas.loc[mask, 'SALDO'] += row['SALDO']
                        df_divisas.loc[mask, 'SALDO NO VENCIDO'] += row['SALDO NO VENCIDO']
                        df_divisas.loc[mask, 'VENCIDO 30'] += row['VENCIDO 30']
                        df_divisas.loc[mask, 'VENCIDO 60'] += row['VENCIDO 60']
                        df_divisas.loc[mask, 'VENCIDO 90'] += row['VENCIDO 90']
                        df_divisas.loc[mask, 'VENCIDO 180'] += row['VENCIDO 180']
                        df_divisas.loc[mask, 'VENCIDO 360'] += row['VENCIDO 360']
                        df_divisas.loc[mask, 'VENCIDO + 360'] += row['VENCIDO + 360']
                        df_divisas.loc[mask, 'DEUDA INCOBRABLE'] += row['DEUDA INCOBRABLE']
                        
                        saldo_despues = safe_float_conversion(df_divisas.loc[mask, 'SALDO'].sum())
                        print(f"    [DEBUG] Compensación divisas aplicada a: {row['DENOMINACION COMERCIAL']}")
                        print(f"    [DEBUG] Antes: {saldo_antes}, Despues: {saldo_despues}, Diferencia: {saldo_despues - saldo_antes}")
                        print(f"    [DEBUG] Total registros divisas afectados: {mask.sum()}")
                    else:
                        # Si no hay cartera para este cliente, agregar el anticipo como nuevo registro
                        # Tomar un registro cualquiera de los anticipos para este cliente
                        cliente_anticipos = df_anticipos_divisas[
                            df_anticipos_divisas['DENOMINACION COMERCIAL'] == row['DENOMINACION COMERCIAL']
                        ]
                        nuevo_registro = cliente_anticipos.iloc[0].copy()
                        df_divisas = pd.concat([df_divisas, pd.DataFrame([nuevo_registro])], ignore_index=True)
                        print(f"    [DEBUG] Agregando anticipo divisas como nuevo registro para: {row['DENOMINACION COMERCIAL']}")
            
            print(f"  [OK] Anticipos divisas compensados: {len(df_anticipos_divisas)} registros procesados")
            print(f"  [OK] Total DIVISAS después de compensación: {len(df_divisas)} registros")
    
    # Agregar columna de NEGOCIO basada en LINEA DE NEGOCIO
    if 'LINEA DE NEGOCIO' in df_divisas.columns:
        df_divisas['NEGOCIO'] = df_divisas['LINEA DE NEGOCIO'].apply(
            lambda x: TABLA_NEGOCIO_CANAL.get(str(x).strip().upper(), {'NEGOCIO': 'OTROS', 'CANAL': str(x)})['NEGOCIO']
        )
        
        # Eliminar columnas origen que ya no son necesarias
        columnas_a_eliminar = ['LINEA DE NEGOCIO', 'EMPRESA', 'ACTIVIDAD']
        for col in columnas_a_eliminar:
            if col in df_divisas.columns:
                df_divisas = df_divisas.drop(columns=[col])
    
    # No filtrar registros con saldos cero o negativos para mantener visibles los anticipos que superan la cartera
    # Los valores negativos deben mostrarse explícitamente en los informes financieros
    print(f"  [DEBUG] Registros DIVISAS antes de filtrar: {len(df_divisas)}")
    print(f"  [DEBUG] Se mantienen todos los registros incluyendo saldos negativos después de compensación")
    
    # Asegurar que los anticipos negativos se mantengan visibles
    if 'SALDO' in df_divisas.columns:
        saldo_negativo_count = len(df_divisas[df_divisas['SALDO'] < 0])
        print(f"  [DEBUG] Registros con saldos negativos en divisas: {saldo_negativo_count}")
        if saldo_negativo_count > 0:
            print(f"  [DEBUG] Ejemplo de saldo negativo: {df_divisas[df_divisas['SALDO'] < 0]['SALDO'].head().tolist()}")
    
    # Agrupar saldos por cliente (DENOMINACION COMERCIAL) para evitar duplicados
    if 'DENOMINACION COMERCIAL' in df_divisas.columns:
        # Mantener las columnas que no deben agruparse por cliente, sino que se tomarán del primer registro
        columnas_a_mantener = ['EMPRESA_2', 'CODIGO AGENTE', 'AGENTE', 'CODIGO COBRADOR', 'COBRADOR', 
                              'CODIGO CLIENTE', 'IDENTIFICACION', 'NOMBRE', 'DIRECCION', 'TELEFONO', 'CIUDAD', 
                              'NUMERO FACTURA', 'TIPO', 'FECHA', 'FECHA VTO', 'VALOR', 'DIAS VENCIDO']
        
        # Definir cómo se deben agrupar las columnas
        agg_dict = {}
        for col in df_divisas.columns:
            # No incluir la columna de agrupación en el diccionario de agregación
            if col == 'DENOMINACION COMERCIAL':
                continue  # Saltar la columna por la que vamos a agrupar
            if col in ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE']:
                agg_dict[col] = 'sum'
            elif col in columnas_a_mantener:
                agg_dict[col] = 'first'  # Tomar el primer valor disponible
            elif col in ['LINEA DE NEGOCIO', 'EMPRESA', 'ACTIVIDAD']:
                agg_dict[col] = 'first'  # Tomar el primer valor disponible
            elif col in ['NEGOCIO', 'CANAL']:
                agg_dict[col] = 'first'  # Tomar el primer valor disponible
            else:
                # Para otras columnas, usar first si no son columnas de saldo
                agg_dict[col] = 'first'
        
    
    # No agrupar por cliente si ya se aplicó compensación, para mantener la compensación por línea de negocio
    # La agrupación se hará en la hoja de vencimientos manteniendo las líneas de negocio
    
    total_divisas = safe_float_conversion(df_divisas['SALDO'].sum()) if 'SALDO' in df_divisas.columns else 0.0
    print(f"  [OK] Saldo total DIVISAS: ${total_divisas:,.0f}")
    
    # Crear copias de los DataFrames para las hojas PESOS y DIVISAS sin las columnas innecesarias
    df_pesos_salida = df_pesos.copy()
    df_divisas_salida = df_divisas.copy()
    
    # Eliminar columnas innecesarias que ya están representadas en otras columnas
    columnas_a_eliminar = ['LINEA DE NEGOCIO', 'EMPRESA', 'ACTIVIDAD']
    for col in columnas_a_eliminar:
        if col in df_pesos_salida.columns:
            df_pesos_salida = df_pesos_salida.drop(columns=[col])
        if col in df_divisas_salida.columns:
            df_divisas_salida = df_divisas_salida.drop(columns=[col])
    
    # Formatear fechas para mostrar solo la fecha sin la hora
    for df in [df_pesos_salida, df_divisas_salida]:
        for col in ['FECHA', 'FECHA VTO']:
            if col in df.columns:
                try:
                    df[col] = pd.to_datetime(df[col]).dt.strftime('%Y-%m-%d')
                except:
                    # Si no se puede convertir a datetime, dejar el valor como está
                    pass
    
    # Mover las columnas NEGOCIO y CANAL al inicio
    if 'NEGOCIO' in df_pesos_salida.columns and 'CANAL' in df_pesos_salida.columns:
        cols = df_pesos_salida.columns.tolist()
        cols = ['NEGOCIO', 'CANAL'] + [col for col in cols if col not in ['NEGOCIO', 'CANAL']]
        df_pesos_salida = df_pesos_salida[cols]
    elif 'NEGOCIO' in df_pesos_salida.columns:
        cols = df_pesos_salida.columns.tolist()
        cols = ['NEGOCIO'] + [col for col in cols if col != 'NEGOCIO']
        df_pesos_salida = df_pesos_salida[cols]
    elif 'CANAL' in df_pesos_salida.columns:
        cols = df_pesos_salida.columns.tolist()
        cols = ['CANAL'] + [col for col in cols if col != 'CANAL']
        df_pesos_salida = df_pesos_salida[cols]
    
    if 'NEGOCIO' in df_divisas_salida.columns and 'CANAL' in df_divisas_salida.columns:
        cols = df_divisas_salida.columns.tolist()
        cols = ['NEGOCIO', 'CANAL'] + [col for col in cols if col not in ['NEGOCIO', 'CANAL']]
        df_divisas_salida = df_divisas_salida[cols]
    elif 'NEGOCIO' in df_divisas_salida.columns:
        cols = df_divisas_salida.columns.tolist()
        cols = ['NEGOCIO'] + [col for col in cols if col != 'NEGOCIO']
        df_divisas_salida = df_divisas_salida[cols]
    elif 'CANAL' in df_divisas_salida.columns:
        cols = df_divisas_salida.columns.tolist()
        cols = ['CANAL'] + [col for col in cols if col != 'CANAL']
        df_divisas_salida = df_divisas_salida[cols]
    
    # HOJA 3: VENCIMIENTOS
    print("\n[6/6] Generando hoja VENCIMIENTOS...")
    df_vencimientos = crear_hoja_vencimientos(df_pesos, df_divisas, trm_dolar, trm_euro)
    print(f"  [OK] Registros en VENCIMIENTOS: {len(df_vencimientos)}")
    
    # Guardar Excel
    output_path = os.path.join(SALIDAS_DIR, output_file)
    print(f"\n[GUARDANDO] {output_path}")
    
    with pd.ExcelWriter(output_path, engine='xlsxwriter') as writer:
        workbook = writer.book
        fmt_miles = workbook.add_format({'num_format': '#,##0.00;-#,##0.00;"-";@'})
        fmt_texto = workbook.add_format({'num_format': '@'})  # Formato de texto
        
        # Hoja PESOS
        if not df_pesos_salida.empty:
            df_pesos_salida.to_excel(writer, sheet_name='PESOS', index=False)
            ws = writer.sheets['PESOS']
            for col in columnas_suma:
                if col in df_pesos_salida.columns:
                    idx = df_pesos_salida.columns.get_loc(col)
                    ws.set_column(idx, idx, 18, fmt_miles)
            # Formato para columnas de texto
            texto_cols = [i for i, col in enumerate(df_pesos_salida.columns) if col not in columnas_suma]
            for idx in texto_cols:
                ws.set_column(idx, idx, 20, fmt_texto)
        
        # Hoja DIVISAS
        if not df_divisas_salida.empty:
            df_divisas_salida.to_excel(writer, sheet_name='DIVISAS', index=False)
            ws = writer.sheets['DIVISAS']
            for col in columnas_suma:
                if col in df_divisas_salida.columns:
                    idx = df_divisas_salida.columns.get_loc(col)
                    ws.set_column(idx, idx, 18, fmt_miles)
            # Formato para columnas de texto
            texto_cols = [i for i, col in enumerate(df_divisas_salida.columns) if col not in columnas_suma]
            for idx in texto_cols:
                ws.set_column(idx, idx, 20, fmt_texto)
        
        # Hoja VENCIMIENTOS
        if not df_vencimientos.empty:
            # Guardar la hoja de vencimientos completa
            df_vencimientos.to_excel(writer, sheet_name='VENCIMIENTOS', index=False)
            ws = writer.sheets['VENCIMIENTOS']
            for col in columnas_suma:
                if col in df_vencimientos.columns:
                    idx = df_vencimientos.columns.get_loc(col)
                    ws.set_column(idx, idx, 18, fmt_miles)
            # Formato para columnas de texto
            texto_cols = [i for i, col in enumerate(df_vencimientos.columns) if col not in columnas_suma]
            for idx in texto_cols:
                ws.set_column(idx, idx, 20, fmt_texto)
    
    print("="*60)
    print(f"[OK] MODELO DE DEUDA GENERADO EXITOSAMENTE")
    print(f"  Archivo: {output_path}")
    print(f"  Total PESOS: ${total_pesos:,.0f}")
    print(f"  Total DIVISAS: ${total_divisas:,.0f}")
    print(f"  TRM USD: ${trm_dolar:,.2f} | EUR: ${trm_euro:,.2f}")
    print("="*60 + "\n")
    
    logging.info(f"Modelo de deuda generado: {output_path}")
    return output_path

def crear_hoja_vencimientos(df_pesos, df_divisas, trm_dolar, trm_euro):
    """Crea la hoja de vencimientos consolidada"""
    registros = []
    
    columnas_venc: list[str] = ['SALDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60',
                     'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360',
                     'DEUDA INCOBRABLE']
    
    # Combinar datos de pesos y divisas para procesamiento unificado
    todos_los_datos = []
    
    # Procesar PESOS
    if not df_pesos.empty:
        for _, row in df_pesos.iterrows():
            # Verificar que la fila tenga datos válidos
            denominacion = ''
            if not pd.isna(row['DENOMINACION COMERCIAL']):
                row_denom = str(row['DENOMINACION COMERCIAL'])
                if row_denom.lower() != 'nan':
                    denominacion = row_denom.strip()
            
            # Solo procesar filas con datos válidos
            if denominacion:
                # Incluir todas las columnas necesarias para la hoja de vencimientos
                registro_pesos = {
                    'DENOMINACION COMERCIAL': str(denominacion) if denominacion else '',
                    'MONEDA': 'pesos',  # Usar 'pesos' en lugar de códigos como 'CO'
                    'EMPRESA_2': str(row.get('EMPRESA_2', '')),
                    'CODIGO AGENTE': str(row.get('CODIGO AGENTE', '')),
                    'AGENTE': str(row.get('AGENTE', '')),
                    'CODIGO COBRADOR': str(row.get('CODIGO COBRADOR', '')),
                    'COBRADOR': str(row.get('COBRADOR', '')),
                    'CODIGO CLIENTE': str(row.get('CODIGO CLIENTE', '')),
                    'IDENTIFICACION': str(row.get('IDENTIFICACION', '')),
                    'NOMBRE': str(row.get('NOMBRE', '')),
                    'DIRECCION': str(row.get('DIRECCION', '')),
                    'TELEFONO': str(row.get('TELEFONO', '')),
                    'CIUDAD': str(row.get('CIUDAD', '')),
                    'NUMERO FACTURA': str(row.get('NUMERO FACTURA', '')),
                    'TIPO': str(row.get('TIPO', '')),
                    'FECHA': str(row.get('FECHA', '')),
                    'FECHA VTO': str(row.get('FECHA VTO', '')),
                    'VALOR': safe_float_conversion(row.get('VALOR', 0.0)),
                    'DIAS VENCIDO': safe_float_conversion(row.get('DIAS VENCIDO', 0.0))
                } # type: dict[str, str | float]
                for col in columnas_venc:
                    try:
                        if col in row and pd.notna(row[col]):
                            val = row[col]
                            try:
                                # Extract scalar value if it's a pandas type
                                if hasattr(val, 'item'):
                                    val = val.item()
                                elif pd.api.types.is_scalar(val):
                                    pass  # Already a scalar
                                else:
                                    val = pd.to_numeric(val, errors='coerce')
                                
                                numeric_val = safe_float_conversion(pd.to_numeric(val, errors='coerce'))
                                if not pd.isna(numeric_val):
                                    registro_pesos[col] = safe_float_conversion(numeric_val)
                                else:
                                    registro_pesos[col] = 0.0
                            except (ValueError, TypeError):
                                registro_pesos[col] = 0.0
                        else:
                            registro_pesos[col] = 0.0
                    except:
                        registro_pesos[col] = 0.0
                todos_los_datos.append(registro_pesos)
    
    # Procesar DIVISAS
    if not df_divisas.empty:
        for _, row in df_divisas.iterrows():
            # Verificar que la fila tenga datos válidos
            denominacion = ''
            if not pd.isna(row['DENOMINACION COMERCIAL']):
                row_denom = str(row['DENOMINACION COMERCIAL'])
                if row_denom.lower() != 'nan':
                    denominacion = row_denom.strip()
            
            # Solo procesar filas con datos válidos
            if denominacion:
                # Incluir todas las columnas necesarias para la hoja de vencimientos
                registro_divisas = {
                    'DENOMINACION COMERCIAL': str(denominacion) if denominacion else '',
                    'MONEDA': 'pesos',  # Usar 'pesos' en lugar de códigos como 'CO'
                    'EMPRESA_2': str(row.get('EMPRESA_2', '')),
                    'CODIGO AGENTE': str(row.get('CODIGO AGENTE', '')),
                    'AGENTE': str(row.get('AGENTE', '')),
                    'CODIGO COBRADOR': str(row.get('CODIGO COBRADOR', '')),
                    'COBRADOR': str(row.get('COBRADOR', '')),
                    'CODIGO CLIENTE': str(row.get('CODIGO CLIENTE', '')),
                    'IDENTIFICACION': str(row.get('IDENTIFICACION', '')),
                    'NOMBRE': str(row.get('NOMBRE', '')),
                    'DIRECCION': str(row.get('DIRECCION', '')),
                    'TELEFONO': str(row.get('TELEFONO', '')),
                    'CIUDAD': str(row.get('CIUDAD', '')),
                    'NUMERO FACTURA': str(row.get('NUMERO FACTURA', '')),
                    'TIPO': str(row.get('TIPO', '')),
                    'FECHA': str(row.get('FECHA', '')),
                    'FECHA VTO': str(row.get('FECHA VTO', '')),
                    'VALOR': safe_float_conversion(row.get('VALOR', 0.0)),
                    'DIAS VENCIDO': safe_float_conversion(row.get('DIAS VENCIDO', 0.0))
                } # type: dict[str, str | float]
                for col in columnas_venc:
                    try:
                        if col in row and pd.notna(row[col]):
                            val = row[col]
                            try:
                                # Extract scalar value if it's a pandas type
                                if hasattr(val, 'item'):
                                    val = val.item()
                                elif pd.api.types.is_scalar(val):
                                    pass  # Already a scalar
                                else:
                                    val = pd.to_numeric(val, errors='coerce')
                                
                                temp_val = safe_float_conversion(pd.to_numeric(val, errors='coerce'))
                                if not pd.isna(temp_val):
                                    valor_original = temp_val
                                    # Convertir a COP solo si es una moneda de divisa
                                    registro_divisas[col] = safe_float_conversion(valor_original)
                                else:
                                    registro_divisas[col] = 0.0
                            except (ValueError, TypeError):
                                registro_divisas[col] = 0.0
                        else:
                            registro_divisas[col] = 0.0
                    except:
                        registro_divisas[col] = 0.0
                todos_los_datos.append(registro_divisas)
    
    if not todos_los_datos:
        # Devolver DataFrame vacío con las columnas correctas
        columnas_vencimientos_1: list[str] = ['EMPRESA_2', 'DENOMINACION COMERCIAL', 'CODIGO AGENTE', 'AGENTE', 'CODIGO COBRADOR', 'COBRADOR', 'CODIGO CLIENTE', 'IDENTIFICACION', 'NOMBRE', 'DIRECCION', 'TELEFONO', 'CIUDAD', 'NUMERO FACTURA', 'TIPO', 'FECHA', 'FECHA VTO', 'VALOR', 'SALDO', 'DIAS VENCIDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE', 'MONEDA', 'NEGOCIO', 'CANAL']
        df_vacio = pd.DataFrame({col: [] for col in columnas_vencimientos_1})
        # Asegurar tipos adecuados para columnas de vencimiento (numéricas)
        for col in columnas_venc:
            if col in df_vacio.columns:
                df_vacio[col] = pd.Series(dtype='float64')
        # Asegurar tipos adecuados para columnas de texto
        for col in ['EMPRESA_2', 'DENOMINACION COMERCIAL', 'CODIGO AGENTE', 'AGENTE', 'CODIGO COBRADOR', 'COBRADOR', 'CODIGO CLIENTE', 'IDENTIFICACION', 'NOMBRE', 'DIRECCION', 'TELEFONO', 'CIUDAD', 'NUMERO FACTURA', 'TIPO', 'FECHA', 'FECHA VTO', 'MONEDA', 'NEGOCIO', 'CANAL']:
            if col in df_vacio.columns:
                df_vacio[col] = pd.Series(dtype='object')
        return df_vacio
    
    # Crear DataFrame de trabajo
    df_trabajo = pd.DataFrame(todos_los_datos)
    
    # Asegurar tipos adecuados para columnas de vencimiento (numéricas)
    for col in columnas_venc:
        if col in df_trabajo.columns:
            df_trabajo[col] = pd.to_numeric(df_trabajo[col], errors='coerce')
    
    print(f"  [DEBUG] Registros de trabajo: {len(df_trabajo)}")
    
    # Agrupar por cliente para consolidar saldos
    if not df_trabajo.empty:
        print(f"  [DEBUG] Columnas de trabajo: {list(df_trabajo.columns)}")
        # Agrupar y sumar solo las columnas de vencimiento
        columnas_agrupacion = [col for col in df_trabajo.columns if col not in columnas_venc]
        agg_dict = {col: 'first' for col in columnas_agrupacion}
        for col in columnas_venc:
            agg_dict[col] = 'sum'
        print(f"  [DEBUG] Agrupando por: {columnas_agrupacion}")
        print(f"  [DEBUG] Agregando: {agg_dict}")
        df_agrupado = df_trabajo.groupby(columnas_agrupacion, as_index=False).agg(agg_dict)
        print(f"  [DEBUG] Registros agrupados: {len(df_agrupado)}")
        
        # Crear registros finales
        for _, row in df_agrupado.iterrows():
            denominacion = ''
            if not safe_isna_check(row['DENOMINACION COMERCIAL']):
                row_denom = str(row['DENOMINACION COMERCIAL'])
                if row_denom.lower() != 'nan':
                    denominacion = row_denom.strip()
            
            registro_final = {
                'EMPRESA_2': str(row.get('EMPRESA_2', '')),
                'DENOMINACION COMERCIAL': str(denominacion) if denominacion else '',
                'CODIGO AGENTE': str(row.get('CODIGO AGENTE', '')),
                'AGENTE': str(row.get('AGENTE', '')),
                'CODIGO COBRADOR': str(row.get('CODIGO COBRADOR', '')),
                'COBRADOR': str(row.get('COBRADOR', '')),
                'CODIGO CLIENTE': str(row.get('CODIGO CLIENTE', '')),
                'IDENTIFICACION': str(row.get('IDENTIFICACION', '')),
                'NOMBRE': str(row.get('NOMBRE', '')),
                'DIRECCION': str(row.get('DIRECCION', '')),
                'TELEFONO': str(row.get('TELEFONO', '')),
                'CIUDAD': str(row.get('CIUDAD', '')),
                'NUMERO FACTURA': str(row.get('NUMERO FACTURA', '')),
                'TIPO': str(row.get('TIPO', '')),
                'FECHA': str(row.get('FECHA', '')),
                'FECHA VTO': str(row.get('FECHA VTO', '')),
                'VALOR': safe_float_conversion(row.get('VALOR', 0.0)),
                'SALDO': safe_float_conversion(row.get('SALDO', 0.0)),
                'DIAS VENCIDO': safe_float_conversion(row.get('DIAS VENCIDO', 0.0)),
                'MONEDA': str(row.get('MONEDA', 'pesos')),
                'NEGOCIO': str(row.get('NEGOCIO', 'OTROS')),
                'CANAL': str(row.get('CANAL', 'OTROS'))
            }
            
            # Agregar columnas de vencimiento
            for col in columnas_venc:
                try:
                    val = row[col]
                    if safe_isna_check(val):
                        registro_final[col] = 0.0
                    else:
                        registro_final[col] = safe_float_conversion(val)
                except:
                    registro_final[col] = 0.0
            
            registros.append(registro_final)
    
    if not registros:
        # Devolver DataFrame vacío con las columnas correctas
        columnas_vencimientos_2: list[str] = ['EMPRESA_2', 'DENOMINACION COMERCIAL', 'CODIGO AGENTE', 'AGENTE', 'CODIGO COBRADOR', 'COBRADOR', 'CODIGO CLIENTE', 'IDENTIFICACION', 'NOMBRE', 'DIRECCION', 'TELEFONO', 'CIUDAD', 'NUMERO FACTURA', 'TIPO', 'FECHA', 'FECHA VTO', 'VALOR', 'SALDO', 'DIAS VENCIDO', 'SALDO NO VENCIDO', 'VENCIDO 30', 'VENCIDO 60', 'VENCIDO 90', 'VENCIDO 180', 'VENCIDO 360', 'VENCIDO + 360', 'DEUDA INCOBRABLE', 'MONEDA', 'NEGOCIO', 'CANAL']
        df_vacio = pd.DataFrame({col: [] for col in columnas_vencimientos_2})
        # Asegurar tipos adecuados para columnas de vencimiento (numéricas)
        for col in columnas_venc:
            if col in df_vacio.columns:
                df_vacio[col] = pd.Series(dtype='float64')
        return df_vacio
    
    df_venc = pd.DataFrame(registros)
    
    # Mover las columnas NEGOCIO y CANAL al final como se requiere
    if 'NEGOCIO' in df_venc.columns and 'CANAL' in df_venc.columns:
        cols = [col for col in df_venc.columns if col not in ['NEGOCIO', 'CANAL']] + ['NEGOCIO', 'CANAL']
        df_venc = df_venc[cols]
    
    # Agregar totales por moneda
    totales = []
    
    if 'MONEDA' in df_venc.columns:
        try:
            for moneda in df_venc['MONEDA'].dropna().unique():
                moneda_str = str(moneda).strip()
                if not moneda_str:
                    continue
                    
                df_moneda = df_venc[df_venc['MONEDA'] == moneda]
                total = {
                    'EMPRESA_2': '', 'DENOMINACION COMERCIAL': '** **',
                    'CODIGO AGENTE': '', 'AGENTE': '', 'CODIGO COBRADOR': '',
                    'COBRADOR': '', 'CODIGO CLIENTE': '', 'IDENTIFICACION': '',
                    'NOMBRE': '', 'DIRECCION': '', 'TELEFONO': '', 'CIUDAD': '',
                    'NUMERO FACTURA': '', 'TIPO': '', 'FECHA': '', 'FECHA VTO': '',
                    'VALOR': 0.0, 'SALDO': 0.0, 'DIAS VENCIDO': 0.0,
                    'MONEDA': moneda_str, 'NEGOCIO': '', 'CANAL': ''
                }
                for col in columnas_venc:
                    try:
                        total[col] = safe_float_conversion(df_moneda[col].sum()) if len(df_moneda) > 0 else 0.0
                    except:
                        total[col] = 0.0
                totales.append(total)
        except Exception as e:
            print(f"[WARN] Error calculando totales por moneda: {e}")
    
    # Total general
    total_general = {
        'EMPRESA_2': '', 'DENOMINACION COMERCIAL': '**Totales**',
        'CODIGO AGENTE': '', 'AGENTE': '', 'CODIGO COBRADOR': '',
        'COBRADOR': '', 'CODIGO CLIENTE': '', 'IDENTIFICACION': '',
        'NOMBRE': '', 'DIRECCION': '', 'TELEFONO': '', 'CIUDAD': '',
        'NUMERO FACTURA': '', 'TIPO': '', 'FECHA': '', 'FECHA VTO': '',
        'VALOR': 0.0, 'SALDO': 0.0, 'DIAS VENCIDO': 0.0,
        'MONEDA': '', 'NEGOCIO': '', 'CANAL': ''
    }
    for col in columnas_venc:
        if col in df_venc.columns:
            try:
                total_general[col] = safe_float_conversion(df_venc[col].sum()) if len(df_venc) > 0 else 0.0
            except:
                total_general[col] = 0.0
        else:
            total_general[col] = 0.0
    totales.append(total_general)
    
    # Concatenar registros con totales
    if totales:
        df_totales = pd.DataFrame(totales)
        return pd.concat([df_venc, df_totales], ignore_index=True)
    else:
        return df_venc

def main():
    """Entrada principal"""
    parser = argparse.ArgumentParser(description="Genera Modelo de Deuda desde cartera y anticipos procesados")
    parser.add_argument("archivo_provision", help="Archivo CARTERA_*.xlsx procesado")
    parser.add_argument("archivo_anticipos", help="Archivo ANTICIPOS_*.xlsx procesado")
    parser.add_argument("-o", "--output-file", help="Nombre del archivo de salida")
    
    args = parser.parse_args()
    
    # Si no se proporciona un nombre de archivo, generar uno con el formato MODELO_DEUDA_fecha_hora
    if not args.output_file:
        from datetime import datetime
        timestamp = datetime.now().strftime("%Y%m%d_%H%M%S")
        args.output_file = f"MODELO_DEUDA_{timestamp}.xlsx"
    
    try:
        resultado = crear_modelo_deuda(args.archivo_provision, args.archivo_anticipos, args.output_file)
        if USE_UNIFIED_LOGGING:
            log_fin_proceso("MODELO_DEUDA", args.output_file)
    except Exception as e:
        if USE_UNIFIED_LOGGING:
            log_error_proceso("MODELO_DEUDA", str(e))
        else:
            print(f"\n[ERROR] ERROR: {e}")
            logging.error(f"Error: {e}", exc_info=True)
        sys.exit(1)

if __name__ == "__main__":
    main()