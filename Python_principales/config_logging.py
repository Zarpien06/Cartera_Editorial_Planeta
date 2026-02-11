# -*- coding: utf-8 -*-
"""
Configuración unificada de logging para todos los scripts del sistema
"""

import os
import logging
import threading
from pathlib import Path
from datetime import datetime

# Directorios
BASE_DIR = Path(__file__).resolve().parent
LOGS_DIR = BASE_DIR / "logs_unificados"
LOGS_DIR.mkdir(exist_ok=True)

# Archivo de log unificado
UNIFIED_LOG_FILE = LOGS_DIR / "procesamiento_completo.log"

# Lock para threading seguro
_lock = threading.Lock()

def setup_unified_logging():
    """Configura el sistema de logging unificado"""
    # Crear logger principal
    logger = logging.getLogger('SISTEMA_CARTERA')
    logger.setLevel(logging.DEBUG)
    
    # Evitar handlers duplicados
    if logger.handlers:
        logger.handlers.clear()
    
    # Handler para archivo unificado
    file_handler = logging.FileHandler(UNIFIED_LOG_FILE, encoding='utf-8')
    file_handler.setLevel(logging.INFO)
    
    # Handler para consola
    console_handler = logging.StreamHandler()
    console_handler.setLevel(logging.INFO)
    
    # Formato sin fecha/hora para mejor legibilidad
    formatter = logging.Formatter(
        '[%(name)s] [%(levelname)s] %(message)s [%(filename)s:%(lineno)d]'
    )
    
    file_handler.setFormatter(formatter)
    console_handler.setFormatter(formatter)
    
    logger.addHandler(file_handler)
    logger.addHandler(console_handler)
    
    return logger

def log_detalle_proceso(componente, proceso, detalles):
    """
    Función para registrar detalles de procesos con mayor claridad
    
    Args:
        componente (str): Nombre del componente (CARTERA, ANTICIPOS, MODELO_DEUDA)
        proceso (str): Nombre del proceso
        detalles (dict): Diccionario con detalles del proceso
    """
    logger = logging.getLogger('SISTEMA_CARTERA')
    
    mensaje = f"[{componente}] [{proceso}] "
    
    # Agregar detalles en formato legible
    detalle_str = " | ".join([f"{k}: {v}" for k, v in detalles.items()])
    mensaje += detalle_str
    
    logger.info(mensaje)

def log_inicio_proceso(componente, archivo_entrada):
    """Registra el inicio de un proceso"""
    log_detalle_proceso(
        componente, 
        "INICIO", 
        {
            "archivo": archivo_entrada,
            "timestamp": datetime.now().strftime("%H:%M:%S")
        }
    )

def log_fin_proceso(componente, archivo_salida, estadisticas=None):
    """Registra la finalización de un proceso"""
    detalles = {
        "archivo_salida": archivo_salida,
        "timestamp": datetime.now().strftime("%H:%M:%S")
    }
    
    if estadisticas:
        detalles.update(estadisticas)
    
    log_detalle_proceso(componente, "FINALIZADO", detalles)

def log_error_proceso(componente, error, contexto=None):
    """Registra un error en el proceso"""
    logger = logging.getLogger('SISTEMA_CARTERA')
    
    mensaje = f"[{componente}] [ERROR] {str(error)}"
    if contexto:
        mensaje += f" | Contexto: {contexto}"
    
    logger.error(mensaje)

# Inicializar logging al importar
logger = setup_unified_logging()