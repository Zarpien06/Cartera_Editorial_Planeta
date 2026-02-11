import os
import sys
import json
import time
from pathlib import Path
from datetime import datetime
from typing import Dict, Any, Tuple, Optional

# Configuración de rutas
BASE_DIR = Path(__file__).parent.absolute()
TRM_FILE = BASE_DIR / 'trm.json'

# Configuración de logging
import logging
logging.basicConfig(
    level=logging.INFO,
    format='%(asctime)s - %(name)s - %(levelname)s - %(message)s',
    handlers=[
        logging.StreamHandler(),
        logging.FileHandler(BASE_DIR / 'trm_config.log')
    ]
)
logger = logging.getLogger('trm_config')

def _ensure_trm_file_exists() -> None:
    """Asegura que el archivo TRM exista con valores por defecto."""
    if not TRM_FILE.exists():
        default_data = {
            'usd': 4000.0,
            'eur': 4500.0,
            'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'actualizado_por': 'sistema',
            'origen': 'default',
            'version': '1.0'
        }
        try:
            TRM_FILE.parent.mkdir(parents=True, exist_ok=True)
            with open(TRM_FILE, 'w', encoding='utf-8') as f:
                json.dump(default_data, f, indent=2, ensure_ascii=False)
            
            # Establecer permisos en sistemas Unix
            if os.name == 'posix':
                TRM_FILE.chmod(0o666)
                
            logger.info(f"Archivo TRM creado en {TRM_FILE}")
        except Exception as e:
            logger.error(f"Error al crear archivo TRM: {e}", exc_info=True)
            raise

def load_trm() -> Dict[str, Any]:
    """
    Carga los valores de TRM desde el archivo JSON.
    
    Returns:
        Dict con las claves: usd, eur, fecha, actualizado_por, origen
    """
    max_attempts = 3
    attempt = 0
    
    while attempt < max_attempts:
        try:
            _ensure_trm_file_exists()
            
            # Limpiar caché del sistema de archivos
            if hasattr(os, 'stat'):
                try:
                    os.stat(str(TRM_FILE))
                except:
                    pass
            
            with open(TRM_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
            
            # Validar y normalizar datos
            if not isinstance(data, dict):
                raise ValueError("Formato de archivo TRM inválido: no es un objeto JSON")
                
            # Validar campos numéricos
            usd = float(data.get('usd', 4000.0))
            eur = float(data.get('eur', 4500.0))
            
            if usd <= 0 or eur <= 0:
                raise ValueError("Valores de TRM deben ser mayores a cero")
            
            result = {
                'usd': usd,
                'eur': eur,
                'fecha': data.get('fecha', datetime.now().strftime('%Y-%m-%d %H:%M:%S')),
                'actualizado_por': data.get('actualizado_por', 'sistema'),
                'origen': data.get('origen', 'json'),
                'version': data.get('version', '1.0')
            }
            
            logger.debug(f"TRM cargada exitosamente: {result}")
            return result
            
        except json.JSONDecodeError as e:
            logger.warning(f"Error de decodificación JSON (intento {attempt + 1}/{max_attempts}): {e}")
            # Si el archivo está corrupto, intentar repararlo
            if TRM_FILE.exists():
                backup_file = TRM_FILE.with_suffix(f'.corrupt.{int(time.time())}.json')
                try:
                    TRM_FILE.rename(backup_file)
                    logger.warning(f"Archivo TRM corrupto respaldado en {backup_file}")
                except Exception as backup_err:
                    logger.error(f"Error al respaldar archivo TRM corrupto: {backup_err}")
            time.sleep(0.1)  # Pequeña pausa entre intentos
            continue
            
        except Exception as e:
            print(f"Error inesperado al cargar TRM: {e}")
            if attempt == max_attempts - 1:  # Último intento
                break
            time.sleep(0.1)  # Pequeña pausa entre intentos
            continue
    
    # Si llegamos aquí, todos los intentos fallaron
    return {
        'usd': 4000.0,
        'eur': 4500.0,
        'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
        'actualizado_por': 'sistema',
        'origen': 'error'
    }

def save_trm(usd, eur, actualizado_por='sistema'):
    """
    Guarda los valores de TRM en el archivo JSON.
    
    Args:
        usd (float): Tasa de cambio USD a COP
        eur (float): Tasa de cambio EUR a COP
        actualizado_por (str): Quién está realizando la actualización
        
    Returns:
        str: Ruta al archivo TRM
    """
    _ensure_trm_file_exists()
    
    try:
        # Cargar datos existentes para preservar otros campos
        try:
            with open(TRM_FILE, 'r', encoding='utf-8') as f:
                data = json.load(f)
        except:
            data = {}
        
        # Actualizar solo los campos necesarios
        data.update({
            'usd': float(usd) if usd is not None else 4000.0,
            'eur': float(eur) if eur is not None else 4500.0,
            'fecha': datetime.now().strftime('%Y-%m-%d %H:%M:%S'),
            'actualizado_por': actualizado_por,
            'origen': 'sistema'
        })
        
        # Guardar en el archivo
        with open(TRM_FILE, 'w', encoding='utf-8') as f:
            json.dump(data, f, ensure_ascii=False, indent=2)
        
        # Asegurar permisos
        os.chmod(TRM_FILE, 0o666)
        
        # Limpiar caché del sistema de archivos
        if hasattr(os, 'stat'):
            try:
                os.stat(str(TRM_FILE))
            except:
                pass
        
        return TRM_FILE
        
    except Exception as e:
        print(f"Error al guardar TRM: {e}")
        raise


def parse_trm_value(text, fallback=None):
    """
    Parsea valores de TRM ingresados por el usuario con formato LATAM,
    convirtiéndolos a float.
    
    Args:
        text (str): Texto a parsear (ej: '4.500,25' o '4500.25')
        fallback: Valor a devolver si el parseo falla
        
    Returns:
        float o valor de fallboack
    """
    if text is None:
        return fallback
        
    if isinstance(text, (int, float)):
        return float(text)
        
    if not isinstance(text, str):
        return fallback
        
    text = text.strip()
    if not text:
        return fallback
        
    try:
        # Intentar convertir directamente (para formato con punto decimal)
        if '.' in text and ',' not in text:
            return float(text)
            
        # Formato latinoamericano (1.234,56)
        if ',' in text and text.count('.') > 0:
            # Eliminar puntos de miles y reemplazar coma por punto
            return float(text.replace('.', '').replace(',', '.'))
            
        # Formato con coma decimal (1234,56)
        if ',' in text:
            return float(text.replace(',', '.'))
            
        # Formato estándar
        return float(text)
        
    except (ValueError, TypeError):
        return fallback


def format_trm_display(value, decimals=2):
    """
    Formatea la TRM para mostrarla con separador de miles y decimales.
    
    Args:
        value: Valor numérico a formatear
        decimals (int): Número de decimales a mostrar (por defecto: 2)
        
    Returns:
        str: Valor formateado (ej: '4.500,25')
    """
    if value is None:
        return "N/A"
        
    try:
        num = float(value)
        # Formatear con separadores de miles y coma decimal
        return "{:,.{}f}".format(num, decimals).replace(",", "X").replace(".", ",").replace("X", ".")
    except (ValueError, TypeError):
        return str(value)


def get_trm_usd():
    """
    Obtiene únicamente el valor de TRM para USD.
    
    Returns:
        float: Valor de TRM para USD o 4000.0 por defecto
    """
    try:
        return load_trm()['usd']
    except:
        return 4000.0


def get_trm_eur():
    """
    Obtiene únicamente el valor de TRM para EUR.
    
    Returns:
        float: Valor de TRM para EUR o 4500.0 por defecto
    """
    try:
        return load_trm()['eur']
    except:
        return 4500.0