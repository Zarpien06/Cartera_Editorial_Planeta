"""
log_bridge - utilitario para enviar logs de los scripts Python
al archivo compartido logs/system_log.txt utilizado por PHP.
"""

from __future__ import annotations

import json
import logging
import threading
from datetime import datetime
from pathlib import Path
from typing import Any, Dict, Optional

BASE_DIR = Path(__file__).resolve().parents[1]
LOG_DIR = BASE_DIR / "front_php" / "logs"
LOG_FILE = LOG_DIR / "system_log.txt"
_lock = threading.Lock()


def _ensure_log_dir() -> None:
    LOG_DIR.mkdir(parents=True, exist_ok=True)


def log_event(source: str, level: str, message: str, extra: Optional[Dict[str, Any]] = None) -> None:
    """
    Escribe una entrada JSON por línea en el log compartido.
    """
    _ensure_log_dir()
    entry = {
        "timestamp": datetime.utcnow().isoformat() + "Z",
        "source": source,
        "level": level.upper(),
        "message": message,
        "context": extra or {},
    }
    line = json.dumps(entry, ensure_ascii=False)
    with _lock:
        with LOG_FILE.open("a", encoding="utf-8") as fh:
            fh.write(line + "\n")


class CentralLogHandler(logging.Handler):
    """
    Handler para conectar el módulo logging estándar con el log compartido.
    """

    def __init__(self, component: str) -> None:
        super().__init__()
        self.component = component

    def emit(self, record: logging.LogRecord) -> None:
        try:
            message = self.format(record)
            if message == record.getMessage():
                # message no contiene nivel/fecha, mantenemos texto limpio
                pass
            extra = {
                "filename": record.filename,
                "lineno": record.lineno,
                "func": record.funcName,
                "module": record.module,
                "pathname": record.pathname,
            }
            # Agregar información adicional del registro si está disponible
            if hasattr(record, 'process'):
                extra["process"] = record.process
            if hasattr(record, 'thread'):
                extra["thread"] = record.thread
                
            log_event(self.component, record.levelname, message, extra)
        except Exception:
            # Evitar que un fallo de logging rompa el procesamiento principal
            pass


def attach_python_logging(component: str, logger: Optional[logging.Logger] = None) -> None:
    """
    Adjunta el handler central al logger indicado (o al logger raíz por defecto).
    """
    handler = CentralLogHandler(component)
    handler.setLevel(logging.INFO)
    # Formato más detallado para incluir toda la información relevante
    formatter = logging.Formatter('[%(levelname)s] %(message)s [%(filename)s:%(lineno)d]')
    handler.setFormatter(formatter)

    target_logger = logger if logger is not None else logging.getLogger()
    # Evitar multiplicar handlers si se llama más de una vez
    already_attached = any(isinstance(h, CentralLogHandler) and h.component == component for h in target_logger.handlers)
    if not already_attached:
        target_logger.addHandler(handler)
