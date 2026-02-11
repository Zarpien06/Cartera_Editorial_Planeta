<?php
// Configuración de la aplicación
define('ROOT_PATH', dirname(__DIR__));

// Ruta al ejecutable de Python en el entorno virtual (ruta relativa para portabilidad)
define('PYTHON_PATH', __DIR__ . '/../../venv/Scripts/python.exe');

// Directorios
define('UPLOAD_DIR', ROOT_PATH . '/uploads');
// Usar el directorio de salidas de Python_principales
define('OUTPUT_DIR', dirname(__DIR__) . '/../Python_principales/salidas');
define('BASE_DIR', dirname(__DIR__) . '/..');

// Asegurar que OUTPUT_DIR exista
if (!file_exists(OUTPUT_DIR)) {
    @mkdir(OUTPUT_DIR, 0777, true);
}

// Configuración de logs
define('LOG_PATH', ROOT_PATH . '/logs/error.log');

// Configuración de la API (no se usa en modo PHP directo)
define('API_KEY', 'no_usado_en_modo_php');

// Rutas a los scripts de Python en Python_principales
define('PY_SCRIPTS', [
    'cartera' => __DIR__ . '/../../Python_principales/procesador_cartera.py',
    'anticipos' => __DIR__ . '/../../Python_principales/procesador_anticipos.py',
    'modelo_deuda' => __DIR__ . '/../../Python_principales/modelo_deuda.py',
    'focus' => __DIR__ . '/../../Python_principales/procesar_y_actualizar_focus.py'
]);

// Parámetros requeridos para cada script
define('SCRIPT_PARAMS', [
    'cartera' => ['input_file', 'fecha_cierre'],
    'anticipos' => ['input_file'],
    'modelo_deuda' => ['archivo_provision', 'archivo_anticipos'],
    'focus' => ['archivo_focus', 'archivo_balance', 'archivo_situacion', 'archivo_modelo']
]);

// Crear directorios si no existen
$directories = [UPLOAD_DIR, OUTPUT_DIR, dirname(LOG_PATH)];
foreach ($directories as $dir) {
    if (!file_exists($dir)) {
        mkdir($dir, 0777, true);
    }
}

// Configuración de errores
ini_set('display_errors', 0);
ini_set('log_errors', 1);
ini_set('error_log', LOG_PATH);

error_reporting(E_ALL);

// Función para registrar errores
function log_error($message) {
    $timestamp = date('Y-m-d H:i:s');
    $log_message = "[$timestamp] $message" . PHP_EOL;
    error_log($log_message, 3, LOG_PATH);
}