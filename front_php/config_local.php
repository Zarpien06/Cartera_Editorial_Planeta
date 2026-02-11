<?php
/**
 * Configuración local para procesamiento PHP directo
 */

// Ruta a Python - CAMBIAR SEGÚN TU INSTALACIÓN
// Para WAMP en Windows, verificar dónde está instalado Python
// Opciones comunes:
// - 'python' o 'python3' si está en el PATH de Windows
// - 'C:\\Python311\\python.exe' - Python 3.11
// - 'C:\\Python39\\python.exe' - Python 3.9
// - 'C:\\Users\\TU_USUARIO\\AppData\\Local\\Programs\\Python\\Python311\\python.exe'

// AUTO-DETECTAR: Intentar usar el entorno virtual del proyecto primero
$venv_python = dirname(__DIR__) . '\\.venv\\Scripts\\python.exe';

if (file_exists($venv_python)) {
    // Usar Python del entorno virtual (RECOMENDADO)
    define('LOCAL_PYTHON_PATH', $venv_python);
} else {
    // Fallback: Intentar usar Python del sistema
    // NOTA: Si esto falla, ejecutar primero: instalar_sistema_completo.bat
    define('LOCAL_PYTHON_PATH', 'python');
}

// Directorios base
define('LOCAL_BASE_DIR', dirname(__DIR__));
define('LOCAL_BACKEND_DIR', LOCAL_BASE_DIR . '/Python_principales');
define('LOCAL_SALIDAS_DIR', LOCAL_BACKEND_DIR . '/salidas');
define('LOCAL_UPLOADS_DIR', __DIR__ . '/uploads');

// Scripts Python
define('LOCAL_PY_SCRIPTS', [
    'cartera' => LOCAL_BACKEND_DIR . '/procesador_cartera.py',
    'anticipos' => LOCAL_BACKEND_DIR . '/procesador_anticipos.py',
    'modelo_deuda' => LOCAL_BACKEND_DIR . '/modelo_deuda.py',
    'focus' => LOCAL_BACKEND_DIR . '/procesar_y_actualizar_focus.py'
]);

// Crear directorios si no existen
if (!is_dir(LOCAL_SALIDAS_DIR)) {
    @mkdir(LOCAL_SALIDAS_DIR, 0777, true);
}
if (!is_dir(LOCAL_UPLOADS_DIR)) {
    @mkdir(LOCAL_UPLOADS_DIR, 0777, true);
}

// Función para verificar que Python esté disponible
function verificar_python() {
    $command = LOCAL_PYTHON_PATH . ' --version 2>&1';
    exec($command, $output, $returnCode);
    
    if ($returnCode !== 0) {
        throw new Exception('Python no est\u00e1 disponible. Verificar LOCAL_PYTHON_PATH en config_local.php');
    }
    
    return implode("\n", $output);
}

// Función para ejecutar script Python
function ejecutar_python($scriptPath, $args = []) {
    if (!file_exists($scriptPath)) {
        throw new Exception("Script no encontrado: $scriptPath");
    }
    
    // En Windows/WAMP, asegurarse de que las rutas usen comillas correctamente
    $isWindows = (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN');
    
    if ($isWindows) {
        // Windows: usar comillas dobles para todas las rutas
        $command = '"' . LOCAL_PYTHON_PATH . '" ' . $scriptPath;
        
        foreach ($args as $arg) {
            // Siempre poner entre comillas los argumentos en Windows
            $command .= ' "' . $arg . '"';
        }
    } else {
        // Linux/Mac: usar escapeshellarg
        $command = escapeshellarg(LOCAL_PYTHON_PATH) . ' ' . escapeshellarg($scriptPath);
        
        foreach ($args as $arg) {
            $command .= ' ' . escapeshellarg($arg);
        }
    }
    
    $command .= ' 2>&1';
    
    // Registrar información de depuración
    error_log("Comando a ejecutar: " . $command);
    
    // Ejecutar y capturar salida
    exec($command, $output, $returnCode);
    
    $outputText = implode("\n", $output);
    
    // Registrar resultados
    error_log("Resultado de ejecución - Código de salida: " . $returnCode);
    error_log("Salida: " . $outputText);
    
    return [
        'success' => ($returnCode === 0),
        'output' => $outputText,
        'exit_code' => $returnCode,
        'command' => $command
    ];
}
