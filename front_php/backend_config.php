<?php
/**
 * Configuración de backends
 * 
 * Prioridad de backends:
 * 1. Backend PHP (local) - Usado para TRM y operaciones rápidas
 * 2. Backend Python (FastAPI) - Usado para procesamiento pesado
 */

// Usar ruta absoluta para evitar problemas
if (!defined('BASE_DIR')) {
    define('BASE_DIR', dirname(__DIR__));
}

// Asegurar que el directorio Python_principales exista
$pythonDir = BASE_DIR . '/Python_principales';
if (!is_dir($pythonDir)) {
    @mkdir($pythonDir, 0755, true);
}

// Asegurar que el archivo TRM exista
try {
    $trmFile = $pythonDir . '/trm.json';
    if (!file_exists($trmFile)) {
        $defaultData = [
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'default'
        ];
        file_put_contents($trmFile, json_encode($defaultData, JSON_PRETTY_PRINT));
        @chmod($trmFile, 0666);
    }
    // Eliminamos el uso de clearstatcache para evitar cache
} catch (Exception $e) {
    error_log('Error al inicializar TRM: ' . $e->getMessage());
}

// Configuración de backends
$backends = [
    'php' => [
        'url' => '', // URL vacía para rutas relativas
        'type' => 'php',
        'priority' => 1,
        'trm_file' => BASE_DIR . '/Python_principales/trm.json' // Ruta al archivo TRM
    ],
    'python' => [
        'url' => getPythonBackendUrl(),
        'type' => 'python',
        'priority' => 2
    ]
];

// Obtener la URL del backend Python
function getPythonBackendUrl() {
    // 1) Variable de entorno
    $envUrl = getenv('BACKEND_URL');
    if (!empty($envUrl)) {
        return rtrim($envUrl, '/');
    }
    
    // 2) Archivo local
    $fileUrl = null;
    $filePath = __DIR__ . DIRECTORY_SEPARATOR . 'backend_url.txt';
    if (is_file($filePath)) {
        $raw = trim(@file_get_contents($filePath));
        if (!empty($raw)) {
            $fileUrl = rtrim($raw, '/');
            return $fileUrl;
        }
    }
    
    // 3) Cookie
    $cookieUrl = null;
    if (isset($_COOKIE['backend_url']) && !empty(trim($_COOKIE['backend_url']))) {
        $cookieUrl = rtrim(trim($_COOKIE['backend_url']), '/');
        return $cookieUrl;
    }
    
    // 4) Fallback: mismo host con puerto configurable (default 8000)
    $isHttps = (!empty($_SERVER['HTTPS']) && $_SERVER['HTTPS'] !== 'off') || 
               (isset($_SERVER['SERVER_PORT']) && $_SERVER['SERVER_PORT'] == 443);
    $scheme = $isHttps ? 'https' : 'http';
    $hostHeader = $_SERVER['HTTP_HOST'] ?? ($_SERVER['SERVER_NAME'] ?? '127.0.0.1');
    $hostname = strpos($hostHeader, ':') !== false ? 
        explode(':', $hostHeader)[0] : $hostHeader;
    
    $port = getenv('BACKEND_PORT') ?: '8000';
    $fallbackUrl = $scheme . '://' . $hostname . ':' . $port;
    
    // Usar la URL de fallback como último recurso
    return $fallbackUrl;
}

// Determinar qué backend usar
function getActiveBackend() {
    global $backends;
    
    // Verificar si el backend PHP está disponible
    $phpBackendAvailable = checkPhpBackendAvailability();
    
    // Si el backend PHP está disponible, usarlo
    if ($phpBackendAvailable) {
        return $backends['php'];
    }
    
    // Si no, usar el backend Python
    return $backends['python'];
}

// Verificar disponibilidad del backend PHP
function checkPhpBackendAvailability() {
    // Verificar si el archivo api.php existe y es accesible
    $apiFile = __DIR__ . '/api.php';
    if (!file_exists($apiFile) || !is_readable($apiFile)) {
        return false;
    }
    
    // Opcional: Hacer una prueba de conexión al backend PHP
    // (puedes descomentar esto si quieres una verificación más estricta)
    /*
    $testUrl = 'http' . (isset($_SERVER['HTTPS']) ? 's' : '') . '://' . 
              $_SERVER['HTTP_HOST'] . dirname($_SERVER['PHP_SELF']) . '/api.php';
    $response = @file_get_contents($testUrl . '?test=1');
    return ($response !== false);
    */
    
    return true;
}

// Función para obtener la TRM actual
function getCurrentTRM() {
    global $backends;
    
    // Ruta al archivo TRM
    $trmFile = $backends['php']['trm_file'];
    
    // Leer directamente del archivo sin usar cache
    // Si el archivo no existe, crearlo con valores por defecto
    if (!file_exists($trmFile)) {
        $defaultTRM = [
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'default'
        ];
        
        // Asegurar que el directorio existe
        $dir = dirname($trmFile);
        if (!is_dir($dir)) {
            mkdir($dir, 0755, true);
        }
        
        // Guardar TRM por defecto
        file_put_contents($trmFile, json_encode($defaultTRM, JSON_PRETTY_PRINT));
        @chmod($trmFile, 0666);
        
        return $defaultTRM;
    }
    
    // Leer el archivo TRM directamente sin cache
    try {
        // Forzar lectura directa del archivo
        $content = file_get_contents($trmFile);
        if ($content === false) {
            throw new Exception('No se pudo leer el archivo TRM');
        }
        
        $trmData = json_decode($content, true);
        if (json_last_error() !== JSON_ERROR_NONE) {
            throw new Exception('Formato de TRM inválido: ' . json_last_error_msg());
        }
        
        // Asegurar que los campos requeridos existen
        $trmData['usd'] = isset($trmData['usd']) ? (float)$trmData['usd'] : 4000.0;
        $trmData['eur'] = isset($trmData['eur']) ? (float)$trmData['eur'] : 4500.0;
        $trmData['fecha'] = $trmData['fecha'] ?? date('Y-m-d H:i:s');
        $trmData['actualizado_por'] = $trmData['actualizado_por'] ?? 'sistema';
        $trmData['origen'] = 'json';
        
        return $trmData;
        
    } catch (Exception $e) {
        // En caso de error, devolver valores por defecto
        return [
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'error',
            'error' => $e->getMessage()
        ];
    }
}

// Obtener el backend activo
$activeBackend = getActiveBackend();

// Obtener TRM actual
$currentTRM = getCurrentTRM();

// Constantes para el frontend
define('CURRENT_TRM_USD', $currentTRM['usd']);
define('CURRENT_TRM_EUR', $currentTRM['eur']);
define('BACKEND_URL', $activeBackend['url']);
define('BACKEND_TYPE', $activeBackend['type']);

/**
 * Obtiene la URL completa de un endpoint de la API
 * 
 * @param string $endpoint Ruta del endpoint (ej: 'api/trm')
 * @param bool $forcePhp Forzar el uso del backend PHP
 * @return string URL completa
 */
function getApiUrl($endpoint = '', $forcePhp = false) {
    global $activeBackend, $backends;
    
    // Para endpoints de TRM, usar siempre rutas relativas
    if ($forcePhp || 
        in_array(strtolower($endpoint), ['trm', 'trm.php', 'api/trm', 'api/trm.php'])) {
        // Usar ruta relativa al archivo trm.php
        return 'trm.php';
    } else {
        $base = rtrim($activeBackend['url'], '/');
        if (empty($base)) {
            return '/' . ltrim($endpoint, '/');
        }
        return $base . '/' . ltrim($endpoint, '/');
    }
}

/**
 * Obtiene la URL para el archivo TRM que pueden usar los scripts de Python
 * 
 * @return string Ruta al archivo TRM
 */
function getTrmFilePath() {
    global $backends;
    return $backends['php']['trm_file'];
}

// Función auxiliar para debug
function log_message($message, $data = null) {
    $logFile = dirname(__DIR__) . '/logs/php_errors.log';
    $logDir = dirname($logFile);
    
    if (!is_dir($logDir)) {
        @mkdir($logDir, 0755, true);
    }
    
    $logMsg = '[' . date('Y-m-d H:i:s') . '] ' . $message . "\n";
    if ($data !== null) {
        $logMsg .= 'Data: ' . print_r($data, true) . "\n\n";
    }
    
    @file_put_contents($logFile, $logMsg, FILE_APPEND);
}