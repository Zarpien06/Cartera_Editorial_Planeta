<?php
require_once 'config.php';

// Si se está llamando al backend PHP directamente, procesar la petición
if (php_sapi_name() !== 'cli' && !empty($_SERVER['REQUEST_METHOD'])) {
    // Configuración
    $uploadDir = __DIR__ . '/uploads/';
    $outputDir = __DIR__ . '/output/';
    $pythonPath = 'C:\\Python39\\python.exe'; // Ajusta la ruta a tu Python

    // Crear directorios si no existen
    if (!file_exists($uploadDir)) mkdir($uploadDir, 0777, true);
    if (!file_exists($outputDir)) mkdir($outputDir, 0777, true);

    // Configurar encabezados para CORS
    header('Access-Control-Allow-Origin: *');
    header('Content-Type: application/json');
    header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
    header('Access-Control-Allow-Headers: Content-Type, X-Requested-With');

    // Manejar solicitud de prueba
    if (isset($_GET['test'])) {
        echo json_encode([
            'status' => 'ok',
            'backend' => 'php',
            'timestamp' => date('Y-m-d H:i:s')
        ]);
        exit;
    }

    // Manejar solicitud OPTIONS para CORS
    if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
        http_response_code(200);
        exit;
    }

    // Verificar que la petición sea POST
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
        http_response_code(405);
        echo json_encode(['error' => 'Método no permitido']);
        exit;
    }

    // Obtener datos de la petición
    $input = [];
    $contentType = isset($_SERVER['CONTENT_TYPE']) ? $_SERVER['CONTENT_TYPE'] : '';

    // Manejar diferentes tipos de contenido
    if (strpos($contentType, 'application/json') !== false) {
        $input = json_decode(file_get_contents('php://input'), true);
    } else {
        $input = $_POST;
    }

    // Obtener datos del formulario
    $action = $_GET['action'] ?? ($input['action'] ?? '');
    $files = $_FILES ?? ($input['files'] ?? []);

    // Validar acción
    $validActions = ['cartera', 'anticipos', 'modelo_deuda', 'focus'];
    if (!in_array($action, $validActions)) {
        http_response_code(400);
        echo json_encode([
            'error' => 'Acción no válida',
            'valid_actions' => $validActions
        ]);
        exit;
    }

    // Función para limpiar archivos antiguos
    function cleanOldFiles($dir, $maxAge = 3600) {
        if (!is_dir($dir)) return;
        
        $files = glob($dir . '*');
        $now = time();
        
        foreach ($files as $file) {
            if (is_file($file) && ($now - filemtime($file) >= $maxAge)) {
                @unlink($file);
            }
        }
    }

    // Limpiar archivos antiguos al finalizar
    register_shutdown_function(function() use ($uploadDir, $outputDir) {
        cleanOldFiles($uploadDir, 3600); // 1 hora
        cleanOldFiles($outputDir, 86400); // 24 horas
    });
}
?>