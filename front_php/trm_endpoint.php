<?php
/**
 * Endpoint TRM simplificado para acceso desde el frontend
 * Lee directamente del archivo trm.json sin usar cache
 */

// Configurar cabeceras
header('Content-Type: application/json; charset=utf-8');
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type');

// Manejar solicitud OPTIONS (preflight)
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit();
}

function respond_with_json($payload, $status = 200) {
    http_response_code($status);
    echo json_encode($payload, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
    exit;
}

// Determinar el método de solicitud
$method = $_SERVER['REQUEST_METHOD'];

// Ruta al archivo TRM (apuntando al nuevo endpoint)
$trmEndpoint = __DIR__ . '/../trm.php';

if ($method === 'GET') {
    // Obtener TRM - leer directamente del archivo sin cache
    // Incluir el archivo TRM para obtener los datos
    $trmFile = __DIR__ . '/../Python_principales/trm.json';
    
    if (!file_exists($trmFile)) {
        // Crear archivo con valores por defecto
        $defaultData = [
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
        file_put_contents($trmFile, json_encode($defaultData, JSON_PRETTY_PRINT));
        @chmod($trmFile, 0666);
        
        respond_with_json($defaultData);
    }
    
    // Leer el archivo TRM directamente sin usar cache
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
        
        respond_with_json($trmData);
        
    } catch (Exception $e) {
        // En caso de error, devolver valores por defecto
        respond_with_json([
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'error',
            'error' => $e->getMessage()
        ], 200);
    }
} 
elseif ($method === 'POST') {
    // Actualizar TRM - redirigir al endpoint principal
    // Obtener datos del cuerpo de la solicitud
    $input = file_get_contents('php://input');
    $data = json_decode($input, true);
    
    // Si no es JSON, intentar obtener de POST normal
    if ($data === null) {
        $usd = filter_input(INPUT_POST, 'usd', FILTER_VALIDATE_FLOAT);
        $eur = filter_input(INPUT_POST, 'eur', FILTER_VALIDATE_FLOAT);
        $updated_by = filter_input(INPUT_POST, 'updated_by', FILTER_SANITIZE_STRING) ?: 'api_web';
    } else {
        $usd = isset($data['usd']) ? filter_var($data['usd'], FILTER_VALIDATE_FLOAT) : null;
        $eur = isset($data['eur']) ? filter_var($data['eur'], FILTER_VALIDATE_FLOAT) : null;
        $updated_by = isset($data['updated_by']) ? filter_var($data['updated_by'], FILTER_SANITIZE_STRING) : 'api_web';
    }
    
    // Validar parámetros
    if ($usd === null || $eur === null || $usd === false || $eur === false) {
        respond_with_json([
            'error' => 'Parámetros inválidos o faltantes (usd, eur)'
        ], 400);
    }
    
    // Crear los datos de TRM
    $trmData = [
        'usd' => floatval($usd),
        'eur' => floatval($eur),
        'fecha' => date('Y-m-d H:i:s'),
        'actualizado_por' => $updated_by,
        'origen' => 'web'
    ];
    
    // Ruta al archivo TRM
    $trmFile = __DIR__ . '/../Python_principales/trm.json';
    
    // Asegurar que el directorio existe
    $dir = dirname($trmFile);
    if (!is_dir($dir)) {
        mkdir($dir, 0755, true);
    }
    
    // Guardar en el archivo
    if (file_put_contents($trmFile, json_encode($trmData, JSON_PRETTY_PRINT)) === false) {
        respond_with_json([
            'error' => 'No se pudo guardar el archivo TRM',
            'path' => $trmFile
        ], 500);
    }
    
    // Asegurar permisos
    @chmod($trmFile, 0666);
    
    respond_with_json([
        'success' => true,
        'message' => 'TRM actualizada correctamente',
        'data' => $trmData
    ]);
} 
else {
    // Método no permitido
    respond_with_json([
        'error' => 'Método no permitido',
        'allowed_methods' => ['GET', 'POST', 'OPTIONS']
    ], 405);
}
?>