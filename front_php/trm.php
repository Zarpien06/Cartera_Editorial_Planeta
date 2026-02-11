<?php
/**
 * Endpoint TRM - API para manejar las Tasas Representativas del Mercado
 * Lee directamente del archivo trm.json sin usar cache
 */

// Habilitar reporte de errores
error_reporting(E_ALL);
ini_set('display_errors', 1);

require_once __DIR__ . '/includes/LogHelper.php';

// Configuración de CORS
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type, Authorization');
header('Access-Control-Max-Age: 3600');

// Manejar solicitud OPTIONS (preflight)
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit();
}

// Configuración de cabeceras - sin cache
header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, no-cache, must-revalidate, max-age=0');
header('Pragma: no-cache');
header('Expires: 0');

function respond_with_json($payload, $status = 200) {
    http_response_code($status);
    header('Content-Type: application/json; charset=utf-8');
    echo json_encode($payload, JSON_PRETTY_PRINT | JSON_UNESCAPED_UNICODE);
    exit;
}

// Ruta al archivo JSON de TRM
$trmFile = __DIR__ . '/../Python_principales/trm.json';

// Asegurar que el directorio existe
$trmDir = dirname($trmFile);
if (!is_dir($trmDir)) {
    mkdir($trmDir, 0755, true);
}

// POST: actualizar TRM en el archivo JSON
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Verificar si los datos vienen en formato JSON
    $input = file_get_contents('php://input');
    $data = json_decode($input, true);
    
    // Si no es JSON, intentar obtener de POST normal
    if ($data === null) {
        $usd = filter_input(INPUT_POST, 'usd', FILTER_VALIDATE_FLOAT);
        $eur = filter_input(INPUT_POST, 'eur', FILTER_VALIDATE_FLOAT);
    } else {
        $usd = isset($data['usd']) ? filter_var($data['usd'], FILTER_VALIDATE_FLOAT) : null;
        $eur = isset($data['eur']) ? filter_var($data['eur'], FILTER_VALIDATE_FLOAT) : null;
    }
    
    if ($usd === null || $eur === null || $usd === false || $eur === false) {
        LogHelper::log('TRM_API', 'WARNING', 'Parámetros inválidos al actualizar TRM', [
            'payload' => $data ?? $_POST ?? []
        ]);
        respond_with_json(["error" => "Parámetros inválidos o faltantes (usd, eur)"], 400);
    }
    
    $trmData = [
        'usd' => $usd,
        'eur' => $eur,
        'fecha' => date('Y-m-d H:i:s'),
        'actualizado_por' => 'sistema_php',
        'origen' => 'json'
    ];
    
    // Asegurar que el directorio existe
    $dir = dirname($trmFile);
    if (!is_dir($dir)) {
        mkdir($dir, 0755, true);
    }
    
    // Guardar en el archivo JSON
    if (file_put_contents($trmFile, json_encode($trmData, JSON_PRETTY_PRINT)) === false) {
        respond_with_json([
            'error' => 'No se pudo guardar el archivo TRM',
            'path' => $trmFile
        ], 500);
    }
    
    // Asegurar permisos
    @chmod($trmFile, 0666);
    
    LogHelper::log('TRM_API', 'INFO', 'TRM actualizada', $trmData);
    
    respond_with_json([
        'success' => true,
        'message' => 'TRM actualizada correctamente',
        'data' => $trmData
    ]);
}

// GET: obtener TRM actual del archivo JSON - leer directamente sin cache
if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    // Si no existe el archivo, devolver valores por defecto
    if (!file_exists($trmFile)) {
        $defaultData = [
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'default'
        ];
        
        // Crear archivo con valores por defecto
        if (file_put_contents($trmFile, json_encode($defaultData, JSON_PRETTY_PRINT)) !== false) {
            @chmod($trmFile, 0666);
        }
        
        LogHelper::log('TRM_API', 'WARNING', 'TRM no encontrada, creando valores por defecto.');
        respond_with_json($defaultData);
    }
    
    try {
        // Leer el archivo JSON directamente sin cache
        $jsonContent = file_get_contents($trmFile);
        if ($jsonContent === false) {
            throw new Exception('No se pudo leer el archivo TRM');
        }
        
        $trmData = json_decode($jsonContent, true, 512, JSON_THROW_ON_ERROR);
        
        // Validar datos
        if (!is_array($trmData)) {
            throw new Exception('Formato de TRM inválido: se esperaba un objeto');
        }
        
        // Asegurar que los campos requeridos existen
        $trmData['usd'] = isset($trmData['usd']) ? (float)$trmData['usd'] : 4000.0;
        $trmData['eur'] = isset($trmData['eur']) ? (float)$trmData['eur'] : 4500.0;
        $trmData['fecha'] = $trmData['fecha'] ?? date('Y-m-d H:i:s');
        $trmData['actualizado_por'] = $trmData['actualizado_por'] ?? 'sistema';
        $trmData['origen'] = 'json';
        
        LogHelper::log('TRM_API', 'INFO', 'TRM consultada', [
            'usd' => $trmData['usd'],
            'eur' => $trmData['eur'],
            'origen' => $trmData['origen'] ?? 'json'
        ]);
        respond_with_json($trmData);
        
    } catch (Exception $e) {
        // En caso de error, devolver valores por defecto
        $defaultData = [
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'error',
            'error' => $e->getMessage()
        ];
        
        LogHelper::log('TRM_API', 'ERROR', 'Error leyendo TRM, entregando valores por defecto', [
            'error' => $e->getMessage()
        ]);
        respond_with_json($defaultData, 200); // 200 para no romper la UI
    }
}

// Si el método no es GET ni POST
LogHelper::log('TRM_API', 'WARNING', 'Método no permitido', [
    'method' => $_SERVER['REQUEST_METHOD'] ?? 'UNKNOWN'
]);
respond_with_json([
    'error' => 'Método no permitido',
    'allowed_methods' => ['GET', 'POST', 'OPTIONS']
], 405);
?>