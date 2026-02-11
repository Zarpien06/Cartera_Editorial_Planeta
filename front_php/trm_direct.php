<?php
/**
 * API TRM Directa - Accede directamente al archivo trm.json
 * Lee directamente del archivo trm.json sin usar cache
 */

// Configuración de cabeceras
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: GET, POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type, Authorization');
header('Content-Type: application/json; charset=utf-8');
header('Cache-Control: no-store, no-cache, must-revalidate, max-age=0');
header('Pragma: no-cache');
header('Expires: 0');

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

// Ruta al archivo TRM
$trmFile = __DIR__ . '/../Python_principales/trm.json';

// POST: Actualizar TRM
if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    // Obtener datos
    $input = file_get_contents('php://input');
    $data = json_decode($input, true);
    
    if ($data === null) {
        $usd = filter_input(INPUT_POST, 'usd', FILTER_VALIDATE_FLOAT);
        $eur = filter_input(INPUT_POST, 'eur', FILTER_VALIDATE_FLOAT);
    } else {
        $usd = isset($data['usd']) ? filter_var($data['usd'], FILTER_VALIDATE_FLOAT) : null;
        $eur = isset($data['eur']) ? filter_var($data['eur'], FILTER_VALIDATE_FLOAT) : null;
    }
    
    if ($usd === null || $eur === null || $usd === false || $eur === false) {
        respond_with_json(["error" => "Parámetros inválidos o faltantes (usd, eur)"], 400);
    }
    
    // Crear datos TRM
    $trmData = [
        'usd' => $usd,
        'eur' => $eur,
        'fecha' => date('Y-m-d H:i:s'),
        'actualizado_por' => 'api_directa',
        'origen' => 'json'
    ];
    
    // Guardar en archivo
    if (file_put_contents($trmFile, json_encode($trmData, JSON_PRETTY_PRINT)) === false) {
        respond_with_json(['error' => 'No se pudo guardar el archivo TRM'], 500);
    }
    
    @chmod($trmFile, 0666);
    
    respond_with_json([
        'success' => true,
        'message' => 'TRM actualizada correctamente',
        'data' => $trmData
    ]);
}

// GET: Obtener TRM - leer directamente sin cache
if ($_SERVER['REQUEST_METHOD'] === 'GET') {
    if (!file_exists($trmFile)) {
        // Crear archivo con valores por defecto
        $defaultData = [
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'default'
        ];
        
        if (file_put_contents($trmFile, json_encode($defaultData, JSON_PRETTY_PRINT)) !== false) {
            @chmod($trmFile, 0666);
        }
        
        respond_with_json($defaultData);
    }
    
    // Leer archivo directamente sin cache
    $content = file_get_contents($trmFile);
    if ($content === false) {
        respond_with_json([
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'error',
            'error' => 'No se pudo leer el archivo TRM'
        ], 200);
    }
    
    $trmData = json_decode($content, true);
    if ($trmData === null) {
        respond_with_json([
            'usd' => 4000.0,
            'eur' => 4500.0,
            'fecha' => date('Y-m-d H:i:s'),
            'actualizado_por' => 'sistema',
            'origen' => 'error',
            'error' => 'Formato JSON inválido'
        ], 200);
    }
    
    // Asegurar que los campos existen
    $trmData['usd'] = isset($trmData['usd']) ? (float)$trmData['usd'] : 4000.0;
    $trmData['eur'] = isset($trmData['eur']) ? (float)$trmData['eur'] : 4500.0;
    $trmData['fecha'] = $trmData['fecha'] ?? date('Y-m-d H:i:s');
    $trmData['actualizado_por'] = $trmData['actualizado_por'] ?? 'sistema';
    $trmData['origen'] = 'json';
    
    respond_with_json($trmData);
}

// Método no permitido
respond_with_json([
    'error' => 'Método no permitido',
    'allowed_methods' => ['GET', 'POST', 'OPTIONS']
], 405);
?>