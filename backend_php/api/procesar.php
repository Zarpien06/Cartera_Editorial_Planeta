<?php
/**
 * Endpoint Procesar - API para manejar el procesamiento de archivos
 */

// Habilitar reporte de errores
error_reporting(E_ALL);
ini_set('display_errors', 1);

// Configuración de CORS
header('Access-Control-Allow-Origin: *');
header('Access-Control-Allow-Methods: POST, OPTIONS');
header('Access-Control-Allow-Headers: Content-Type, Authorization');
header('Access-Control-Max-Age: 3600');

// Manejar solicitud OPTIONS (preflight)
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    http_response_code(200);
    exit();
}

// Solo permitir POST para procesamiento
if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
    http_response_code(405);
    header('Content-Type: application/json');
    echo json_encode([
        'error' => 'Método no permitido',
        'allowed_methods' => ['POST', 'OPTIONS']
    ]);
    exit;
}

// Incluir la configuración local
require_once __DIR__ . '/../../front_php/config_local.php';

// Incluir el procesador principal
require_once __DIR__ . '/../../front_php/procesar.php';
?>