<?php
require_once __DIR__ . '/../config/config.php';
require_once __DIR__ . '/../includes/Response.php';
require_once __DIR__ . '/../includes/FileHandler.php';
require_once __DIR__ . '/../includes/PythonRunner.php';
require_once __DIR__ . '/../includes/ProcessorHandler.php';

// Configurar manejo de errores
set_error_handler(function($errno, $errstr, $errfile, $errline) {
    Response::error(500, "Error en el servidor: $errstr en $errfile en la línea $errline");
});

set_exception_handler(function($e) {
    Response::error(500, "Excepción no capturada: " . $e->getMessage());
});

// Obtener método y ruta de la solicitud
$method = $_SERVER['REQUEST_METHOD'];
$path = parse_url($_SERVER['REQUEST_URI'], PHP_URL_PATH);
$path = trim($path, '/');

// Rutas de la API
$routes = [
    'POST' => [
        'api/process' => 'handleProcessRequest',
    ]
];

// Enrutamiento
if (isset($routes[$method][$path])) {
    $handler = $routes[$method][$path];
    $handler();
} else {
    Response::error(404, 'Ruta no encontrada');
}

/**
 * Maneja las solicitudes de procesamiento
 */
function handleProcessRequest() {
    // Verificar autenticación
    $headers = getallheaders();
    $apiKey = $headers['X-API-KEY'] ?? '';
    
    if ($apiKey !== API_KEY) {
        Response::error(401, 'No autorizado');
    }

    // Obtener datos de la solicitud
    $input = json_decode(file_get_contents('php://input'), true);
    
    if (json_last_error() !== JSON_ERROR_NONE) {
        Response::error(400, 'JSON inválido: ' . json_last_error_msg());
    }

    // Validar acción
    $action = $input['action'] ?? '';
    $validActions = ['cartera', 'anticipos', 'modelo_deuda', 'focus'];
    
    if (!in_array($action, $validActions)) {
        Response::error(400, 'Acción no válida. Acciones válidas: ' . implode(', ', $validActions));
    }

    // Manejar archivos subidos
    $fileHandler = new FileHandler(UPLOAD_DIR);
    $uploadedFiles = $fileHandler->handleUploads($_FILES);
    
    // Si no hay archivos en $_FILES, verificar si vienen en el JSON
    if (empty($uploadedFiles) && !empty($input['files'])) {
        $uploadedFiles = $fileHandler->saveBase64Files($input['files']);
    }

    // Crear manejador de procesadores
    $processor = new ProcessorHandler(PYTHON_PATH, OUTPUT_DIR);
    
    try {
        $result = null;
        
        switch ($action) {
            case 'cartera':
                $result = processCartera($processor, $uploadedFiles, $input);
                break;
                
            case 'anticipos':
                $result = processAnticipos($processor, $uploadedFiles, $input);
                break;
                
            case 'modelo_deuda':
                $result = processModeloDeuda($processor, $uploadedFiles, $input);
                break;
                
            case 'focus':
                $result = processFocus($processor, $uploadedFiles, $input);
                break;
        }
        
        if ($result['success']) {
            // Buscar archivo de salida
            $outputFile = $processor->getLatestOutputFile();
            
            if ($outputFile && file_exists($outputFile)) {
                Response::success([
                    'message' => 'Procesamiento completado exitosamente',
                    'output_file' => basename($outputFile),
                    'download_url' => '/backend_php/download.php?file=' . urlencode(basename($outputFile)),
                    'output' => $result['output']
                ]);
            } else {
                Response::success([
                    'message' => 'Procesamiento completado',
                    'output' => $result['output']
                ]);
            }
        } else {
            Response::error(500, 'Error al procesar: ' . ($result['error'] ?? 'Error desconocido'));
        }
        
    } catch (Exception $e) {
        Response::error(500, 'Excepción al procesar: ' . $e->getMessage());
    }
}

/**
 * Procesa cartera
 */
function processCartera($processor, $files, $params) {
    if (empty($files)) {
        Response::error(400, 'Se requiere un archivo de entrada');
    }
    
    $inputFile = $files[0]['path'];
    $fechaCierre = $params['fecha_cierre'] ?? $params['fecha_corte'] ?? date('Y-m-d');
    
    // Parámetros opcionales
    $processorParams = [];
    if (!empty($params['trm_usd'])) {
        $processorParams['trm_usd'] = $params['trm_usd'];
    }
    if (!empty($params['trm_eur'])) {
        $processorParams['trm_eur'] = $params['trm_eur'];
    }
    
    return $processor->processCartera($inputFile, $fechaCierre, $processorParams);
}

/**
 * Procesa anticipos
 */
function processAnticipos($processor, $files, $params) {
    if (empty($files)) {
        Response::error(400, 'Se requiere un archivo de entrada');
    }
    
    $inputFile = $files[0]['path'];
    return $processor->processAnticipos($inputFile, $params);
}

/**
 * Procesa modelo de deuda
 */
function processModeloDeuda($processor, $files, $params) {
    if (count($files) < 2) {
        Response::error(400, 'Se requieren 2 archivos: provisión y anticipos');
    }
    
    $archivoProvision = $files[0]['path'];
    $archivoAnticipos = $files[1]['path'];
    
    return $processor->processModeloDeuda($archivoProvision, $archivoAnticipos, $params);
}

/**
 * Procesa FOCUS
 */
function processFocus($processor, $files, $params) {
    // FOCUS puede buscar archivos automáticamente o usar los proporcionados
    $focusParams = [];
    
    if (!empty($files)) {
        // Mapear archivos según el orden o nombres
        foreach ($files as $i => $file) {
            $fileName = strtolower($file['name']);
            
            if (strpos($fileName, 'focus') !== false) {
                $focusParams['archivo_focus'] = $file['path'];
            } elseif (strpos($fileName, 'balance') !== false) {
                $focusParams['archivo_balance'] = $file['path'];
            } elseif (strpos($fileName, 'situacion') !== false) {
                $focusParams['archivo_situacion'] = $file['path'];
            } elseif (strpos($fileName, 'modelo') !== false || strpos($fileName, 'deuda') !== false) {
                $focusParams['archivo_modelo'] = $file['path'];
            }
        }
    }
    
    return $processor->processFocus($focusParams);
}

/**
 * Maneja la descarga de archivos generados
 */
function handleDownloadRequest() {
    $file = $_GET['file'] ?? '';
    if (empty($file)) {
        Response::error(400, 'Nombre de archivo no proporcionado');
    }

    $filePath = OUTPUT_DIR . '/' . basename($file);
    
    if (!file_exists($filePath)) {
        Response::error(404, 'Archivo no encontrado');
    }

    // Enviar el archivo para descarga
    header('Content-Description: File Transfer');
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="' . basename($filePath) . '"');
    header('Expires: 0');
    header('Cache-Control: must-revalidate');
    header('Pragma: public');
    header('Content-Length: ' . filesize($filePath));
    readfile($filePath);
    exit;
}
