<?php
/**
 * Procesador principal de archivos
 * Ejecuta los scripts Python correspondientes según el tipo de proceso
 */

error_reporting(E_ALL);
ini_set('display_errors', 1);

// Incluir configuración local
require_once 'config_local.php';
require_once __DIR__ . '/includes/LogHelper.php';

try {
    LogHelper::log('PHP_PROCESAR', 'INFO', 'Inicio de procesamiento', [
        'request_method' => $_SERVER['REQUEST_METHOD'] ?? 'UNKNOWN',
        'content_type' => $_SERVER['CONTENT_TYPE'] ?? 'UNKNOWN',
    ]);
    // Verificar método HTTP
    if ($_SERVER['REQUEST_METHOD'] !== 'POST') {
        throw new Exception('Método no permitido. Solo se acepta POST.');
    }

    // Obtener tipo de proceso
    $tipo = $_POST['tipo'] ?? $_GET['tipo'] ?? '';
    LogHelper::log('PHP_PROCESAR', 'INFO', 'Tipo de proceso detectado', ['tipo' => $tipo]);

    // Validar tipo
    $tiposValidos = ['cartera', 'anticipos', 'modelo_deuda', 'focus'];
    if (!in_array($tipo, $tiposValidos)) {
        throw new Exception('Tipo de proceso no válido. Tipos válidos: ' . implode(', ', $tiposValidos));
    }

    // Verificar que Python esté disponible
    try {
        $pythonVersion = verificar_python();
    } catch (Exception $e) {
        throw new Exception('Python no está disponible: ' . $e->getMessage());
    }

    // Manejar archivos según el tipo
    $uploadedFiles = [];
    $scriptArgs = [];

    switch ($tipo) {
        case 'cartera':
            // Procesar cartera
            if (!isset($_FILES['archivo'])) {
                $errorMsg = 'Archivo de cartera no proporcionado. $_FILES está vacío.';
                LogHelper::log('PHP_PROCESAR', 'ERROR', $errorMsg, [
                    '_FILES_keys' => array_keys($_FILES),
                    '_POST_keys' => array_keys($_POST)
                ]);
                throw new Exception($errorMsg);
            }
            
            $uploadError = $_FILES['archivo']['error'] ?? UPLOAD_ERR_NO_FILE;
            if ($uploadError !== UPLOAD_ERR_OK) {
                $errorMessages = [
                    UPLOAD_ERR_INI_SIZE => 'El archivo excede upload_max_filesize en php.ini',
                    UPLOAD_ERR_FORM_SIZE => 'El archivo excede MAX_FILE_SIZE del formulario',
                    UPLOAD_ERR_PARTIAL => 'El archivo se subió parcialmente',
                    UPLOAD_ERR_NO_FILE => 'No se subió ningún archivo',
                    UPLOAD_ERR_NO_TMP_DIR => 'Falta la carpeta temporal',
                    UPLOAD_ERR_CANT_WRITE => 'Error al escribir el archivo en disco',
                    UPLOAD_ERR_EXTENSION => 'Una extensión de PHP detuvo la subida'
                ];
                $errorMsg = $errorMessages[$uploadError] ?? "Error desconocido de carga (código: $uploadError)";
                LogHelper::log('PHP_PROCESAR', 'ERROR', 'Error en carga de archivo cartera', [
                    'upload_error_code' => $uploadError,
                    'upload_error_message' => $errorMsg,
                    'file_name' => $_FILES['archivo']['name'] ?? 'N/A',
                    'file_size' => $_FILES['archivo']['size'] ?? 'N/A',
                    'file_type' => $_FILES['archivo']['type'] ?? 'N/A'
                ]);
                throw new Exception("Error al cargar archivo de cartera: $errorMsg");
            }

            $fechaCierre = $_POST['fecha_cierre'] ?? date('Y-m-d');
            $uploadedFiles[] = $_FILES['archivo'];
            $scriptArgs = [$fechaCierre];
            break;

        case 'anticipos':
            // Procesar anticipos
            if (!isset($_FILES['archivo']) || $_FILES['archivo']['error'] !== UPLOAD_ERR_OK) {
                throw new Exception('Archivo de anticipos no proporcionado o error en la carga');
            }

            $uploadedFiles[] = $_FILES['archivo'];
            break;

        case 'modelo_deuda':
            // Procesar modelo de deuda
            if (!isset($_FILES['provision']) || $_FILES['provision']['error'] !== UPLOAD_ERR_OK ||
                !isset($_FILES['anticipos']) || $_FILES['anticipos']['error'] !== UPLOAD_ERR_OK) {
                throw new Exception('Archivos de modelo de deuda no proporcionados o error en la carga');
            }

            $uploadedFiles[] = $_FILES['provision'];
            $uploadedFiles[] = $_FILES['anticipos'];
            break;

        case 'focus':
            // Procesar FOCUS
            $requiredFiles = ['balance', 'situacion', 'focus', 'acumulado', 'modelo'];
            foreach ($requiredFiles as $fileKey) {
                if (!isset($_FILES[$fileKey]) || $_FILES[$fileKey]['error'] !== UPLOAD_ERR_OK) {
                    throw new Exception("Archivo $fileKey no proporcionado o error en la carga");
                }
                $uploadedFiles[] = $_FILES[$fileKey];
            }
            break;
    }

    // Verificar que el directorio de uploads exista y sea escribible
    if (!is_dir(LOCAL_UPLOADS_DIR)) {
        @mkdir(LOCAL_UPLOADS_DIR, 0777, true);
    }
    if (!is_writable(LOCAL_UPLOADS_DIR)) {
        throw new Exception("El directorio de uploads no es escribible: " . LOCAL_UPLOADS_DIR);
    }

    // Mover archivos a directorio de uploads con nombres únicos
    $tempFiles = [];
    foreach ($uploadedFiles as $index => $fileInfo) {
        $tmpName = $fileInfo['tmp_name'] ?? null;
        $originalName = $fileInfo['name'] ?? 'archivo';
        
        if (!$tmpName || !is_uploaded_file($tmpName)) {
            LogHelper::log('PHP_PROCESAR', 'ERROR', 'Archivo temporal inválido', [
                'tmp_name' => $tmpName,
                'original_name' => $originalName,
                'is_uploaded_file' => $tmpName ? is_uploaded_file($tmpName) : false
            ]);
            throw new Exception("Archivo temporal inválido o no subido correctamente: $originalName");
        }
        
        $extension = strtolower(pathinfo($originalName, PATHINFO_EXTENSION));
        $extension = $extension ? '.' . $extension : '';
        $baseName = pathinfo($originalName, PATHINFO_FILENAME);
        $baseName = preg_replace('/[^A-Za-z0-9_\-]/', '_', $baseName);
        if ($baseName === '' || $baseName === null) {
            $baseName = 'archivo';
        }
        $uniqueName = $baseName . '_' . uniqid('', true) . $extension;
        $targetPath = LOCAL_UPLOADS_DIR . '/' . $uniqueName;
        
        if (!move_uploaded_file($tmpName, $targetPath)) {
            $lastError = error_get_last();
            LogHelper::log('PHP_PROCESAR', 'ERROR', 'Error moviendo archivo', [
                'tmp_name' => $tmpName,
                'target_path' => $targetPath,
                'original_name' => $originalName,
                'last_error' => $lastError,
                'uploads_dir_exists' => is_dir(LOCAL_UPLOADS_DIR),
                'uploads_dir_writable' => is_writable(LOCAL_UPLOADS_DIR)
            ]);
            throw new Exception("Error moviendo archivo a directorio de uploads. Verifique permisos y espacio en disco.");
        }
        
        $tempFiles[] = $targetPath;
        LogHelper::log('PHP_PROCESAR', 'DEBUG', 'Archivo preparado', [
            'original' => $originalName,
            'destino' => $targetPath,
            'tamaño' => filesize($targetPath)
        ]);
    }

    // Convertir rutas temporales a absolutas
    $absoluteTempFiles = [];
    foreach ($tempFiles as $tempFile) {
        $absolutePath = realpath($tempFile);
        if ($absolutePath === false) {
            throw new Exception("No se pudo resolver la ruta del archivo temporal: $tempFile");
        }
        $absoluteTempFiles[] = $absolutePath;
    }
    
    // Agregar rutas de archivos temporales a los argumentos del script
    $scriptArgs = array_merge($absoluteTempFiles, $scriptArgs);

    // Obtener ruta del script Python
    $scriptPath = LOCAL_PY_SCRIPTS[$tipo] ?? '';
    if (empty($scriptPath) || !file_exists($scriptPath)) {
        throw new Exception("Script Python no encontrado para tipo: $tipo");
    }

    // Registrar información de depuración
    LogHelper::log('PHP_PROCESAR', 'DEBUG', 'Ejecutando script Python', [
        'script_path' => $scriptPath,
        'script_args' => $scriptArgs
    ]);
    
    // Ejecutar script Python
    $result = ejecutar_python($scriptPath, $scriptArgs);
    
    // Registrar resultado
    LogHelper::log('PHP_PROCESAR', 'DEBUG', 'Resultado de ejecución Python', [
        'success' => $result['success'],
        'exit_code' => $result['exit_code'],
        'output' => $result['output']
    ]);
    
    // Verificar resultado
    if ($result['success']) {
        // Buscar archivo de salida priorizando patrones por tipo
        $patternMap = [
            'cartera' => ['CARTERA*.xlsx', 'CARTERA*.xls'],
            'anticipos' => ['ANTICIPOS*.xlsx', 'ANTICIPOS*.xls'],
            'modelo_deuda' => ['MODELO_DEUDA*.xlsx', 'MODELO_DEUDA*.xls'],
            'focus' => ['FOCUS*.xlsx', 'FOCUS*.xls']
        ];

        $outputFiles = [];
        if (isset($patternMap[$tipo])) {
            foreach ($patternMap[$tipo] as $pattern) {
                $matches = glob(LOCAL_SALIDAS_DIR . '/' . $pattern);
                if (!empty($matches)) {
                    $outputFiles = array_merge($outputFiles, $matches);
                }
            }
        }

        // Fallback general si no se encontró con patrones
        if (empty($outputFiles)) {
            $outputFiles = glob(LOCAL_SALIDAS_DIR . '/*.xlsx') ?: [];
            if (empty($outputFiles)) {
                $outputFiles = glob(LOCAL_SALIDAS_DIR . '/*.xls') ?: [];
            }
            if (empty($outputFiles)) {
                $outputFiles = glob(LOCAL_SALIDAS_DIR . '/*.csv') ?: [];
            }
        }

        if (!empty($outputFiles)) {
            // Tomar el archivo más reciente
            $outputFile = $outputFiles[0];
            foreach ($outputFiles as $file) {
                if (filemtime($file) > filemtime($outputFile)) {
                    $outputFile = $file;
                }
            }

            // Verificar que el archivo exista y no esté vacío
            if (file_exists($outputFile) && filesize($outputFile) > 0) {
                $fileSize = filesize($outputFile);
                // Configurar headers para descarga
                header('Content-Type: application/octet-stream');
                header('Content-Disposition: attachment; filename="' . basename($outputFile) . '"');
                header('Content-Length: ' . $fileSize);
                header('Cache-Control: no-cache, must-revalidate');
                header('Expires: 0');

                // Enviar archivo
                readfile($outputFile);
                
                // Limpiar archivo de salida después de enviarlo
                @unlink($outputFile);
                
                LogHelper::log('PHP_PROCESAR', 'INFO', 'Archivo entregado al usuario', [
                    'tipo' => $tipo,
                    'salida' => $outputFile,
                    'tamaño' => $fileSize,
                ]);
            } else {
                throw new Exception('Archivo de salida vacío o no encontrado');
            }
        } else {
            throw new Exception('No se generó archivo de salida. Revisar logs de Python.');
        }
    } else {
        throw new Exception('Error al ejecutar Python: ' . $result['output']);
    }
    
} catch (Exception $e) {
    LogHelper::logException('PHP_PROCESAR', $e, [
        'tipo' => $tipo ?? 'desconocido',
    ]);
    
    // Enviar respuesta JSON de error
    header('Content-Type: application/json');
    http_response_code(500);
    
    // Intentar obtener más información del error
    $errorDetails = [
        'error' => $e->getMessage(),
        'tipo' => $tipo ?? 'desconocido',
        'file' => basename($e->getFile()),
        'line' => $e->getLine()
    ];
    
    // Si hay resultado parcial, incluirlo
    if (isset($result)) {
        $errorDetails['python_output'] = $result['output'] ?? '';
        $errorDetails['python_exit_code'] = $result['exit_code'] ?? '';
        $errorDetails['python_command'] = $result['command'] ?? '';
    }
    
    // Verificar directorios
    if (isset($outputFile)) {
        $errorDetails['output_file_expected'] = $outputFile;
        $errorDetails['output_file_exists'] = file_exists($outputFile);
    }
    
    $errorDetails['salidas_dir_exists'] = is_dir(LOCAL_SALIDAS_DIR);
    $errorDetails['salidas_dir_writable'] = is_writable(LOCAL_SALIDAS_DIR);
    
    // Archivos procesados
    if (isset($uploadedFiles)) {
        $errorDetails['uploaded_files'] = array_map(function ($info) {
            return $info['name'] ?? 'desconocido';
        }, $uploadedFiles);
    }
    
    echo json_encode($errorDetails);
}

LogHelper::log('PHP_PROCESAR', 'INFO', 'Fin de procesamiento', ['tipo' => $tipo ?? 'desconocido']);
?>