<?php
// Configuración básica
error_reporting(E_ALL);
ini_set('display_errors', 1);

// Directorios
$uploadDir = __DIR__ . '/uploads/';
$outputDir = __DIR__ . '/output/';

// Crear directorios si no existen
if (!file_exists($uploadDir)) mkdir($uploadDir, 0777, true);
if (!file_exists($outputDir)) mkdir($outputDir, 0777, true);

// Ruta a Python (ajustar según tu instalación)
$pythonPath = 'python';

// Procesar formulario si se envió
$message = '';
$downloadLink = '';

if ($_SERVER['REQUEST_METHOD'] === 'POST') {
    try {
        // Verificar que se subió un archivo
        if (!isset($_FILES['archivo']) || $_FILES['archivo']['error'] !== UPLOAD_ERR_OK) {
            throw new Exception('Error al subir el archivo');
        }

        $file = $_FILES['archivo'];
        $fileName = uniqid() . '_' . basename($file['name']);
        $filePath = $uploadDir . $fileName;
        
        // Mover el archivo subido
        if (!move_uploaded_file($file['tmp_name'], $filePath)) {
            throw new Exception('No se pudo guardar el archivo');
        }

        // Determinar qué script ejecutar según la opción seleccionada
        $tipoProceso = $_POST['tipo_proceso'];
        $outputFile = $outputDir . 'resultado_' . $tipoProceso . '_' . time() . '.xlsx';
        
        // Comando base de Python
        $command = '"C:\\Python39\\python.exe" ';  // Ajusta la ruta a tu Python
        
        switch ($tipoProceso) {
            case 'cartera':
                $script = __DIR__ . '/../back-end/procesador_cartera.py';
                $command .= '"' . $script . '" "' . $filePath . '" --output "' . $outputFile . '"';
                break;
                
            case 'anticipos':
                $script = __DIR__ . '/../back-end/procesador_anticipos.py';
                $command .= '"' . $script . '" "' . $filePath . '" --output "' . $outputFile . '"';
                break;
                
            case 'modelo_deuda':
                // Para modelo deuda necesitamos dos archivos
                if (!isset($_FILES['archivo2']) || $_FILES['archivo2']['error'] !== UPLOAD_ERR_OK) {
                    throw new Exception('Para el modelo deuda se necesitan dos archivos');
                }
                
                $file2 = $_FILES['archivo2'];
                $fileName2 = uniqid() . '_' . basename($file2['name']);
                $filePath2 = $uploadDir . $fileName2;
                
                if (!move_uploaded_file($file2['tmp_name'], $filePath2)) {
                    throw new Exception('Error al guardar el segundo archivo');
                }
                
                $script = __DIR__ . '/../back-end/modelo_deuda.py';
                $command .= '"' . $script . '" "' . $filePath . '" "' . $filePath2 . '" --output "' . $outputFile . '"';
                break;
                
            case 'focus':
                $script = __DIR__ . '/../back-end/procesar_y_actualizar_focus.py';
                $command .= '"' . $script . '" --input "' . $filePath . '" --output "' . $outputFile . '"';
                break;
                
            default:
                throw new Exception('Tipo de proceso no válido');
        }

        // Ejecutar el comando
        $output = [];
        $returnVar = 0;
        
        exec($command . ' 2>&1', $output, $returnVar);
        
        if ($returnVar !== 0) {
            throw new Exception('Error al procesar el archivo: ' . implode("\n", $output));
        }
        
        if (!file_exists($outputFile)) {
            throw new Exception('No se generó el archivo de salida. Verifica los permisos.');
        }
        
        $downloadLink = 'download.php?file=' . basename($outputFile);
        $message = '¡Archivo procesado correctamente! <a href="' . $downloadLink . '">Descargar resultado</a>';
        
    } catch (Exception $e) {
        $message = 'Error: ' . $e->getMessage();
    }
}
?>

<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Procesador de Archivos</title>
    <style>
        body {
            font-family: Arial, sans-serif;
            max-width: 800px;
            margin: 0 auto;
            padding: 20px;
            line-height: 1.6;
        }
        .container {
            background: #f9f9f9;
            padding: 20px;
            border-radius: 8px;
            box-shadow: 0 2px 4px rgba(0,0,0,0.1);
        }
        .form-group {
            margin-bottom: 15px;
        }
        label {
            display: block;
            margin-bottom: 5px;
            font-weight: bold;
        }
        select, input[type="file"] {
            width: 100%;
            padding: 8px;
            margin-bottom: 10px;
            border: 1px solid #ddd;
            border-radius: 4px;
        }
        button {
            background: #4CAF50;
            color: white;
            padding: 10px 15px;
            border: none;
            border-radius: 4px;
            cursor: pointer;
            font-size: 16px;
        }
        button:hover {
            background: #45a049;
        }
        .message {
            padding: 10px;
            margin: 10px 0;
            border-radius: 4px;
        }
        .success {
            background: #dff0d8;
            color: #3c763d;
            border: 1px solid #d6e9c6;
        }
        .error {
            background: #f2dede;
            color: #a94442;
            border: 1px solid #ebccd1;
        }
        .file2-group {
            display: none;
            margin-top: 10px;
        }
    </style>
</head>
<body>
    <div class="container">
        <h1>Procesador de Archivos</h1>
        
        <?php if ($message): ?>
            <div class="message <?php echo strpos($message, 'Error') === 0 ? 'error' : 'success'; ?>">
                <?php echo $message; ?>
            </div>
        <?php endif; ?>
        
        <form method="post" enctype="multipart/form-data" id="uploadForm">
            <div class="form-group">
                <label for="tipo_proceso">Tipo de Proceso:</label>
                <select name="tipo_proceso" id="tipo_proceso" required>
                    <option value="">Seleccione un proceso</option>
                    <option value="cartera">Procesar Cartera</option>
                    <option value="anticipos">Procesar Anticipos</option>
                    <option value="modelo_deuda">Modelo de Deuda (2 archivos)</option>
                    <option value="focus">Procesar FOCUS</option>
                </select>
            </div>
            
            <div class="form-group">
                <label for="archivo">Archivo de Entrada:</label>
                <input type="file" name="archivo" id="archivo" accept=".xlsx,.xls" required>
            </div>
            
            <div class="form-group file2-group" id="file2-group">
                <label for="archivo2">Segundo Archivo (solo para Modelo de Deuda):</label>
                <input type="file" name="archivo2" id="archivo2" accept=".xlsx,.xls">
            </div>
            
            <button type="submit">Procesar Archivo</button>
        </form>
    </div>

    <script>
        // Mostrar/ocultar el segundo campo de archivo según la selección
        document.getElementById('tipo_proceso').addEventListener('change', function() {
            const file2Group = document.getElementById('file2-group');
            const file2Input = document.getElementById('archivo2');
            
            if (this.value === 'modelo_deuda') {
                file2Group.style.display = 'block';
                file2Input.required = true;
            } else {
                file2Group.style.display = 'none';
                file2Input.required = false;
            }
        });
    </script>
</body>
</html>
