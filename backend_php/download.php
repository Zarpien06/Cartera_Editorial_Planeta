<?php
// Configuración básica
error_reporting(0);

// Directorio de salida
$outputDir = __DIR__ . '/output/';

// Obtener el nombre del archivo
$fileName = $_GET['file'] ?? '';

// Validar el nombre del archivo (solo caracteres seguros)
if (!preg_match('/^[a-zA-Z0-9_\-\.]+\.xlsx?$/i', $fileName)) {
    die('Nombre de archivo no válido');
}

$filePath = $outputDir . $fileName;

// Verificar que el archivo exista y sea legible
if (!file_exists($filePath) || !is_readable($filePath)) {
    die('El archivo solicitado no existe o no está disponible');
}

// Obtener información del archivo
$fileSize = filesize($filePath);
$fileInfo = pathinfo($filePath);
$fileExtension = strtolower($fileInfo['extension']);

// Configurar las cabeceras para la descarga
header('Content-Description: File Transfer');
header('Content-Type: application/octet-stream');
header('Content-Disposition: attachment; filename="' . $fileName . '"');
header('Content-Length: ' . $fileSize);
header('Expires: 0');
header('Cache-Control: must-revalidate');
header('Pragma: public');

// Limpiar el búfer de salida y enviar el archivo
ob_clean();
flush();
readfile($filePath);

exit;
