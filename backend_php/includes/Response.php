<?php

class Response {
    /**
     * Envía una respuesta JSON exitosa
     */
    public static function success($data = null, $statusCode = 200) {
        http_response_code($statusCode);
        header('Content-Type: application/json');
        
        $response = [
            'success' => true,
            'data' => $data
        ];
        
        echo json_encode($response, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
        exit;
    }
    
    /**
     * Envía una respuesta de error
     */
    public static function error($statusCode, $message = '') {
        http_response_code($statusCode);
        header('Content-Type: application/json');
        
        $statusTexts = [
            400 => 'Solicitud incorrecta',
            401 => 'No autorizado',
            403 => 'Prohibido',
            404 => 'No encontrado',
            405 => 'Método no permitido',
            500 => 'Error interno del servidor',
            503 => 'Servicio no disponible'
        ];
        
        $statusText = $statusTexts[$statusCode] ?? 'Error';
        
        $response = [
            'success' => false,
            'error' => [
                'code' => $statusCode,
                'message' => $message ?: $statusText,
                'status' => $statusText
            ]
        ];
        
        // Registrar error
        error_log("$statusCode $statusText: $message");
        
        echo json_encode($response, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
        exit;
    }
    
    /**
     * Envía un archivo para descarga
     */
    public static function file($filePath, $filename = null) {
        if (!file_exists($filePath)) {
            self::error(404, 'Archivo no encontrado');
        }
        
        $filename = $filename ?: basename($filePath);
        $mimeType = mime_content_type($filePath);
        
        header('Content-Description: File Transfer');
        header('Content-Type: ' . $mimeType);
        header('Content-Disposition: attachment; filename="' . $filename . '"');
        header('Content-Length: ' . filesize($filePath));
        
        readfile($filePath);
        exit;
    }
}
