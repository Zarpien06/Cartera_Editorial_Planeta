<?php

class FileHandler {
    private $uploadDir;
    private $allowedExtensions = ['xls', 'xlsx', 'csv', 'txt'];
    private $maxFileSize = 50 * 1024 * 1024; // 50MB
    
    public function __construct($uploadDir) {
        $this->uploadDir = rtrim($uploadDir, '/') . '/';
        
        // Asegurar que el directorio de subidas exista
        if (!file_exists($this->uploadDir)) {
            mkdir($this->uploadDir, 0777, true);
        }
    }
    
    /**
     * Maneja la subida de archivos a través de $_FILES
     */
    public function handleUploads($files) {
        $uploadedFiles = [];
        
        if (empty($files)) {
            return $uploadedFiles;
        }
        
        // Normalizar el array de archivos
        $normalizedFiles = $this->normalizeFilesArray($files);
        
        foreach ($normalizedFiles as $file) {
            if ($file['error'] === UPLOAD_ERR_OK) {
                $uploadedFile = $this->processUploadedFile($file);
                if ($uploadedFile) {
                    $uploadedFiles[] = $uploadedFile;
                }
            }
        }
        
        return $uploadedFiles;
    }
    
    /**
     * Guarda archivos codificados en base64
     */
    public function saveBase64Files($files) {
        $savedFiles = [];
        
        foreach ($files as $fileData) {
            if (empty($fileData['content']) || empty($fileData['name'])) {
                continue;
            }
            
            $fileContent = $fileData['content'];
            $fileName = $fileData['name'];
            $filePath = $this->generateUniqueFilename($this->uploadDir, $fileName);
            
            // Decodificar contenido base64
            $fileContent = str_replace('data:application/octet-stream;base64,', '', $fileContent);
            $fileContent = str_replace(' ', '+', $fileContent);
            $decodedContent = base64_decode($fileContent, true);
            
            if ($decodedContent === false) {
                continue;
            }
            
            // Guardar archivo
            if (file_put_contents($filePath, $decodedContent) !== false) {
                $savedFiles[] = [
                    'name' => $fileName,
                    'path' => $filePath,
                    'size' => filesize($filePath),
                    'type' => mime_content_type($filePath)
                ];
            }
        }
        
        return $savedFiles;
    }
    
    /**
     * Procesa un archivo subido
     */
    private function processUploadedFile($file) {
        // Validar tamaño
        if ($file['size'] > $this->maxFileSize) {
            return null;
        }
        
        // Validar extensión
        $extension = strtolower(pathinfo($file['name'], PATHINFO_EXTENSION));
        if (!in_array($extension, $this->allowedExtensions)) {
            return null;
        }
        
        // Generar nombre único
        $fileName = $this->generateUniqueFilename($this->uploadDir, $file['name']);
        
        // Mover archivo a la ubicación final
        if (move_uploaded_file($file['tmp_name'], $fileName)) {
            return [
                'name' => $file['name'],
                'path' => $fileName,
                'size' => $file['size'],
                'type' => $file['type']
            ];
        }
        
        return null;
    }
    
    /**
     * Normaliza el array de archivos de $_FILES
     */
    private function normalizeFilesArray($files) {
        $normalized = [];
        
        if (!is_array($files['name'])) {
            return [$files];
        }
        
        foreach ($files['name'] as $key => $name) {
            if (is_array($name)) {
                // Manejar múltiples archivos con el mismo nombre
                foreach ($name as $i => $n) {
                    $normalized[] = [
                        'name' => $n,
                        'type' => $files['type'][$key][$i],
                        'tmp_name' => $files['tmp_name'][$key][$i],
                        'error' => $files['error'][$key][$i],
                        'size' => $files['size'][$key][$i]
                    ];
                }
            } else {
                $normalized[] = [
                    'name' => $name,
                    'type' => $files['type'][$key],
                    'tmp_name' => $files['tmp_name'][$key],
                    'error' => $files['error'][$key],
                    'size' => $files['size'][$key]
                ];
            }
        }
        
        return $normalized;
    }
    
    /**
     * Genera un nombre de archivo único
     */
    private function generateUniqueFilename($directory, $originalName) {
        $extension = pathinfo($originalName, PATHINFO_EXTENSION);
        $basename = pathinfo($originalName, PATHINFO_FILENAME);
        $counter = 1;
        
        $filename = $directory . $basename . '.' . $extension;
        
        while (file_exists($filename)) {
            $filename = $directory . $basename . '_' . $counter++ . '.' . $extension;
        }
        
        return $filename;
    }
    
    /**
     * Limpia archivos antiguos del directorio de subidas
     */
    public function cleanupOldFiles($maxAge = 3600) {
        $files = glob($this->uploadDir . '*');
        $now = time();
        $deleted = 0;
        
        foreach ($files as $file) {
            if (is_file($file) && ($now - filemtime($file) >= $maxAge)) {
                @unlink($file);
                $deleted++;
            }
        }
        
        return $deleted;
    }
}
