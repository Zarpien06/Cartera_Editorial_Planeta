<?php

class PythonRunner {
    private $pythonPath;
    private $timeout = 300; // 5 minutos por defecto
    
    public function __construct($pythonPath = 'python3') {
        $this->pythonPath = $pythonPath;
    }
    
    /**
     * Ejecuta un script de Python con los argumentos proporcionados
     */
    public function execute($scriptPath, array $args = []) {
        if (!file_exists($scriptPath)) {
            throw new Exception("El script de Python no existe en: $scriptPath");
        }
        
        // Construir el comando
        $command = escapeshellarg($this->pythonPath) . ' ' . escapeshellarg($scriptPath);
        
        // Agregar argumentos
        foreach ($args as $arg) {
            $command .= ' ' . escapeshellarg($arg);
        }
        
        // Redirigir salida estándar y de error
        $command .= ' 2>&1';
        
        // Ejecutar el comando con tiempo de espera
        $descriptors = [
            0 => ['pipe', 'r'], // stdin
            1 => ['pipe', 'w'], // stdout
            2 => ['pipe', 'w']  // stderr
        ];
        
        $process = proc_open($command, $descriptors, $pipes, dirname($scriptPath));
        
        if (!is_resource($process)) {
            throw new Exception("No se pudo ejecutar el comando: $command");
        }
        
        // Configurar el proceso como no bloqueante
        stream_set_blocking($pipes[1], false);
        stream_set_blocking($pipes[2], false);
        
        // Leer la salida con tiempo de espera
        $output = '';
        $errorOutput = '';
        $startTime = time();
        $outputFile = null;
        
        while (true) {
            // Verificar tiempo de espera
            if ((time() - $startTime) > $this->timeout) {
                $this->terminateProcess($process);
                throw new Exception("Tiempo de espera agotado (más de {$this->timeout} segundos)");
            }
            
            // Leer de stdout
            $stdout = stream_get_contents($pipes[1]);
            if ($stdout !== false) {
                $output .= $stdout;
                
                // Intentar detectar la ruta del archivo de salida con múltiples patrones
                if (preg_match('/(?:Archivo guardado en|generado|✅ Archivo generado|Ruta de salida):\s*(.+?)(?:[\r\n]|$)/i', $stdout, $matches)) {
                    $outputFile = trim($matches[1]);
                }
                // También buscar rutas absolutas de archivos Excel
                if (preg_match('/((?:[A-Z]:|\/)(?:[^\/]+\/)*[^\s]+\.xlsx?)(?:[\s\r\n]|$)/i', $stdout, $matches)) {
                    $potentialFile = trim($matches[1]);
                    if (file_exists($potentialFile)) {
                        $outputFile = $potentialFile;
                    }
                }
            }
            
            // Leer de stderr
            $stderr = stream_get_contents($pipes[2]);
            if ($stderr !== false) {
                $errorOutput .= $stderr;
            }
            
            // Verificar si el proceso ha terminado
            $status = proc_get_status($process);
            if (!$status['running']) {
                // Leer cualquier salida restante
                $output .= stream_get_contents($pipes[1]);
                $errorOutput .= stream_get_contents($pipes[2]);
                break;
            }
            
            // Esperar un poco antes de la siguiente iteración
            usleep(100000); // 100ms
        }
        
        // Cerrar los pipes
        fclose($pipes[0]);
        fclose($pipes[1]);
        fclose($pipes[2]);
        
        // Cerrar el proceso
        $exitCode = proc_close($process);
        
        // Procesar la salida
        if ($exitCode !== 0 || !empty($errorOutput)) {
            $errorMessage = "Error al ejecutar el script (Código: $exitCode)\n";
            $errorMessage .= "Salida de error:\n$errorOutput";
            
            // Intentar extraer el mensaje de error específico
            if (preg_match('/Error: (.+?)[\r\n]/', $errorOutput, $matches)) {
                $errorMessage = $matches[1];
            } elseif (preg_match('/Exception: (.+?)[\r\n]/', $errorOutput, $matches)) {
                $errorMessage = $matches[1];
            } elseif (preg_match('/Traceback \(most recent call last\):[\s\S]+?([\s\S]+?)(?=\n\w|$)/', $errorOutput, $matches)) {
                $errorMessage = trim($matches[1]);
            }
            
            return [
                'success' => false,
                'error' => $errorMessage,
                'output' => $output,
                'error_output' => $errorOutput,
                'exit_code' => $exitCode
            ];
        }
        
        return [
            'success' => true,
            'output' => $output,
            'output_file' => $outputFile
        ];
    }
    
    /**
     * Establece el tiempo máximo de ejecución en segundos
     */
    public function setTimeout($seconds) {
        $this->timeout = (int) $seconds;
        return $this;
    }
    
    /**
     * Mata un proceso y todos sus hijos
     */
    private function terminateProcess($process) {
        if (function_exists('proc_terminate')) {
            // En Windows
            if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {
                $status = proc_get_status($process);
                if ($status['running']) {
                    // En Windows, necesitamos usar taskkill para asegurarnos de matar el proceso y sus hijos
                    exec("taskkill /F /T /PID {$status['pid']}");
                }
            } else {
                // En Unix/Linux
                $status = proc_get_status($process);
                if ($status['running']) {
                    // Enviamos SIGTERM al grupo de procesos
                    posix_kill($status['pid'], SIGTERM);
                    
                    // Esperamos un poco
                    usleep(100000); // 100ms
                    
                    // Si aún está en ejecución, forzamos la terminación
                    $status = proc_get_status($process);
                    if ($status['running']) {
                        posix_kill($status['pid'], SIGKILL);
                    }
                }
            }
        }
        
        // Cerramos el proceso de todos modos
        if (is_resource($process)) {
            proc_terminate($process);
        }
    }
}
