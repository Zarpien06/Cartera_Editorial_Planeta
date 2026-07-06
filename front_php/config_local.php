<?php
// Normalizar separadores para Windows
function normalizarRuta($ruta) {
    if (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN') {
        return str_replace('/', '\\', $ruta);
    }
    return $ruta;
}

// Python: usar solo forward slashes o solo backslashes, no mezclar
$venv_python = normalizarRuta(dirname(__DIR__) . '/.venv/Scripts/python.exe');

if (file_exists($venv_python)) {
    define('LOCAL_PYTHON_PATH', $venv_python);
} else {
    define('LOCAL_PYTHON_PATH', 'python');
}

// Directorios — usar DIRECTORY_SEPARATOR para consistencia
define('LOCAL_BASE_DIR',     normalizarRuta(dirname(__DIR__)));
define('LOCAL_BACKEND_DIR',  normalizarRuta(LOCAL_BASE_DIR . '/Python_principales'));
define('LOCAL_SALIDAS_DIR',  normalizarRuta(LOCAL_BACKEND_DIR . '/salidas'));
define('LOCAL_UPLOADS_DIR',  normalizarRuta(__DIR__ . '/uploads'));

// Scripts Python
define('LOCAL_PY_SCRIPTS', [
    'cartera'      => normalizarRuta(LOCAL_BACKEND_DIR . '/procesador_cartera.py'),
    'anticipos'    => normalizarRuta(LOCAL_BACKEND_DIR . '/procesador_anticipos.py'),
    'modelo_deuda' => normalizarRuta(LOCAL_BACKEND_DIR . '/modelo_deuda.py'),
    'focus'        => normalizarRuta(LOCAL_BACKEND_DIR . '/procesar_y_actualizar_focus.py'),
]);

// Crear directorios si no existen
if (!is_dir(LOCAL_SALIDAS_DIR)) {
    @mkdir(LOCAL_SALIDAS_DIR, 0777, true);
}
if (!is_dir(LOCAL_UPLOADS_DIR)) {
    @mkdir(LOCAL_UPLOADS_DIR, 0777, true);
}

// Función para verificar que Python esté disponible
function verificar_python() {
    $command = LOCAL_PYTHON_PATH . ' --version 2>&1';
    exec($command, $output, $returnCode);
    
    if ($returnCode !== 0) {
        throw new Exception('Python no est\u00e1 disponible. Verificar LOCAL_PYTHON_PATH en config_local.php');
    }
    
    return implode("\n", $output);
}

// Función para ejecutar script Python
function ejecutar_python($scriptPath, $args = []) {
    if (!file_exists($scriptPath)) {
        throw new Exception("Script no encontrado: $scriptPath");
    }

    $isWindows = (strtoupper(substr(PHP_OS, 0, 3)) === 'WIN');

    if ($isWindows) {
        $command = 'chcp 65001 > nul & ';
        $command .= '"' . LOCAL_PYTHON_PATH . '" "' . $scriptPath . '"';
        foreach ($args as $arg) {
            $command .= ' "' . $arg . '"';
        }
    } else {
        $command = escapeshellarg(LOCAL_PYTHON_PATH) . ' ' . escapeshellarg($scriptPath);
        foreach ($args as $arg) {
            $command .= ' ' . escapeshellarg($arg);
        }
    }

    $command .= ' 2>&1';

    error_log("[PYTHON] Comando: " . $command);

    // ── NUEVO: usar proc_open para capturar stdout + stderr de forma fiable ──
    $descriptors = [
        0 => ['pipe', 'r'],  // stdin
        1 => ['pipe', 'w'],  // stdout
        2 => ['pipe', 'w'],  // stderr
    ];

    $env = null; // hereda el entorno del proceso padre
    $process = proc_open($command, $descriptors, $pipes, dirname($scriptPath), $env);

    if (!is_resource($process)) {
        return [
            'success'   => false,
            'output'    => 'proc_open() falló — exec() puede estar deshabilitado en php.ini',
            'exit_code' => -1,
            'command'   => $command,
        ];
    }

    fclose($pipes[0]); // no necesitamos stdin

    $stdout = stream_get_contents($pipes[1]);
    $stderr = stream_get_contents($pipes[2]);
    fclose($pipes[1]);
    fclose($pipes[2]);

    $returnCode = proc_close($process);

    // Combinar stdout y stderr para no perder nada
    $outputText = trim($stdout . "\n" . $stderr);

    error_log("[PYTHON] Exit code: " . $returnCode);
    error_log("[PYTHON] STDOUT: " . $stdout);
    error_log("[PYTHON] STDERR: " . $stderr);

    return [
        'success'   => ($returnCode === 0),
        'output'    => $outputText,
        'exit_code' => $returnCode,
        'command'   => $command,
    ];
}