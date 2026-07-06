<?php
require_once 'config_local.php';

echo "<pre>";

// 1. ¿Qué ruta de Python está usando?
echo "Python path: " . LOCAL_PYTHON_PATH . "\n";
echo "¿Existe el ejecutable? " . (file_exists(LOCAL_PYTHON_PATH) ? "SÍ" : "NO") . "\n\n";

// 2. ¿Existe el script?
$script = LOCAL_PY_SCRIPTS['focus'];
echo "Script focus: " . $script . "\n";
echo "¿Existe el script? " . (file_exists($script) ? "SÍ" : "NO") . "\n\n";

// 3. ¿proc_open está disponible?
echo "proc_open disponible: " . (function_exists('proc_open') ? "SÍ" : "NO") . "\n";
echo "exec disponible: " . (function_exists('exec') ? "SÍ" : "NO") . "\n\n";

// 4. Probar Python directamente
$cmd = '"' . LOCAL_PYTHON_PATH . '" --version 2>&1';
echo "Comando: $cmd\n";
exec($cmd, $out, $code);
echo "Exit code: $code\n";
echo "Output: " . implode("\n", $out) . "\n\n";

// 5. Directorios
echo "LOCAL_SALIDAS_DIR: " . LOCAL_SALIDAS_DIR . "\n";
echo "¿Existe salidas? " . (is_dir(LOCAL_SALIDAS_DIR) ? "SÍ" : "NO") . "\n";
echo "LOCAL_UPLOADS_DIR: " . LOCAL_UPLOADS_DIR . "\n";
echo "¿Existe uploads? " . (is_dir(LOCAL_UPLOADS_DIR) ? "SÍ" : "NO") . "\n";

echo "</pre>";