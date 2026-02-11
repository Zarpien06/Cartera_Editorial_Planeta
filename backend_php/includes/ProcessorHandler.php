<?php

/**
 * Manejador de procesadores Python
 * Gestiona la ejecución de los 4 procesadores principales
 */
class ProcessorHandler {
    private $pythonRunner;
    private $fileHandler;
    private $outputDir;
    
    public function __construct($pythonPath, $outputDir) {
        $this->pythonRunner = new PythonRunner($pythonPath);
        $this->fileHandler = new FileHandler(UPLOAD_DIR);
        $this->outputDir = $outputDir;
    }
    
    /**
     * Procesa cartera
     */
    public function processCartera($inputFile, $fechaCierre, $params = []) {
        $scriptPath = PY_SCRIPTS['cartera'];
        
        // El script espera: python procesador_cartera.py <input_file> [output_file] [fecha_cierre]
        $outputFile = $this->outputDir . '/CARTERA_' . date('Y-m-d_H-i-s') . '.xlsx';
        
        $args = [
            $inputFile,
            $outputFile,
            $fechaCierre
        ];
        
        // Agregar TRM si se proporciona
        if (!empty($params['trm_usd'])) {
            $args[] = $params['trm_usd'];
        }
        if (!empty($params['trm_eur']) && !empty($params['trm_usd'])) {
            $args[] = $params['trm_eur'];
        }
        
        $result = $this->pythonRunner->execute($scriptPath, $args);
        
        // Agregar la ruta del archivo de salida al resultado
        if ($result['success'] && file_exists($outputFile)) {
            $result['output_file'] = $outputFile;
        }
        
        return $result;
    }
    
    /**
     * Procesa anticipos
     */
    public function processAnticipos($inputFile, $params = []) {
        // Usar el script principal de procesamiento de anticipos
        $scriptPath = PY_SCRIPTS['anticipos'];
        
        // Generar nombre de archivo de salida
        $outputFile = $this->outputDir . '/ANTICIPOS_' . date('Y-m-d_H-i-s') . '.xlsx';
        
        // Argumentos: input_file [output_file]
        $args = [realpath($inputFile)];
        if ($outputFile) {
            $args[] = $outputFile;
        }
        
        $result = $this->pythonRunner->execute($scriptPath, $args);
        
        // Agregar la ruta del archivo de salida al resultado
        if ($result['success'] && file_exists($outputFile)) {
            $result['output_file'] = $outputFile;
        }
        
        return $result;
    }
    
    /**
     * Procesa modelo de deuda
     */
    public function processModeloDeuda($archivoProvision, $archivoAnticipos, $params = []) {
        $scriptPath = PY_SCRIPTS['modelo_deuda'];
        
        // El script modelo_deuda.py espera ser llamado como módulo
        // Necesitamos crear un wrapper temporal
        $wrapperScript = $this->createModeloDeudaWrapper(
            $archivoProvision, 
            $archivoAnticipos, 
            $params
        );
        
        try {
            $result = $this->pythonRunner->execute($wrapperScript, []);
            
            // Limpiar wrapper temporal
            @unlink($wrapperScript);
            
            return $result;
        } catch (Exception $e) {
            @unlink($wrapperScript);
            throw $e;
        }
    }
    
    /**
     * Procesa FOCUS
     */
    public function processFocus($params = []) {
        $scriptPath = PY_SCRIPTS['focus'];
        
        // El script procesar_y_actualizar_focus.py busca archivos automáticamente
        // pero podemos pasarle rutas específicas si se proporcionan
        $args = [];
        
        if (!empty($params['archivo_focus'])) {
            $args[] = '--archivo_focus';
            $args[] = $params['archivo_focus'];
        }
        if (!empty($params['archivo_balance'])) {
            $args[] = '--archivo_balance';
            $args[] = $params['archivo_balance'];
        }
        if (!empty($params['archivo_situacion'])) {
            $args[] = '--archivo_situacion';
            $args[] = $params['archivo_situacion'];
        }
        if (!empty($params['archivo_modelo'])) {
            $args[] = '--archivo_modelo';
            $args[] = $params['archivo_modelo'];
        }
        
        return $this->pythonRunner->execute($scriptPath, $args);
    }
    
    /**
     * Crea un script wrapper temporal para modelo_deuda.py
     */
    private function createModeloDeudaWrapper($archivoProvision, $archivoAnticipos, $params) {
        $trmFile = BASE_DIR . '/Python_principales/trm.json';
        $outputFile = !empty($params['output_file']) ? 
            $params['output_file'] : 
            $this->outputDir . '/modelo_deuda_' . date('Ymd_His') . '.xlsx';
        
        $pythonCode = <<<PYTHON
#!/usr/bin/env python3
# -*- coding: utf-8 -*-
import sys
import os

# Agregar el directorio back-end al path
sys.path.insert(0, os.path.dirname(os.path.abspath(__file__)))

# Importar la función
from modelo_deuda import crear_modelo_deuda

# Ejecutar
try:
    archivo_provision = r'{$archivoProvision}'
    archivo_anticipos = r'{$archivoAnticipos}'
    output_file = r'{$outputFile}'
    trm_file = r'{$trmFile}'
    
    print(f"Procesando modelo de deuda...")
    print(f"  Provisión: {archivo_provision}")
    print(f"  Anticipos: {archivo_anticipos}")
    print(f"  Salida: {output_file}")
    
    df_pesos, df_divisas, df_vencimientos = crear_modelo_deuda(
        archivo_provision=archivo_provision,
        archivo_anticipos=archivo_anticipos,
        output_file=output_file,
        trm_file=trm_file
    )
    
    print(f"✅ Modelo de deuda generado exitosamente: {output_file}")
    print(f"Archivo guardado en: {output_file}")
    
except Exception as e:
    print(f"❌ Error: {str(e)}", file=sys.stderr)
    import traceback
    traceback.print_exc()
    sys.exit(1)
PYTHON;
        
        $wrapperPath = sys_get_temp_dir() . '/modelo_deuda_wrapper_' . uniqid() . '.py';
        file_put_contents($wrapperPath, $pythonCode);
        
        return $wrapperPath;
    }
    
    /**
     * Obtiene el archivo de salida más reciente del directorio de salidas
     */
    public function getLatestOutputFile($pattern = '*.xlsx') {
        $files = glob($this->outputDir . '/' . $pattern);
        if (empty($files)) {
            return null;
        }
        
        usort($files, function($a, $b) {
            return filemtime($b) - filemtime($a);
        });
        
        return $files[0];
    }
    
    /**
     * Limpia archivos antiguos
     */
    public function cleanupOldFiles($maxAge = 86400) {
        return $this->fileHandler->cleanupOldFiles($maxAge);
    }
}
