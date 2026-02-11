<?php
/**
 * Visor sencillo de logs JSON centralizados.
 * Lee el archivo logs/system_log.txt y permite filtros básicos.
 */

$logFile = __DIR__ . '/system_log.txt';
$entries = [];
$stats = [
    'INFO' => 0,
    'WARNING' => 0,
    'ERROR' => 0,
    'DEBUG' => 0,
];

// Filtros mejorados
$levelFilter = isset($_GET['level']) ? strtoupper(trim($_GET['level'])) : '';
$sourceFilter = isset($_GET['source']) ? trim($_GET['source']) : '';
$searchFilter = isset($_GET['q']) ? trim($_GET['q']) : '';
$processFilter = isset($_GET['process']) ? trim($_GET['process']) : ''; // Nuevo filtro por proceso

if (file_exists($logFile)) {
    $lines = @file($logFile, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES);
    if ($lines !== false) {
        // Mostrar los últimos 2000 eventos para mejor rendimiento y más contexto
        $lines = array_slice($lines, -2000);
        foreach ($lines as $line) {
            $decoded = json_decode($line, true);
            if (!is_array($decoded)) {
                continue;
            }

            $matchesLevel = $levelFilter ? strtoupper($decoded['level'] ?? '') === $levelFilter : true;
            $matchesSource = $sourceFilter ? stripos($decoded['source'] ?? '', $sourceFilter) !== false : true;
            $matchesSearch = true;
            if ($searchFilter) {
                $haystack = json_encode($decoded, JSON_UNESCAPED_UNICODE);
                $matchesSearch = stripos($haystack, $searchFilter) !== false;
            }
            
            // Nuevo filtro por proceso
            $matchesProcess = true;
            if ($processFilter) {
                if ($processFilter === 'cartera') {
                    $matchesProcess = stripos($decoded['source'] ?? '', 'CARTERA') !== false || 
                                     stripos($decoded['source'] ?? '', 'PY_PROCESADOR_CARTERA') !== false;
                } elseif ($processFilter === 'anticipos') {
                    $matchesProcess = stripos($decoded['source'] ?? '', 'ANTICIPOS') !== false || 
                                     stripos($decoded['source'] ?? '', 'PY_PROCESADOR_ANTICIPOS') !== false;
                } elseif ($processFilter === 'modelo_deuda') {
                    $matchesProcess = stripos($decoded['source'] ?? '', 'MODELO_DEUDA') !== false || 
                                     stripos($decoded['source'] ?? '', 'PY_MODELO_DEUDA') !== false;
                } elseif ($processFilter === 'focus') {
                    $matchesProcess = stripos($decoded['source'] ?? '', 'FOCUS') !== false || 
                                     stripos($decoded['source'] ?? '', 'PY_FOCUS') !== false;
                }
            }

            if ($matchesLevel && $matchesSource && $matchesSearch && $matchesProcess) {
                $entries[] = $decoded;
                $levelKey = strtoupper($decoded['level'] ?? '');
                if (isset($stats[$levelKey])) {
                    $stats[$levelKey]++;
                }
            }
        }
        // Ordenar por timestamp descendente
        usort($entries, fn($a, $b) => strcmp($b['timestamp'] ?? '', $a['timestamp'] ?? ''));
    }
}

function h(?string $value): string {
    return htmlspecialchars($value ?? '', ENT_QUOTES, 'UTF-8');
}

function renderContext(array $entry): string {
    $parts = [];
    
    // Verificar si el contexto contiene información adicional de los scripts Python
    if (!empty($entry['context'])) {
        // Mostrar información de ubicación para mensajes de Python si está disponible
        if (isset($entry['context']['filename']) && isset($entry['context']['lineno'])) {
            $locationInfo = [];
            if (!empty($entry['context']['filename'])) {
                $locationInfo[] = "Archivo: " . h($entry['context']['filename']);
            }
            if (!empty($entry['context']['lineno'])) {
                $locationInfo[] = "Línea: " . h($entry['context']['lineno']);
            }
            if (!empty($entry['context']['func'])) {
                $locationInfo[] = "Función: " . h($entry['context']['func']);
            }
            
            if (!empty($locationInfo)) {
                $parts[] = '<div class="location-info">' . implode(" | ", $locationInfo) . '</div>';
            }
        }
        
        // Mostrar el contexto completo en formato JSON si hay más información
        $filteredContext = array_filter($entry['context'], function($key) {
            return !in_array($key, ['filename', 'lineno', 'func', 'module', 'pathname', 'process', 'thread']);
        }, ARRAY_FILTER_USE_KEY);
        
        if (!empty($filteredContext)) {
            $parts[] = '<div class="context-block"><h4>Contexto Adicional</h4><pre>' .
                h(json_encode($filteredContext, JSON_UNESCAPED_UNICODE | JSON_PRETTY_PRINT)) .
                '</pre></div>';
        }
    }
    
    if (!empty($entry['request'])) {
        $parts[] = '<div class="context-block"><h4>Request</h4><ul>' .
            '<li><strong>Método:</strong> ' . h($entry['request']['method'] ?? '') . '</li>' .
            '<li><strong>URI:</strong> ' . h($entry['request']['uri'] ?? '') . '</li>' .
            '<li><strong>IP:</strong> ' . h($entry['request']['ip'] ?? '') . '</li>' .
            '<li><strong>User-Agent:</strong> ' . h($entry['request']['user_agent'] ?? '') . '</li>' .
            '</ul></div>';
    }
    
    if (empty($parts)) {
        return '<em>Sin datos adicionales</em>';
    }
    return implode('', $parts);
}

// Función para formatear la fecha y hora de forma más legible
function formatTimestamp($timestamp) {
    // Convertir a objeto DateTime
    $dt = new DateTime($timestamp);
    // Formatear en formato legible: DD/MM/YYYY HH:MM:SS
    return $dt->format('d/m/Y H:i:s');
}

// Función para determinar si un mensaje contiene información de verificación de sumas
function isSumVerification($message) {
    $sumKeywords = [
        'SALDO TOTAL',
        'TOTAL DE REGISTROS',
        'REGISTROS VENCIDOS',
        'REGISTROS POR VENCER',
        'SUMA DE MONTOS',
        'VERIFICANDO',
        'VALIDANDO',
        'COMPARANDO',
        'DISCREPANCIA',
        'CORREGIDO'
    ];
    
    foreach ($sumKeywords as $keyword) {
        if (stripos($message, $keyword) !== false) {
            return true;
        }
    }
    return false;
}

// Función para determinar si un mensaje contiene información de errores
function isErrorOrWarning($message) {
    return stripos($message, 'ERROR') !== false || stripos($message, 'WARNING') !== false || stripos($message, 'ADVERTENCIA') !== false;
}
?>
<!DOCTYPE html>
<html lang="es">
<head>
    <meta charset="UTF-8">
    <title>Visor de Logs</title>
    <style>
        :root {
            --error: #f87171;
            --warning: #fbbf24;
            --info: #3b82f6;
            --debug: #6b7280;
            --cartera: #8b5cf6;
            --anticipos: #10b981;
            --modelo-deuda: #0ea5e9;
            --focus: #8b5cf6;
            --php: #f97316;
            --sum-verification: #ec4899;
            --error-warning: #ef4444;
        }
        body { font-family: 'Inter', Arial, sans-serif; background: #0f172a; margin: 0; color: #e2e8f0; }
        header { padding: 20px 30px; border-bottom: 1px solid #1f2937; }
        h1 { margin: 0; font-size: 1.5rem; }
        header p { margin: 4px 0 0; color: #94a3b8; font-size: 0.9rem; }
        .cards { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 15px; }
        .card { flex: 1; min-width: 140px; background: #1e293b; border-radius: 8px; padding: 12px; border: 1px solid #0f172a; position: relative; overflow: hidden; }
        .card::after { content: ""; position: absolute; inset: 0; background: linear-gradient(90deg, transparent, rgba(255,255,255,0.08), transparent); transform: translateX(-100%); transition: transform .6s; }
        .card:hover::after { transform: translateX(100%); }
        .card-title { font-size: 0.75rem; color: #94a3b8; margin-bottom: 3px; text-transform: uppercase; letter-spacing: 0.05em; }
        .card-value { font-size: 1.5rem; font-weight: 700; color: #f8fafc; }
        .filters { display: flex; gap: 8px; flex-wrap: wrap; padding: 15px 30px; background: #0b1120; border-bottom: 1px solid #1f2937; }
        .filters input, .filters select, .filters button, .filters a { padding: 8px 10px; border-radius: 6px; border: 1px solid #1f2937; background: #111827; color: #e2e8f0; font-size: 0.85rem; }
        .filters button { background: #2563eb; border-color: #2563eb; cursor: pointer; }
        .filters a { text-decoration: none; line-height: 30px; }
        .timeline { padding: 15px 30px 40px; display: flex; flex-direction: column; gap: 8px; }
        .event { border-left: 3px solid #1f2937; background: #0b1120; padding: 10px 12px; border-radius: 6px; box-shadow: 0 2px 6px rgba(15,23,42,0.5); }
        .event-error { border-color: var(--error); }
        .event-warning { border-color: var(--warning); }
        .event-info { border-color: var(--info); }
        .event-debug { border-color: var(--debug); }
        .event-cartera { border-color: var(--cartera); }
        .event-anticipos { border-color: var(--anticipos); }
        .event-modelo-deuda { border-color: var(--modelo-deuda); }
        .event-focus { border-color: var(--focus); }
        .event-header { display: flex; justify-content: space-between; align-items: center; gap: 6px; flex-wrap: wrap; }
        .event-time { font-size: 0.75rem; color: #94a3b8; }
        .event-message { margin: 6px 0; font-size: 0.9rem; }
        .event-body { border-top: 1px solid #1f2937; padding-top: 8px; margin-top: 6px; }
        .badge { padding: 2px 6px; border-radius: 999px; font-size: 0.65rem; font-weight: 600; color: #0f172a; }
        .badge.INFO { background: #3b82f6; }
        .badge.ERROR { background: #f87171; }
        .badge.WARNING { background: #facc15; color: #111827; }
        .badge.DEBUG { background: #9ca3af; color: #111827; }
        .badge.CARTERA { background: #8b5cf6; }
        .badge.ANTICIPOS { background: #10b981; }
        .badge.MODELO_DEUDA { background: #0ea5e9; }
        .badge.FOCUS { background: #8b5cf6; }
        .badge.PHP_PROCESAR { background: #f97316; }
        .context { display: grid; gap: 8px; }
        .context-block { background: #111827; padding: 8px; border-radius: 5px; border: 1px solid #1f2937; }
        .context-block h4 { margin: 0 0 4px; font-size: 0.75rem; color: #cbd5f5; letter-spacing: 0.05em; text-transform: uppercase; }
        pre { margin: 0; font-size: 0.75rem; color: #f8fafc; white-space: pre-wrap; }
        ul { margin: 0; padding-left: 14px; font-size: 0.75rem; }
        .process-filter { display: flex; gap: 8px; margin-bottom: 12px; }
        .process-btn { padding: 6px 12px; border-radius: 5px; border: 1px solid #1f2937; background: #111827; color: #e2e8f0; cursor: pointer; text-decoration: none; font-size: 0.85rem; }
        .process-btn.active { background: #2563eb; border-color: #2563eb; }
        .process-btn.cartera { color: #c4b5fd; }
        .process-btn.anticipos { color: #6ee7b7; }
        .process-btn.modelo-deuda { color: #7dd3fc; }
        .process-btn.focus { color: #c4b5fd; }
        .event-source { font-size: 0.8rem; color: #94a3b8; }
        /* Estilos para mostrar logs más compactos */
        .compact-view .event { padding: 6px 10px; }
        .compact-view .event-message { margin: 4px 0; font-size: 0.85rem; }
        .compact-view .event-header { gap: 4px; }
        .compact-view .event-time { font-size: 0.7rem; }
        .compact-view .badge { padding: 1px 5px; font-size: 0.6rem; }
        .compact-view .event-body { padding-top: 6px; margin-top: 4px; }
        .compact-view .context-block { padding: 6px; }
        .compact-view .context-block h4 { font-size: 0.7rem; margin: 0 0 3px; }
        .compact-view pre, .compact-view ul { font-size: 0.7rem; }
        
        /* Estilos para agrupación por proceso */
        .process-group { margin-bottom: 30px; }
        .process-header { 
            padding: 12px 15px; 
            background: #1e293b; 
            border-radius: 8px; 
            margin-bottom: 15px; 
            display: flex; 
            justify-content: space-between; 
            align-items: center; 
            border-left: 4px solid #4b5563;
        }
        .process-header.process-cartera { border-left-color: var(--cartera); }
        .process-header.process-anticipos { border-left-color: var(--anticipos); }
        .process-header.process-modelo-deuda { border-left-color: var(--modelo-deuda); }
        .process-header.process-focus { border-left-color: var(--focus); }
        .process-header.process-php { border-left-color: var(--php); }
        .process-header.process-other { border-left-color: #6b7280; }
        .process-count { 
            font-size: 0.8rem; 
            color: #94a3b8; 
            background: #374151; 
            padding: 2px 8px; 
            border-radius: 999px; 
        }
        .process-entries { display: flex; flex-direction: column; gap: 8px; }
        
        /* Resaltado para verificaciones de sumas y errores */
        .sum-verification { 
            border-left-color: var(--sum-verification) !important; 
            background: rgba(236, 72, 153, 0.05) !important; 
        }
        .error-warning { 
            border-left-color: var(--error-warning) !important; 
            background: rgba(239, 68, 68, 0.05) !important; 
        }
        .sum-verification .event-message, 
        .error-warning .event-message { 
            font-weight: bold; 
            color: #f8fafc; 
        }
        
        /* Mejoras visuales para la información de ubicación */
        .location-info {
            font-size: 0.75rem;
            color: #94a3b8;
            background: #111827;
            padding: 6px 10px;
            border-radius: 4px;
            border: 1px solid #1f2937;
            margin-bottom: 8px;
            font-family: 'Consolas', 'Courier New', monospace;
        }
        .event-anticipo { border-left-color: #10b981 !important; background: rgba(16,185,129,0.08) !important; }
    </style>
</head>
<body>
    <header>
        <h1>Panel de Diagnóstico</h1>
        <p>Revisa los últimos eventos registrados por PHP y Python.</p>
    </header>
    
    <!-- Filtros por proceso -->
    <div class="filters">
        <div class="process-filter">
            <a href="viewer.php" class="process-btn <?php echo !$processFilter ? 'active' : ''; ?>">Todos</a>
            <a href="viewer.php?process=cartera" class="process-btn cartera <?php echo $processFilter === 'cartera' ? 'active' : ''; ?>">Cartera</a>
            <a href="viewer.php?process=anticipos" class="process-btn anticipos <?php echo $processFilter === 'anticipos' ? 'active' : ''; ?>">Anticipos</a>
            <a href="viewer.php?process=modelo_deuda" class="process-btn modelo-deuda <?php echo $processFilter === 'modelo_deuda' ? 'active' : ''; ?>">Modelo Deuda</a>
            <a href="viewer.php?process=focus" class="process-btn focus <?php echo $processFilter === 'focus' ? 'active' : ''; ?>">FOCUS</a>
        </div>
    </div>
    
    <form class="filters" method="get">
        <?php if ($processFilter): ?>
            <input type="hidden" name="process" value="<?php echo h($processFilter); ?>">
        <?php endif; ?>
        <select name="level">
            <option value="">Nivel (todos)</option>
            <?php foreach (['INFO', 'WARNING', 'ERROR', 'DEBUG'] as $level): ?>
                <option value="<?php echo $level; ?>" <?php echo $levelFilter === $level ? 'selected' : ''; ?>>
                    <?php echo $level; ?>
                </option>
            <?php endforeach; ?>
        </select>
        <input type="text" name="source" placeholder="Filtrar por componente" value="<?php echo h($sourceFilter); ?>">
        <input type="text" name="q" placeholder="Buscar texto..." value="<?php echo h($searchFilter); ?>">
        <button type="submit">Filtrar</button>
        <a href="viewer.php<?php echo $processFilter ? '?process=' . h($processFilter) : ''; ?>">Limpiar filtros</a>
        <a href="viewer.php?q=anticipo" class="process-btn anticipos" style="margin-left:10px;">Solo anticipos</a>
    </form>

    <div class="cards">
        <div class="card">
            <div class="card-title">Eventos mostrados</div>
            <div class="card-value"><?php echo count($entries); ?></div>
        </div>
        <?php foreach ($stats as $label => $value): ?>
            <div class="card">
                <div class="card-title"><?php echo $label; ?></div>
                <div class="card-value"><?php echo $value; ?></div>
            </div>
        <?php endforeach; ?>
    </div>

    <?php if (empty($entries)): ?>
        <p style="padding: 15px 30px;">No hay eventos para mostrar con los filtros actuales.</p>
    <?php else: ?>
        <div class="timeline compact-view">
            <?php 
            // Variables para agrupar logs por proceso
            $currentProcess = '';
            $processGroups = [];
            
            // Agrupar entradas por proceso
            foreach ($entries as $entry) {
                $source = $entry['source'] ?? '';
                $process = 'OTROS';
                
                if (stripos($source, 'CARTERA') !== false || stripos($source, 'PY_PROCESADOR_CARTERA') !== false) {
                    $process = 'CARTERA';
                } elseif (stripos($source, 'ANTICIPOS') !== false || stripos($source, 'PY_PROCESADOR_ANTICIPOS') !== false) {
                    $process = 'ANTICIPOS';
                } elseif (stripos($source, 'MODELO_DEUDA') !== false || stripos($source, 'PY_MODELO_DEUDA') !== false) {
                    $process = 'MODELO_DEUDA';
                } elseif (stripos($source, 'FOCUS') !== false || stripos($source, 'PY_FOCUS') !== false) {
                    $process = 'FOCUS';
                } elseif (stripos($source, 'PHP_PROCESAR') !== false) {
                    $process = 'PHP_PROCESAR';
                }
                
                $processGroups[$process][] = $entry;
            }
            
            // Mostrar logs agrupados
            foreach ($processGroups as $processName => $processEntries): 
                // Ordenar entradas del proceso por timestamp
                usort($processEntries, fn($a, $b) => strcmp($b['timestamp'] ?? '', $a['timestamp'] ?? ''));
                ?>
                
                <div class="process-group">
                    <h3 class="process-header 
                        <?php 
                        if ($processName === 'CARTERA') {
                            echo 'process-cartera';
                        } elseif ($processName === 'ANTICIPOS') {
                            echo 'process-anticipos';
                        } elseif ($processName === 'MODELO_DEUDA') {
                            echo 'process-modelo-deuda';
                        } elseif ($processName === 'FOCUS') {
                            echo 'process-focus';
                        } elseif ($processName === 'PHP_PROCESAR') {
                            echo 'process-php';
                        } else {
                            echo 'process-other';
                        }
                        ?>">
                        <?php echo h($processName); ?>
                        <span class="process-count"><?php echo count($processEntries); ?> eventos</span>
                    </h3>
                    
                    <div class="process-entries">
                        <?php foreach ($processEntries as $entry): 
                            $levelClass = h(strtolower($entry['level'] ?? 'info'));
                            $source = h($entry['source'] ?? '');
                            $message = $entry['message'] ?? '';
                            
                            // Determinar clases especiales
                            $specialClasses = [];
                            if (isSumVerification($message)) {
                                $specialClasses[] = 'sum-verification';
                            }
                            if (isErrorOrWarning($message)) {
                                $specialClasses[] = 'error-warning';
                            }
                            if (stripos($message, 'anticipo') !== false || stripos($source, 'anticipo') !== false) {
                                $specialClasses[] = 'event-anticipo';
                            }
                            $specialClass = implode(' ', $specialClasses);
                            
                            // Determinar clase de proceso para el borde
                            $processClass = '';
                            if (stripos($source, 'CARTERA') !== false || stripos($source, 'PY_PROCESADOR_CARTERA') !== false) {
                                $processClass = 'event-cartera';
                            } elseif (stripos($source, 'ANTICIPOS') !== false || stripos($source, 'PY_PROCESADOR_ANTICIPOS') !== false) {
                                $processClass = 'event-anticipos';
                            } elseif (stripos($source, 'MODELO_DEUDA') !== false || stripos($source, 'PY_MODELO_DEUDA') !== false) {
                                $processClass = 'event-modelo-deuda';
                            } elseif (stripos($source, 'FOCUS') !== false || stripos($source, 'PY_FOCUS') !== false) {
                                $processClass = 'event-focus';
                            }
                        ?>
                        <article class="event event-<?php echo $levelClass; ?> <?php echo $processClass; ?> <?php echo $specialClass; ?>">
                            <div class="event-header">
                                <span class="badge <?php echo h($entry['level'] ?? ''); ?>"><?php echo h($entry['level'] ?? ''); ?></span>
                                <?php if ($processFilter !== 'cartera' && $processFilter !== 'anticipos' && $processFilter !== 'modelo_deuda' && $processFilter !== 'focus'): ?>
                                    <span class="badge <?php 
                                        if (stripos($source, 'CARTERA') !== false || stripos($source, 'PY_PROCESADOR_CARTERA') !== false) {
                                            echo 'CARTERA';
                                        } elseif (stripos($source, 'ANTICIPOS') !== false || stripos($source, 'PY_PROCESADOR_ANTICIPOS') !== false) {
                                            echo 'ANTICIPOS';
                                        } elseif (stripos($source, 'MODELO_DEUDA') !== false || stripos($source, 'PY_MODELO_DEUDA') !== false) {
                                            echo 'MODELO_DEUDA';
                                        } elseif (stripos($source, 'FOCUS') !== false || stripos($source, 'PY_FOCUS') !== false) {
                                            echo 'FOCUS';
                                        } else {
                                            echo h($entry['level'] ?? '');
                                        }
                                    ?>"><?php echo $source; ?></span>
                                <?php endif; ?>
                                <span class="event-time"><?php echo formatTimestamp($entry['timestamp'] ?? ''); ?></span>
                            </div>
                            <p class="event-message"><?php echo h($message); ?></p>
                            <div class="event-body">
                                <?php echo renderContext($entry); ?>
                            </div>
                        </article>
                        <?php endforeach; ?>
                    </div>
                </div>
            <?php endforeach; ?>
        </div>
    <?php endif; ?>

    <p style="padding: 0 30px 15px; font-size: 0.8rem;">
        Archivo origen: <?php echo h($logFile); ?> —
        <a href="system_log.txt" download>Descargar log completo</a>
    </p>
</body>
</html>