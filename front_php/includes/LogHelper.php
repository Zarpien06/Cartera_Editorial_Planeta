<?php
/**
 * LogHelper - utilitario de logging centralizado para PHP
 * Escribe entradas en formato JSON por línea dentro de /logs/system_log.txt
 */

final class LogHelper
{
    private const LOG_DIR = __DIR__ . '/../logs';
    private const LOG_FILE = __DIR__ . '/../logs/system_log.txt';

    /**
     * Registra un evento en el log central.
     *
     * @param string $source  Componente (p. ej. PHP_PROCESAR, TRM_API)
     * @param string $level   Nivel (INFO, ERROR, DEBUG, WARNING)
     * @param string $message Mensaje principal
     * @param array  $context Datos adicionales opcionales
     */
    public static function log(string $source, string $level, string $message, array $context = []): void
    {
        self::ensureLogDir();

        $entry = [
            'timestamp' => date('c'),
            'source' => $source,
            'level' => strtoupper($level),
            'message' => $message,
            'context' => $context,
            'request' => self::getRequestMeta(),
        ];

        $line = json_encode($entry, JSON_UNESCAPED_UNICODE) . PHP_EOL;

        $fh = @fopen(self::LOG_FILE, 'ab');
        if ($fh === false) {
            return;
        }

        if (flock($fh, LOCK_EX)) {
            fwrite($fh, $line);
            fflush($fh);
            flock($fh, LOCK_UN);
        }

        fclose($fh);
    }

    /**
     * Registra rápidamente una excepción.
     */
    public static function logException(string $source, Throwable $exception, array $context = []): void
    {
        $context['exception'] = [
            'type' => get_class($exception),
            'message' => $exception->getMessage(),
            'file' => $exception->getFile(),
            'line' => $exception->getLine(),
            'trace' => $exception->getTraceAsString(),
        ];

        self::log($source, 'ERROR', $exception->getMessage(), $context);
    }

    private static function ensureLogDir(): void
    {
        if (!is_dir(self::LOG_DIR)) {
            @mkdir(self::LOG_DIR, 0777, true);
        }
    }

    private static function getRequestMeta(): array
    {
        return [
            'method' => $_SERVER['REQUEST_METHOD'] ?? null,
            'uri' => $_SERVER['REQUEST_URI'] ?? null,
            'ip' => $_SERVER['REMOTE_ADDR'] ?? null,
            'user_agent' => $_SERVER['HTTP_USER_AGENT'] ?? null,
        ];
    }
}

