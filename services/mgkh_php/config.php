<?php
declare(strict_types=1);

/**
 * Читает .env из корня проекта и отдаёт конфиг МГКХ.
 * Пока НЕ делает никаких API-запросов — только загрузка настроек.
 */
function mgkh_load_env(): array
{
    $root = dirname(__DIR__, 2);          // .../unified_dashboard
    $envPath = $root . '/.env';
    $env = [];

    if (is_readable($envPath)) {
        foreach (file($envPath, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
            $line = trim($line);
            if ($line === '' || $line[0] === '#') {
                continue;
            }
            if (!str_contains($line, '=')) {
                continue;
            }
            [$k, $v] = explode('=', $line, 2);
            $env[trim($k)] = trim($v, " \t\"'");
        }
    }

    return [
        'url'         => $env['MGKH_URL']         ?? '',
        'api_key'     => $env['MGKH_API_KEY']     ?? '',
        'query_id'    => (int)($env['MGKH_QUERY_ID']    ?? 23),
        'extend_days' => (int)($env['MGKH_EXTEND_DAYS'] ?? 30),

        // Date-поля, которые считаем ошибкой при значении 0001-01-01.
        // Имена должны точно совпадать с "name" кастомных полей в Redmine.
        'checked_fields' => [
            'Установка датчиков мутности',
            'Срок решения',
        ],
    ];
}

return mgkh_load_env();
