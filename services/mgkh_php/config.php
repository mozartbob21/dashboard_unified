<?php
// Загрузка .env из корня проекта + константы модуля.

function wr_load_env(): array {
    $root = dirname(__DIR__, 2); // .../unified_dashboard
    $file = $root . DIRECTORY_SEPARATOR . '.env';
    $env  = [];
    if (is_file($file)) {
        foreach (file($file, FILE_IGNORE_NEW_LINES | FILE_SKIP_EMPTY_LINES) as $line) {
            $line = trim($line);
            if ($line === '' || $line[0] === '#') continue;
            $pos = strpos($line, '=');
            if ($pos === false) continue;
            $k = trim(substr($line, 0, $pos));
            $v = trim(substr($line, $pos + 1));
            if (strlen($v) >= 2 && ($v[0] === '"' || $v[0] === "'")) {
                $v = substr($v, 1, -1);
            }
            $env[$k] = $v;
        }
    }
    // приоритет реальным переменным окружения
    foreach (['MGKH_URL','MGKH_API_KEY','MGKH_QUERY_ID','MGKH_EXTEND_DAYS'] as $k) {
        $g = getenv($k);
        if ($g !== false && $g !== '') $env[$k] = $g;
    }
    return $env;
}

$ENV = wr_load_env();

define('WR_RM_URL',      rtrim($ENV['MGKH_URL'] ?? 'https://mgkh.rm.mosreg.ru', '/'));
define('WR_RM_KEY',      $ENV['MGKH_API_KEY'] ?? '');
define('WR_QUERY_ID',    (int)($ENV['MGKH_QUERY_ID'] ?? 23));
define('WR_EXTEND_DAYS', (int)($ENV['MGKH_EXTEND_DAYS'] ?? 30));

// Ссылка на онлайн-таблицу жалоб (Яндекс.Диск public key)
define('WR_COMP_URL', $ENV['WATER_COMP_URL'] ?? 'https://disk.yandex.ru/i/NQTPtLj_LVuymw');

// ID пользовательского поля «Причина возврата»
define('WR_CF_REASON', (int)($ENV['WATER_CF_REASON'] ?? 264));

// Соответствие числовых значений «Тип проблемы»
$WR_TYPE_MAP = [
    '334' => 'Резонансные обращения',
    '335' => 'Системный адрес',
    '336' => 'Системные отключения',
];