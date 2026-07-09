<?php
declare(strict_types=1);

require __DIR__ . '/Redmine.php';
require __DIR__ . '/Checker.php';

function log_stderr(string $msg): void
{
    fwrite(STDERR, 'STAGE:' . $msg . PHP_EOL);
}

function out(array $payload): void
{
    echo json_encode($payload, JSON_UNESCAPED_UNICODE | JSON_UNESCAPED_SLASHES);
    exit(0);
}

try {
    $config = require __DIR__ . '/config.php';

    $url     = (string)($config['url'] ?? '');
    $apiKey  = (string)($config['api_key'] ?? '');
    $queryId = (int)($config['query_id'] ?? 0);

    if ($url === '' || $apiKey === '') {
        out([
            'ok'      => false,
            'error'   => 'Не заданы MGKH_URL или MGKH_API_KEY в .env',
            'buckets' => ['close' => [], 'extend' => [], 'rework' => []],
            'metrics' => ['total' => 0, 'close' => 0, 'extend' => 0, 'rework' => 0],
        ]);
    }

    log_stderr('Авторизация и запрос списка задач');

    $redmine = new Redmine($url, $apiKey);

    $limit  = 100;
    $offset = 0;
    $allIssues = [];

    do {
        $resp = $redmine->getIssues($queryId, $limit, $offset);

        if (!$resp['ok']) {
            out([
                'ok'      => false,
                'error'   => $resp['error'] ?? ('HTTP ' . ($resp['http_code'] ?? '?')),
                'buckets' => ['close' => [], 'extend' => [], 'rework' => []],
                'metrics' => ['total' => 0, 'close' => 0, 'extend' => 0, 'rework' => 0],
            ]);
        }

        $issues     = $resp['data']['issues']     ?? [];
        $totalCount = (int)($resp['data']['total_count'] ?? count($issues));

        foreach ($issues as $issue) {
            $allIssues[] = $issue;
        }

        $offset += $limit;
    } while ($offset < $totalCount && !empty($issues));

    log_stderr('Классификация задач');

    $buckets = (new Checker())->classify($allIssues);

    $close  = $buckets['close']  ?? [];
    $extend = $buckets['extend'] ?? [];
    $rework = $buckets['rework'] ?? [];

    $addUrl = static function (array $rows) use ($url) {
        foreach ($rows as &$r) {
            $id = $r['id'] ?? 0;
            $r['url'] = rtrim($url, '/') . '/issues/' . $id;
        }
        unset($r);
        return $rows;
    };

    log_stderr('Формирование результата');

    out([
        'ok'          => true,
        'error'       => '',
        'created_at'  => date('Y-m-d H:i:s'),
        'redmine_url' => rtrim($url, '/') . '/issues?query_id=' . $queryId,
        'buckets'     => [
            'close'  => $addUrl($close),
            'extend' => $addUrl($extend),
            'rework' => $addUrl($rework),
        ],
        'metrics'     => [
            'total'  => count($allIssues),
            'close'  => count($close),
            'extend' => count($extend),
            'rework' => count($rework),
        ],
    ]);

} catch (\Throwable $e) {
    out([
        'ok'      => false,
        'error'   => 'PHP exception: ' . $e->getMessage(),
        'buckets' => ['close' => [], 'extend' => [], 'rework' => []],
        'metrics' => ['total' => 0, 'close' => 0, 'extend' => 0, 'rework' => 0],
    ]);
}
