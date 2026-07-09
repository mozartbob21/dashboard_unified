<?php
declare(strict_types=1);
require __DIR__ . '/Redmine.php';

$config = require __DIR__ . '/config.php';
$redmine = new Redmine($config['url'], $config['api_key']);
$resp = $redmine->getIssues((int)$config['query_id'], 3, 0);

if (!$resp['ok']) {
    fwrite(STDERR, "Ошибка: " . ($resp['error'] ?? '?') . PHP_EOL);
    exit(1);
}

foreach (($resp['data']['issues'] ?? []) as $issue) {
    echo "=== Задача #{$issue['id']} ===\n";
    echo "subject: " . ($issue['subject'] ?? '') . "\n";
    echo "status: " . ($issue['status']['name'] ?? '') . "\n";
    echo "due_date: " . ($issue['due_date'] ?? '(нет)') . "\n";
    echo "start_date: " . ($issue['start_date'] ?? '(нет)') . "\n";
    echo "--- custom_fields ---\n";
    foreach (($issue['custom_fields'] ?? []) as $cf) {
        $val = is_array($cf['value'] ?? '') ? json_encode($cf['value'], JSON_UNESCAPED_UNICODE) : ($cf['value'] ?? '');
        echo "  [{$cf['id']}] «{$cf['name']}» = «{$val}»\n";
    }
    echo "\n";
}