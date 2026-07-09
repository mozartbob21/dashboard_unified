<?php
require __DIR__ . '/services/mgkh_php/Redmine.php';
require __DIR__ . '/services/mgkh_php/Checker.php';

$config  = require __DIR__ . '/services/mgkh_php/config.php';
$redmine = new Redmine($config['url'], $config['api_key']);
$resp    = $redmine->getIssues((int)$config['query_id'], 100, 0);

$buckets = (new Checker())->classify($resp['data']['issues'] ?? []);

echo "НА ЗАКРЫТИЕ: " . count($buckets['close']) . "\n";
echo "ПРОДЛЕНИЕ:   " . count($buckets['extend']) . "\n";
echo "ВОЗВРАТ:     " . count($buckets['rework']) . "\n\n";

foreach ($buckets['rework'] as $r) {
    echo "#{$r['id']}: {$r['comment']}\n";
}
