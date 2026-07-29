<?php
// Работа с Redmine REST API (чтение задач/статусов + запись PUT).

function wr_rm_request(string $method, string $path, ?array $body = null): array {
    $url = WR_RM_URL . $path;
    $ch  = curl_init($url);

    $headers = [
        'X-Redmine-API-Key: ' . WR_RM_KEY,
        'Content-Type: application/json',
        'Accept: application/json',
    ];

    $opts = [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_TIMEOUT        => 120,
        CURLOPT_HTTPHEADER     => $headers,
        CURLOPT_CUSTOMREQUEST  => $method,
        CURLOPT_SSL_VERIFYPEER => true,
    ];
    if ($body !== null) {
        $opts[CURLOPT_POSTFIELDS] = json_encode($body, JSON_UNESCAPED_UNICODE);
    }
    curl_setopt_array($ch, $opts);

    $raw  = curl_exec($ch);
    $err  = curl_error($ch);
    $code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);

    if ($raw === false) {
        throw new RuntimeException('RM запрос не прошёл: ' . $err);
    }
    if ($code === 401) throw new RuntimeException('RM 401: неверный API-ключ');
    if ($code === 403) throw new RuntimeException('RM 403: нет прав');
    if ($code >= 400) {
        throw new RuntimeException('RM HTTP ' . $code . ' — ' . $path . ' · ' . substr($raw, 0, 300));
    }
    if ($code === 204 || $raw === '') return [];
    $js = json_decode($raw, true);
    return is_array($js) ? $js : [];
}

function wr_deiso($v): string {
    // "2026-05-01" -> "01.05.2026"
    $s = (string)($v ?? '');
    return preg_replace('/\b(\d{4})-(\d{2})-(\d{2})\b/', '$3.$2.$1', $s);
}

function wr_rm_load_issues(): array {
    global $WR_TYPE_MAP;
    $rows = [];
    $offset = 0;
    $total  = PHP_INT_MAX;

    while ($offset < $total) {
        $js = wr_rm_request('GET', '/issues.json?query_id=' . WR_QUERY_ID . '&limit=100&offset=' . $offset);
        $total  = (int)($js['total_count'] ?? 0);
        $issues = $js['issues'] ?? [];
        if (!$issues) break;

        foreach ($issues as $is) {
            $row = [
                '#'       => (string)$is['id'],
                'Трекер'  => $is['tracker']['name'] ?? '',
                'Проект'  => $is['project']['name'] ?? '',
                'Тема'    => $is['subject'] ?? '',
                'Создано' => wr_deiso(substr((string)($is['created_on'] ?? ''), 0, 10)),
            ];
            foreach (($is['custom_fields'] ?? []) as $cf) {
                $v = $cf['value'] ?? '';
                if (is_array($v)) $v = implode(', ', $v);
                $v = (string)$v;
                if (($cf['name'] ?? '') === 'Тип проблемы' && isset($WR_TYPE_MAP[$v])) {
                    $v = $WR_TYPE_MAP[$v];
                }
                $row[$cf['name']] = wr_deiso($v);
            }
            $rows[] = $row;
        }
        $offset += count($issues);
    }

    // журналы: ищем просьбы «продлить до ДД.ММ.ГГГГ»
    $extReq = [];
    foreach ($rows as $row) {
        $id = $row['#'];
        try {
            $js = wr_rm_request('GET', '/issues/' . $id . '.json?include=journals');
            foreach (($js['issue']['journals'] ?? []) as $j) {
                $notes = $j['notes'] ?? '';
                if (preg_match('/продл[а-яё]*[\s\S]{0,120}?до\s*(\d{1,2}\.\d{1,2}\.\d{2,4})/iu', $notes, $m)) {
                    $extReq[(int)$id] = wr_pdate($m[1]);
                }
            }
        } catch (Throwable $e) { /* пропускаем */ }
    }

    // собираем полный список заголовков
    $headers = [];
    foreach ($rows as $r) {
        foreach (array_keys($r) as $k) {
            if (!in_array($k, $headers, true)) $headers[] = $k;
        }
    }

    return ['rows' => $rows, 'headers' => $headers, 'ext_req' => $extReq];
}

function wr_rm_statuses(): array {
    $js = wr_rm_request('GET', '/issue_statuses.json');
    $list = $js['issue_statuses'] ?? [];
    $closed = null; $rework = null; $names = [];
    foreach ($list as $s) {
        $names[] = $s['name'];
        $ln = mb_strtolower($s['name'], 'UTF-8');
        if ($closed === null && mb_strpos($ln, 'закрыт') !== false) $closed = (int)$s['id'];
        if ($rework === null && mb_strpos($ln, 'доработ') !== false) $rework = (int)$s['id'];
    }
    if (!$closed || !$rework) {
        throw new RuntimeException('Не нашёл статусы «Закрыта»/«…на доработку». Есть: ' . implode(', ', $names));
    }
    return ['closed' => $closed, 'rework' => $rework];
}

function wr_rm_update(int $id, array $issueBody): void {
    wr_rm_request('PUT', '/issues/' . $id . '.json', ['issue' => $issueBody]);
}