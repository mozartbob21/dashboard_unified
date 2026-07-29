<?php
// Логика классификации задач «Качество воды».

function wr_norm($t): string {
    if ($t === null) return '';
    $t = str_replace(["\r", "\n"], ' ', (string)$t);
    return trim(preg_replace('/\s+/u', ' ', $t));
}
function wr_low($t): string { return mb_strtolower(wr_norm($t), 'UTF-8'); }

$WR_OTP = ['все решено','причин не выявлено','причин не выявленно','не требуется','нет','-','—',
           'вода восстановлена','решено','отсутствует','нет проблемы','проблема отсутствует'];

function wr_is_otp($t): bool {
    global $WR_OTP;
    $n = wr_norm($t);
    if ($n === '') return true;
    if (in_array(rtrim(mb_strtolower($n,'UTF-8'), '.'), $WR_OTP, true)) return true;
    return mb_strlen(str_replace(' ', '', $n), 'UTF-8') < 5;
}

function wr_pdate($s): ?string {
    if (!preg_match('/(\d{1,2})\.(\d{1,2})\.(\d{2,4})/', wr_norm($s), $m)) return null;
    $d = (int)$m[1]; $mo = (int)$m[2]; $y = (int)$m[3];
    if ($y < 100) $y += 2000;
    if (!checkdate($mo, $d, $y)) return null;
    return sprintf('%04d-%02d-%02d', $y, $mo, $d);
}
function wr_dlog($s): bool {
    $iso = wr_pdate($s);
    return $iso !== null && (int)substr($iso, 0, 4) >= 2026;
}
function wr_fmt(?string $iso): string {
    if (!$iso) return '';
    [$y,$m,$d] = explode('-', $iso);
    return "$d.$m.$y";
}
function wr_today(): string { return date('Y-m-d'); }
function wr_add_days(string $iso, int $n): string {
    return date('Y-m-d', strtotime($iso . ' +' . $n . ' days'));
}

function wr_sysok($s): bool {
    $parts = array_values(array_filter(preg_split('/[,;\s]+/u', wr_norm($s))));
    if (!$parts) return false;
    foreach ($parts as $p) if (!preg_match('/^\d{3,4}$/', $p)) return false;
    return true;
}

/* ---------- адреса ---------- */
$WR_ADDR_STOP = ['московская','область','обл','россия','город','гор','мкр','мкрн','микрорайон','ул',
 'улица','пр-кт','проспект','пер','переулок','б-р','бульвар','ш','шоссе','проезд','пр-д','наб',
 'набережная','пл','площадь','дом','корп','корпус','стр','строение','кв','квартира','рп','дп',
 'поселок','посёлок','пос','село','деревня','дер','снт','тер','территория','городской','округ','го'];

function wr_addr_key($s): ?array {
    global $WR_ADDR_STOP;
    $t = str_replace('ё', 'е', wr_low($s));
    if ($t === '') return null;
    $house = '';
    if (preg_match('/(?:^|[\s,.(])(?:д|дом)\s*\.?\s*№?\s*(\d+\s*[а-я]?(?:[\/-]\d+)?)/u', $t, $m)) {
        $house = preg_replace('/\s+/u', '', $m[1]);
    } elseif (preg_match('/,\s*(\d+[а-я]?(?:[\/-]\d+)?)\s*(?:,|$)/u', $t, $m)) {
        $house = $m[1];
    }
    $clean  = preg_replace('/[^а-яa-z0-9\s-]/u', ' ', $t);
    $tokens = [];
    foreach (preg_split('/\s+/u', $clean) as $w) {
        if ($w === '' || mb_strlen($w,'UTF-8') < 3) continue;
        if (in_array($w, $WR_ADDR_STOP, true)) continue;
        if (preg_match('/^\d+[а-я]?([\/-]\d+)?$/u', $w)) continue;
        $tokens[] = $w;
    }
    if (!$tokens) return null;
    $street = array_pop($tokens);
    return ['house' => $house, 'street' => $street, 'rest' => $tokens];
}

function wr_comp_rows_for(int $id, string $subject, array $comp): array {
    $byId = $comp['byId'][$id] ?? [];
    $k = wr_addr_key($subject);
    $exact = []; $street = [];
    if ($k) {
        foreach ($comp['list'] as $rec) {
            $rk = $rec['key'];
            if (!$rk || $rk['street'] !== $k['street']) continue;
            $restOk = (!$rk['rest']) || (!$k['rest']) || count(array_intersect($rk['rest'], $k['rest'])) > 0;
            if (!$restOk) continue;
            if ($k['house'] && $rk['house'] && $rk['house'] === $k['house']) $exact[] = $rec;
            else $street[] = $rec;
        }
    }
    $rows = $byId;
    foreach ($exact as $e) if (!in_array($e, $byId, true)) $rows[] = $e;
    return $rows ?: $street;
}

/* ---------- текстовые эвристики ---------- */
$WR_ACTIONS = ['разъясн','оповещ','уведом','информир','проинформир','довед','обзвон','опрош','опрос','сообщен','связал'];
$WR_INVEST  = ['обследован','замер','разъясн','опрос','выезд','выход','провед','устранен','наладк','восстанов','информир'];
$WR_CAUSE_STUB = ['выясняется','устанавливается','уточняется','не установлена','не установлено',
                  'не выявлена','не выявлено','не определена','неизвестн','информация уточняется'];

function wr_has_any(string $low, array $keys): bool {
    foreach ($keys as $k) if (mb_strpos($low, $k) !== false) return true;
    return false;
}

/* ---------- РЕЗОНАНСНЫЕ ---------- */
function wr_check_rez(array $row): array {
    global $WR_ACTIONS, $WR_INVEST, $WR_CAUSE_STUB;

    $cause  = $row['Причина'] ?? ($row['Причина проблемы'] ?? '');
    $action = $row['Принятые меры'] ?? ($row['Меры'] ?? '');
    $result = $row['Результат'] ?? ($row['Итог'] ?? '');

    $reasons = [];

    if (wr_is_otp($cause) && wr_is_otp($action) && wr_is_otp($result)) {
        return ['decision' => 'close', 'reasons' => ['всё решено / причин нет']];
    }

    $cl = wr_low($cause); $al = wr_low($action); $rl = wr_low($result);

    if (wr_norm($cause) === '' || wr_has_any($cl, $WR_CAUSE_STUB)) {
        $reasons[] = 'причина не установлена / формулировка-заглушка';
    }
    if (wr_norm($action) === '' || (!wr_has_any($al, $WR_ACTIONS) && !wr_has_any($al, $WR_INVEST))) {
        $reasons[] = 'нет описания принятых мер';
    }
    if (wr_norm($result) === '') {
        $reasons[] = 'не заполнен результат';
    }

    if (!$reasons) {
        return ['decision' => 'close', 'reasons' => ['данные заполнены корректно']];
    }
    return ['decision' => 'rework', 'reasons' => $reasons];
}

/* ---------- СИСТЕМНЫЕ ---------- */
function wr_check_sys(array $row): array {
    $codes = $row['Коды'] ?? ($row['Системные коды'] ?? ($row['Мероприятия'] ?? ''));
    $plan  = $row['Плановая дата'] ?? ($row['План'] ?? ($row['Срок'] ?? ''));

    $reasons = [];
    if (!wr_sysok($codes)) {
        $reasons[] = 'коды мероприятий не заполнены / некорректны (ожидались 3–4 цифры)';
    }
    if (!wr_dlog($plan)) {
        $reasons[] = 'плановая дата отсутствует или ранее 2026';
    }

    if (!$reasons) return ['decision' => 'close', 'reasons' => ['коды и плановая дата корректны']];
    return ['decision' => 'rework', 'reasons' => $reasons];
}

/* ---------- ГЛАВНЫЙ АНАЛИЗ ---------- */
function wr_analyze(array $data, array $comp, int $extendDays): array {
    $out = [];
    $extReq = $data['ext_req'] ?? [];

    foreach ($data['rows'] as $row) {
        $id      = (int)$row['#'];
        $subject = $row['Тема'] ?? '';
        $type    = $row['Тип проблемы'] ?? ($row['Тип'] ?? '');

        // просьба о продлении из журнала — приоритет
        if (isset($extReq[$id])) {
            $out[] = [
                'id' => $id, 'subject' => $subject, 'type' => $type,
                'decision' => 'extend',
                'extend_to' => $extReq[$id],
                'reasons' => ['в комментариях запрошено продление до ' . wr_fmt($extReq[$id])],
            ];
            continue;
        }

        $tl = wr_low($type);
        if (mb_strpos($tl, 'систем') !== false) {
            $res = wr_check_sys($row);
        } elseif (mb_strpos($tl, 'резонанс') !== false) {
            $res = wr_check_rez($row);
        } else {
            // по умолчанию — как резонансное
            $res = wr_check_rez($row);
        }

        $item = [
            'id' => $id, 'subject' => $subject, 'type' => $type,
            'decision' => $res['decision'],
            'reasons'  => $res['reasons'],
        ];
        if ($res['decision'] === 'rework') {
            $item['extend_to'] = wr_add_days(wr_today(), $extendDays);
        }
        $out[] = $item;
    }
    return $out;
}