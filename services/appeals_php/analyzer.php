<?php
$CONFIG = [
    "W_LOW" => 0.25, "W_MED" => 0.50, "W_HIGH" => 0.75, "W_EXT" => 1.00,
    "COLLECTIVE_MIN" => 0.2,
    "COLLECTIVE_COEF" => 15.0,
    "SYSTEMATICITY_COEF" => 90.0,
];

$CRITICALITY_PATTERNS = [
    "MAXIMAL" => [
        'погиб\w*', 'смерть', 'летальн\w*', 'реанимаци\w*', 'сбили реб[её]нка',
        'дтп с пострадавшими', 'труп\w*', 'ударило током', 'утечка газа',
        'пожар\w*', 'взрыв\w*', 'искрит', 'электричество\s+искрит', 'проводка\s+искрит',
        'обрыв лэп', 'отравлени\w*',
    ],
    "HIGH" => [
        'открытый люк', 'провал\w*', 'аварийност\w*', 'перелом\w*', 'госпитализаци\w*',
        'сильное кровотечение', 'реб[её]нок пострадал', 'провалился в люк',
        'сбила машина', 'сотрясени\w*', 'нет воды', 'нет хвс', 'нет гвс', 'нет отопления',
        'затоплени\w*', 'прорыв\w*трубы', 'обрушени\w*',
    ],
    "MEDIUM" => [
        'упал\w*', 'травм\w*', 'ушиб\w*', 'порез\w*', 'ожог\w*', 'опасно\b',
        'угроза\b', 'слабый напор', 'нет напора', 'ржавая вода', 'грязная вода',
        'коричневая вода', 'запах канализации', 'воняет', 'черви', 'плесен\w*',
        'протечк\w*', 'залива\w*', 'течь\b',
    ],
    "LOW" => [
        'может травмироваться', 'опасно ходить', 'риск\b', 'небезопасно',
    ],
];

$DANGER_OBJECTS = ['люк', 'яма', 'провал', 'голол[её]д', 'сосульк\w*', 'провод\w*', 'крыш\w*'];
$HARM_CONSEQUENCES = ['упал', 'травм\w*', 'сломал\w*', 'кровь\w*', 'пострадал\w*', 'разбил\w*'];

function clamp01($x) {
    return max(0.0, min(1.0, floatval($x)));
}

function any_pattern_match($tl, $patterns) {
    foreach ($patterns as $pat) {
        if (preg_match('/' . $pat . '/u', $tl)) {
            return true;
        }
    }
    return false;
}

function detect_criticality_score_and_level($text) {
    global $CONFIG, $CRITICALITY_PATTERNS, $DANGER_OBJECTS, $HARM_CONSEQUENCES;

    if (empty($text)) {
        return [0.0, "НИЗКИЙ"];
    }

    $tl = str_replace("ё", "е", mb_strtolower($text, 'UTF-8'));
    $base_score = 0.0;
    $level = "НИЗКИЙ";

    $has_max  = any_pattern_match($tl, $CRITICALITY_PATTERNS["MAXIMAL"]);
    $has_high = any_pattern_match($tl, $CRITICALITY_PATTERNS["HIGH"]);
    $has_med  = any_pattern_match($tl, $CRITICALITY_PATTERNS["MEDIUM"]);
    $has_low  = any_pattern_match($tl, $CRITICALITY_PATTERNS["LOW"]);

    if ($has_max) {
        $base_score = $CONFIG["W_EXT"];
        $level = "МАКСИМАЛЬНЫЙ";
    } elseif ($has_high) {
        $base_score = $CONFIG["W_HIGH"];
        $level = "ВЫСОКИЙ";
    } elseif ($has_med) {
        $base_score = $CONFIG["W_MED"];
        $level = "СРЕДНИЙ";
    } elseif ($has_low) {
        $base_score = $CONFIG["W_LOW"];
        $level = "НИЗКИЙ";
    }

    $bonus = 0.0;
    $has_object = any_pattern_match($tl, $DANGER_OBJECTS);
    $has_harm   = any_pattern_match($tl, $HARM_CONSEQUENCES);

    if ($has_object && $has_harm) {
        $bonus = 0.10;
    } elseif ($has_object || $has_harm) {
        $bonus = 0.05;
    }

    $final_score = clamp01($base_score + $bonus);

    if ($final_score >= 0.90) {
        $level = "МАКСИМАЛЬНЫЙ";
    } elseif ($final_score >= 0.65) {
        $level = "ВЫСОКИЙ";
    } elseif ($final_score >= 0.35) {
        $level = "СРЕДНИЙ";
    } else {
        $level = "НИЗКИЙ";
    }

    return [$final_score, $level];
}


function calculate_systematicity_days_and_score($text, $address = null, $date_history = null) {
    global $CONFIG;
    $systematicity_days = 0;
    $tl = $text ? mb_strtolower($text, 'UTF-8') : "";
    $text_detected = false;
    $systematic_markers = [
        "постоянно", "каждый год", "уже не первый раз", "из года в год",
        "регулярно", "систематически", "каждый сезон", "каждую весну",
        "каждую осень", "каждую зиму", "каждое лето", "ежегодно",
        "хроническ", "перманентн",
    ];
    foreach ($systematic_markers as $word) {
        if (mb_strpos($tl, $word) !== false) {
            $text_detected = true;
            $systematicity_days = 14;
            break;
        }
    }
    if ($address && $date_history && count($date_history) > 1) {
        sort($date_history);
        $first = new DateTime($date_history[0]);
        $last  = new DateTime($date_history[count($date_history) - 1]);
        $delta = (int)$first->diff($last)->days;
        $systematicity_days = max($systematicity_days, $delta);
    }
    $score = 1.0 - exp(-$systematicity_days / $CONFIG["SYSTEMATICITY_COEF"]);
    if ($text_detected) {
        $score = clamp01($score + 0.08);
    }
    return [$systematicity_days, $score];
}

function detect_collective_complaint($text, $is_collective_status = false, $n_signers = 0) {
    global $CONFIG;
    if ($is_collective_status) {
        $score = 1.0 - exp(-$n_signers / $CONFIG["COLLECTIVE_COEF"]);
        return [true, $n_signers, $score];
    }
    if (empty($text)) {
        return [false, 0, 0.0];
    }
    $tl = mb_strtolower($text, 'UTF-8');
    $collective_patterns = [
        '\bмы\b', '\bнам\b', '\bу нас\b', '\bнаш\w*\b', '\bсосед[ия]\b',
        '\bжител[ияь]\w*\b', '\bжильц[ыа]\w*\b', 'всем домом', 'весь дом',
        'наш подъезд', 'наша улица', 'коллективное обращение', 'коллективная жалоба',
        'от лица жителей', 'просим от имени жителей', 'инициативная группа',
        'собрание жильцов', 'совет дома', 'все соседи', 'весь подъезд',
        'все жильцы', 'общедомов\w*',
    ];
    $anti_patterns = [
        'мы с мужем', 'мы с ребенком', 'мы с мамой', 'мы с женой',
        'мы с девушкой', 'мы семья', 'мы вдвоем', 'мы вдвоём',
        'мы с супруг\w*', 'мы с братом', 'мы с сестрой',
    ];
    $has_collective = any_pattern_match($tl, $collective_patterns);
    $has_anti       = any_pattern_match($tl, $anti_patterns);
    $is_collective = false;
    $score = 0.0;
    $detected_signers = 0;
    if ($has_collective) {
        if ($has_anti) {
            $strong_markers = [
                'коллективн\w*', 'инициативная группа', 'собрание жильцов',
                'жители дома', 'подписи', 'весь дом', 'весь подъезд',
            ];
            if (any_pattern_match($tl, $strong_markers)) {
                $is_collective = true;
            }
        } else {
            $is_collective = true;
        }
    }
    if ($is_collective) {
        if (preg_match('/(\d+)\s*(?:подпис(?:ей|и|антов)|жителей|человек|квартир)/u', $tl, $m)) {
            $detected_signers = (int)$m[1];
            $score = 1.0 - exp(-$detected_signers / $CONFIG["COLLECTIVE_COEF"]);
        } else {
            $score = $CONFIG["COLLECTIVE_MIN"];
            $detected_signers = 3;
        }
    }
    return [$is_collective, $detected_signers, $score];
}

// --- Тесты систематичности ---
$s1 = calculate_systematicity_days_and_score("во дворе постоянно грязь, из года в год");
$s2 = calculate_systematicity_days_and_score("просто разовая жалоба");
// --- Тесты коллективных жалоб ---
$k1 = detect_collective_complaint("мы жители дома требуем убрать свалку, 25 подписей");
$k2 = detect_collective_complaint("мы с мужем купили квартиру");
$k3 = detect_collective_complaint("во дворе яма");

$c1 = detect_criticality_score_and_level("на детской площадке пожар, есть погибшие");
$c2 = detect_criticality_score_and_level("нет горячей воды уже неделю");
$c3 = detect_criticality_score_and_level("во дворе открытый люк, ребёнок упал и разбил колено");
$c4 = detect_criticality_score_and_level("ржавая вода из крана");
$c5 = detect_criticality_score_and_level("небольшой риск поскользнуться");
$c6 = detect_criticality_score_and_level("просто спасибо за работу");

$result = [
    "created_by"  => "php_analyzer_v7",
    "c1_pozhar"   => ["score" => $c1[0], "level" => $c1[1]],
    "c2_voda"     => ["score" => $c2[0], "level" => $c2[1]],
    "c3_luk"      => ["score" => $c3[0], "level" => $c3[1]],
    "c4_rzhavaya" => ["score" => $c4[0], "level" => $c4[1]],
    "c5_risk"     => ["score" => $c5[0], "level" => $c5[1]],
    "c6_spasibo"  => ["score" => $c6[0], "level" => $c6[1]],
    "systematicity_1" => ["days" => $s1[0], "score" => $s1[1]],
    "systematicity_2" => ["days" => $s2[0], "score" => $s2[1]],
    "collective_1"    => ["is_collective" => $k1[0], "signers" => $k1[1], "score" => $k1[2]],
    "collective_2"    => ["is_collective" => $k2[0], "signers" => $k2[1], "score" => $k2[2]],
    "collective_3"    => ["is_collective" => $k3[0], "signers" => $k3[1], "score" => $k3[2]],
];

echo json_encode($result, JSON_UNESCAPED_UNICODE);
