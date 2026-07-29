<?php
// Минимальный парсер XLSX: возвращает [sheetName => [ [ячейки строки], ... ], ...]

function wr_xlsx_parse(string $bytes): array {
    $tmp = tempnam(sys_get_temp_dir(), 'wrxlsx');
    file_put_contents($tmp, $bytes);

    $zip = new ZipArchive();
    if ($zip->open($tmp) !== true) {
        @unlink($tmp);
        throw new RuntimeException('Не удалось открыть XLSX (ZipArchive).');
    }

    // 1. Общие строки (sharedStrings)
    $shared = [];
    $ssXml = $zip->getFromName('xl/sharedStrings.xml');
    if ($ssXml !== false) {
        $ss = simplexml_load_string($ssXml);
        if ($ss !== false) {
            foreach ($ss->si as $si) {
                $shared[] = wr_xlsx_si_text($si);
            }
        }
    }

    // 2. Имена листов из workbook.xml + связи
    $sheetsMeta = wr_xlsx_sheet_files($zip);

    $result = [];
    foreach ($sheetsMeta as $meta) {
        $name = $meta['name'];
        $path = $meta['path'];
        $xml  = $zip->getFromName($path);
        if ($xml === false) continue;
        $sheet = simplexml_load_string($xml);
        if ($sheet === false) continue;
        $result[$name] = wr_xlsx_rows($sheet, $shared);
    }

    $zip->close();
    @unlink($tmp);
    return $result;
}

function wr_xlsx_si_text(SimpleXMLElement $si): string {
    if (isset($si->t)) return (string)$si->t;
    $out = '';
    foreach ($si->r as $r) {
        if (isset($r->t)) $out .= (string)$r->t;
    }
    return $out;
}

function wr_xlsx_sheet_files(ZipArchive $zip): array {
    $wbXml = $zip->getFromName('xl/workbook.xml');
    $relXml = $zip->getFromName('xl/_rels/workbook.xml.rels');
    if ($wbXml === false) return [];

    $wb = simplexml_load_string($wbXml);
    $rels = [];
    if ($relXml !== false) {
        $rx = simplexml_load_string($relXml);
        if ($rx !== false) {
            foreach ($rx->Relationship as $rel) {
                $rels[(string)$rel['Id']] = (string)$rel['Target'];
            }
        }
    }

    $wb->registerXPathNamespace('r', 'http://schemas.openxmlformats.org/officeDocument/2006/relationships');
    $out = [];
    foreach ($wb->sheets->sheet as $sh) {
        $name = (string)$sh['name'];
        $rid  = '';
        foreach ($sh->attributes('http://schemas.openxmlformats.org/officeDocument/2006/relationships') as $k => $v) {
            if ($k === 'id') $rid = (string)$v;
        }
        $target = $rels[$rid] ?? '';
        if ($target === '') continue;
        $target = ltrim($target, '/');
        if (strpos($target, 'xl/') !== 0) $target = 'xl/' . $target;
        $out[] = ['name' => $name, 'path' => $target];
    }
    return $out;
}

function wr_xlsx_rows(SimpleXMLElement $sheet, array $shared): array {
    $rows = [];
    if (!isset($sheet->sheetData)) return $rows;

    foreach ($sheet->sheetData->row as $row) {
        $cells = [];
        $maxCol = 0;
        foreach ($row->c as $c) {
            $ref = (string)$c['r'];         // напр. "B3"
            $col = wr_col_index($ref);
            $type = (string)$c['t'];
            $val = '';
            if ($type === 's') {
                $idx = (int)$c->v;
                $val = $shared[$idx] ?? '';
            } elseif ($type === 'inlineStr') {
                $val = isset($c->is) ? wr_xlsx_si_text($c->is) : '';
            } else {
                $val = isset($c->v) ? (string)$c->v : '';
            }
            $cells[$col] = $val;
            if ($col > $maxCol) $maxCol = $col;
        }
        // выравниваем в плотный массив 0..maxCol
        $line = [];
        for ($i = 0; $i <= $maxCol; $i++) {
            $line[$i] = $cells[$i] ?? '';
        }
        $rows[] = $line;
    }
    return $rows;
}

function wr_col_index(string $ref): int {
    // "B3" -> 1 (0-based)
    $letters = '';
    for ($i = 0; $i < strlen($ref); $i++) {
        $ch = $ref[$i];
        if ($ch >= 'A' && $ch <= 'Z') $letters .= $ch;
        elseif ($ch >= 'a' && $ch <= 'z') $letters .= strtoupper($ch);
        else break;
    }
    $n = 0;
    for ($i = 0; $i < strlen($letters); $i++) {
        $n = $n * 26 + (ord($letters[$i]) - ord('A') + 1);
    }
    return $n - 1;
}