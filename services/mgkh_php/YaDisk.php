<?php
// Скачивание публичного файла с Яндекс.Диска (XLSX-байты).

function wr_yadisk_download(string $publicKey): string {
    $api = 'https://cloud-api.yandex.net/v1/disk/public/resources/download?public_key='
         . urlencode($publicKey);

    $meta = wr_http_get_json($api);
    if (!isset($meta['href'])) {
        throw new RuntimeException('Яндекс.Диск: не получен href для скачивания.');
    }

    $bytes = wr_http_get_raw($meta['href']);
    if ($bytes === '' ) {
        throw new RuntimeException('Яндекс.Диск: пустой ответ при скачивании файла.');
    }
    return $bytes;
}

function wr_http_get_json(string $url): array {
    $raw = wr_http_get_raw($url);
    $js  = json_decode($raw, true);
    if (!is_array($js)) {
        throw new RuntimeException('Некорректный JSON от Яндекс.Диска.');
    }
    return $js;
}

function wr_http_get_raw(string $url): string {
    $ch = curl_init($url);
    curl_setopt_array($ch, [
        CURLOPT_RETURNTRANSFER => true,
        CURLOPT_FOLLOWLOCATION => true,
        CURLOPT_TIMEOUT        => 120,
        CURLOPT_SSL_VERIFYPEER => true,
        CURLOPT_USERAGENT      => 'unified-dashboard/water-rm',
    ]);
    $body = curl_exec($ch);
    $err  = curl_error($ch);
    $code = curl_getinfo($ch, CURLINFO_HTTP_CODE);
    curl_close($ch);
    if ($body === false) {
        throw new RuntimeException('HTTP ошибка: ' . $err);
    }
    if ($code >= 400) {
        throw new RuntimeException('HTTP ' . $code . ' при запросе к ' . $url);
    }
    return $body;
}