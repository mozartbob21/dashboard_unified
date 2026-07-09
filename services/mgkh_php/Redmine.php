<?php
declare(strict_types=1);

/**
 * Клиент Redmine REST API.
 * Читает список задач и отдельные задачи. Записи пока нет.
 */
class Redmine
{
    private string $url;
    private string $apiKey;

    public function __construct(string $url, string $apiKey)
    {
        $this->url    = rtrim($url, '/');
        $this->apiKey = $apiKey;
    }

    /** GET-запрос. Возвращает [ok, http_code, data|error]. */
    private function get(string $path, array $query = []): array
    {
        $qs   = $query ? ('?' . http_build_query($query)) : '';
        $full = $this->url . $path . $qs;

        $ch = curl_init($full);
        curl_setopt_array($ch, [
            CURLOPT_RETURNTRANSFER => true,
            CURLOPT_TIMEOUT        => 30,
            CURLOPT_HTTPHEADER     => [
                'X-Redmine-API-Key: ' . $this->apiKey,
                'Content-Type: application/json',
            ],
            CURLOPT_SSL_VERIFYPEER => true,
        ]);

        $body = curl_exec($ch);
        $err  = curl_error($ch);
        $code = (int)curl_getinfo($ch, CURLINFO_HTTP_CODE);

        if ($body === false) {
            return ['ok' => false, 'http_code' => $code, 'error' => "cURL: $err"];
        }

        $data = json_decode($body, true);
        if ($data === null && $code >= 400) {
            return ['ok' => false, 'http_code' => $code, 'error' => "HTTP $code: " . substr($body, 0, 300)];
        }

        return ['ok' => $code < 400, 'http_code' => $code, 'data' => $data];
    }

    /** Список задач по сохранённому фильтру (query_id). */
    public function getIssues(int $queryId, int $limit = 100, int $offset = 0): array
    {
        return $this->get('/issues.json', [
            'query_id' => $queryId,
            'limit'    => $limit,
            'offset'   => $offset,
        ]);
    }

    /** Одна задача целиком (с кастомными полями и, при желании, журналами). */
    public function getIssue(int $id, array $include = ['journals']): array
    {
        return $this->get("/issues/{$id}.json", [
            'include' => implode(',', $include),
        ]);
    }
}
