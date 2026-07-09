<?php
declare(strict_types=1);

/**
 * Классификатор задач контроля качества воды.
 * Раскладывает задачи по 3 корзинам: close / extend / rework
 */
final class Checker
{
    // ID кастомных полей Redmine
    private const F_FAST_SOLUTION = 222; // Быстрое решение (тут ищем "не требуется")
    private const F_FLUSH_MKD     = 43;  // Промывка МКД ХВС и ГВС — дата промывки
    private const F_TURBIDITY     = 229; // Установка датчиков мутности
    private const F_SYS           = 257; // SYS номер в ZULUGIS
    private const F_DEADLINE      = 226; // Срок решения

    private const STUB_DATE       = '0001-01-01';
    private const OBSERVE_DAYS    = 30;

    /**
     * @param array $issues массив задач из Redmine API
     * @return array{close: array, extend: array, rework: array}
     */
    public function classify(array $issues): array
    {
        $result = ['close' => [], 'extend' => [], 'rework' => []];
        $today  = new DateTimeImmutable('today');

        foreach ($issues as $issue) {
            $cf       = $this->indexCustomFields($issue['custom_fields'] ?? []);
            $problems = $this->findProblems($cf);

            $flushDate = $this->parseDate($this->cfVal($cf, self::F_FLUSH_MKD));

            // --- Возврат: есть ошибки ИЛИ нет даты промывки ---
            if ($problems !== [] || $flushDate === null) {
                if ($flushDate === null && $problems === []) {
                    $problems[] = 'нет даты промывки (поле «Промывка МКД ХВС и ГВС»)';
                }
                $result['rework'][] = [
                    'id'      => $issue['id'],
                    'subject' => $issue['subject'] ?? '',
                    'comment' => implode('; ', $problems),
                ];
                continue;
            }

            // --- Ошибок нет, промывка была: считаем окно наблюдения ---
            $deadline = $flushDate->modify('+' . self::OBSERVE_DAYS . ' days');

            if ($today >= $deadline) {
                $result['close'][] = [
                    'id'      => $issue['id'],
                    'subject' => $issue['subject'] ?? '',
                ];
            } else {
                $result['extend'][] = [
                    'id'          => $issue['id'],
                    'subject'     => $issue['subject'] ?? '',
                    'flush_date'  => $flushDate->format('d.m.Y'),
                    'due_date'    => $deadline->format('d.m.Y'),
                    'comment'     => sprintf(
                        'промывка %s, окно наблюдения до %s',
                        $flushDate->format('d.m.Y'),
                        $deadline->format('d.m.Y')
                    ),
                ];
            }
        }

        return $result;
    }

    /** Ищет все проблемы в задаче. Пустой массив = ошибок нет. */
    private function findProblems(array $cf): array
    {
        $problems = [];

        // 1. Датчики мутности: заглушка без обоснования "не требуется"
        $turbidity = $this->cfVal($cf, self::F_TURBIDITY);
        if ($turbidity === self::STUB_DATE) {
            $fast = mb_strtolower($this->cfVal($cf, self::F_FAST_SOLUTION));
            $hasWaiver = str_contains($fast, 'датчик') && str_contains($fast, 'не требуется');
            if (!$hasWaiver) {
                $problems[] = 'датчики мутности: заглушка (01.01.0001) без обоснования «не требуется»';
            }
        }

        // 2. SYS номер: пусто или 0
        $sys = trim($this->cfVal($cf, self::F_SYS));
        if ($sys === '' || $sys === '0') {
            $problems[] = 'SYS в ZuluGIS не привязан: «' . ($sys === '' ? '(пусто)' : $sys) . '»';
        }

        // 3. Срок решения: заглушка
        if ($this->cfVal($cf, self::F_DEADLINE) === self::STUB_DATE) {
            $problems[] = 'срок решения невалиден (01.01.0001)';
        }

        // 4. Жалоба после промывки (если дату удалось вытащить из текста)
        $complaintNote = $this->checkComplaintAfterFlush($cf);
        if ($complaintNote !== null) {
            $problems[] = $complaintNote;
        }

        return $problems;
    }

    /**
     * Проверяет, была ли жалоба ПОСЛЕ промывки.
     * @return string|null текст проблемы или null (нет проблемы / дата неизвестна)
     */
    private function checkComplaintAfterFlush(array $cf): ?string
    {
        $flush = $this->parseDate($this->cfVal($cf, self::F_FLUSH_MKD));
        if ($flush === null) {
            return null;
        }

        $text = $this->cfVal($cf, 265); // Текст жалобы
        $complaintDate = $this->extractDateFromText($text);
        if ($complaintDate === null) {
            return null; // дата неизвестна — не считаем ошибкой
        }

        if ($complaintDate > $flush) {
            return sprintf(
                'жалоба %s уже ПОСЛЕ промывки %s — мероприятие не помогло, требуется иное решение',
                $complaintDate->format('d.m.Y'),
                $flush->format('d.m.Y')
            );
        }

        return null;
    }

    /** Вытаскивает первую дату формата ДД.ММ.ГГГГ из текста. */
    private function extractDateFromText(string $text): ?DateTimeImmutable
    {
        if (preg_match('/\b(\d{2})\.(\d{2})\.(\d{4})\b/u', $text, $m)) {
            $d = DateTimeImmutable::createFromFormat('!d.m.Y', "{$m[1]}.{$m[2]}.{$m[3]}");
            return $d ?: null;
        }
        return null;
    }

    /** Парсит дату Redmine (Y-m-d). Заглушку 0001-01-01 считает отсутствием даты. */
    private function parseDate(string $val): ?DateTimeImmutable
    {
        $val = trim($val);
        if ($val === '' || $val === self::STUB_DATE) {
            return null;
        }
        $d = DateTimeImmutable::createFromFormat('!Y-m-d', $val);
        return $d ?: null;
    }

    /** Индексирует custom_fields по id для быстрого доступа. */
    private function indexCustomFields(array $fields): array
    {
        $out = [];
        foreach ($fields as $f) {
            $val = $f['value'] ?? '';
            if (is_array($val)) {
                $val = implode(', ', $val);
            }
            $out[(int)($f['id'] ?? 0)] = (string)$val;
        }
        return $out;
    }

    /** Значение поля по id (пустая строка если нет). */
    private function cfVal(array $cf, int $id): string
    {
        return $cf[$id] ?? '';
    }
}