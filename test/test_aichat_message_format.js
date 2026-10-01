// Run with: node --test test/test_aichat_message_format.js
const assert = require('node:assert/strict');
const {test} = require('node:test');
const {readFileSync} = require('node:fs');
const {join} = require('node:path');
const vm = require('node:vm');
const context = vm.createContext({});
vm.runInContext(readFileSync(join(__dirname, '../static/aichat-message-format.js'), 'utf8'), context);
const render = context.NeuronaChatFormat.render;

test('renders report headings, paragraphs and emphasis', () => {
  const html = render('## Отчёт по округу\r\n\r\n**Главное:** 12 объектов.\n*Проверено сегодня.*\n\nСледующий абзац.');
  assert.match(html, /<h3 class="ac-text-heading">Отчёт по округу<\/h3>/);
  assert.match(html, /<p><strong>Главное:<\/strong> 12 объектов\.<br><em>Проверено сегодня\.<\/em><\/p>/);
  assert.match(html, /<p>Следующий абзац\.<\/p>/);
  assert.equal(render('# Пример C#'), '<h3 class="ac-text-heading">Пример C#</h3>');
  assert.equal(render('## Раздел ##'), '<h3 class="ac-text-heading">Раздел</h3>');
});
test('groups semantic lists with nested items and ordered start', () => {
  const html = render('- Один\n- Два\n  - Вложенный\n    пояснение\n\n3. Проверить\n4. Отправить');
  assert.equal(html, '<ul><li>Один</li><li>Два<ul><li>Вложенный<br>пояснение</li></ul></li></ul>' +
    '<ol start="3"><li>Проверить</li><li>Отправить</li></ol>');
  assert.equal(render('- Родитель\n  - Подпункт\n  Продолжение'),
    '<ul><li>Родитель<ul><li>Подпункт</li></ul><br>Продолжение</li></ul>');
});
test('keeps code literal, escaped and separate from formatted text', () => {
  const html = render('Поле `**status**`\n\n```html\n<script>alert("x")</script>\n**plain**\n```\n\nГотово.');
  assert.match(html, /<p>Поле <code>\*\*status\*\*<\/code><\/p>/);
  assert.match(html, /<pre class="ac-code"><code>&lt;script&gt;alert\(&quot;x&quot;\)&lt;\/script&gt;\n\*\*plain\*\*<\/code><\/pre>/);
  assert.doesNotMatch(html, /<script>/);
  assert.match(render('~~~python\nprint(1)'), /<pre class="ac-code"><code>print\(1\)<\/code><\/pre>/);
});
test('renders comparison tables, alignment and escaped pipe content', () => {
  const html = render('| Округ | Заявок |\n| :--- | ---: |\n| **Химки** | 12 |\n| `А|Б` | 3 |\n| А\\|Б | 4 |\n\nПосле таблицы.');
  assert.match(html, /tabindex="0" role="region" aria-label="Таблица в ответе"/);
  assert.match(html, /<th scope="col" class="ac-align-end">Заявок<\/th>/);
  assert.match(html, /<td class="ac-align-start"><strong>Химки<\/strong><\/td>/);
  assert.match(html, /<td class="ac-align-start"><code>А\|Б<\/code><\/td>/);
  assert.match(html, /<td class="ac-align-start">А\|Б<\/td>/);
  assert.match(html, /<p>После таблицы\.<\/p>$/);
});
test('malformed and unsupported markdown stay readable without lost rows', () => {
  const html = render('| Округ | Количество |\n| не разделитель | 2 |\n\n**незакрыто\n\n[Источник](https://example.org)');
  assert.doesNotMatch(html, /<table>|<a /);
  assert.match(html, /не разделитель/);
  assert.match(html, /\*\*незакрыто/);
  assert.match(html, /\[Источник\]\(https:\/\/example.org\)/);
  assert.match(render('| А | Б |\n| --- | --- |\n| 1 | 2 | 3 |'), /<p>\| 1 \| 2 \| 3 \|<\/p>/);
});
test('escapes model HTML without active URLs, images or event handlers', () => {
  const payloads = ['<img src=x onerror=alert(1)>', '<script>alert(1)</script>', '**<svg onload=alert(1)>**',
    '[click](javascript:alert(1)) ![remote](https://example.org/pixel)',
    '| <iframe src="https://example.org"> | x |\n| --- | --- |\n| <button onclick="x()"> | y |',
    '# </h3><img src=x>', '```html\n</code></pre><script>alert(1)</script>\n```'];
  for (const text of payloads) {
    const html = render(text);
    assert.doesNotMatch(html, /<(?:script|img|svg|iframe|button|a)(?:\s|>)/i);
    assert.doesNotMatch(html, /<[^>]+\s(?:src|href|on\w+)\s*=/i);
  }
  assert.equal(render('&lt;img src=x&gt;'), '<p>&amp;lt;img src=x&amp;gt;</p>');
});
test('preserves long text and unbroken values, handles empty inputs', () => {
  const long = 'Солнечногорск '.repeat(20000) + 'a'.repeat(12000);
  assert.equal(render(long), '<p>' + long + '</p>');
  assert.equal(render(null), '');
  assert.equal(render('  \n\n'), '');
  assert.equal(render(0), '<p>0</p>');
});
