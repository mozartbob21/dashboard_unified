/* Dependency-free report formatting. Model HTML stays text; rendering never
 * loads remote content or creates untrusted links or images. */
(function (root) {
  'use strict';
  function escapeHtml(value) {
    return String(value).replace(/[&<>"']/g, function (char) {
      return { '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[char];
    });
  }

  function inline(text, depth) {
    depth = depth || 0;
    var out = '', plain = '', i = 0;
    function flush() { out += escapeHtml(plain); plain = ''; }
    while (i < text.length) {
      var char = text.charAt(i);
      if (char === '\\' && /[\\`*_{}\[\]()#+.!|>-]/.test(text.charAt(i + 1))) {
        plain += text.charAt(i + 1); i += 2; continue;
      }
      if (char === '`') {
        var run = 1;
        while (text.charAt(i + run) === '`') run++;
        var marker = text.slice(i, i + run), end = text.indexOf(marker, i + run);
        if (end > i + run) {
          flush(); out += '<code>' + escapeHtml(text.slice(i + run, end)) + '</code>';
          i = end + run; continue;
        }
        plain += marker; i += run; continue;
      }
      if (char === '*' && depth < 4) {
        var strong = text.charAt(i + 1) === '*', delimiter = strong ? '**' : '*';
        var close = text.indexOf(delimiter, i + delimiter.length);
        if (close > i + delimiter.length && !/\s/.test(text.charAt(i + delimiter.length)) &&
            !/\s/.test(text.charAt(close - 1))) {
          flush(); var tag = strong ? 'strong' : 'em';
          out += '<' + tag + '>' + inline(text.slice(i + delimiter.length, close), depth + 1) + '</' + tag + '>';
          i = close + delimiter.length; continue;
        }
      }
      if (char === '\n') { flush(); out += '<br>'; i++; continue; }
      plain += char; i++;
    }
    flush(); return out;
  }

  function listMarker(line) {
    var match = /^(\s*)([-+*•]|\d{1,9}[.)])\s+(.*)$/.exec(line);
    if (!match) return null;
    return { indent: match[1].replace(/\t/g, '    ').length,
      ordered: /^\d/.test(match[2]), number: parseInt(match[2], 10), text: match[3] };
  }

  function tableCells(line) {
    var value = line.trim();
    if (value.indexOf('|') < 0) return null;
    var cells = [], cell = '', codeMarker = '', i = 0;
    while (i < value.length) {
      var char = value.charAt(i);
      if (char === '\\' && i + 1 < value.length) {
        cell += value.slice(i, i + 2); i += 2; continue;
      }
      if (char === '`') {
        var end = i + 1;
        while (value.charAt(end) === '`') end++;
        var marker = value.slice(i, end);
        if (!codeMarker) codeMarker = marker;
        else if (codeMarker === marker) codeMarker = '';
        cell += marker; i = end; continue;
      }
      if (char === '|' && !codeMarker) { cells.push(cell.trim()); cell = ''; }
      else cell += char;
      i++;
    }
    cells.push(cell.trim());
    if (value.charAt(0) === '|') cells.shift();
    if (value.charAt(value.length - 1) === '|' && !codeMarker && /(^|[^\\])\|$/.test(value)) cells.pop();
    return cells.length > 1 ? cells : null;
  }

  function tableHeader(lines, index) {
    if (index + 1 >= lines.length) return null;
    var cells = tableCells(lines[index]), rule = tableCells(lines[index + 1]);
    if (!cells || !rule || cells.length !== rule.length || !rule.every(function (cell) {
      return /^:?-{3,}:?$/.test(cell);
    })) return null;
    return { cells: cells, alignments: rule.map(function (cell) {
      return cell.charAt(cell.length - 1) === ':' ? (cell.charAt(0) === ':' ? 'center' : 'end') : 'start';
    }) };
  }
  function fence(line) { return /^\s{0,3}(`{3,}|~{3,})[^`~]*$/.exec(line); }
  function heading(line) {
    var match = /^\s{0,3}#{1,6}\s+(.+?)\s*$/.exec(line);
    if (match) match[1] = match[1].replace(/\s+#+\s*$/, '');
    return match;
  }

  function render(value) {
    var lines = String(value == null ? '' : value).replace(/\r\n?/g, '\n').split('\n');
    var result = [], i = 0;
    function list(start, depth) {
      var first = listMarker(lines[start]), index = start, items = [], tag = first.ordered ? 'ol' : 'ul';
      while (index < lines.length) {
        var item = listMarker(lines[index]);
        if (!item || item.indent !== first.indent || item.ordered !== first.ordered) break;
        var content = item.text, children = '';
        index++;
        while (index < lines.length && lines[index].trim()) {
          var next = listMarker(lines[index]);
          if (next) {
            if (next.indent > first.indent && depth < 12) {
              var nested = list(index, depth + 1); children += nested.html; index = nested.next; continue;
            }
            break;
          }
          var indentation = (/^\s*/.exec(lines[index])[0] || '').replace(/\t/g, '    ').length;
          if (indentation <= first.indent || fence(lines[index]) || heading(lines[index])) break;
          if (children) children += '<br>' + inline(lines[index].trim());
          else content += '\n' + lines[index].trim();
          index++;
        }
        items.push('<li>' + inline(content) + children + '</li>');
        var afterBlank = index;
        while (afterBlank < lines.length && !lines[afterBlank].trim()) afterBlank++;
        var following = afterBlank < lines.length ? listMarker(lines[afterBlank]) : null;
        if (following && following.indent === first.indent && following.ordered === first.ordered) index = afterBlank;
      }
      var startAttribute = first.ordered && first.number !== 1 ? ' start="' + first.number + '"' : '';
      return { html: '<' + tag + startAttribute + '>' + items.join('') + '</' + tag + '>', next: index };
    }
    while (i < lines.length) {
      if (!lines[i].trim()) { i++; continue; }
      var opening = fence(lines[i]);
      if (opening) {
        var code = [], marker = opening[1], closing = new RegExp('^\\s{0,3}' + marker.charAt(0) + '{' + marker.length + ',}\\s*$');
        i++;
        while (i < lines.length && !closing.test(lines[i])) { code.push(lines[i]); i++; }
        if (i < lines.length) i++;
        result.push('<pre class="ac-code"><code>' + escapeHtml(code.join('\n')) + '</code></pre>'); continue;
      }
      var title = heading(lines[i]);
      if (title) { result.push('<h3 class="ac-text-heading">' + inline(title[1]) + '</h3>'); i++; continue; }
      var header = tableHeader(lines, i);
      if (header) {
        var rows = [], head = header.cells.map(function (cell, column) {
          return '<th scope="col" class="ac-align-' + header.alignments[column] + '">' + inline(cell) + '</th>';
        }).join('');
        i += 2;
        while (i < lines.length && lines[i].trim()) {
          var cells = tableCells(lines[i]);
          if (!cells || cells.length !== header.cells.length) break;
          rows.push('<tr>' + cells.map(function (cell, column) {
            return '<td class="ac-align-' + header.alignments[column] + '">' + inline(cell) + '</td>';
          }).join('') + '</tr>'); i++;
        }
        result.push('<div class="ac-table-wrap" tabindex="0" role="region" aria-label="Таблица в ответе"><table><thead><tr>' +
          head + '</tr></thead><tbody>' + rows.join('') + '</tbody></table></div>'); continue;
      }
      if (listMarker(lines[i])) { var group = list(i, 0); result.push(group.html); i = group.next; continue; }
      var paragraph = [lines[i++]];
      while (i < lines.length && lines[i].trim() && !fence(lines[i]) && !heading(lines[i]) &&
             !listMarker(lines[i]) && !tableHeader(lines, i)) paragraph.push(lines[i++]);
      result.push('<p>' + inline(paragraph.join('\n')) + '</p>');
    }
    return result.join('');
  }
  root.NeuronaChatFormat = Object.freeze({ render: render });
})(typeof window !== 'undefined' ? window : globalThis);
