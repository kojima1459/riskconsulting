(function (root) {
    'use strict';
    function escapeHtml(value) {
        return String(value === null || typeof value === 'undefined' ? '' : value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&#39;');
    }
    function safeLink(url) {
        return /^(https?:\/\/|mailto:)/i.test(url) && !/[\x00-\x20\x7f]/.test(url);
    }
    function inline(text, depth) {
        var out = '', i = 0, end, mark, n, inner, match, target, alt;
        text = String(text); depth = depth || 0;
        if (depth > 8) { return escapeHtml(text); }
        while (i < text.length) {
            if (text.charAt(i) === '\\' && /[\\`*_{}\[\]()#+.!>|~-]/.test(text.charAt(i + 1))) { out += escapeHtml(text.charAt(i + 1)); i += 2; continue; }
            if (text.charAt(i) === '`') {
                match = /^`+/.exec(text.substring(i)); mark = match[0]; end = text.indexOf(mark, i + mark.length);
                if (end >= 0) { out += '<code>' + escapeHtml(text.substring(i + mark.length, end).replace(/\n/g, ' ')) + '</code>'; i = end + mark.length; continue; }
            }
            if (text.substring(i, i + 2) === '![' || text.charAt(i) === '[') {
                n = text.charAt(i) === '!' ? 2 : 1; end = text.indexOf('](', i + n);
                if (end >= 0) {
                    var finish = end + 2, level = 1;
                    while (finish < text.length && level) { if (text.charAt(finish) === '(') { level++; } if (text.charAt(finish) === ')') { level--; } if (level) { finish++; } }
                    if (!level) {
                        target = text.substring(end + 2, finish).replace(/^\s+|\s+$/g, '').replace(/\s+["'][\s\S]*["']$/, '');
                        if (target.charAt(0) === '<' && target.charAt(target.length - 1) === '>') { target = target.substring(1, target.length - 1); }
                        alt = text.substring(i + n, end);
                        if (n === 2) { out += '<span class="image-reference">[画像: ' + escapeHtml(alt) + ' / ' + escapeHtml(target) + ']</span>'; }
                        else if (safeLink(target)) { out += '<a href="#" data-link="' + escapeHtml(target) + '">' + inline(alt, depth + 1) + '</a>'; }
                        else { out += inline(alt, depth + 1) + ' <span class="image-reference">(' + escapeHtml(target) + ')</span>'; }
                        i = finish + 1; continue;
                    }
                }
            }
            mark = text.substring(i, i + 2);
            if (mark === '**' || mark === '__' || mark === '~~') {
                end = text.indexOf(mark, i + 2);
                if (end > i + 2) { inner = inline(text.substring(i + 2, end), depth + 1); out += mark === '~~' ? '<del>' + inner + '</del>' : '<strong>' + inner + '</strong>'; i = end + 2; continue; }
            }
            mark = text.charAt(i);
            if ((mark === '*' || mark === '_') && !(mark === '_' && /\w/.test(text.charAt(i - 1)))) {
                end = text.indexOf(mark, i + 1);
                if (end > i + 1) { out += '<em>' + inline(text.substring(i + 1, end), depth + 1) + '</em>'; i = end + 1; continue; }
            }
            out += text.charAt(i) === '\n' ? '<br>' : escapeHtml(text.charAt(i)); i++;
        }
        return out;
    }
    function tableCells(line) {
        var result = [], current = '', escaped = false, code = false, i, ch;
        line = line.replace(/^\s*\|/, '').replace(/\|\s*$/, '');
        for (i = 0; i < line.length; i++) {
            ch = line.charAt(i);
            if (escaped) { current += '\\' + ch; escaped = false; }
            else if (ch === '\\') { escaped = true; }
            else if (ch === '`') { code = !code; current += ch; }
            else if (ch === '|' && !code) { result.push(current.replace(/^\s+|\s+$/g, '')); current = ''; }
            else { current += ch; }
        }
        if (escaped) { current += '\\'; }
        result.push(current.replace(/^\s+|\s+$/g, '')); return result;
    }
    function tableDivider(line) {
        var cells = tableCells(line), i;
        if (line.indexOf('|') < 0 || cells.length < 1) { return false; }
        for (i = 0; i < cells.length; i++) { if (!/^:?-{3,}:?$/.test(cells[i])) { return false; } }
        return true;
    }
    function listMatch(line) { return /^( *)([-+*]|\d+[.)])\s+(.+|)$/.exec(line); }
    function isBlock(lines, i) {
        return /^( {0,3}#{1,6}\s| {0,3}(?:`{3,}|~{3,})| {0,3}>| {0,3}(?:[-*_]\s*){3,}$)/.test(lines[i]) || !!listMatch(lines[i]) || (i + 1 < lines.length && tableDivider(lines[i + 1]));
    }
    function blocks(lines, depth) {
        var result = '', i = 0, line, match, group, j, fence, lang, cells, aligns, k, type, baseIndent, item, next, itemText, start;
        depth = depth || 0;
        if (depth > 20) { return '<p>' + escapeHtml(lines.join('\n')) + '</p>'; }
        while (i < lines.length) {
            line = lines[i];
            if (!/\S/.test(line)) { i++; continue; }
            match = /^ {0,3}(`{3,}|~{3,})(.*)$/.exec(line);
            if (match) {
                fence = match[1]; lang = match[2].replace(/^\s+|\s+$/g, ''); group = []; i++;
                while (i < lines.length && !(new RegExp('^ {0,3}' + fence.charAt(0) + '{' + fence.length + ',}\\s*$')).test(lines[i])) { group.push(lines[i++]); }
                if (i < lines.length) { i++; }
                result += '<pre><code' + (lang ? ' data-language="' + escapeHtml(lang) + '"' : '') + '>' + escapeHtml(group.join('\n')) + '</code></pre>'; continue;
            }
            match = /^ {0,3}(#{1,6})[ \t]+(.*)$/.exec(line);
            if (match) { result += '<h' + match[1].length + '>' + inline(match[2].replace(/[ \t]+#+[ \t]*$/, '').replace(/[ \t]+$/, '')) + '</h' + match[1].length + '>'; i++; continue; }
            if (i + 1 < lines.length && /^\s*(=+|-{3,})\s*$/.test(lines[i + 1]) && !listMatch(line)) {
                type = lines[i + 1].indexOf('=') >= 0 ? 'h1' : 'h2'; result += '<' + type + '>' + inline(line) + '</' + type + '>'; i += 2; continue;
            }
            if (/^ {0,3}(?:(?:-\s*){3,}|(?:\*\s*){3,}|(?:_\s*){3,})$/.test(line)) { result += '<hr>'; i++; continue; }
            if (/^ {0,3}>/.test(line)) {
                group = []; while (i < lines.length && /^ {0,3}>/.test(lines[i])) { group.push(lines[i++].replace(/^ {0,3}> ?/, '')); }
                result += '<blockquote>' + blocks(group, depth + 1) + '</blockquote>'; continue;
            }
            if (i + 1 < lines.length && line.indexOf('|') >= 0 && tableDivider(lines[i + 1])) {
                cells = tableCells(line); aligns = tableCells(lines[i + 1]); result += '<div class="table-wrap"><table><thead><tr>';
                for (j = 0; j < cells.length; j++) { result += '<th>' + inline(cells[j]) + '</th>'; }
                result += '</tr></thead><tbody>'; i += 2;
                while (i < lines.length && /\S/.test(lines[i]) && lines[i].indexOf('|') >= 0) {
                    group = tableCells(lines[i++]); result += '<tr>';
                    for (j = 0; j < cells.length; j++) { k = aligns[j] || ''; type = /^:.*:$/.test(k) ? 'center' : /:$/.test(k) ? 'right' : 'left'; result += '<td style="text-align:' + type + '">' + inline(group[j] || '') + '</td>'; }
                    result += '</tr>';
                }
                result += '</tbody></table></div>'; continue;
            }
            match = listMatch(line);
            if (match) {
                baseIndent = match[1].length; type = /\d/.test(match[2].charAt(0)) ? 'ol' : 'ul'; start = type === 'ol' ? parseInt(match[2], 10) : 1;
                result += '<' + type + (type === 'ol' && start !== 1 ? ' start="' + start + '"' : '') + '>';
                while (i < lines.length) {
                    match = listMatch(lines[i]);
                    if (!match || match[1].length !== baseIndent || (/\d/.test(match[2].charAt(0)) ? 'ol' : 'ul') !== type) { break; }
                    item = [match[3]]; i++;
                    while (i < lines.length) {
                        next = listMatch(lines[i]);
                        if (next && next[1].length <= baseIndent) { break; }
                        if (/\S/.test(lines[i]) && lines[i].match(/^ */)[0].length <= baseIndent) { break; }
                        if (!/\S/.test(lines[i])) { if (i + 1 >= lines.length || lines[i + 1].match(/^ */)[0].length <= baseIndent) { break; } }
                        item.push(lines[i].substring(Math.min(baseIndent + 2, lines[i].match(/^ */)[0].length))); i++;
                    }
                    match = /^\[([ xX])\]\s*(.*)$/.exec(item[0]);
                    if (match) { item[0] = match[2]; itemText = blocks(item, depth + 1); itemText = itemText.replace(/^<p>/, '<p><span class="task-box' + (match[1] !== ' ' ? ' task-done' : '') + '" aria-label="' + (match[1] !== ' ' ? '完了' : '未完了') + '">' + (match[1] !== ' ' ? '&#10003;' : '') + '</span>'); }
                    else { itemText = blocks(item, depth + 1); }
                    result += '<li>' + itemText + '</li>';
                }
                result += '</' + type + '>'; continue;
            }
            group = [line]; i++;
            while (i < lines.length && /\S/.test(lines[i]) && !isBlock(lines, i)) { group.push(lines[i++]); }
            result += '<p>' + inline(group.join('\n')) + '</p>';
        }
        return result;
    }
    function render(text) { return blocks(String(text || '').replace(/^\ufeff/, '').replace(/\r\n?/g, '\n').replace(/\t/g, '    ').split('\n')); }
    function diff(previous, current) {
        var a = String(previous || '').replace(/\r\n?/g, '\n').split('\n'), b = String(current || '').replace(/\r\n?/g, '\n').split('\n'), rows = [], i, j, matrix, limit = 500000, prefix = 0, suffix = 0;
        while (prefix < a.length && prefix < b.length && a[prefix] === b[prefix]) { prefix++; }
        while (suffix < a.length - prefix && suffix < b.length - prefix && a[a.length - 1 - suffix] === b[b.length - 1 - suffix]) { suffix++; }
        for (i = 0; i < prefix; i++) { rows.push({type:'same',text:a[i]}); }
        var aa = a.slice(prefix, a.length - suffix), bb = b.slice(prefix, b.length - suffix);
        if (aa.length * bb.length <= limit) {
            matrix = [];
            for (i = 0; i <= aa.length; i++) { matrix[i] = []; for (j = 0; j <= bb.length; j++) { matrix[i][j] = 0; } }
            for (i = aa.length - 1; i >= 0; i--) { for (j = bb.length - 1; j >= 0; j--) { matrix[i][j] = aa[i] === bb[j] ? matrix[i + 1][j + 1] + 1 : Math.max(matrix[i + 1][j], matrix[i][j + 1]); } }
            i = 0; j = 0;
            while (i < aa.length || j < bb.length) {
                if (i < aa.length && j < bb.length && aa[i] === bb[j]) { rows.push({type:'same',text:aa[i++]}); j++; }
                else if (i < aa.length && (j >= bb.length || matrix[i + 1][j] >= matrix[i][j + 1])) { rows.push({type:'remove',text:aa[i++]}); }
                else { rows.push({type:'add',text:bb[j++]}); }
            }
        } else {
            for (i = 0; i < aa.length; i++) { rows.push({type:'remove',text:aa[i]}); }
            for (i = 0; i < bb.length; i++) { rows.push({type:'add',text:bb[i]}); }
        }
        for (i = a.length - suffix; i < a.length; i++) { rows.push({type:'same',text:a[i]}); }
        return rows;
    }
    function renderDiff(previous, current) {
        var rows = diff(previous, current), html = '<div class="diff-intro">緑：追加　赤：削除</div><div class="diff-lines">', i;
        for (i = 0; i < rows.length; i++) { html += '<div class="diff-line diff-' + rows[i].type + '"><span class="diff-prefix">' + (rows[i].type === 'add' ? '+' : rows[i].type === 'remove' ? '−' : ' ') + '</span><pre>' + escapeHtml(rows[i].text) + '</pre></div>'; }
        return html + '</div>';
    }
    root.reviewMarkdown = {render:render,inline:inline,escapeHtml:escapeHtml,safeLink:safeLink,diff:diff,renderDiff:renderDiff};
    if (typeof module !== 'undefined' && module.exports) { module.exports = root.reviewMarkdown; }
}(typeof window !== 'undefined' ? window : this));
