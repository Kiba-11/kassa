// Выгрузка в Excel без внешних библиотек: XLSX — это ZIP-архив с XML-листами.
// Структура книги совпадает с компьютерной версией (app/app.py → build_xlsx).
(function (global) {
  'use strict';

  const enc = new TextEncoder();
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const NSR = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';
  const XML_HEAD = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\n';

  // ─── ZIP (без сжатия) ─────────────────────────────────────────────────────
  const CRC = (() => {
    const t = new Uint32Array(256);
    for (let n = 0; n < 256; n++) {
      let c = n;
      for (let k = 0; k < 8; k++) c = c & 1 ? 0xEDB88320 ^ (c >>> 1) : c >>> 1;
      t[n] = c >>> 0;
    }
    return t;
  })();
  const crc32 = b => {
    let c = 0xFFFFFFFF;
    for (let i = 0; i < b.length; i++) c = CRC[(c ^ b[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  };

  function zip(files) {
    const now = new Date();
    const time = (now.getHours() << 11) | (now.getMinutes() << 5) | (now.getSeconds() >> 1);
    const date = ((now.getFullYear() - 1980) << 9) | ((now.getMonth() + 1) << 5) | now.getDate();
    const local = [], central = [];
    let offset = 0;
    for (const f of files) {
      const name = enc.encode(f.name), data = enc.encode(f.text), crc = crc32(data);
      const lh = new DataView(new ArrayBuffer(30));
      lh.setUint32(0, 0x04034b50, true);
      lh.setUint16(4, 20, true);
      lh.setUint16(6, 0x0800, true);            // имена в UTF-8
      lh.setUint16(10, time, true);
      lh.setUint16(12, date, true);
      lh.setUint32(14, crc, true);
      lh.setUint32(18, data.length, true);
      lh.setUint32(22, data.length, true);
      lh.setUint16(26, name.length, true);
      local.push(new Uint8Array(lh.buffer), name, data);

      const ch = new DataView(new ArrayBuffer(46));
      ch.setUint32(0, 0x02014b50, true);
      ch.setUint16(4, 20, true);
      ch.setUint16(6, 20, true);
      ch.setUint16(8, 0x0800, true);
      ch.setUint16(12, time, true);
      ch.setUint16(14, date, true);
      ch.setUint32(16, crc, true);
      ch.setUint32(20, data.length, true);
      ch.setUint32(24, data.length, true);
      ch.setUint16(28, name.length, true);
      ch.setUint32(42, offset, true);
      central.push(new Uint8Array(ch.buffer), name);
      offset += 30 + name.length + data.length;
    }
    const cdSize = central.reduce((a, b) => a + b.length, 0);
    const end = new DataView(new ArrayBuffer(22));
    end.setUint32(0, 0x06054b50, true);
    end.setUint16(8, files.length, true);
    end.setUint16(10, files.length, true);
    end.setUint32(12, cdSize, true);
    end.setUint32(16, offset, true);
    const parts = [...local, ...central, new Uint8Array(end.buffer)];
    const out = new Uint8Array(parts.reduce((a, b) => a + b.length, 0));
    let p = 0;
    for (const part of parts) { out.set(part, p); p += part.length; }
    return out;
  }

  // ─── Листы ────────────────────────────────────────────────────────────────
  const x = s => String(s).replace(/[\u0000-\u0008\u000B\u000C\u000E-\u001F]/g, '')
    .replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;');
  const col = c => { let s = ''; while (c > 0) { const m = (c - 1) % 26; s = String.fromCharCode(65 + m) + s; c = Math.floor((c - 1) / 26); } return s; };
  const serial = iso => { const [y, m, d] = iso.split('-').map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };

  // Стили (индексы cellXfs в styles.xml)
  const ST = { head: 1, text: 2, money: 3, date: 4, center: 5, group: 6, boldText: 7, boldMoney: 8, title: 9, subtitle: 10, note: 11 };

  class Sheet {
    constructor(name, widths) {
      Object.assign(this, { name, widths, rows: new Map(), heights: {}, merges: [], freeze: 0, printTitles: '' });
    }
    put(r, c, cell) { (this.rows.get(r) || this.rows.set(r, new Map()).get(r)).set(c, cell); }
    val(r, c, v, s = 0) { this.put(r, c, { v, s }); }
    fx(r, c, f, s = 0) { this.put(r, c, { f, s }); }
    blank(r, c, s) { this.put(r, c, { s }); }
    merge(r1, c1, r2, c2) { this.merges.push(`${col(c1)}${r1}:${col(c2)}${r2}`); }
    xml() {
      const cellXml = (r, c, { v, f, s }) => {
        const ref = `${col(c)}${r}`, st = s ? ` s="${s}"` : '';
        if (f) return `<c r="${ref}"${st}><f>${x(f)}</f></c>`;
        if (typeof v === 'number') return `<c r="${ref}"${st}><v>${v}</v></c>`;
        if (v == null || v === '') return `<c r="${ref}"${st}/>`;
        return `<c r="${ref}"${st} t="inlineStr"><is><t xml:space="preserve">${x(v)}</t></is></c>`;
      };
      const rows = [...this.rows.keys()].sort((a, b) => a - b).map(r => {
        const cells = [...this.rows.get(r)].sort((a, b) => a[0] - b[0]).map(([c, cell]) => cellXml(r, c, cell)).join('');
        const h = this.heights[r];
        return `<row r="${r}"${h ? ` ht="${h}" customHeight="1"` : ''}>${cells}</row>`;
      }).join('');
      const cols = this.widths.map((w, i) => `<col min="${i + 1}" max="${i + 1}" width="${w}" customWidth="1"/>`).join('');
      const pane = this.freeze ? `<pane ySplit="${this.freeze}" topLeftCell="A${this.freeze + 1}" activePane="bottomLeft" state="frozen"/>` : '';
      const merges = this.merges.length ? `<mergeCells count="${this.merges.length}">${this.merges.map(m => `<mergeCell ref="${m}"/>`).join('')}</mergeCells>` : '';
      return XML_HEAD + `<worksheet xmlns="${NS}" xmlns:r="${NSR}"><sheetPr><pageSetUpPr fitToPage="1"/></sheetPr>`
        + `<sheetViews><sheetView workbookViewId="0">${pane}</sheetView></sheetViews><sheetFormatPr defaultRowHeight="15"/>`
        + `<cols>${cols}</cols><sheetData>${rows}</sheetData>${merges}`
        + '<pageMargins left="0.5" right="0.5" top="0.75" bottom="0.75" header="0.3" footer="0.3"/>'
        + '<pageSetup paperSize="9" orientation="portrait" fitToWidth="1" fitToHeight="0"/></worksheet>';
    }
  }

  const STYLES = XML_HEAD + `<styleSheet xmlns="${NS}">`
    + '<numFmts count="2"><numFmt numFmtId="164" formatCode="#,##0.00\\ &quot;₽&quot;"/><numFmt numFmtId="165" formatCode="dd\\.mm\\.yyyy"/></numFmts>'
    + '<fonts count="4"><font><sz val="11"/><name val="Calibri"/></font><font><b/><sz val="11"/><name val="Calibri"/></font>'
    + '<font><b/><sz val="14"/><name val="Calibri"/></font><font><i/><sz val="11"/><color rgb="FF666666"/><name val="Calibri"/></font></fonts>'
    + '<fills count="4"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDF0"/><bgColor indexed="64"/></patternFill></fill>'
    + '<fill><patternFill patternType="solid"><fgColor rgb="FFF5F5F7"/><bgColor indexed="64"/></patternFill></fill></fills>'
    + '<borders count="2"><border><left/><right/><top/><bottom/><diagonal/></border>'
    + '<border>' + ['left', 'right', 'top', 'bottom'].map(s => `<${s} style="thin"><color rgb="FFA0A0A0"/></${s}>`).join('') + '<diagonal/></border></borders>'
    + '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>'
    + '<cellXfs count="12">'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>'
    + '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1"/>'
    + '<xf numFmtId="164" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1"/>'
    + '<xf numFmtId="165" fontId="0" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
    + '<xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" applyBorder="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
    + '<xf numFmtId="0" fontId="1" fillId="3" borderId="1" xfId="0" applyFont="1" applyFill="1" applyBorder="1" applyAlignment="1"><alignment horizontal="center" vertical="center"/></xf>'
    + '<xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" applyFont="1" applyBorder="1"/>'
    + '<xf numFmtId="164" fontId="1" fillId="0" borderId="1" xfId="0" applyNumberFormat="1" applyFont="1" applyBorder="1"/>'
    + '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center"/></xf>'
    + '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1" applyAlignment="1"><alignment horizontal="center" vertical="center" wrapText="1"/></xf>'
    + '<xf numFmtId="0" fontId="3" fillId="0" borderId="0" xfId="0" applyFont="1"/>'
    + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';

  function header(sheet, r, titles) {
    titles.forEach((t, i) => sheet.val(r, i + 1, t, ST.head));
  }

  function build(state) {
    const s = state.settings;
    const byDateAsc = (a, b) => a.date.localeCompare(b.date) || a.id - b.id;
    const paid = new Map(), last = new Map();
    for (const i of state.incomes) {
      if (!i.member_id) continue;
      paid.set(i.member_id, Math.round(((paid.get(i.member_id) || 0) + i.amount) * 100) / 100);
      if ((last.get(i.member_id) || '') < i.date) last.set(i.member_id, i.date);
    }
    const groups = new Map();
    for (const m of [...state.members].sort((a, b) => a.sort - b.sort || a.id - b.id)) {
      if (m.active || paid.get(m.id)) (groups.get(m.grp) || groups.set(m.grp, []).get(m.grp)).push(m);
    }

    // Ведомость
    const v = new Sheet('Ведомость', [7, 30, 13, 13, 13, 12, 16]);
    [['Ведомость', ST.title], [`на развитие: ${s.org}`, ST.subtitle], [`${s.year} учебный год`, ST.subtitle]].forEach(([t, st], i) => {
      v.merge(i + 1, 1, i + 1, 7);
      v.val(i + 1, 1, t, st);
    });
    v.heights[2] = 32;
    header(v, 5, ['№ п/п', 'ФИО', 'К сдаче', 'Сдано', 'Долг', 'Дата сдачи', 'Подпись']);
    let r = 6;
    const subtotals = [];
    for (const [grp, members] of groups) {
      v.merge(r, 1, r, 7);
      v.val(r, 1, grp, ST.group);
      for (let c = 2; c <= 7; c++) v.blank(r, c, ST.group);
      r++;
      const first = r;
      members.forEach((m, n) => {
        v.val(r, 1, n + 1, ST.center);
        v.val(r, 2, m.name, ST.text);
        v.val(r, 3, m.fee, ST.money);
        paid.get(m.id) ? v.val(r, 4, paid.get(m.id), ST.money) : v.blank(r, 4, ST.money);
        v.fx(r, 5, `MAX(C${r}-D${r},0)`, ST.money);
        last.get(m.id) ? v.val(r, 6, serial(last.get(m.id)), ST.date) : v.blank(r, 6, ST.date);
        v.blank(r, 7, ST.text);
        v.heights[r] = 20;
        r++;
      });
      v.blank(r, 1, ST.boldText);
      v.val(r, 2, `Итого: ${grp.toLowerCase()}`, ST.boldText);
      ['C', 'D', 'E'].forEach((L, i) => v.fx(r, 3 + i, `SUM(${L}${first}:${L}${r - 1})`, ST.boldMoney));
      v.blank(r, 6, ST.boldText);
      v.blank(r, 7, ST.boldText);
      subtotals.push(r);
      r++;
    }
    v.blank(r, 1, ST.boldText);
    v.val(r, 2, 'ИТОГО', ST.boldText);
    ['C', 'D', 'E'].forEach((L, i) => v.fx(r, 3 + i, subtotals.length ? `SUM(${subtotals.map(n => L + n).join(',')})` : '0', ST.boldMoney));
    v.blank(r, 6, ST.boldText);
    v.blank(r, 7, ST.boldText);
    const totalRow = r;
    v.freeze = 5;
    v.printTitles = "'Ведомость'!$5:$5";

    // Приход
    const names = new Map(state.members.map(m => [m.id, m.name]));
    const inc = new Sheet('Приход', [12, 32, 14, 13, 44]);
    header(inc, 1, ['Дата', 'От кого', 'Сумма', 'Способ', 'Комментарий']);
    r = 2;
    for (const i of [...state.incomes].sort(byDateAsc)) {
      inc.val(r, 1, serial(i.date), ST.date);
      inc.val(r, 2, names.get(i.member_id) || i.source, ST.text);
      inc.val(r, 3, i.amount, ST.money);
      inc.val(r, 4, i.method, ST.text);
      inc.val(r, 5, i.comment, ST.text);
      r++;
    }
    inc.val(r, 2, 'Итого', ST.boldText);
    inc.fx(r, 3, `SUM(C2:C${Math.max(r - 1, 2)})`, ST.boldMoney);
    const incomeTotal = `'Приход'!C${r}`;
    inc.freeze = 1;
    inc.printTitles = "'Приход'!$1:$1";

    // Расход
    const exp = new Sheet('Расход', [12, 34, 24, 14, 7, 18, 34]);
    header(exp, 1, ['Дата', 'Наименование', 'Категория', 'Сумма', 'Чек', 'Кто купил', 'Комментарий']);
    r = 2;
    for (const e of [...state.expenses].sort(byDateAsc)) {
      exp.val(r, 1, serial(e.date), ST.date);
      exp.val(r, 2, e.title, ST.text);
      exp.val(r, 3, e.category, ST.text);
      exp.val(r, 4, e.amount, ST.money);
      exp.val(r, 5, e.receipt ? 'да' : 'нет', ST.center);
      exp.val(r, 6, e.buyer, ST.text);
      exp.val(r, 7, e.comment, ST.text);
      r++;
    }
    exp.val(r, 2, 'Итого', ST.boldText);
    exp.fx(r, 4, `SUM(D2:D${Math.max(r - 1, 2)})`, ST.boldMoney);
    const expenseTotal = `'Расход'!D${r}`;
    exp.freeze = 1;
    exp.printTitles = "'Расход'!$1:$1";

    // Сводка
    const sum = new Sheet('Сводка', [40, 18]);
    sum.val(1, 1, `Сводка за ${s.year} учебный год`, ST.title);
    const rows = [
      ['Остаток на начало года', { v: s.opening_balance }],
      ['Поступило', { f: incomeTotal }],
      ['Израсходовано', { f: expenseTotal }],
      ['Остаток в кассе', { f: 'B3+B4-B5' }],
      ['Ожидается к сбору (долг)', { f: `'Ведомость'!E${totalRow}` }],
    ];
    rows.forEach(([label, cell], i) => {
      const bold = i === 3;
      sum.val(3 + i, 1, label, bold ? ST.boldText : ST.text);
      sum.put(3 + i, 2, { ...cell, s: bold ? ST.boldMoney : ST.money });
    });
    if (s.opening_note) sum.val(9, 1, `Примечание: ${s.opening_note}`, ST.note);
    const now = new Date();
    const pad = n => String(n).padStart(2, '0');
    sum.val(10, 1, `Сформировано ${pad(now.getDate())}.${pad(now.getMonth() + 1)}.${now.getFullYear()} ${pad(now.getHours())}:${pad(now.getMinutes())}`, ST.note);

    const sheets = [sum, v, inc, exp];
    const defined = sheets.map((sh, i) => sh.printTitles ? `<definedName name="_xlnm.Print_Titles" localSheetId="${i}">${x(sh.printTitles)}</definedName>` : '').join('');
    const files = [
      {
        name: '[Content_Types].xml',
        text: XML_HEAD + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">'
          + '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>'
          + '<Default Extension="xml" ContentType="application/xml"/>'
          + '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>'
          + '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>'
          + sheets.map((_, i) => `<Override PartName="/xl/worksheets/sheet${i + 1}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>`).join('')
          + '</Types>',
      },
      {
        name: '_rels/.rels',
        text: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/></Relationships>',
      },
      {
        name: 'xl/workbook.xml',
        text: XML_HEAD + `<workbook xmlns="${NS}" xmlns:r="${NSR}"><sheets>`
          + sheets.map((sh, i) => `<sheet name="${x(sh.name)}" sheetId="${i + 1}" r:id="rId${i + 1}"/>`).join('')
          + `</sheets><definedNames>${defined}</definedNames><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>`,
      },
      {
        name: 'xl/_rels/workbook.xml.rels',
        text: XML_HEAD + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          + sheets.map((_, i) => `<Relationship Id="rId${i + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet${i + 1}.xml"/>`).join('')
          + `<Relationship Id="rId${sheets.length + 1}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/></Relationships>`,
      },
      { name: 'xl/styles.xml', text: STYLES },
      ...sheets.map((sh, i) => ({ name: `xl/worksheets/sheet${i + 1}.xml`, text: sh.xml() })),
    ];
    return zip(files);
  }

  global.KassaXlsx = { build };
})(window);
