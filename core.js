/* Browser-local Excel I/O. No formula execution and no remote requests. */
(function () {
  'use strict';

  const MAX_BYTES = 15 * 1024 * 1024;
  const MAX_ROWS = 30000;
  const MAX_COLS = 100;
  const MAX_SHEETS = 40;
  const XML_LIMIT = 40 * 1024 * 1024;
  const MIME = 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet';
  const NS = 'http://schemas.openxmlformats.org/spreadsheetml/2006/main';
  const REL_NS = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships';

  function zipLibrary() {
    if (!window.JSZip) throw new Error('Excel 组件未加载，请刷新页面后重试。');
    return window.JSZip;
  }

  function cleanText(value) {
    return String(value == null ? '' : value).replace(/[\u0000-\u0008\u000b\u000c\u000e-\u001f\ufffe\uffff]/g, '');
  }

  function xmlEscape(value) {
    return cleanText(value).replace(/&/g, '&amp;').replace(/</g, '&lt;').replace(/>/g, '&gt;').replace(/"/g, '&quot;').replace(/'/g, '&apos;');
  }

  function parseXml(text, label) {
    if (text.length > XML_LIMIT) throw new Error('表格解压后的内容过大，请拆分为较小文件。');
    if (/<!DOCTYPE|<!ENTITY/i.test(text)) throw new Error('文件包含不支持的 XML 声明，请另存为标准 .xlsx 后导入。');
    const doc = new DOMParser().parseFromString(text, 'application/xml');
    if (doc.getElementsByTagName('parsererror').length || doc.getElementsByTagNameNS('*', 'parsererror').length) {
      throw new Error((label || 'Excel') + ' XML 数据损坏，请用 Excel 重新保存后导入。');
    }
    return doc;
  }

  const elements = (node, name) => Array.from(node.getElementsByTagNameNS('*', name));
  const directChildren = (node, name) => Array.from(node.children || []).filter(child => child.localName === name);
  const textOf = (node, name) => { const found = elements(node, name)[0]; return found ? found.textContent || '' : ''; };

  function normalizePath(base, target) {
    if (!target || /^[a-z]+:/i.test(target)) return null;
    const parts = (target.startsWith('/') ? target.slice(1) : base.slice(0, base.lastIndexOf('/') + 1) + target).replace(/\\/g, '/').split('/');
    const result = [];
    for (const part of parts) {
      if (!part || part === '.') continue;
      if (part === '..') { if (!result.length) return null; result.pop(); }
      else result.push(part);
    }
    return result.join('/');
  }

  function columnIndex(reference) {
    const letters = String(reference || '').match(/^\$?([A-Za-z]+)/);
    if (!letters) return -1;
    let index = 0;
    for (const letter of letters[1].toUpperCase()) index = index * 26 + letter.charCodeAt(0) - 64;
    return index - 1;
  }

  function columnName(index) {
    let n = index + 1;
    let result = '';
    while (n > 0) { const rem = (n - 1) % 26; result = String.fromCharCode(65 + rem) + result; n = Math.floor((n - 1) / 26); }
    return result;
  }

  async function readPart(zip, name, required) {
    const entry = name && zip.file(name);
    if (!entry) {
      if (required) throw new Error('Excel 缺少必要的数据文件，请重新保存为 .xlsx 后导入。');
      return null;
    }
    const data = await entry.async('string');
    if (data.length > XML_LIMIT) throw new Error('表格解压后的内容过大，请拆分后导入。');
    return data;
  }

  function dateFormat(code) {
    const stripped = String(code || '').replace(/"[^"]*"/g, '').replace(/\\./g, '').replace(/\[(?!h\]|m\]|s\])[^\]]*\]/gi, '');
    return /[ymd]/i.test(stripped) && !/general|@/i.test(stripped);
  }

  function serialDate(raw, date1904, includeTime) {
    const n = Number(raw);
    if (!Number.isFinite(n) || n < 0 || n > 2958465) return raw;
    const base = Date.UTC(date1904 ? 1904 : 1899, date1904 ? 0 : 11, date1904 ? 1 : 30);
    const date = new Date(base + Math.round(n * 86400000));
    if (Number.isNaN(date.getTime())) return raw;
    const iso = date.toISOString();
    return includeTime && n % 1 ? iso.slice(0, 19).replace('T', ' ') : iso.slice(0, 10);
  }

  function getStyleFormats(stylesText) {
    if (!stylesText) return [];
    const doc = parseXml(stylesText, '样式');
    const custom = new Map(elements(doc, 'numFmt').map(n => [Number(n.getAttribute('numFmtId')), n.getAttribute('formatCode') || '']));
    const cellXfs = elements(doc, 'cellXfs')[0];
    if (!cellXfs) return [];
    const builtInDates = new Set([14, 15, 16, 17, 22, 27, 28, 29, 30, 31, 34, 35, 36, 50, 51, 52, 53, 54, 55, 56, 57, 58]);
    return directChildren(cellXfs, 'xf').map(xf => {
      const id = Number(xf.getAttribute('numFmtId') || 0);
      const code = custom.get(id) || '';
      return { isDate: builtInDates.has(id) || dateFormat(code), includeTime: id === 22 || /h|s/i.test(code.replace(/"[^"]*"/g, '')) };
    });
  }

  function parseSheet(text, sharedStrings, formats, date1904, warnings, name) {
    const doc = parseXml(text, '工作表“' + name + '”');
    const sheetData = elements(doc, 'sheetData')[0];
    const rows = [];
    let maxCol = 0;
    let formulaCount = 0;
    let errorCount = 0;
    let previousRow = -1;
    for (const rowNode of sheetData ? directChildren(sheetData, 'row') : []) {
      const rowAttr = Number(rowNode.getAttribute('r'));
      const rowIndex = rowAttr > 0 ? rowAttr - 1 : previousRow + 1;
      previousRow = rowIndex;
      const cells = directChildren(rowNode, 'c');
      // Excel often stores styled blank rows far below the data. Ignore them.
      const hasContent = cells.some(c => elements(c, 'v').length || elements(c, 'is').length || elements(c, 'f').length);
      if (!hasContent) continue;
      if (rowIndex >= MAX_ROWS) throw new Error('工作表“' + name + '”超出 30,000 行，请拆分后导入。');
      const row = [];
      let previousCol = -1;
      for (const cell of cells) {
        const reference = cell.getAttribute('r');
        const col = reference ? columnIndex(reference) : previousCol + 1;
        previousCol = col;
        if (col < 0) continue;
        const type = cell.getAttribute('t') || '';
        const raw = textOf(cell, 'v');
        const formula = elements(cell, 'f')[0];
        if (formula) formulaCount++;
        let value = raw;
        if (type === 's') value = sharedStrings[Number(raw)] == null ? '' : sharedStrings[Number(raw)];
        else if (type === 'inlineStr') value = elements(cell, 't').map(t => t.textContent || '').join('');
        else if (type === 'b') value = raw === '1' ? 'TRUE' : 'FALSE';
        else if (type === 'e') { value = raw; errorCount++; }
        else if (type === 'd') value = raw;
        else {
          const format = formats[Number(cell.getAttribute('s') || 0)];
          if (format && format.isDate && raw !== '') value = serialDate(raw, date1904, format.includeTime);
          // All remaining numeric cells stay as their raw XML text. Do not round IDs/phones.
        }
        if (col >= MAX_COLS) {
          if (value !== '') throw new Error('工作表“' + name + '”超出 100 列，请删除无关列后导入。');
          continue;
        }
        if (value !== '' || raw !== '') { row[col] = value; maxCol = Math.max(maxCol, col + 1); }
      }
      if (row.some(value => value !== '' && value != null)) rows[rowIndex] = row;
    }
    while (rows.length && !rows[rows.length - 1]) rows.pop();
    for (let i = 0; i < rows.length; i++) {
      const row = rows[i] || [];
      for (let c = 0; c < maxCol; c++) if (row[c] == null) row[c] = '';
      rows[i] = row;
    }
    const widths = [];
    for (const node of elements(doc, 'col')) {
      const min = Math.max(1, Number(node.getAttribute('min')) || 1);
      const max = Math.min(MAX_COLS, Number(node.getAttribute('max')) || min);
      const width = Number(node.getAttribute('width'));
      if (Number.isFinite(width)) for (let c = min - 1; c < max; c++) widths[c] = width;
    }
    if (formulaCount) warnings.push('“' + name + '”含 ' + formulaCount + ' 个公式，已读取文件中保存的计算结果；未执行公式。');
    if (errorCount) warnings.push('“' + name + '”含 ' + errorCount + ' 个 Excel 错误值，请核对后处理。');
    return { name, rows, widths, merges: elements(doc, 'mergeCell').map(n => n.getAttribute('ref') || '') };
  }

  async function readXlsx(file) {
    const JSZip = zipLibrary();
    const bytes = await file.arrayBuffer();
    let zip;
    try { zip = await JSZip.loadAsync(bytes); }
    catch (_) { throw new Error('无法打开此 Excel 文件，可能已损坏、加密或不是标准 .xlsx 格式。'); }
    if (Object.keys(zip.files).length > 5000) throw new Error('工作簿包含过多内部文件，请另存为简化的 .xlsx。');
    let workbookPath = 'xl/workbook.xml';
    const rootRelsText = await readPart(zip, '_rels/.rels', false);
    if (rootRelsText) {
      const rootRels = parseXml(rootRelsText, '工作簿关系');
      const officeDoc = elements(rootRels, 'Relationship').find(r => /\/officeDocument$/.test(r.getAttribute('Type') || ''));
      if (officeDoc && officeDoc.getAttribute('TargetMode') !== 'External') workbookPath = normalizePath('_root', officeDoc.getAttribute('Target')) || workbookPath;
    }
    const workbookText = await readPart(zip, workbookPath, true);
    const workbook = parseXml(workbookText, '工作簿');
    const base = workbookPath.slice(0, workbookPath.lastIndexOf('/') + 1);
    const fileName = workbookPath.slice(workbookPath.lastIndexOf('/') + 1);
    const relsText = await readPart(zip, base + '_rels/' + fileName + '.rels', true);
    const rels = elements(parseXml(relsText, '工作表关系'), 'Relationship');
    const relation = new Map(rels.filter(r => r.getAttribute('TargetMode') !== 'External').map(r => [r.getAttribute('Id'), normalizePath(workbookPath, r.getAttribute('Target'))]));
    const stringRelation = rels.find(r => /\/sharedStrings$/.test(r.getAttribute('Type') || ''));
    const styleRelation = rels.find(r => /\/styles$/.test(r.getAttribute('Type') || ''));
    const stringsText = await readPart(zip, stringRelation ? relation.get(stringRelation.getAttribute('Id')) : base + 'sharedStrings.xml', false);
    const sharedStrings = stringsText ? elements(parseXml(stringsText, '共享文本'), 'si').map(si => elements(si, 't').map(t => t.textContent || '').join('')) : [];
    const formats = getStyleFormats(await readPart(zip, styleRelation ? relation.get(styleRelation.getAttribute('Id')) : base + 'styles.xml', false));
    const properties = elements(workbook, 'workbookPr')[0];
    const date1904 = !!properties && /^(1|true)$/i.test(properties.getAttribute('date1904') || '');
    const sheetNodes = elements(workbook, 'sheet');
    if (sheetNodes.length > MAX_SHEETS) throw new Error('工作簿超过 40 个工作表，请拆分后导入。');
    const sheets = [];
    const warnings = [];
    let totalRows = 0;
    for (const node of sheetNodes) {
      const id = node.getAttributeNS(REL_NS, 'id') || node.getAttribute('r:id');
      const path = relation.get(id);
      const relationNode = rels.find(r => r.getAttribute('Id') === id);
      if (relationNode && !/\/worksheet$/.test(relationNode.getAttribute('Type') || '')) continue;
      if (!path) throw new Error('工作表关系缺失，请用 Excel 重新保存后导入。');
      const sheetText = await readPart(zip, path, true);
      const sheet = parseSheet(sheetText, sharedStrings, formats, date1904, warnings, node.getAttribute('name') || '工作表');
      totalRows += sheet.rows.filter(row => row.some(v => v !== '')).length;
      if (totalRows > MAX_ROWS) throw new Error('工作簿有效数据合计超过 30,000 行，请分批导入。');
      sheets.push(sheet);
    }
    if (!sheets.length) throw new Error('此文件没有可读取的工作表。');
    return { name: file.name, sheets, warnings };
  }

  function parseCsv(text, delimiter) {
    const rows = [];
    let row = [];
    let value = '';
    let quoted = false;
    for (let i = 0; i < text.length; i++) {
      const ch = text[i];
      if (ch === '"') {
        if (quoted && text[i + 1] === '"') { value += '"'; i++; }
        else if (quoted || value === '') quoted = !quoted;
        else value += ch;
      } else if (ch === delimiter && !quoted) {
        row.push(value); value = '';
        if (row.length >= MAX_COLS) throw new Error('CSV 超出 100 列，请删除无关列后导入。');
      } else if ((ch === '\r' || ch === '\n') && !quoted) {
        if (ch === '\r' && text[i + 1] === '\n') i++;
        row.push(value); rows.push(row); row = []; value = '';
        if (rows.length > MAX_ROWS) throw new Error('CSV 超出 30,000 行，请拆分后导入。');
      } else value += ch;
    }
    if (quoted) throw new Error('CSV 中存在未闭合的双引号，请检查文件格式。');
    if (value !== '' || row.length) { row.push(value); rows.push(row); }
    while (rows.length && rows[rows.length - 1].every(v => v.trim() === '')) rows.pop();
    if (rows.length > MAX_ROWS) throw new Error('CSV 超出 30,000 行，请拆分后导入。');
    return rows;
  }

  async function readCsv(file) {
    const bytes = new Uint8Array(await file.arrayBuffer());
    let text;
    try {
      text = bytes[0] === 0xff && bytes[1] === 0xfe ? new TextDecoder('utf-16le', { fatal: true }).decode(bytes) : new TextDecoder('utf-8', { fatal: true }).decode(bytes);
    } catch (_) { throw new Error('CSV 编码无法识别，请另存为“CSV UTF-8”后导入。'); }
    text = text.replace(/^\ufeff/, '');
    const firstLine = text.split(/\r?\n/)[0] || '';
    const candidates = [',', '\t', ';'];
    const delimiter = candidates.sort((a, b) => firstLine.split(b).length - firstLine.split(a).length)[0];
    return { name: file.name, sheets: [{ name: 'CSV', rows: parseCsv(text, delimiter), widths: [] }], warnings: [] };
  }

  async function read(file) {
    if (!file || typeof file.arrayBuffer !== 'function') throw new Error('请选择 Excel 或 CSV 文件。');
    if (file.size > MAX_BYTES) throw new Error('文件超过 15 MB，请拆分后导入。');
    if (!file.size) throw new Error('文件为空，请检查后重新选择。');
    const extension = String(file.name || '').split('.').pop().toLowerCase();
    if (extension === 'csv' || extension === 'tsv') return readCsv(file);
    if (extension === 'xls') throw new Error('暂不支持旧版 .xls，请用 Excel 另存为 .xlsx 后导入。');
    if (extension !== 'xlsx') throw new Error('请选择 .xlsx 或 UTF-8 .csv 文件。');
    return readXlsx(file);
  }

  function detectHeader(rows) {
    const groups = [
      /订单|单号|订单编号|订单号/,
      /商品|产品|货品|品名|sku/i,
      /数量|件数|订购|购买数量|采购数量/,
      /收货|收件|买家|客户/,
      /电话|手机|联系/,
      /地址|省份|城市|区县/,
      /规格|型号|编码|条码/,
      /单价|金额|备注|发货|供应商/
    ];
    let bestIndex = 0;
    let bestScore = -1;
    (rows || []).slice(0, 20).forEach((row, index) => {
      const values = (row || []).map(v => String(v == null ? '' : v).trim()).filter(Boolean);
      if (!values.length) return;
      const short = values.filter(v => v.length <= 24);
      const score = groups.reduce((sum, re) => sum + (short.some(v => re.test(v)) ? 5 : 0), 0) + Math.min(short.length, 16) * 0.3 - values.filter(v => v.length > 60).length;
      if (score > bestScore) { bestScore = score; bestIndex = index; }
    });
    return bestIndex;
  }

  function identifierHeader(header) { return /单号|订单|电话|手机|编码|编号|条码|邮编|身份证|id$|sku/i.test(String(header)); }
  function quantityHeader(header) { return /^(数量|件数|订购数量|购买数量|采购数量|商品数量|发货数量|qty|quantity)$/i.test(String(header).trim()); }

  function defaultWidth(header) {
    const text = String(header || '');
    if (/地址/.test(text)) return 48;
    if (/商品名称|产品名称|货品名称|来源文件|备注/.test(text)) return 32;
    if (/编码|编号|单号|订单|电话|手机/.test(text)) return 25;
    return Math.max(12, Math.min(26, text.length * 2 + 4));
  }

  function cellXml(value, ref, header, isHeader) {
    if (isHeader) return '<c r="' + ref + '" s="1" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(value) + '</t></is></c>';
    if (value == null || value === '') return '<c r="' + ref + '" s="2"/>';
    if (value instanceof Date && !Number.isNaN(value.getTime())) {
      return '<c r="' + ref + '" s="4"><v>' + (value.getTime() / 86400000 + 25569) + '</v></c>';
    }
    if (!identifierHeader(header)) {
      const isNumeric = typeof value === 'number' || (quantityHeader(header) && /^-?\d+(\.\d+)?$/.test(String(value).trim()));
      const numeric = Number(value);
      if (isNumeric && Number.isFinite(numeric)) return '<c r="' + ref + '" s="3"><v>' + numeric + '</v></c>';
      if (typeof value === 'boolean') return '<c r="' + ref + '" t="b"><v>' + (value ? '1' : '0') + '</v></c>';
    }
    return '<c r="' + ref + '" s="2" t="inlineStr"><is><t xml:space="preserve">' + xmlEscape(value) + '</t></is></c>';
  }

  async function buildWorkbook(options) {
    const { headers, rows = [], widths = [] } = options || {};
    if (!Array.isArray(headers) || !headers.length) throw new Error('请先配置导出列。');
    if (headers.length > MAX_COLS) throw new Error('导出列数不能超过 100 列。');
    if (!Array.isArray(rows) || rows.length > MAX_ROWS) throw new Error('单次最多导出 30,000 行，请分批处理。');
    const JSZip = zipLibrary();
    const zip = new JSZip();
    const name = cleanText(options.sheetName || '订单').replace(/[\\/?*\[\]:]/g, '').replace(/^'+|'+$/g, '').slice(0, 31) || '订单';
    const end = columnName(headers.length - 1) + (rows.length + 1);
    const columnXml = headers.map((h, i) => '<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' + Math.max(5, Math.min(80, Number(widths[i]) || defaultWidth(h))) + '" customWidth="1"/>').join('');
    const headerXml = '<row r="1" ht="28" customHeight="1">' + headers.map((h, c) => cellXml(h, columnName(c) + '1', h, true)).join('') + '</row>';
    const dataXml = rows.map((row, index) => '<row r="' + (index + 2) + '" ht="22" customHeight="1">' + headers.map((h, c) => cellXml(Array.isArray(row) ? row[c] : row[h], columnName(c) + (index + 2), h, false)).join('') + '</row>').join('');
    zip.file('[Content_Types].xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/><Default Extension="xml" ContentType="application/xml"/><Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/></Types>');
    zip.file('_rels/.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL_NS + '/officeDocument" Target="xl/workbook.xml"/></Relationships>');
    zip.file('xl/workbook.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><workbook xmlns="' + NS + '" xmlns:r="' + REL_NS + '"><bookViews><workbookView/></bookViews><sheets><sheet name="' + xmlEscape(name) + '" sheetId="1" r:id="rId1"/></sheets></workbook>');
    zip.file('xl/_rels/workbook.xml.rels', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Id="rId1" Type="' + REL_NS + '/worksheet" Target="worksheets/sheet1.xml"/><Relationship Id="rId2" Type="' + REL_NS + '/styles" Target="styles.xml"/></Relationships>');
    zip.file('xl/styles.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><styleSheet xmlns="' + NS + '"><numFmts count="1"><numFmt numFmtId="164" formatCode="yyyy-mm-dd"/></numFmts><fonts count="2"><font><sz val="11"/><color rgb="FF24344B"/><name val="等线"/><family val="2"/></font><font><b/><sz val="11"/><color rgb="FFFFFFFF"/><name val="等线"/><family val="2"/></font></fonts><fills count="3"><fill><patternFill patternType="none"/></fill><fill><patternFill patternType="gray125"/></fill><fill><patternFill patternType="solid"><fgColor rgb="FF183B56"/><bgColor indexed="64"/></patternFill></fill></fills><borders count="1"><border><left/><right/><top/><bottom/><diagonal/></border></borders><cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs><cellXfs count="5"><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/><xf numFmtId="0" fontId="1" fillId="2" borderId="0" xfId="0" applyFont="1" applyFill="1" applyAlignment="1"><alignment vertical="center" wrapText="1"/></xf><xf numFmtId="49" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1" applyAlignment="1"><alignment vertical="center"/></xf><xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/></cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>');
    zip.file('xl/worksheets/sheet1.xml', '<?xml version="1.0" encoding="UTF-8" standalone="yes"?><worksheet xmlns="' + NS + '"><dimension ref="A1:' + end + '"/><sheetViews><sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/></sheetView></sheetViews><sheetFormatPr defaultRowHeight="22"/><cols>' + columnXml + '</cols><sheetData>' + headerXml + dataXml + '</sheetData><autoFilter ref="A1:' + end + '"/><pageMargins left="0.3" right="0.3" top="0.5" bottom="0.5" header="0.2" footer="0.2"/></worksheet>');
    return zip.generateAsync({ type: 'uint8array', compression: 'DEFLATE', compressionOptions: { level: 6 } });
  }

  function downloadBlob(blob, filename) {
    const file = blob instanceof Blob ? blob : new Blob([blob], { type: MIME });
    const url = URL.createObjectURL(file);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = cleanText(filename || '订单.xlsx').replace(/[\\/:*?"<>|]/g, '_');
    anchor.style.display = 'none';
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }

  async function downloadWorkbook(options) {
    const bytes = await buildWorkbook(options);
    const name = options.filename || '订单.xlsx';
    downloadBlob(new Blob([bytes], { type: MIME }), /\.xlsx$/i.test(name) ? name : name + '.xlsx');
    return bytes;
  }

  function toBase64(data) {
    const bytes = data instanceof Uint8Array ? data : new Uint8Array(data);
    let binary = '';
    for (let i = 0; i < bytes.length; i += 8192) binary += String.fromCharCode.apply(null, bytes.subarray(i, i + 8192));
    return btoa(binary);
  }

  function fromBase64(text) {
    const binary = atob(String(text));
    return Uint8Array.from(binary, ch => ch.charCodeAt(0));
  }

  window.OrderExcel = Object.freeze({ read, detectHeader, buildWorkbook, downloadWorkbook, downloadBlob, toBase64, fromBase64, limits: Object.freeze({ bytes: MAX_BYTES, rows: MAX_ROWS, cols: MAX_COLS }) });
})();
