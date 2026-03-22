(function (global) {

  const HEADERS = ["Cod", "Secc", "Curso", "Tipo", "Horario", "Aula", "Docente"];
  const TITLE = "Matrícula UNI | Horarios";
  const TABLE_NAME = "HorariosUNI";
  const BRAND_HEX = "1F4E78";
  const XML_HEADER = '<?xml version="1.0" encoding="utf-8"?>';

  const STATIC_FILES = {
    "[Content_Types].xml": "\uFEFF" + XML_HEADER + '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types"><Default Extension="xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" /><Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml" /><Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" /><Override PartName="/xl/theme/theme1.xml" ContentType="application/vnd.openxmlformats-officedocument.theme+xml" /><Override PartName="/xl/worksheets/sheet1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" /><Override PartName="/xl/tables/table1.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml" /><Override PartName="/xl/worksheets/sheet2.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" /></Types>',
    "_rels/.rels": "\uFEFF" + XML_HEADER + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="/xl/workbook.xml" Id="R1d92fb803f4942f4" /></Relationships>',
    "xl/workbook.xml": XML_HEADER + '<x:workbook xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheets><x:sheet name="Horarios" sheetId="1" r:id="Refa401af0ca34cb4" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" /><x:sheet name="Fuente" sheetId="2" r:id="R3948f4d4306b41eb" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" /></x:sheets></x:workbook>',
    "xl/_rels/workbook.xml.rels": "\uFEFF" + XML_HEADER + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="/xl/styles.xml" Id="R20f7cd23889b490d" /><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Target="/xl/theme/theme1.xml" Id="R0ab13dcc6938484b" /><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet1.xml" Id="Refa401af0ca34cb4" /><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="/xl/worksheets/sheet2.xml" Id="R3948f4d4306b41eb" /></Relationships>',
    "xl/worksheets/_rels/sheet1.xml.rels": "\uFEFF" + XML_HEADER + '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships"><Relationship Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/table" Target="/xl/tables/table1.xml" Id="R0726a36dc5ec4342" /></Relationships>',
    "xl/styles.xml": XML_HEADER + '<x:styleSheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:fonts count="6"><x:font><x:sz val="11" /><x:color theme="1" /><x:name val="Calibri" /><x:family val="2" /><x:scheme val="minor" /></x:font><x:font><x:b /><x:sz val="14" /></x:font><x:font><x:i /><x:color rgb="00666666" /></x:font><x:font><x:b /><x:color rgb="00FFFFFF" /></x:font><x:font><x:color rgb="00666666" /></x:font><x:font><x:b /></x:font></x:fonts><x:fills count="3"><x:fill><x:patternFill patternType="none" /></x:fill><x:fill><x:patternFill patternType="gray125" /></x:fill><x:fill><x:patternFill patternType="solid"><x:fgColor rgb="001F4E78" /></x:patternFill></x:fill></x:fills><x:borders count="2"><x:border><x:left /><x:right /><x:top /><x:bottom /><x:diagonal /></x:border><x:border><x:bottom style="thin"><x:color rgb="00D9D9D9" /></x:bottom></x:border></x:borders><x:cellStyleXfs count="1"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" /></x:cellStyleXfs><x:cellXfs count="9"><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" /><x:xf numFmtId="0" fontId="1" fillId="0" borderId="1" xfId="0" /><x:xf numFmtId="0" fontId="0" fillId="0" borderId="1" xfId="0" /><x:xf numFmtId="0" fontId="2" fillId="0" borderId="1" xfId="0" /><x:xf numFmtId="0" fontId="4" fillId="0" borderId="0" xfId="0" /><x:xf numFmtId="0" fontId="3" fillId="2" borderId="0" xfId="0" applyAlignment="1"><x:alignment horizontal="center" vertical="center" /></x:xf><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><x:alignment horizontal="center" vertical="top" /></x:xf><x:xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0" applyAlignment="1"><x:alignment vertical="top" wrapText="1" /></x:xf><x:xf numFmtId="0" fontId="5" fillId="0" borderId="0" xfId="0" /></x:cellXfs><x:cellStyles count="1"><x:cellStyle name="Normal" xfId="0" /></x:cellStyles></x:styleSheet>',
    "xl/theme/theme1.xml": XML_HEADER + '<a:theme xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" name="Office Theme"><a:themeElements><a:clrScheme name="Office"><a:dk1><a:sysClr val="windowText" lastClr="000000" /></a:dk1><a:lt1><a:sysClr val="window" lastClr="FFFFFF" /></a:lt1><a:dk2><a:srgbClr val="1F497D" /></a:dk2><a:lt2><a:srgbClr val="EEECE1" /></a:lt2><a:accent1><a:srgbClr val="4F81BD" /></a:accent1><a:accent2><a:srgbClr val="C0504D" /></a:accent2><a:accent3><a:srgbClr val="9BBB59" /></a:accent3><a:accent4><a:srgbClr val="8064A2" /></a:accent4><a:accent5><a:srgbClr val="4BACC6" /></a:accent5><a:accent6><a:srgbClr val="F79646" /></a:accent6><a:hlink><a:srgbClr val="0000FF" /></a:hlink><a:folHlink><a:srgbClr val="800080" /></a:folHlink></a:clrScheme><a:fontScheme name="Office"><a:majorFont><a:latin typeface="Cambria" /><a:ea typeface="" /><a:cs typeface="" /></a:majorFont><a:minorFont><a:latin typeface="Calibri" /><a:ea typeface="" /><a:cs typeface="" /></a:minorFont></a:fontScheme><a:fmtScheme name="Office"><a:fillStyleLst><a:solidFill><a:schemeClr val="phClr" /></a:solidFill><a:gradFill rotWithShape="1"><a:gsLst><a:gs pos="0"><a:schemeClr val="phClr"><a:tint val="50000" /><a:satMod val="300000" /></a:schemeClr></a:gs><a:gs pos="35000"><a:schemeClr val="phClr"><a:tint val="37000" /><a:satMod val="300000" /></a:schemeClr></a:gs><a:gs pos="100000"><a:schemeClr val="phClr"><a:tint val="15000" /><a:satMod val="350000" /></a:schemeClr></a:gs></a:gsLst><a:lin ang="16200000" scaled="1" /></a:gradFill></a:fillStyleLst><a:lnStyleLst><a:ln w="9525" cap="flat" cmpd="sng" algn="ctr"><a:solidFill><a:schemeClr val="phClr"><a:shade val="95000" /><a:satMod val="105000" /></a:schemeClr></a:solidFill><a:prstDash val="solid" /></a:ln></a:lnStyleLst><a:effectStyleLst><a:effectStyle><a:effectLst /></a:effectStyle></a:effectStyleLst><a:bgFillStyleLst><a:solidFill><a:schemeClr val="phClr" /></a:solidFill></a:bgFillStyleLst></a:fmtScheme></a:themeElements><a:objectDefaults /><a:extraClrSchemeLst /></a:theme>'
  };

  function escapeXml(value) {
    return String(value ?? "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/\"/g, "&quot;")
      .replace(/'/g, "&apos;");
  }

  function normalizeWhitespace(value) {
    return String(value ?? "").replace(/\s+/g, " ").trim();
  }

  function decodeHtmlEntities(value) {
    const entityMap = {
      amp: "&",
      lt: "<",
      gt: ">",
      quot: '"',
      apos: "'",
      nbsp: " "
    };

    return String(value ?? "")
      .replace(/&#(\d+);/g, (_, code) => String.fromCodePoint(Number(code)))
      .replace(/&#x([\da-f]+);/gi, (_, code) => String.fromCodePoint(parseInt(code, 16)))
      .replace(/&([a-z]+);/gi, (full, name) => entityMap[name.toLowerCase()] ?? full);
  }

  function stripTags(value) {
    return normalizeWhitespace(
      decodeHtmlEntities(String(value ?? "").replace(/<br\s*\/?>/gi, " ").replace(/<[^>]+>/g, " "))
    );
  }

  function extractRowsFromTable(tableHtml) {
    const tbodyMatch = tableHtml.match(/<tbody[^>]*>([\s\S]*?)<\/tbody>/i);
    const rowSource = tbodyMatch ? tbodyMatch[1] : tableHtml;
    const rowMatches = rowSource.match(/<tr[^>]*>[\s\S]*?<\/tr>/gi) || [];

    return rowMatches
      .map((rowHtml) => {
        const cells = [...rowHtml.matchAll(/<t[dh][^>]*>([\s\S]*?)<\/t[dh]>/gi)].map((match) => stripTags(match[1]));
        if (cells.length < HEADERS.length) {
          return null;
        }
        return HEADERS.reduce((accumulator, header, index) => {
          accumulator[header] = cells[index] ?? "";
          return accumulator;
        }, {});
      })
      .filter(Boolean);
  }

  function extractRowsFromHtml(htmlText) {
    const tables = htmlText.match(/<table[^>]*>[\s\S]*?<\/table>/gi) || [];

    for (const tableHtml of tables) {
      const headerMatches = [...tableHtml.matchAll(/<th[^>]*>([\s\S]*?)<\/th>/gi)].map((match) => stripTags(match[1]));
      const signature = headerMatches.slice(0, HEADERS.length);
      if (HEADERS.every((header, index) => signature[index] === header)) {
        return extractRowsFromTable(tableHtml);
      }
    }

    throw new Error("No se encontró una tabla con las columnas esperadas en el HTML.");
  }

  function columnLetter(index) {
    let value = "";
    let current = index;
    while (current > 0) {
      const remainder = (current - 1) % 26;
      value = String.fromCharCode(65 + remainder) + value;
      current = Math.floor((current - 1) / 26);
    }
    return value;
  }

  function inlineCell(reference, styleId, value) {
    return `<x:c r="${reference}" s="${styleId}" t="inlineStr"><x:is><x:t xml:space="preserve">${escapeXml(value)}</x:t></x:is></x:c>`;
  }

  function numberCell(reference, styleId, value) {
    return `<x:c r="${reference}" s="${styleId}" t="n"><x:v>${value}</x:v></x:c>`;
  }

  function buildSheet1Xml(rows, sourceSummary) {
    const sheetRows = [];
    sheetRows.push(`<x:row r="1" ht="22" customHeight="1">${inlineCell("A1", 1, TITLE)}${["B", "C", "D", "E", "F", "G"].map((column) => `<x:c r="${column}1" s="2" t="n" />`).join("")}</x:row>`);
    sheetRows.push(`<x:row r="2">${inlineCell("A2", 3, `Extraído de: ${sourceSummary}`)}${["B", "C", "D", "E", "F", "G"].map((column) => `<x:c r="${column}2" s="2" t="n" />`).join("")}</x:row>`);
    sheetRows.push(`<x:row r="3">${inlineCell("A3", 4, `Total de filas extraídas: ${rows.length}`)}</x:row>`);
    sheetRows.push(`<x:row r="4" ht="22" customHeight="1">${HEADERS.map((header, index) => inlineCell(`${columnLetter(index + 1)}4`, 5, header)).join("")}</x:row>`);

    rows.forEach((row, index) => {
      const rowNumber = index + 5;
      const values = HEADERS.map((header) => row[header] ?? "");
      const cells = values.map((value, columnIndex) => {
        const column = columnLetter(columnIndex + 1);
        const styleId = columnIndex <= 4 ? 6 : 7;
        return inlineCell(`${column}${rowNumber}`, styleId, value);
      });
      sheetRows.push(`<x:row r="${rowNumber}">${cells.join("")}</x:row>`);
    });

    const endRow = rows.length + 4;

    return XML_HEADER +
      `<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetFormatPr baseColWidth="8" defaultRowHeight="15" /><x:cols><x:col min="1" max="1" width="12" customWidth="1" /><x:col min="2" max="2" width="8" customWidth="1" /><x:col min="3" max="3" width="38" customWidth="1" /><x:col min="4" max="4" width="16" customWidth="1" /><x:col min="5" max="5" width="16" customWidth="1" /><x:col min="6" max="6" width="24" customWidth="1" /><x:col min="7" max="7" width="36" customWidth="1" /></x:cols><x:sheetData>${sheetRows.join("")}</x:sheetData><x:mergeCells><x:mergeCell ref="A3:G3" /><x:mergeCell ref="A2:G2" /><x:mergeCell ref="A1:G1" /></x:mergeCells><x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" /><x:tableParts count="1"><x:tablePart r:id="R0726a36dc5ec4342" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" /></x:tableParts></x:worksheet>`;
  }

  function buildSheet2Xml(files, rowCount) {
    const rows = [
      `<x:row r="1"><x:c r="A1" s="8" t="inlineStr"><x:is><x:t xml:space="preserve">Archivo origen</x:t></x:is></x:c><x:c r="B1" s="8" t="inlineStr"><x:is><x:t xml:space="preserve">${escapeXml(files.map((file) => file.name).join(", "))}</x:t></x:is></x:c></x:row>`,
      `<x:row r="2"><x:c r="A2" t="inlineStr"><x:is><x:t xml:space="preserve">Filas extraídas</x:t></x:is></x:c>${numberCell("B2", 0, rowCount)}</x:row>`,
      `<x:row r="3"><x:c r="A3" t="inlineStr"><x:is><x:t xml:space="preserve">Columnas</x:t></x:is></x:c><x:c r="B3" t="inlineStr"><x:is><x:t xml:space="preserve">${escapeXml(HEADERS.join(", "))}</x:t></x:is></x:c></x:row>`
    ];

    files.forEach((file, index) => {
      const rowNumber = index + 4;
      rows.push(`<x:row r="${rowNumber}"><x:c r="A${rowNumber}" t="inlineStr"><x:is><x:t xml:space="preserve">Archivo ${index + 1}</x:t></x:is></x:c><x:c r="B${rowNumber}" t="inlineStr"><x:is><x:t xml:space="preserve">${escapeXml(file.name)}</x:t></x:is></x:c></x:row>`);
    });

    return XML_HEADER + `<x:worksheet xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:sheetFormatPr baseColWidth="8" defaultRowHeight="15" /><x:cols><x:col min="1" max="1" width="20" customWidth="1" /><x:col min="2" max="2" width="50" customWidth="1" /></x:cols><x:sheetData>${rows.join("")}</x:sheetData><x:pageMargins left="0.7" right="0.7" top="0.75" bottom="0.75" header="0.3" footer="0.3" /></x:worksheet>`;
  }

  function buildTableXml(rowCount) {
    const endRow = rowCount + 4;
    return XML_HEADER + `<x:table id="1" name="${TABLE_NAME}" displayName="${TABLE_NAME}" ref="A4:G${endRow}" headerRowCount="1" xmlns:x="http://schemas.openxmlformats.org/spreadsheetml/2006/main"><x:autoFilter ref="A4:G${endRow}" /><x:tableColumns count="7">${HEADERS.map((header, index) => `<x:tableColumn id="${index + 1}" name="${escapeXml(header)}" />`).join("")}</x:tableColumns><x:tableStyleInfo name="TableStyleMedium2" showRowStripes="1" /></x:table>`;
  }

  function makeSourceSummary(files) {
    if (files.length === 1) {
      return files[0].name;
    }
    return `${files.length} archivos (${files.map((file) => file.name).join(", ")})`;
  }

  function createWorkbookFiles(files, rows) {
    return {
      ...STATIC_FILES,
      "xl/worksheets/sheet1.xml": buildSheet1Xml(rows, makeSourceSummary(files)),
      "xl/worksheets/sheet2.xml": buildSheet2Xml(files, rows.length),
      "xl/tables/table1.xml": buildTableXml(rows.length)
    };
  }

  function createCrcTable() {
    const table = new Uint32Array(256);
    for (let i = 0; i < 256; i += 1) {
      let current = i;
      for (let bit = 0; bit < 8; bit += 1) {
        current = (current & 1) ? (0xedb88320 ^ (current >>> 1)) : (current >>> 1);
      }
      table[i] = current >>> 0;
    }
    return table;
  }

  const CRC_TABLE = createCrcTable();
  const encoder = new TextEncoder();

  function crc32(bytes) {
    let crc = 0xffffffff;
    for (let i = 0; i < bytes.length; i += 1) {
      crc = CRC_TABLE[(crc ^ bytes[i]) & 0xff] ^ (crc >>> 8);
    }
    return (crc ^ 0xffffffff) >>> 0;
  }

  function writeUint16(view, offset, value) {
    view.setUint16(offset, value, true);
  }

  function writeUint32(view, offset, value) {
    view.setUint32(offset, value >>> 0, true);
  }

  function createZip(fileMap) {
    const entries = Object.entries(fileMap).map(([name, content]) => {
      const nameBytes = encoder.encode(name);
      const dataBytes = encoder.encode(content);
      return {
        name,
        nameBytes,
        dataBytes,
        crc: crc32(dataBytes)
      };
    });

    let localSize = 0;
    let centralSize = 0;
    entries.forEach((entry) => {
      localSize += 30 + entry.nameBytes.length + entry.dataBytes.length;
      centralSize += 46 + entry.nameBytes.length;
    });

    const output = new Uint8Array(localSize + centralSize + 22);
    const view = new DataView(output.buffer);
    let offset = 0;

    entries.forEach((entry) => {
      entry.localOffset = offset;
      writeUint32(view, offset, 0x04034b50);
      writeUint16(view, offset + 4, 20);
      writeUint16(view, offset + 6, 0);
      writeUint16(view, offset + 8, 0);
      writeUint16(view, offset + 10, 0);
      writeUint16(view, offset + 12, 0);
      writeUint32(view, offset + 14, entry.crc);
      writeUint32(view, offset + 18, entry.dataBytes.length);
      writeUint32(view, offset + 22, entry.dataBytes.length);
      writeUint16(view, offset + 26, entry.nameBytes.length);
      writeUint16(view, offset + 28, 0);
      output.set(entry.nameBytes, offset + 30);
      output.set(entry.dataBytes, offset + 30 + entry.nameBytes.length);
      offset += 30 + entry.nameBytes.length + entry.dataBytes.length;
    });

    const centralOffset = offset;

    entries.forEach((entry) => {
      writeUint32(view, offset, 0x02014b50);
      writeUint16(view, offset + 4, 20);
      writeUint16(view, offset + 6, 20);
      writeUint16(view, offset + 8, 0);
      writeUint16(view, offset + 10, 0);
      writeUint16(view, offset + 12, 0);
      writeUint16(view, offset + 14, 0);
      writeUint32(view, offset + 16, entry.crc);
      writeUint32(view, offset + 20, entry.dataBytes.length);
      writeUint32(view, offset + 24, entry.dataBytes.length);
      writeUint16(view, offset + 28, entry.nameBytes.length);
      writeUint16(view, offset + 30, 0);
      writeUint16(view, offset + 32, 0);
      writeUint16(view, offset + 34, 0);
      writeUint16(view, offset + 36, 0);
      writeUint32(view, offset + 38, 0);
      writeUint32(view, offset + 42, entry.localOffset);
      output.set(entry.nameBytes, offset + 46);
      offset += 46 + entry.nameBytes.length;
    });

    const centralLength = offset - centralOffset;
    writeUint32(view, offset, 0x06054b50);
    writeUint16(view, offset + 4, 0);
    writeUint16(view, offset + 6, 0);
    writeUint16(view, offset + 8, entries.length);
    writeUint16(view, offset + 10, entries.length);
    writeUint32(view, offset + 12, centralLength);
    writeUint32(view, offset + 16, centralOffset);
    writeUint16(view, offset + 20, 0);

    return output;
  }

  function buildWorkbook(files, rows) {
    return createZip(createWorkbookFiles(files, rows));
  }

  function downloadWorkbook(files, rows) {
    const data = buildWorkbook(files, rows);
    const blob = new Blob([data], {
      type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
    });
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement("a");
    anchor.href = url;
    anchor.download = files.length === 1 ? `${files[0].name.replace(/\.(html?|HTML?)$/, "")}.xlsx` : "horarios_uni_combinado.xlsx";
    anchor.click();
    setTimeout(() => URL.revokeObjectURL(url), 1000);
  }

  function createState() {
    return {
      files: [],
      rows: []
    };
  }

  async function readFiles(fileList) {
    const htmlFiles = [...fileList].filter((file) => /\.html?$/i.test(file.name));
    if (!htmlFiles.length) {
      throw new Error("Selecciona al menos un archivo HTML válido.");
    }

    const processed = [];
    const rows = [];

    for (const file of htmlFiles) {
      const text = await file.text();
      const fileRows = extractRowsFromHtml(text);
      processed.push({ name: file.name, rows: fileRows.length });
      rows.push(...fileRows);
    }

    return { files: processed, rows };
  }

  function renderPreviewRows(tbody, rows) {
    if (!rows.length) {
      tbody.innerHTML = '<tr><td colspan="7" class="empty-state empty-state--table">Arrastra un archivo HTML para ver la tabla.</td></tr>';
      return;
    }

    const fragment = rows
      .slice(0, 50)
      .map((row) => `<tr>${HEADERS.map((header) => `<td>${escapeXml(row[header])}</td>`).join("")}</tr>`)
      .join("");

    tbody.innerHTML = fragment;
  }

  function renderSourceList(listElement, files) {
    if (!files.length) {
      listElement.innerHTML = '<li class="empty-state">Todavía no hay archivos procesados.</li>';
      return;
    }

    listElement.innerHTML = files
      .map((file) => `<li><strong>${escapeXml(file.name)}</strong> — ${file.rows} filas</li>`)
      .join("");
  }

  function setMessage(messageElement, text, tone = "") {
    messageElement.textContent = text;
    if (tone) {
      messageElement.dataset.tone = tone;
    } else {
      delete messageElement.dataset.tone;
    }
  }

  function setupTheme() {
    const toggle = document.getElementById("theme-toggle");
    if (!toggle) return;

    const ICONS = { light: "🌙", dark: "☀️" };
    const LABELS = { light: "Cambiar a modo oscuro", dark: "Cambiar a modo claro" };
    const STORAGE_KEY = "theme";

    function applyTheme(theme) {
      document.documentElement.dataset.theme = theme;
      toggle.textContent = ICONS[theme];
      toggle.setAttribute("aria-label", LABELS[theme]);
    }

    const saved = localStorage.getItem(STORAGE_KEY);
    const systemDark = window.matchMedia("(prefers-color-scheme: dark)").matches;
    applyTheme(saved || (systemDark ? "dark" : "light"));

    toggle.addEventListener("click", () => {
      const next = document.documentElement.dataset.theme === "dark" ? "light" : "dark";
      localStorage.setItem(STORAGE_KEY, next);
      applyTheme(next);
    });

    window.matchMedia("(prefers-color-scheme: dark)").addEventListener("change", (e) => {
      if (!localStorage.getItem(STORAGE_KEY)) {
        applyTheme(e.matches ? "dark" : "light");
      }
    });
  }

  function setupApp() {
    if (typeof document === "undefined") {
      return;
    }

    setupTheme();
    const state = createState();
    const dropZone = document.getElementById("drop-zone");
    const fileInput = document.getElementById("file-input");
    const downloadButton = document.getElementById("download-button");
    const resetButton = document.getElementById("reset-button");
    const rowCount = document.getElementById("row-count");
    const fileCount = document.getElementById("file-count");
    const statusText = document.getElementById("status-text");
    const previewBody = document.getElementById("preview-body");
    const sourceList = document.getElementById("source-list");
    const tableCaption = document.getElementById("table-caption");
    const message = document.getElementById("message");

    function refreshView() {
      fileCount.textContent = String(state.files.length);
      rowCount.textContent = String(state.rows.length);
      statusText.textContent = state.rows.length ? "Listo para exportar" : "Esperando archivos";
      downloadButton.disabled = !state.rows.length;
      resetButton.disabled = !state.files.length;
      tableCaption.textContent = state.rows.length ? `Mostrando ${Math.min(state.rows.length, 50)} de ${state.rows.length} filas.` : "No hay datos todavía.";
      renderPreviewRows(previewBody, state.rows);
      renderSourceList(sourceList, state.files);
    }

    async function handleFiles(fileList) {
      try {
        setMessage(message, "Procesando archivos...", "");
        const result = await readFiles(fileList);
        state.files = result.files;
        state.rows = result.rows;
        refreshView();
        setMessage(message, `Se procesaron ${state.files.length} archivo(s) y ${state.rows.length} fila(s).`, "success");
      } catch (error) {
        state.files = [];
        state.rows = [];
        refreshView();
        setMessage(message, error.message, "error");
      }
    }

    function resetState() {
      state.files = [];
      state.rows = [];
      fileInput.value = "";
      refreshView();
      setMessage(message, "Se limpiaron los datos cargados.", "");
    }

    ["dragenter", "dragover"].forEach((eventName) => {
      dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropZone.classList.add("is-active");
      });
    });

    ["dragleave", "drop"].forEach((eventName) => {
      dropZone.addEventListener(eventName, (event) => {
        event.preventDefault();
        dropZone.classList.remove("is-active");
      });
    });

    dropZone.addEventListener("drop", (event) => {
      if (event.dataTransfer?.files?.length) {
        handleFiles(event.dataTransfer.files);
      }
    });

    fileInput.addEventListener("change", (event) => {
      if (event.target.files?.length) {
        handleFiles(event.target.files);
      }
    });

    downloadButton.addEventListener("click", () => {
      if (state.rows.length) {
        downloadWorkbook(state.files, state.rows);
        setMessage(message, "El archivo XLSX se descargó correctamente.", "success");
      }
    });

    resetButton.addEventListener("click", resetState);
    refreshView();
  }

  const exported = {
    HEADERS,
    buildWorkbook,
    createWorkbookFiles,
    decodeHtmlEntities,
    extractRowsFromHtml,
    makeSourceSummary,
    stripTags
  };

  if (typeof module !== "undefined" && module.exports) {
    module.exports = exported;
  } else {
    global.HtmlToXlsxApp = exported;
    setupApp();
  }
})(typeof window !== "undefined" ? window : globalThis);
