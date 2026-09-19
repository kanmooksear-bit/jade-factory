/* xlsx.js — เขียนไฟล์ Excel (.xlsx) ขึ้นมาเองในเบราว์เซอร์
   ไม่ต้องพึ่งไลบรารีข้างนอก โหลดเฉพาะตอนกดปุ่ม "โหลด Excel"

   วิธีใช้
     var blob = XLSXMini.build([{ name: "ชีต1", cols: [{w:12}], rows: [[...]] }]);
   ช่องหนึ่งช่อง (cell) ใส่ได้แบบนี้
     null / ""            ช่องว่าง
     "ข้อความ"             ข้อความ
     123                  ตัวเลข
     { f: "SUM(A1:A9)" }  สูตร (ไม่ต้องมี = ข้างหน้า)
     { d: "2026-09-18" }  วันที่
     { v: ..., s: 2 }     กำหนดสไตล์เอง
   สไตล์: 0 ปกติ · 1 หัวเรื่องใหญ่ · 2 หัวตาราง · 3 วันที่ · 4 ตัวเลขทศนิยม · 5 ตัวหนา
*/
(function () {
  "use strict";

  /* ---------- CRC32 ---------- */
  var TAB = (function () {
    var t = new Uint32Array(256), c, n, k;
    for (n = 0; n < 256; n++) {
      c = n;
      for (k = 0; k < 8; k++) c = (c & 1) ? (0xEDB88320 ^ (c >>> 1)) : (c >>> 1);
      t[n] = c >>> 0;
    }
    return t;
  })();
  function crc32(buf) {
    var c = 0xFFFFFFFF;
    for (var i = 0; i < buf.length; i++) c = TAB[(c ^ buf[i]) & 0xFF] ^ (c >>> 8);
    return (c ^ 0xFFFFFFFF) >>> 0;
  }

  /* ---------- ZIP แบบไม่บีบอัด (stored) ---------- */
  function zip(files) {
    var enc = new TextEncoder(), parts = [], central = [], off = 0;
    files.forEach(function (f) {
      var nm = enc.encode(f.name);
      var data = typeof f.data === "string" ? enc.encode(f.data) : f.data;
      var crc = crc32(data);

      var lh = new Uint8Array(30 + nm.length);
      var dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);        // version needed
      dv.setUint16(6, 0x0800, true);    // ธงบอกว่าชื่อไฟล์เป็น UTF-8
      dv.setUint16(8, 0, true);         // method 0 = stored
      dv.setUint16(10, 0, true); dv.setUint16(12, 0x2821, true); // เวลา/วันที่คงที่
      dv.setUint32(14, crc, true);
      dv.setUint32(18, data.length, true);
      dv.setUint32(22, data.length, true);
      dv.setUint16(26, nm.length, true);
      dv.setUint16(28, 0, true);
      lh.set(nm, 30);
      parts.push(lh, data);

      var ch = new Uint8Array(46 + nm.length);
      var cv = new DataView(ch.buffer);
      cv.setUint32(0, 0x02014b50, true);
      cv.setUint16(4, 20, true); cv.setUint16(6, 20, true);
      cv.setUint16(8, 0x0800, true);
      cv.setUint16(10, 0, true);
      cv.setUint16(12, 0, true); cv.setUint16(14, 0x2821, true);
      cv.setUint32(16, crc, true);
      cv.setUint32(20, data.length, true);
      cv.setUint32(24, data.length, true);
      cv.setUint16(28, nm.length, true);
      cv.setUint32(42, off, true);
      ch.set(nm, 46);
      central.push(ch);

      off += lh.length + data.length;
    });

    var csize = central.reduce(function (a, b) { return a + b.length; }, 0);
    var eo = new Uint8Array(22);
    var ev = new DataView(eo.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, csize, true);
    ev.setUint32(16, off, true);

    return new Blob(parts.concat(central, [eo]),
      { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" });
  }

  /* ---------- ตัวช่วย ---------- */
  function esc(s) {
    return String(s).replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/\r/g, "");
  }
  // Excel นับวันจาก 30 ธ.ค. 1899
  function serial(iso) {
    var p = String(iso).slice(0, 10).split("-");
    var ms = Date.UTC(+p[0], +p[1] - 1, +p[2]);
    return Math.round(ms / 86400000) + 25569;
  }
  function colName(n) {                 // 0 -> A, 26 -> AA
    var s = "";
    n += 1;
    while (n > 0) { var r = (n - 1) % 26; s = String.fromCharCode(65 + r) + s; n = (n - r - 1) / 26; }
    return s;
  }

  /* ---------- ชีตหนึ่งชีต ---------- */
  function sheetXml(sh) {
    var out = ['<?xml version="1.0" encoding="UTF-8" standalone="yes"?>',
      '<worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">'];
    if (sh.freeze) {
      out.push('<sheetViews><sheetView workbookViewId="0"><pane ySplit="' + sh.freeze +
        '" topLeftCell="A' + (sh.freeze + 1) + '" activePane="bottomLeft" state="frozen"/>' +
        '</sheetView></sheetViews>');
    }
    if (sh.cols && sh.cols.length) {
      out.push("<cols>");
      sh.cols.forEach(function (c, i) {
        out.push('<col min="' + (i + 1) + '" max="' + (i + 1) + '" width="' +
          (c && c.w ? c.w : 12) + '" customWidth="1"/>');
      });
      out.push("</cols>");
    }
    out.push("<sheetData>");

    (sh.rows || []).forEach(function (row, ri) {
      var r = ri + 1, cells = [];
      (row || []).forEach(function (cell, ci) {
        if (cell === null || cell === undefined || cell === "") return;
        var ref = colName(ci) + r, s = 0, body = "", tAttr = "";

        if (typeof cell === "object") {
          if (cell.s !== undefined) s = cell.s;
          if (cell.f !== undefined) {
            body = "<f>" + esc(cell.f) + "</f>";
          } else if (cell.d !== undefined) {
            if (cell.d === null || cell.d === "") return;
            if (s === 0) s = 3;
            body = "<v>" + serial(cell.d) + "</v>";
          } else if (cell.v === null || cell.v === undefined || cell.v === "") {
            if (!s) return;
          } else if (typeof cell.v === "number") {
            body = "<v>" + cell.v + "</v>";
          } else {
            tAttr = ' t="inlineStr"';
            body = "<is><t xml:space=\"preserve\">" + esc(cell.v) + "</t></is>";
          }
        } else if (typeof cell === "number") {
          body = "<v>" + cell + "</v>";
        } else {
          tAttr = ' t="inlineStr"';
          body = "<is><t xml:space=\"preserve\">" + esc(cell) + "</t></is>";
        }

        cells.push('<c r="' + ref + '"' + (s ? ' s="' + s + '"' : "") + tAttr + ">" + body + "</c>");
      });
      if (cells.length) out.push('<row r="' + r + '">' + cells.join("") + "</row>");
    });

    out.push("</sheetData></worksheet>");
    return out.join("");
  }

  var STYLES = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
    '<styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">' +
    '<numFmts count="2">' +
      '<numFmt numFmtId="164" formatCode="dd/mm/yyyy"/>' +
      '<numFmt numFmtId="165" formatCode="#,##0.00"/></numFmts>' +
    '<fonts count="3">' +
      '<font><sz val="11"/><name val="Tahoma"/></font>' +
      '<font><b/><sz val="11"/><name val="Tahoma"/></font>' +
      '<font><b/><sz val="14"/><name val="Tahoma"/></font></fonts>' +
    '<fills count="3">' +
      '<fill><patternFill patternType="none"/></fill>' +
      '<fill><patternFill patternType="gray125"/></fill>' +
      '<fill><patternFill patternType="solid"><fgColor rgb="FFEDEDED"/>' +
        '<bgColor indexed="64"/></patternFill></fill></fills>' +
    '<borders count="2"><border/>' +
      '<border><left/><right/><top/><bottom style="thin">' +
        '<color rgb="FF999999"/></bottom><diagonal/></border></borders>' +
    '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
    '<cellXfs count="6">' +
      '<xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>' +
      '<xf numFmtId="0" fontId="2" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
      '<xf numFmtId="0" fontId="1" fillId="2" borderId="1" xfId="0" ' +
        'applyFont="1" applyFill="1" applyBorder="1"/>' +
      '<xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>' +
      '<xf numFmtId="0" fontId="1" fillId="0" borderId="0" xfId="0" applyFont="1"/>' +
    '</cellXfs>' +
    '<cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles>' +
    "</styleSheet>";

  function build(sheets) {
    var files = [];

    files.push({ name: "[Content_Types].xml", data:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
      '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
      '<Default Extension="xml" ContentType="application/xml"/>' +
      '<Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>' +
      sheets.map(function (s, i) {
        return '<Override PartName="/xl/worksheets/sheet' + (i + 1) +
          '.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>';
      }).join("") +
      '<Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>' +
      "</Types>" });

    files.push({ name: "_rels/.rels", data:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>' +
      "</Relationships>" });

    files.push({ name: "xl/workbook.xml", data:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" ' +
      'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships"><sheets>' +
      sheets.map(function (s, i) {
        return '<sheet name="' + esc(s.name) + '" sheetId="' + (i + 1) +
          '" r:id="rId' + (i + 1) + '"/>';
      }).join("") +
      '</sheets><calcPr calcId="0" fullCalcOnLoad="1"/></workbook>' });

    files.push({ name: "xl/_rels/workbook.xml.rels", data:
      '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
      '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
      sheets.map(function (s, i) {
        return '<Relationship Id="rId' + (i + 1) +
          '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet' +
          (i + 1) + '.xml"/>';
      }).join("") +
      '<Relationship Id="rId' + (sheets.length + 1) +
      '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
      "</Relationships>" });

    files.push({ name: "xl/styles.xml", data: STYLES });
    sheets.forEach(function (s, i) {
      files.push({ name: "xl/worksheets/sheet" + (i + 1) + ".xml", data: sheetXml(s) });
    });

    return zip(files);
  }

  function save(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  window.XLSXMini = { build: build, save: save, serial: serial };
})();
