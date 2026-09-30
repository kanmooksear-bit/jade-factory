/* docx.js — เขียนไฟล์ Word (.docx) ขึ้นมาเองในเบราว์เซอร์
   ไม่พึ่งไลบรารีข้างนอก โหลดเฉพาะตอนกดปุ่ม "โหลด Word"

   วิธีใช้
     var d = DOCXMini.doc();          // เปิดเอกสารใหม่ A4 แนวตั้ง
     d.title("หัวเรื่อง");
     d.p("ย่อหน้า", { size: 14 });
     d.table({ head: ["ก","ข"], rows: [["1","2"]], widths: [60,40] });
     d.image(dataUrl, { w: 8 });      // กว้าง 8 ซม. สูงตามสัดส่วน
     d.pageBreak();
     DOCXMini.save(d.blob(), "ชื่อไฟล์.docx");

   หมายเหตุเรื่องภาษาไทยใน Word
     ไทยเป็น "complex script" ขนาดตัวอักษรคุมด้วย w:szCs ไม่ใช่ w:sz
     ตัวหนาคุมด้วย w:bCs ไม่ใช่ w:b — ต้องใส่คู่กันทุกที่ ไม่งั้นไทยจะไม่เปลี่ยน
*/
(function () {
  "use strict";

  /* ---------- CRC32 + ZIP แบบไม่บีบอัด ---------- */
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

  function zip(files) {
    var enc = new TextEncoder(), parts = [], central = [], off = 0;
    files.forEach(function (f) {
      var nm = enc.encode(f.name);
      var data = typeof f.data === "string" ? enc.encode(f.data) : f.data;
      var crc = crc32(data);

      var lh = new Uint8Array(30 + nm.length), dv = new DataView(lh.buffer);
      dv.setUint32(0, 0x04034b50, true);
      dv.setUint16(4, 20, true);
      dv.setUint16(6, 0x0800, true);
      dv.setUint16(8, 0, true);
      dv.setUint16(10, 0, true); dv.setUint16(12, 0x2821, true);
      dv.setUint32(14, crc, true);
      dv.setUint32(18, data.length, true);
      dv.setUint32(22, data.length, true);
      dv.setUint16(26, nm.length, true);
      lh.set(nm, 30);
      parts.push(lh, data);

      var ch = new Uint8Array(46 + nm.length), cv = new DataView(ch.buffer);
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
    var eo = new Uint8Array(22), ev = new DataView(eo.buffer);
    ev.setUint32(0, 0x06054b50, true);
    ev.setUint16(8, files.length, true);
    ev.setUint16(10, files.length, true);
    ev.setUint32(12, csize, true);
    ev.setUint32(16, off, true);

    return new Blob(parts.concat(central, [eo]),
      { type: "application/vnd.openxmlformats-officedocument.wordprocessingml.document" });
  }

  function esc(s) {
    return String(s == null ? "" : s)
      .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;").replace(/\r/g, "");
  }

  function b64bytes(b64) {
    var bin = atob(b64), n = bin.length, u = new Uint8Array(n);
    for (var i = 0; i < n; i++) u[i] = bin.charCodeAt(i);
    return u;
  }

  /* ---------- ค่าคงที่หน้ากระดาษ ---------- */
  var EMU_CM = 360000;          // 1 ซม. = 360000 EMU
  var TW_CM = 566.929;          // 1 ซม. = 566.929 twip
  var A4_W = 11906, A4_H = 16838;                 // A4 แนวตั้ง (twip)
  var MAR = { top: 1134, right: 1134, bottom: 1134, left: 1134 };   // 2 ซม.
  var BODY_TW = A4_W - MAR.left - MAR.right;      // ความกว้างพื้นที่พิมพ์
  var FONT = "Tahoma";          // มีอยู่ทุกเครื่อง อ่านไทยได้ชัด

  /* ---------- ตัวสร้างเอกสาร ---------- */
  function doc(opts) {
    opts = opts || {};
    var body = [], media = [], rels = [], relN = 0, imgN = 0, brk = false;

    function rPr(o) {
      o = o || {};
      var sz = Math.round((o.size || 11) * 2);
      var x = '<w:rPr><w:rFonts w:ascii="' + FONT + '" w:hAnsi="' + FONT +
              '" w:cs="' + FONT + '"/>';
      if (o.b) x += "<w:b/><w:bCs/>";
      if (o.i) x += "<w:i/><w:iCs/>";
      if (o.color) x += '<w:color w:val="' + o.color + '"/>';
      x += '<w:sz w:val="' + sz + '"/><w:szCs w:val="' + sz + '"/>';
      return x + "</w:rPr>";
    }

    function pPr(o) {
      o = o || {};
      var x = "<w:pPr>";
      // ขึ้นหน้าใหม่ด้วย pageBreakBefore ของย่อหน้าถัดไป ไม่ใช่ย่อหน้าว่างที่มีแต่ <w:br>
      // แบบหลังทำให้ได้หน้าว่างคั่นเวลาเนื้อหาหน้าก่อนพอดีขอบ
      if (brk) { x += "<w:pageBreakBefore/>"; brk = false; }
      if (o.align) x += '<w:jc w:val="' + o.align + '"/>';
      x += '<w:spacing w:before="' + (o.before == null ? 0 : o.before) +
           '" w:after="' + (o.after == null ? 80 : o.after) +
           '" w:line="' + (o.line || 260) + '" w:lineRule="auto"/>';
      if (o.shade) x += '<w:shd w:val="clear" w:fill="' + o.shade + '"/>';
      if (o.rule) {
        x += '<w:pBdr><w:bottom w:val="single" w:sz="' + (o.ruleSz || 6) +
             '" w:space="2" w:color="' + (o.ruleColor || "1E6B4F") + '"/></w:pBdr>';
      }
      if (o.keep) x += "<w:keepNext/>";
      if (o.ind) x += '<w:ind w:left="' + Math.round(o.ind * TW_CM) + '"/>';
      return x + "</w:pPr>";
    }

    // ย่อหน้าที่มีหลายช่วงคนละสไตล์: p([["ตัวหนา",{b:1}], " ตามด้วยปกติ"])
    function runs(txt, o) {
      var list = Array.isArray(txt) ? txt : [[txt, o]];
      return list.map(function (it) {
        var t = Array.isArray(it) ? it[0] : it;
        var s = Array.isArray(it) ? (it[1] || o) : o;
        // ขึ้นบรรทัดใหม่ในย่อหน้าเดียวกันด้วย <w:br/> ไม่ใช่ตัด \n ทิ้ง
        var body = String(t == null ? "" : t).split("\n").map(function (ln) {
          return '<w:t xml:space="preserve">' + esc(ln) + "</w:t>";
        }).join("<w:br/>");
        return "<w:r>" + rPr(s) + body + "</w:r>";
      }).join("");
    }

    function imgRun(dataUrl, o) {
        o = o || {};
        var m = /^data:image\/(png|jpe?g);base64,(.+)$/i.exec(dataUrl || "");
        if (!m) return "";
        var ext = m[1].toLowerCase() === "png" ? "png" : "jpeg";
        imgN++; relN++;
        var name = "image" + imgN + "." + ext;
        media.push({ name: "word/media/" + name, data: b64bytes(m[2]) });
        rels.push('<Relationship Id="rId' + (100 + relN) + '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Target="media/' + name + '"/>');
        var cx = Math.round((o.w || 7) * EMU_CM);
        var cy = Math.round(cx * (o.ratio || 0.75));
        return '<w:r><w:drawing><wp:inline distT="0" distB="0" distL="0" distR="0">' +
          '<wp:extent cx="' + cx + '" cy="' + cy + '"/>' +
          '<wp:effectExtent l="0" t="0" r="0" b="0"/>' +
          '<wp:docPr id="' + imgN + '" name="รูป ' + imgN + '"/>' +
          "<wp:cNvGraphicFramePr><a:graphicFrameLocks noChangeAspect=\"1\"/></wp:cNvGraphicFramePr>" +
          "<a:graphic><a:graphicData uri=\"http://schemas.openxmlformats.org/drawingml/2006/picture\">" +
          "<pic:pic><pic:nvPicPr><pic:cNvPr id=\"" + imgN + '" name="' + name + '"/>' +
          "<pic:cNvPicPr/></pic:nvPicPr>" +
          '<pic:blipFill><a:blip r:embed="rId' + (100 + relN) + '"/>' +
          "<a:stretch><a:fillRect/></a:stretch></pic:blipFill>" +
          '<pic:spPr><a:xfrm><a:off x="0" y="0"/><a:ext cx="' + cx + '" cy="' + cy + '"/></a:xfrm>' +
          "<a:prstGeom prst=\"rect\"><a:avLst/></a:prstGeom></pic:spPr>" +
          "</pic:pic></a:graphicData></a:graphic></wp:inline></w:drawing></w:r>";
    }

    var api = {
      p: function (txt, o) {
        body.push("<w:p>" + pPr(o) + runs(txt, o) + "</w:p>");
        return api;
      },
      title: function (txt, sub) {
        body.push("<w:p>" + pPr({ after: 40, line: 240 }) +
          runs(txt, { size: 21, b: 1, color: "14452F" }) + "</w:p>");
        if (sub) {
          body.push("<w:p>" + pPr({ after: 200, rule: true, line: 240 }) +
            runs(sub, { size: 10.5, color: "6B7A72" }) + "</w:p>");
        }
        return api;
      },
      h2: function (txt) {
        body.push("<w:p>" + pPr({ before: 260, after: 90, keep: true, rule: true, ruleSz: 4 }) +
          runs(txt, { size: 14.5, b: 1, color: "1E6B4F" }) + "</w:p>");
        return api;
      },
      h3: function (txt) {
        body.push("<w:p>" + pPr({ before: 160, after: 50, keep: true }) +
          runs(txt, { size: 12, b: 1, color: "14452F" }) + "</w:p>");
        return api;
      },
      spacer: function (h) {
        body.push('<w:p><w:pPr><w:spacing w:after="' + (h || 80) +
                  '" w:line="20" w:lineRule="exact"/></w:pPr></w:p>');
        return api;
      },

      /* กล่องตัวเลขสรุป — ตารางแถวเดียว ช่องละตัวเลข */
      cards: function (items) {
        var n = items.length || 1;
        var w = Math.floor(BODY_TW / n);
        var head = items.map(function () {
          return '<w:gridCol w:w="' + w + '"/>';
        }).join("");
        var cells = items.map(function (it) {
          return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' +
            '<w:shd w:val="clear" w:fill="' + (it.bad ? "FBECEC" : it.warn ? "FDF4E3" : "EFF5F1") + '"/>' +
            '<w:tcMar><w:top w:w="110" w:type="dxa"/><w:bottom w:w="110" w:type="dxa"/>' +
            '<w:left w:w="110" w:type="dxa"/><w:right w:w="110" w:type="dxa"/></w:tcMar>' +
            "</w:tcPr>" +
            "<w:p>" + pPr({ after: 20, align: "center", line: 240 }) +
              runs(it.label, { size: 9.5, color: "6B7A72" }) + "</w:p>" +
            "<w:p>" + pPr({ after: 20, align: "center", line: 240 }) +
              runs(String(it.value), { size: 19, b: 1,
                color: it.bad ? "A03028" : it.warn ? "8A5A00" : "14452F" }) + "</w:p>" +
            "<w:p>" + pPr({ after: 0, align: "center", line: 220 }) +
              runs(it.note || "", { size: 9, color: "6B7A72" }) + "</w:p>" +
            "</w:tc>";
        }).join("");
        body.push('<w:tbl><w:tblPr><w:tblW w:w="' + BODY_TW + '" w:type="dxa"/>' +
          '<w:tblBorders>' +
          ['top','left','bottom','right','insideH','insideV'].map(function (s) {
            return '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="FFFFFF"/>';
          }).join("") + "</w:tblBorders>" +
          '<w:tblCellSpacing w:w="30" w:type="dxa"/></w:tblPr>' +
          "<w:tblGrid>" + head + "</w:tblGrid><w:tr>" + cells + "</w:tr></w:tbl>");
        api.spacer(60);
        return api;
      },

      /* ตาราง: head = ชื่อคอลัมน์, rows = แถว, widths = %, align = ["l","r",...] */
      table: function (t) {
        var head = t.head || [], rows = t.rows || [];
        var n = head.length || (rows[0] || []).length;
        var pct = t.widths || [];
        var ws = [];
        for (var i = 0; i < n; i++) {
          ws.push(Math.round(BODY_TW * (pct[i] != null ? pct[i] : 100 / n) / 100));
        }
        var al = t.align || [];
        function cell(v, o) {
          o = o || {};
          var a = o.align === "r" ? "right" : o.align === "c" ? "center" : "left";
          // ช่องหนึ่งใส่รูปก็ได้ — ส่ง { img: dataUrl, w: ซม., ratio: สูง/กว้าง } มาแทนข้อความ
          var inner;
          if (v && typeof v === "object" && !Array.isArray(v) && v.img) {
            inner = "<w:p>" + pPr({ after: 0, align: "center", line: 240 }) +
                    (imgRun(v.img, { w: v.w || 4, ratio: v.ratio }) || "") + "</w:p>";
          } else {
            var lines = (typeof v === "string" && v.indexOf("\n") >= 0) ? v.split("\n") : [v];
            inner = lines.map(function (ln, li) {
              return "<w:p>" + pPr({ after: li === lines.length - 1 ? 0 : 50, align: a, line: 230 }) +
                runs(ln == null ? "" : ln, { size: o.size || 10, b: o.b, color: o.color }) + "</w:p>";
            }).join("");
          }
          return '<w:tc><w:tcPr><w:tcW w:w="' + o.w + '" w:type="dxa"/>' +
            (o.fill ? '<w:shd w:val="clear" w:fill="' + o.fill + '"/>' : "") +
            '<w:tcMar><w:top w:w="70" w:type="dxa"/><w:bottom w:w="70" w:type="dxa"/>' +
            '<w:left w:w="90" w:type="dxa"/><w:right w:w="90" w:type="dxa"/></w:tcMar>' +
            '<w:vAlign w:val="' + (o.vtop ? "top" : "center") + '"/></w:tcPr>' +
            inner + "</w:tc>";
        }
        var bcol = t.border === false ? "FFFFFF" : "D8E2DC";
        var x = '<w:tbl><w:tblPr><w:tblW w:w="' + BODY_TW + '" w:type="dxa"/>' +
          "<w:tblBorders>" +
          ['top','left','bottom','right','insideH','insideV'].map(function (s) {
            return '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="' + bcol + '"/>';
          }).join("") + "</w:tblBorders></w:tblPr><w:tblGrid>" +
          ws.map(function (w) { return '<w:gridCol w:w="' + w + '"/>'; }).join("") +
          "</w:tblGrid>";

        if (head.length) {
          x += '<w:tr><w:trPr><w:tblHeader/></w:trPr>' + head.map(function (h, i) {
            return cell(h, { w: ws[i], b: 1, fill: "1E6B4F", color: "FFFFFF",
                             align: al[i], size: 10 });
          }).join("") + "</w:tr>";
        }
        rows.forEach(function (r, ri) {
          var bad = t.badRow && t.badRow(r, ri);
          x += "<w:tr>" + r.map(function (v, i) {
            return cell(v, { w: ws[i], align: al[i], vtop: t.vtop, size: t.size,
              fill: bad ? "FBECEC" : (t.plain ? null : (ri % 2 ? "F5F8F6" : null)),
              color: bad && i === 0 ? "A03028" : null,
              b: bad && i === 0 });
          }).join("") + "</w:tr>";
        });
        if (brk) { body.push("<w:p>" + pPr({ after: 0, line: 20 }) + "</w:p>"); }
        body.push(x + "</w:tbl>");
        api.spacer(80);
        return api;
      },

      /* รูป: w = ความกว้างเป็นเซนติเมตร, ratio = สูง/กว้าง */
      image: function (dataUrl, o) {
        var r = imgRun(dataUrl, o);
        if (r) body.push("<w:p>" + pPr({ after: 80, align: "center" }) + r + "</w:p>");
        return api;
      },


      /* แถวรูปเรียงข้างกัน พร้อมคำบรรยายใต้รูป */
      imageRow: function (pics) {
        if (!pics.length) return api;
        // รูปเดียวไม่ต้องยืดเต็มหน้า — กว้างสุด 8 ซม. กำลังดูออกและไฟล์ไม่เทอะทะ
        var n = Math.max(pics.length, 2), w = Math.floor(BODY_TW / n);
        var cmW = Math.min(8, (BODY_TW / n) / TW_CM - 0.5);
        var cells = pics.map(function (p) {
          var r = imgRun(p.data, { w: cmW, ratio: p.ratio });
          return '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/>' +
            '<w:tcMar><w:top w:w="60" w:type="dxa"/><w:bottom w:w="60" w:type="dxa"/>' +
            '<w:left w:w="60" w:type="dxa"/><w:right w:w="60" w:type="dxa"/></w:tcMar></w:tcPr>' +
            "<w:p>" + pPr({ after: 30, align: "center", line: 240 }) + (r || "") + "</w:p>" +
            "<w:p>" + pPr({ after: 0, align: "center", line: 220 }) +
              runs(p.caption || "", { size: 8.5, color: "6B7A72" }) + "</w:p></w:tc>";
        }).join("");
        var grid = [], pad = "";
        for (var gi = 0; gi < n; gi++) grid.push('<w:gridCol w:w="' + w + '"/>');
        for (var pi = pics.length; pi < n; pi++) {
          pad += '<w:tc><w:tcPr><w:tcW w:w="' + w + '" w:type="dxa"/></w:tcPr>' +
                 "<w:p>" + pPr({ after: 0 }) + "</w:p></w:tc>";
        }
        body.push('<w:tbl><w:tblPr><w:tblW w:w="' + BODY_TW + '" w:type="dxa"/>' +
          "<w:tblBorders>" +
          ['top','left','bottom','right','insideH','insideV'].map(function (s) {
            return '<w:' + s + ' w:val="single" w:sz="4" w:space="0" w:color="E4ECE7"/>';
          }).join("") + "</w:tblBorders></w:tblPr><w:tblGrid>" +
          grid.join("") + "</w:tblGrid><w:tr>" + cells + pad + "</w:tr></w:tbl>");
        api.spacer(80);
        return api;
      },

      pageBreak: function () { brk = true; return api; },

      blob: function () {
        var foot = opts.footer || "";
        var sect = "<w:sectPr>" +
          '<w:headerReference w:type="default" r:id="rIdHdr"/>' +
          '<w:footerReference w:type="default" r:id="rIdFtr"/>' +
          '<w:pgSz w:w="' + A4_W + '" w:h="' + A4_H + '"/>' +
          '<w:pgMar w:top="' + (MAR.top + 340) + '" w:right="' + MAR.right +
            '" w:bottom="' + (MAR.bottom + 280) + '" w:left="' + MAR.left +
            '" w:header="567" w:footer="567" w:gutter="0"/>' +
          "</w:sectPr>";

        var NS = 'xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main" ' +
          'xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships" ' +
          'xmlns:wp="http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing" ' +
          'xmlns:a="http://schemas.openxmlformats.org/drawingml/2006/main" ' +
          'xmlns:pic="http://schemas.openxmlformats.org/drawingml/2006/picture"';

        var document_ = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          "<w:document " + NS + "><w:body>" + body.join("") + sect + "</w:body></w:document>";

        var hdr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          "<w:hdr " + NS + "><w:p>" +
          pPr({ after: 0, rule: true, ruleSz: 4, ruleColor: "D8E2DC", line: 240 }) +
          runs(opts.header || "", { size: 8.5, color: "8A968F" }) + "</w:p></w:hdr>";

        var ftr = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          "<w:ftr " + NS + "><w:p>" + pPr({ after: 0, align: "center", line: 240 }) +
          runs(foot ? foot + "  ·  หน้า " : "หน้า ", { size: 8.5, color: "8A968F" }) +
          "<w:r>" + rPr({ size: 8.5, color: "8A968F" }) +
          '<w:fldChar w:fldCharType="begin"/></w:r>' +
          "<w:r>" + rPr({ size: 8.5, color: "8A968F" }) +
          '<w:instrText xml:space="preserve"> PAGE </w:instrText></w:r>' +
          "<w:r>" + rPr({ size: 8.5, color: "8A968F" }) +
          '<w:fldChar w:fldCharType="end"/></w:r>' + "</w:p></w:ftr>";

        var docRels = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rIdSty" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>' +
          '<Relationship Id="rIdHdr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/header" Target="header1.xml"/>' +
          '<Relationship Id="rIdFtr" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/footer" Target="footer1.xml"/>' +
          rels.join("") + "</Relationships>";

        var styles = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          "<w:styles " + NS + "><w:docDefaults><w:rPrDefault><w:rPr>" +
          '<w:rFonts w:ascii="' + FONT + '" w:hAnsi="' + FONT + '" w:cs="' + FONT + '"/>' +
          '<w:sz w:val="22"/><w:szCs w:val="22"/><w:lang w:bidi="th-TH"/>' +
          "</w:rPr></w:rPrDefault><w:pPrDefault><w:pPr>" +
          '<w:spacing w:after="80" w:line="260" w:lineRule="auto"/>' +
          "</w:pPr></w:pPrDefault></w:docDefaults>" +
          '<w:style w:type="paragraph" w:default="1" w:styleId="Normal">' +
          '<w:name w:val="Normal"/></w:style></w:styles>';

        var types = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">' +
          '<Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>' +
          '<Default Extension="xml" ContentType="application/xml"/>' +
          '<Default Extension="jpeg" ContentType="image/jpeg"/>' +
          '<Default Extension="png" ContentType="image/png"/>' +
          '<Override PartName="/word/document.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml"/>' +
          '<Override PartName="/word/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.styles+xml"/>' +
          '<Override PartName="/word/header1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml"/>' +
          '<Override PartName="/word/footer1.xml" ContentType="application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml"/>' +
          "</Types>";

        var root = '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>' +
          '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">' +
          '<Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="word/document.xml"/>' +
          "</Relationships>";

        return zip([
          { name: "[Content_Types].xml", data: types },
          { name: "_rels/.rels", data: root },
          { name: "word/document.xml", data: document_ },
          { name: "word/_rels/document.xml.rels", data: docRels },
          { name: "word/styles.xml", data: styles },
          { name: "word/header1.xml", data: hdr },
          { name: "word/footer1.xml", data: ftr }
        ].concat(media));
      }
    };
    return api;
  }

  function save(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click(); a.remove();
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
  }

  // อ่านสัดส่วนรูปจาก data URL เอาไว้คิดความสูงตอนวางลง Word
  function ratio(dataUrl) {
    return new Promise(function (res) {
      var im = new Image();
      im.onload = function () { res(im.naturalHeight / im.naturalWidth || 0.75); };
      im.onerror = function () { res(0.75); };
      im.src = dataUrl;
    });
  }

  window.DOCXMini = { doc: doc, save: save, ratio: ratio };
})();
