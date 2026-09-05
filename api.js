/* ตัวกลางคุยกับ Supabase — ทุกหน้าเรียกผ่านไฟล์นี้
   หน้าเว็บเรียกได้เฉพาะฟังก์ชันที่ฐานข้อมูลอนุญาต ตารางถูกล็อกด้วย RLS */
window.JADE = (function () {
  "use strict";

  var CFG = window.JADE_CONFIG || {};
  var BASE = String(CFG.url || "").replace(/\/+$/, "") + "/rest/v1/rpc/";
  var KEY = CFG.anonKey || "";
  var QKEY = "jade.queue";

  function configured() {
    return /^https:\/\/.+\.supabase\.co$/.test(String(CFG.url || "")) && KEY.length > 20;
  }

  function rpc(fn, args) {
    if (!configured()) {
      return Promise.reject(new Error("ยังไม่ได้ตั้งค่า config.js — ใส่ PROJECT URL และ ANON KEY ก่อน"));
    }
    return fetch(BASE + fn, {
      method: "POST",
      headers: { apikey: KEY, Authorization: "Bearer " + KEY, "Content-Type": "application/json" },
      body: JSON.stringify(args || {})
    }).then(function (res) {
      return res.text().then(function (txt) {
        var data = null;
        try { data = txt ? JSON.parse(txt) : null; } catch (e) { data = null; }
        if (!res.ok) {
          var m = (data && (data.message || data.hint || data.details)) || ("เชื่อมต่อไม่สำเร็จ (" + res.status + ")");
          throw new Error(m);
        }
        return data;
      });
    });
  }

  /* ---------- จำผู้ใช้ในเครื่องนี้ ---------- */
  var me = {
    get: function () {
      try {
        return { name: localStorage.getItem("jade.name") || "", pin: localStorage.getItem("jade.pin") || "",
                 role: localStorage.getItem("jade.role") || "worker" };
      } catch (e) { return { name: "", pin: "", role: "worker" }; }
    },
    set: function (name, pin, role) {
      try {
        localStorage.setItem("jade.name", name);
        localStorage.setItem("jade.pin", pin);
        localStorage.setItem("jade.role", role || "worker");
      } catch (e) {}
    },
    clear: function () {
      try { ["jade.name", "jade.pin", "jade.role"].forEach(function (k) { localStorage.removeItem(k); }); } catch (e) {}
    },
    ok: function () { var m = me.get(); return !!(m.name && m.pin); }
  };

  /* ---------- คิวออฟไลน์: เก็บไว้ในเครื่อง ส่งเองเมื่อเน็ตกลับมา ---------- */
  function queue() {
    try { return JSON.parse(localStorage.getItem(QKEY) || "[]"); } catch (e) { return []; }
  }
  function setQueue(q) {
    try { localStorage.setItem(QKEY, JSON.stringify(q)); } catch (e) {}
    if (typeof onQueue === "function") onQueue(q.length);
  }
  var onQueue = null;

  function submit(args) {
    return rpc("app_submit", args).catch(function (err) {
      if (!navigator.onLine || /Failed to fetch|NetworkError|เชื่อมต่อไม่สำเร็จ/i.test(err.message)) {
        var q = queue();
        q.push(args);
        setQueue(q);
        return { ok: true, queued: true };
      }
      throw err;
    });
  }

  function flush() {
    var q = queue();
    if (!q.length || !navigator.onLine) return Promise.resolve(0);
    var sent = 0;
    return q.reduce(function (chain, item) {
      return chain.then(function () {
        return rpc("app_submit", item).then(function () { sent++; },
          function (err) { if (!/Failed to fetch|NetworkError/i.test(err.message)) sent++; });
      });
    }, Promise.resolve()).then(function () {
      setQueue(q.slice(sent));
      return sent;
    });
  }

  window.addEventListener("online", function () { flush(); });

  /* ---------- ตัวช่วยทั่วไป ---------- */
  var TH_M = ["ม.ค.","ก.พ.","มี.ค.","เม.ย.","พ.ค.","มิ.ย.","ก.ค.","ส.ค.","ก.ย.","ต.ค.","พ.ย.","ธ.ค."];
  function thDate(iso, withYear) {
    if (!iso) return "—";
    var p = String(iso).slice(0, 10).split("-");
    if (p.length !== 3) return String(iso);
    var s = Number(p[2]) + " " + TH_M[Number(p[1]) - 1];
    return withYear === false ? s : s + " " + (Number(p[0]) + 543);
  }
  function today() {
    var d = new Date();
    return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
  }
  var nf = new Intl.NumberFormat("th-TH");
  function esc(s) {
    return String(s == null ? "" : s).replace(/&/g, "&amp;").replace(/</g, "&lt;")
      .replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  }
  function msg(el, text, kind) {
    if (!el) return;
    if (!text) { el.hidden = true; return; }
    el.hidden = false;
    el.className = "msg " + (kind || "ok");
    el.textContent = text;
  }

  /* ---------- ช่องวันที่แบบไทย วว/ดด/ปปปป (พ.ศ.) ----------
     เปลี่ยนช่อง <input type="date"> ทุกช่องให้พิมพ์เป็น วว/ดด/ปปปป
     โค้ดส่วนอื่นยังอ่าน/เขียน .value เป็น YYYY-MM-DD เหมือนเดิม   */

  var TH_FULL = ["มกราคม","กุมภาพันธ์","มีนาคม","เมษายน","พฤษภาคม","มิถุนายน",
                 "กรกฎาคม","สิงหาคม","กันยายน","ตุลาคม","พฤศจิกายน","ธันวาคม"];
  var NATIVE = Object.getOwnPropertyDescriptor(HTMLInputElement.prototype, "value");

  function pad2(n) { return String(n).padStart(2, "0"); }

  function allDigits(s) {
    if (!s.length) return false;
    for (var i = 0; i < s.length; i++) { var c = s.charAt(i); if (c < "0" || c > "9") return false; }
    return true;
  }
  function onlyDigits(s, max) {
    var out = "";
    for (var i = 0; i < s.length && out.length < max; i++) {
      var c = s.charAt(i);
      if (c >= "0" && c <= "9") out += c;
    }
    return out;
  }

  // "05/09/2569" (หรือ ค.ศ.) -> "2026-09-05" ; ไม่ถูกต้องคืน ""
  function parseThai(s) {
    var p = String(s || "").split("/");
    if (p.length !== 3) return "";
    var ds = p[0].trim(), ms = p[1].trim(), ys = p[2].trim();
    if (!allDigits(ds) || !allDigits(ms) || !allDigits(ys)) return "";
    if (ds.length > 2 || ms.length > 2 || ys.length !== 4) return "";
    var d = +ds, mo = +ms, y = +ys;
    if (y > 2400) y -= 543;
    if (mo < 1 || mo > 12 || d < 1 || d > 31) return "";
    var dt = new Date(y, mo - 1, d);
    if (dt.getFullYear() !== y || dt.getMonth() !== mo - 1 || dt.getDate() !== d) return "";
    return y + "-" + pad2(mo) + "-" + pad2(d);
  }

  // "2026-09-05" -> "05/09/2569"
  function formatThai(iso) {
    var p = String(iso || "").slice(0, 10).split("-");
    if (p.length !== 3) return "";
    return pad2(+p[2]) + "/" + pad2(+p[1]) + "/" + (+p[0] + 543);
  }

  function thLong(iso) {
    var p = String(iso || "").slice(0, 10).split("-");
    if (p.length !== 3) return "";
    return +p[2] + " " + TH_FULL[+p[1] - 1] + " " + (+p[0] + 543);
  }

  function enhanceDate(inp) {
    if (!inp || inp.dataset.thaiDate) return;
    inp.dataset.thaiDate = "1";
    inp.type = "text";
    inp.setAttribute("inputmode", "numeric");
    inp.setAttribute("autocomplete", "off");
    inp.setAttribute("maxlength", "10");
    inp.placeholder = "วว/ดด/ปปปป";

    var iso = "";

    var row = document.createElement("div");
    row.className = "daterow";
    inp.parentNode.insertBefore(row, inp);
    row.appendChild(inp);

    var btn = document.createElement("button");
    btn.type = "button";
    btn.className = "btn btn-sm daybtn";
    btn.textContent = "วันนี้";
    row.appendChild(btn);

    var hint = document.createElement("span");
    hint.className = "datehint";
    row.parentNode.insertBefore(hint, row.nextSibling);

    function paint() {
      var raw = NATIVE.get.call(inp);
      if (iso) { hint.textContent = thLong(iso); hint.className = "datehint"; }
      else if (raw) { hint.textContent = "ยังไม่ครบ — พิมพ์เป็น วว/ดด/ปปปป เช่น 05/09/2569"; hint.className = "datehint bad"; }
      else { hint.textContent = ""; hint.className = "datehint"; }
    }

    Object.defineProperty(inp, "value", {
      configurable: true,
      get: function () { return iso; },
      set: function (v) {
        iso = v ? String(v).slice(0, 10) : "";
        NATIVE.set.call(inp, iso ? formatThai(iso) : "");
        paint();
      }
    });

    inp.addEventListener("input", function () {
      var digits = onlyDigits(NATIVE.get.call(inp), 8);
      var out = digits.slice(0, 2);
      if (digits.length > 2) out += "/" + digits.slice(2, 4);
      if (digits.length > 4) out += "/" + digits.slice(4, 8);
      NATIVE.set.call(inp, out);
      iso = parseThai(out);
      paint();
    });

    inp.addEventListener("blur", function () {
      if (iso) NATIVE.set.call(inp, formatThai(iso));
      paint();
    });

    btn.addEventListener("click", function () {
      inp.value = today();
      inp.dispatchEvent(new Event("change", { bubbles: true }));
    });

    paint();
  }

  function enhanceDates(root) {
    var list = (root || document).querySelectorAll('input[type="date"]');
    Array.prototype.forEach.call(list, enhanceDate);
  }

  if (document.readyState === "loading") {
    document.addEventListener("DOMContentLoaded", function () { enhanceDates(); });
  } else {
    enhanceDates();
  }

  return {
    rpc: rpc, me: me, submit: submit, flush: flush,
    queueLength: function () { return queue().length; },
    onQueueChange: function (fn) { onQueue = fn; },
    configured: configured,
    thDate: thDate, today: today, nf: nf, esc: esc, msg: msg,
    parseThai: parseThai, formatThai: formatThai, thLong: thLong, enhanceDates: enhanceDates,
    factory: CFG.factory || "โรงงานหยก"
  };
})();
