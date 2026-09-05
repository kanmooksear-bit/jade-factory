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

  return {
    rpc: rpc, me: me, submit: submit, flush: flush,
    queueLength: function () { return queue().length; },
    onQueueChange: function (fn) { onQueue = fn; },
    configured: configured,
    thDate: thDate, today: today, nf: nf, esc: esc, msg: msg,
    factory: CFG.factory || "โรงงานหยก"
  };
})();
