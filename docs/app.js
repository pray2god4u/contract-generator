/* 계약서 초안 생성기 — 문서 파일 생성 엔진 포함 (DOCX·HWPX·TXT) */
(function () {
  "use strict";
  var G = typeof globalThis !== "undefined" ? globalThis : window;
  var IDX = G.CAT_INDEX;
  var TPL = (G.CAT_TEMPLATES || {}).templates || [];
  var docxLib = typeof window !== "undefined" ? window.docx : require("./vendor/docx.umd.js");
  var JSZipLib = typeof window !== "undefined" ? window.JSZip : require("./vendor/jszip.min.js");
  var ASSETS = G.HWPX_ASSETS;

  var tplById = {};
  TPL.forEach(function (t) { tplById[t.id] = t; });

  var $ = function (id) { return document.getElementById(id); };
  var state = { selected: null, skeleton: null, vals: null, model: null, filled: false };

  /* ---------- HWPX 정적 XML ---------- */
  var CONTAINER_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n' +
    '<ocf:container xmlns:ocf="urn:oasis:names:tc:opendocument:xmlns:container" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf">' +
    '<ocf:rootfiles><ocf:rootfile full-path="Contents/content.hpf" media-type="application/hwpml-package+xml"/></ocf:rootfiles></ocf:container>';
  var HPF_XML = '<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>\n' +
    '<opf:package xmlns:opf="http://www.idpf.org/2007/opf/" xmlns:hpf="http://www.hancom.co.kr/schema/2011/hpf" xmlns:hh="http://www.hancom.co.kr/hwpml/2011/head">' +
    '<opf:metadata><opf:meta name="generator" content="contract-generator"/><opf:meta name="kordoc-layout" content="default"/></opf:metadata>' +
    '<opf:manifest><opf:item id="header" href="Contents/header.xml" media-type="application/xml"/>' +
    '<opf:item id="section0" href="Contents/section0.xml" media-type="application/xml"/></opf:manifest>' +
    '<opf:spine><opf:itemref idref="header" linear="no"/><opf:itemref idref="section0" linear="yes"/></opf:spine></opf:package>';

  /* ---------- 초기화 ---------- */
  function init() {
    var s = IDX.stats || {};
    var metrics = [
      { b: fmt(s.total || IDX.records.length), span: "색인된 서식 문서" },
      { b: Object.keys(IDX.categories).length, span: "대분류" },
      { b: TPL.length, span: "구조 템플릿" },
      { b: "3", span: "저장 형식 (DOCX·HWPX·TXT)" }
    ];
    $("metrics").innerHTML = metrics.map(function (m) {
      return '<div class="metric"><b>' + m.b + "</b><span>" + m.span + "</span></div>";
    }).join("");

    var mcSel = $("mc");
    Object.keys(IDX.categories).sort().forEach(function (mc) {
      var o = document.createElement("option");
      o.value = mc; o.textContent = mc + " (" + sumSc(mc) + ")";
      mcSel.appendChild(o);
    });

    ["q", "mc", "sc", "ds"].forEach(function (id) {
      $(id).addEventListener("input", render);
      $(id).addEventListener("change", render);
    });
    $("btn-generate").addEventListener("click", generate);
    $("btn-ai").addEventListener("click", aiGenerate);
    $("btn-save").addEventListener("click", saveDoc);
    $("btn-connect-folder").addEventListener("click", connectFolder);
    $("btn-copy").addEventListener("click", copyDraft);
    $("btn-print").addEventListener("click", function () { window.print(); });
    restoreFolder();
    restoreAiCfg();
    render();
  }

  /* ---------- 작업 폴더 연결 (OneAI 작업 폴더 지정) ---------- */
  var dirHandle = null;
  function idbOpen() {
    return new Promise(function (res, rej) {
      var r = indexedDB.open("contract-generator", 1);
      r.onupgradeneeded = function () { r.result.createObjectStore("handles"); };
      r.onsuccess = function () { res(r.result); };
      r.onerror = function () { rej(r.error); };
    });
  }
  function idbSet(key, val) {
    return idbOpen().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction("handles", "readwrite");
        tx.objectStore("handles").put(val, key);
        tx.oncomplete = function () { res(); };
        tx.onerror = function () { rej(tx.error); };
      });
    });
  }
  function idbGet(key) {
    return idbOpen().then(function (db) {
      return new Promise(function (res, rej) {
        var tx = db.transaction("handles", "readonly");
        var rq = tx.objectStore("handles").get(key);
        rq.onsuccess = function () { res(rq.result); };
        rq.onerror = function () { rej(rq.error); };
      });
    });
  }
  function updateFolderStatus() {
    var el = $("save-folder-status");
    if (dirHandle) el.textContent = "📁 저장 폴더 연결됨: " + dirHandle.name + " — 생성한 문서가 이 폴더에 바로 저장됩니다.";
    else el.textContent = "저장 폴더 미연결 — 「저장 폴더 연결」로 OneAI 작업 폴더를 지정하면 문서가 그곳에 저장됩니다.";
  }
  function connectFolder() {
    if (!window.showDirectoryPicker) {
      toast("이 브라우저는 폴더 연결을 지원하지 않습니다. 저장 시 다운로드로 대체됩니다.");
      return;
    }
    window.showDirectoryPicker({ mode: "readwrite" }).then(function (h) {
      dirHandle = h;
      return idbSet("workdir", h).then(function () {
        updateFolderStatus();
        toast("작업 폴더가 연결되었습니다: " + h.name);
      });
    }).catch(function (e) {
      if (e && e.name !== "AbortError") toast("폴더 연결 실패: " + e.message);
    });
  }
  function restoreFolder() {
    if (!window.showDirectoryPicker || !indexedDB) { updateFolderStatus(); return; }
    idbGet("workdir").then(function (h) {
      if (h) { dirHandle = h; }
      updateFolderStatus();
    }).catch(function () { updateFolderStatus(); });
  }
  function ensureDirPermission() {
    if (!dirHandle) return Promise.resolve(false);
    return dirHandle.queryPermission({ mode: "readwrite" }).then(function (p) {
      if (p === "granted") return true;
      return dirHandle.requestPermission({ mode: "readwrite" }).then(function (p2) { return p2 === "granted"; });
    });
  }
  function saveToDir(blob, filename) {
    return ensureDirPermission().then(function (ok) {
      if (!ok) return false;
      return dirHandle.getFileHandle(filename, { create: true }).then(function (fh) {
        return fh.createWritable().then(function (w) {
          return w.write(blob).then(function () { return w.close(); });
        });
      }).then(function () { return true; });
    }).catch(function () { return false; });
  }

  /* ---------- AI 연결 설정 (사내 GenAI) ---------- */
  var AI_DEFAULTS = {
    url: "https://genai.douzone.com/api/etos-120b/v1/chat/completions",
    model: "dsm-120b",
    temp: "0.3",
    maxtok: "4096"
  };
  function aiCfg() {
    return {
      url: ($("ai-url").value || AI_DEFAULTS.url).trim(),
      model: ($("ai-model").value || AI_DEFAULTS.model).trim(),
      temp: parseFloat($("ai-temp").value || AI_DEFAULTS.temp) || 0.3,
      maxtok: parseInt($("ai-maxtok").value || AI_DEFAULTS.maxtok, 10) || 4096
    };
  }
  function saveAiCfg() {
    try {
      localStorage.setItem("cg_ai_cfg", JSON.stringify({
        url: $("ai-url").value, model: $("ai-model").value,
        temp: $("ai-temp").value, maxtok: $("ai-maxtok").value
      }));
    } catch (e) { /* 사설모드 등 */ }
  }
  function restoreAiCfg() {
    var el = $("ai-url");
    if (!el) return;
    el.value = AI_DEFAULTS.url;
    $("ai-model").value = AI_DEFAULTS.model;
    $("ai-temp").value = AI_DEFAULTS.temp;
    $("ai-maxtok").value = AI_DEFAULTS.maxtok;
    try {
      var saved = JSON.parse(localStorage.getItem("cg_ai_cfg") || "{}");
      if (saved.url) el.value = saved.url;
      if (saved.model) $("ai-model").value = saved.model;
      if (saved.temp) $("ai-temp").value = saved.temp;
      if (saved.maxtok) $("ai-maxtok").value = saved.maxtok;
    } catch (e) { /* 무시 */ }
    ["ai-url", "ai-model", "ai-temp", "ai-maxtok"].forEach(function (id) {
      $(id).addEventListener("change", saveAiCfg);
    });
  }

  /* ---------- AI 초안 생성 (사내 GenAI) ---------- */
  function buildAiPrompt(rec, secs, vals) {
    var structure = secs.map(function (s) {
      var art = parseInt(s.art, 10);
      var line = "- [" + s.ft + "]";
      if (!isNaN(art) && art > 0) line += " 제" + art + "조";
      if (s.hint) line += " " + s.hint;
      return line;
    }).join("\n");
    var sys = "당신은 대한민국 법률 문서 작성 전문가입니다. 주어진 문서 구조와 기본 정보를 바탕으로 실제 사용 가능한 수준의 계약서/법률 문서 초안을 한국어로 작성합니다.\n" +
      "규칙:\n" +
      "1. 첫 줄에 문서 제목만 출력하고, 각 조항은 '제N조 (조문제목)' 형태의 줄로 시작하며 그 아래 본문 문단을 작성합니다.\n" +
      "2. 조문 번호는 주어진 구조의 순서를 따르며 누락 없이 재번호화합니다. 항 번호는 ①②③을 사용합니다.\n" +
      "3. 계약서라면 서두(갑·을 표시와 계약 체결 문언)와 말미(계약일, 서명란)를 반드시 포함합니다.\n" +
      "4. 입력된 기본 정보(당사자, 금액, 기간 등)를 문장에 정확히 반영하고, 미입력값은 '____' 빈칸으로 둡니다.\n" +
      "5. 특약사항 입력이 있으면 마지막 조로 포함합니다.\n" +
      "6. 설명·주석·사족 없이 문서 본문만 출력합니다. 마크다운 기호(#, *, -)를 쓰지 않습니다.";
    var user = "문서 유형: " + (rec.ds === "71835" ? "계약 외 법률 문서" : "계약서") + "\n" +
      "서식명: " + rec.t + "\n" +
      "분류: " + rec.mc + " > " + rec.sc + " > " + rec.dc + "\n" +
      "관련법령: " + (rec.law || "미명시") + "\n\n" +
      "기본 정보:\n" +
      "- 갑(발주자/매도인): " + vals.a + " (대표자 " + vals.ra + ")\n" +
      "- 을(수급자/매수인): " + vals.b + " (대표자 " + vals.rb + ")\n" +
      "- 계약금액: " + vals.amount + "\n" +
      "- 계약기간: " + vals.period + "\n" +
      "- 계약일: " + koDate(vals.date) + " / 체결 장소: " + vals.place + "\n" +
      "- 특약사항: " + (vals.extra || "없음") + "\n\n" +
      "문서 구조(이 순서와 역할을 따르세요):\n" + structure;
    return [{ role: "system", content: sys }, { role: "user", content: user }];
  }

  function callLLM(cfg, messages, onDelta) {
    return fetch(cfg.url, {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        model: cfg.model,
        messages: messages,
        max_tokens: cfg.maxtok,
        temperature: cfg.temp,
        top_p: 0.8,
        stream: true,
        reasoning_effort: "low"
      })
    }).then(function (res) {
      if (!res.ok) {
        return res.text().then(function (t) {
          throw new Error("HTTP " + res.status + " " + String(t).slice(0, 120));
        });
      }
      var reader = res.body.getReader();
      var dec = new TextDecoder("utf-8");
      var buf = "", full = "";
      function pump() {
        return reader.read().then(function (r) {
          if (r.done) return full;
          buf += dec.decode(r.value, { stream: true });
          var lines = buf.split("\n");
          buf = lines.pop();
          lines.forEach(function (ln) {
            ln = ln.trim();
            if (ln.indexOf("data:") !== 0) return;
            var payload = ln.slice(5).trim();
            if (!payload || payload === "[DONE]") return;
            try {
              var j = JSON.parse(payload);
              var d = j.choices && j.choices[0] && j.choices[0].delta;
              if (d && d.content) { full += d.content; if (onDelta) onDelta(full); }
            } catch (e) { /* 아직 온전하지 않은 조각 무시 */ }
          });
          return pump();
        });
      }
      return pump();
    });
  }

  function parseAiDraft(text, rec) {
    var lines = String(text).replace(/\r/g, "").split("\n");
    var paras = [];
    var title = rec.t;
    var i = 0;
    while (i < lines.length && !lines[i].trim()) i++;
    if (i < lines.length) { title = lines[i].trim().replace(/^#+\s*/, ""); i++; }
    for (; i < lines.length; i++) {
      var ln = lines[i].trim();
      if (!ln) continue;
      ln = ln.replace(/^#+\s*/, "").replace(/^\*\*|\*\*$/g, "");
      if (/^제\d+조/.test(ln)) paras.push({ kind: "article", text: ln });
      else paras.push({ kind: "body", text: ln });
    }
    return { title: title, paras: paras };
  }

  function aiGenerate() {
    if (!state.selected) { toast("먼저 왼쪽에서 서식을 선택하세요."); return; }
    var cfg = aiCfg();
    state.vals = collectVals();
    var rec = state.selected;
    if (!state.skeleton) state.skeleton = tplById[rec.id] ? tplById[rec.id].sections : genericSections(rec);
    var messages = buildAiPrompt(rec, state.skeleton, state.vals);
    var btn = $("btn-ai");
    btn.disabled = true;
    $("doc").innerHTML = "<h3>AI 초안 생성 중…</h3><p class='doc-sub'>사내 GenAI(" + esc(cfg.model) +
      ")가 조문을 작성하고 있어요. 최대 수십 초가 걸릴 수 있습니다.</p><pre class='ai-stream' id='ai-stream'></pre>";
    var started = Date.now();
    callLLM(cfg, messages, function (partial) {
      var el = $("ai-stream");
      if (el) el.textContent = partial.slice(-1200);
    }).then(function (full) {
      btn.disabled = false;
      var parsed = parseAiDraft(full, rec);
      if (!parsed.paras.length) throw new Error("AI 응답에서 조문을 찾지 못했습니다.");
      if (!parsed.title || parsed.title.length < 5 || /^(계약서|계약서\s?초안|문서|draft)$/i.test(parsed.title)) {
        parsed.title = rec.t.replace(/_\d+$/, "");
      }
      state.filled = true;
      state.model = {
        title: parsed.title,
        metaLine: "분류: " + rec.mc + " › " + rec.sc + " › " + rec.dc +
          " · 관련법령: " + (rec.law || "미명시") + " · 출처: AI 허브(" + rec.ds + ", aihub.or.kr) · AI 초안(" +
          cfg.model + ", " + Math.round((Date.now() - started) / 1000) + "초)",
        paras: parsed.paras,
        isContract: rec.ds !== "71835"
      };
      renderModel();
      toast("AI 초안 생성 완료! 문서 파일로 저장할 수 있어요.");
    }).catch(function (e) {
      btn.disabled = false;
      $("doc").innerHTML = "<h3>AI 생성 실패</h3><p class='hint'>" + esc(e.message) +
        "</p><p class='hint'>사내망에서 접속했는지, AI 연결 설정의 요청 URL이 맞는지 확인하세요. 「초안 생성」(규칙 기반)은 계속 사용할 수 있습니다.</p>";
    });
  }

  function sumSc(mc) {
    var scs = IDX.categories[mc] || {};
    return Object.keys(scs).reduce(function (a, k) { return a + scs[k]; }, 0);
  }
  function fmt(n) { return String(n).replace(/\B(?=(\d{3})+(?!\d))/g, ","); }

  /* ---------- 검색·목록 ---------- */
  function filtered() {
    var q = $("q").value.trim().toLowerCase();
    var mc = $("mc").value, sc = $("sc").value, ds = $("ds").value;
    return IDX.records.filter(function (r) {
      if (mc && r.mc !== mc) return false;
      if (sc && r.sc !== sc) return false;
      if (ds && r.ds !== ds) return false;
      if (q) {
        var hay = (r.t + " " + r.law + " " + r.dc + " " + r.sc).toLowerCase();
        if (hay.indexOf(q) === -1) return false;
      }
      return true;
    });
  }

  function render() {
    var mc = $("mc").value;
    var scSel = $("sc");
    scSel.innerHTML = '<option value="">전체</option>';
    if (mc) {
      Object.keys(IDX.categories[mc] || {}).sort().forEach(function (sc) {
        var o = document.createElement("option");
        o.value = sc; o.textContent = sc + " (" + IDX.categories[mc][sc] + ")";
        scSel.appendChild(o);
      });
    }
    var rows = filtered();
    $("count").textContent = "총 " + fmt(rows.length) + "건" + (rows.length > 100 ? " · 100건만 표시" : "");
    var list = rows.slice(0, 100).map(function (r) {
      var hasTpl = tplById[r.id] ? " · <span class='badge purple'>구조템플릿</span>" : "";
      return '<div class="result-item" data-id="' + esc(r.id) + '"><b>' + esc(r.t) + "</b>" +
        "<div class='meta'><span class='badge'>" + esc(r.ds) + "</span>" +
        "<span>" + esc(r.mc) + " › " + esc(r.sc) + " › " + esc(r.dc) + "</span>" +
        "<span>" + esc(r.law || "법령 미명시") + " · " + r.nsec + "개 구획</span>" + hasTpl + "</div></div>";
    }).join("");
    $("results").innerHTML = list || "<p class='count'>검색 결과가 없습니다.</p>";
    Array.prototype.forEach.call($("results").children, function (el) {
      el.addEventListener("click", function () { select(el.getAttribute("data-id")); });
    });
  }

  function esc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&#39;" }[c];
    });
  }
  function xmlEsc(s) {
    return String(s == null ? "" : s).replace(/[&<>"']/g, function (c) {
      return { "&": "&amp;", "<": "&lt;", ">": "&gt;", '"': "&quot;", "'": "&apos;" }[c];
    });
  }

  /* ---------- 선택·입력 폼 ---------- */
  function select(id) {
    state.selected = null;
    for (var i = 0; i < IDX.records.length; i++) {
      if (IDX.records[i].id === id) { state.selected = IDX.records[i]; break; }
    }
    Array.prototype.forEach.call($("results").children, function (el) {
      el.classList.toggle("active", el.getAttribute("data-id") === id);
    });
    $("fields-panel").hidden = false;
    var today = new Date().toISOString().slice(0, 10);
    var fields = [
      ["f-party-a", "갑 (발주자·매도인 등)"], ["f-rep-a", "갑 대표자"],
      ["f-party-b", "을 (수급자·매수인 등)"], ["f-rep-b", "을 대표자"],
      ["f-amount", "계약금액 (원)"], ["f-period", "계약기간"],
      ["f-date", "계약일"], ["f-place", "계약 체결 장소"]
    ];
    $("field-grid").innerHTML = fields.map(function (f) {
      var v = f[0] === "f-date" ? today : "";
      return "<div><label for='" + f[0] + "'>" + f[1] + "</label>" +
        "<input id='" + f[0] + "' value=\"" + esc(v) + "\"></div>";
    }).join("");
    $("field-grid").innerHTML +=
      "<div style='grid-column:1/-1'><label for='f-extra'>특약사항 (선택)</label>" +
      "<textarea id='f-extra' rows='3' placeholder='추가로 정할 조건을 쉼표로 구분해 입력'></textarea></div>";
    showSkeleton();
  }

  function showSkeleton() {
    state.skeleton = tplById[state.selected.id] ? tplById[state.selected.id].sections : genericSections(state.selected);
    state.filled = false;
    state.model = buildDocModel(state.selected, state.skeleton, null, false);
    renderModel();
  }

  function genericSections(r) {
    var isNonContract = r.ds === "71835";
    if (isNonContract) {
      return [
        { ft: "제목", art: null, hint: "문서 제목 (예: 내용증명, 위임장)" },
        { ft: "본문", art: 1, hint: "발신인·수신인 및 문서의 목적을 기술" },
        { ft: "본문", art: 2, hint: "요청 내용 또는 위임 내용을 구체적으로 기술" },
        { ft: "본문", art: 3, hint: "근거 사실·법령·기한을 기술" },
        { ft: "본문", art: 4, hint: "후속 조치 및 회신 요청" },
        { ft: "날짜·서명", art: null, hint: "작성일과 발신인 성명·주소·연락처" }
      ];
    }
    return [
      { ft: "제목", art: null, hint: "계약서 제목" },
      { ft: "전문", art: null, hint: "갑·을이 아래와 같이 계약을 체결한다" },
      { ft: "본문", art: 1, hint: "계약의 목적" },
      { ft: "본문", art: 2, hint: "계약 대상물(용역)의 표시 및 범위" },
      { ft: "본문", art: 3, hint: "계약금액 및 지급 방법" },
      { ft: "본문", art: 4, hint: "계약 이행 및 납품(완료) 기한" },
      { ft: "본문", art: 5, hint: "계약의 해제·해지 사유" },
      { ft: "본문", art: 6, hint: "손해배상 및 분쟁 해결" },
      { ft: "본문", art: 7, hint: "기타 사항 (특약)" },
      { ft: "날짜·서명", art: null, hint: "계약일과 갑·을 서명 날인" }
    ];
  }

  /* ---------- 문서 모델 (모든 형식의 공급원) ---------- */
  function fieldVal(id) { var el = $(id); return el ? el.value.trim() : ""; }

  function collectVals() {
    return {
      a: fieldVal("f-party-a") || "갑(상호)",
      ra: fieldVal("f-rep-a") || "(대표자)",
      b: fieldVal("f-party-b") || "을(상호)",
      rb: fieldVal("f-rep-b") || "(대표자)",
      amount: fieldVal("f-amount") || "금 ________원",
      period: fieldVal("f-period") || "____년 __월 __일부터 ____년 __월 __일까지",
      date: fieldVal("f-date") || "____년 __월 __일",
      place: fieldVal("f-place") || "________________",
      extra: fieldVal("f-extra")
    };
  }

  function koDate(d) {
    var m = /^(\d{4})-(\d{2})-(\d{2})$/.exec(d);
    return m ? m[1] + "년 " + parseInt(m[2], 10) + "월 " + parseInt(m[3], 10) + "일" : d;
  }

  function fillText(hint, v) {
    if (/금액|대금/.test(hint)) return "계약금액은 " + v.amount + "으로 하며, 지급 방법과 시기는 별도 합의에 따른다.";
    if (/기간|기한/.test(hint)) return "계약기간은 " + v.period + (v.period.indexOf("까지") > -1 ? "로 한다." : "으로 한다.");
    if (/목적/.test(hint)) return "본 계약은 갑·을 상호 간의 권리와 의무를 명확히 하기 위하여 체결한다.";
    if (/해제|해지/.test(hint)) return "일방이 채무를 이행하지 아니하는 경우 상대방은 상당한 기간을 정하여 최고한 후 계약을 해제(해지)할 수 있다.";
    if (/손해배상|분쟁/.test(hint)) return "계약 관련 분쟁이 발생할 경우 당사자 간 협의로 해결하며, 협의가 이루어지지 않으면 관할 법원에 제소할 수 있다.";
    if (/대상|범위|표시/.test(hint)) return "계약 대상의 품목·규격·수량 등은 별도 명세서에 따른다.";
    if (/이행|납품|완료/.test(hint)) return "을은 계약기간 내에 계약 내용을 완료하여 갑에게 납품(통보)하여야 한다.";
    if (/기타|특약/.test(hint)) return "본 계약에 정하지 아니한 사항은 민법 등 관련 법령과 관례에 따른다.";
    return hint + " ______________________";
  }

  function buildDocModel(rec, secs, vals, filled) {
    var v = vals || {};
    var m = {
      title: rec.t,
      metaLine: "분류: " + rec.mc + " › " + rec.sc + " › " + rec.dc +
        " · 관련법령: " + (rec.law || "미명시") + " · 출처: AI 허브(" + rec.ds + ", aihub.or.kr)",
      paras: [],
      isContract: rec.ds !== "71835"
    };
var artNo = 0, lastArtKey = null, itemNo = 0;
    secs.forEach(function (s) {
      if (s.ft === "제목") return;
      var artN = parseInt(s.art, 10);
      var hasArt = !isNaN(artN) && artN > 0;
      if (filled && !hasArt) return; /* 전문·서명 등 비조문 힌트는 서두/말미 블록이 대체 */
      if (hasArt && s.art !== lastArtKey) {
        artNo++;
        m.paras.push({ kind: "article", text: "제" + artNo + "조 (" + artTitleOf(s) + ")" });
        lastArtKey = s.art; itemNo = 0;
      }
      if (!hasArt) lastArtKey = null;
      itemNo++;
      var text = filled ? fillText(s.hint, v) : s.hint;
      var prefix = hasArt ? circled(itemNo) + " " : "";
      m.paras.push({ kind: "body", text: prefix + text, indent: hasArt });
    });
    if (filled) {
      var extraItems = (v.extra || "").split(/[,\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
      if (extraItems.length) {
        artNo++;
        m.paras.push({ kind: "article", text: "제" + artNo + "조 (특약사항)" });
        extraItems.forEach(function (t, i) {
          m.paras.push({ kind: "body", text: circled(i + 1) + " " + t, indent: true });
        });
      }
      if (m.isContract) {
        m.paras.unshift({ kind: "body", text: "계약서 번호: 제 ______ 호" });
        m.paras.splice(1, 0, {
          kind: "body",
          text: v.a + "(대표자 " + v.ra + ", 이하 \"갑\")과 " + v.b + "(대표자 " + v.rb +
            ", 이하 \"을\")은 " + koDate(v.date) + " " + v.place + "에서 아래와 같이 계약을 체결한다."
        });
        m.paras.push({ kind: "body", text: "이 계약의 성립을 증명하기 위하여 계약서 2부를 작성하고 갑·을이 각각 서명 날인 후 1부씩 보관한다." });
        m.paras.push({ kind: "body", text: koDate(v.date), right: true });
        m.paras.push({ kind: "body", text: "갑(발주자): " + v.a + " · 대표자 " + v.ra + " (서명·인)" });
        m.paras.push({ kind: "body", text: "을(수급자): " + v.b + " · 대표자 " + v.rb + " (서명·인)" });
      } else {
        m.paras.push({ kind: "body", text: koDate(v.date) });
        m.paras.push({ kind: "body", text: "발신인: " + v.a + " " + v.ra + " (서명·인)" });
      }
      m.paras.push({ kind: "meta", text: "※ 본 초안은 AI Hub 법률 문서 서식 데이터의 구조를 참고한 예시이며, 법적 효력 검토를 변호사 등 전문가와 하시기 바랍니다." });
    } else {
      m.paras.push({ kind: "meta", text: "※ 위 구획은 원문 데이터의 조문 구조를 요약한 안내입니다. 기본 정보를 입력하고 「초안 생성」을 누르면 계약서 초안으로 변환되며, DOCX·HWPX 문서 파일로 저장할 수 있습니다." });
    }
    return m;
  }

  function artTitleOf(s) {
    var lb = (s.lb || []).filter(function (x) { return x && x !== "기타"; });
    if (lb.length) return lb[0];
    var h = (s.hint || "").replace(/^기재 항목:\s*/, "").split(",")[0].trim();
    if (h && h.length <= 22 && !/기술|작성|해당|기재/.test(h)) return h;
return "계약조건";
  }

  function circled(n) {
    var c = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮⑯⑰⑱⑲⑳";
    return c[n - 1] || n + ".";
  }

  /* ---------- 화면 렌더 ---------- */
  function generate() {
    if (!state.selected) return;
    state.vals = collectVals();
    state.filled = true;
    state.model = buildDocModel(state.selected, state.skeleton, state.vals, true);
    renderModel();
    toast("초안이 생성되었습니다. 문서 파일로 저장할 수 있어요.");
  }

  function renderModel() {
    var m = state.model;
    var html = "<h3>" + esc(m.title) + (state.filled ? " (초안)" : " — 구조 미리보기") + "</h3>";
    html += "<p class='doc-sub'>" + esc(m.metaLine) + "</p>";
    m.paras.forEach(function (p) {
      if (p.kind === "article") html += "<div class='art-title'>" + esc(p.text) + "</div>";
      else if (p.kind === "meta") html += "<p class='hint'>" + esc(p.text) + "</p>";
      else html += "<p class='sec'" + (p.right ? " style='text-align:right'" : "") + ">" + esc(p.text) + "</p>";
    });
    $("doc").innerHTML = html;
  }

  /* ---------- DOCX 생성 ---------- */
  function buildDocxChildren(model) {
    var D = docxLib;
    function run(text, opts) {
      var o = { text: text, font: "맑은 고딕", size: 20 };
      if (opts) for (var k in opts) o[k] = opts[k];
      return new D.TextRun(o);
    }
    function para(children, opts) {
      var o = { children: children };
      if (opts) for (var k in opts) o[k] = opts[k];
      return new D.Paragraph(o);
    }
    var out = [];
    out.push(para([run(model.title, { bold: true, size: 32 })],
      { alignment: D.AlignmentType.CENTER, spacing: { after: 160 } }));
    out.push(para([run(model.metaLine, { size: 16, color: "647693" })],
      { alignment: D.AlignmentType.CENTER, spacing: { after: 360 } }));
    model.paras.forEach(function (p) {
      if (p.kind === "article") {
        out.push(para([run(p.text, { bold: true, size: 28 })], { spacing: { before: 240, after: 100 } }));
      } else if (p.kind === "meta") {
        out.push(para([run(p.text, { size: 16, color: "8a97ab" })], { spacing: { before: 200 } }));
      } else {
        var opts = { spacing: { after: 80 } };
        if (p.indent) opts.indent = { left: 420 };
        if (p.right) opts.alignment = D.AlignmentType.RIGHT;
        out.push(para([run(p.text)], opts));
      }
    });
    return out;
  }

  /* ---------- HWPX 생성 ---------- */
  function hwpxPara(p, isFirst) {
    var run = '<hp:run charPrIDRef="' + (p.kind === "article" ? "6" : "0") + '">';
    if (isFirst) run += ASSETS.secPr + ASSETS.ctrl;
    return '<hp:p paraPrIDRef="' + (p.kind === "article" ? "2" : "0") + '" styleIDRef="0">' +
      run + "<hp:t>" + xmlEsc(p.text) + "</hp:t></hp:run></hp:p>";
  }

  function buildSectionXml(model) {
    var lines = ['<?xml version="1.0" encoding="UTF-8" standalone="yes" ?>',
      '<hs:sec xmlns:hs="http://www.hancom.co.kr/hwpml/2011/section" xmlns:hp="http://www.hancom.co.kr/hwpml/2011/paragraph">'];
    var titlePara = { kind: "title", text: model.title };
    lines.push('<hp:p paraPrIDRef="1" styleIDRef="0"><hp:run charPrIDRef="5">' +
      ASSETS.secPr + ASSETS.ctrl + "<hp:t>" + xmlEsc(titlePara.text) + "</hp:t></hp:run></hp:p>");
    lines.push('<hp:p paraPrIDRef="0" styleIDRef="0"><hp:run charPrIDRef="0"><hp:t>' +
      xmlEsc(model.metaLine) + "</hp:t></hp:run></hp:p>");
    model.paras.forEach(function (p) { lines.push(hwpxPara(p, false)); });
    lines.push("</hs:sec>");
    return lines.join("\n");
  }

  function prvText(model) {
    var lines = [model.title, model.metaLine];
    model.paras.forEach(function (p) { lines.push(p.text); });
    return lines.join("\r\n");
  }

  function makeHwpxFiles(model) {
    return [
      ["mimetype", "application/hwp+zip"],
      ["META-INF/container.xml", CONTAINER_XML],
      ["Contents/content.hpf", HPF_XML],
      ["Contents/header.xml", ASSETS.header],
      ["Contents/section0.xml", buildSectionXml(model)],
      ["Preview/PrvText.txt", prvText(model)]
    ];
  }

  function makeHwpxBlob(model) {
    var zip = new JSZipLib();
    makeHwpxFiles(model).forEach(function (f) {
      zip.file(f[0], f[1], f[0] === "mimetype" ? { compression: "STORE" } : {});
    });
    return zip.generateAsync({ type: "blob", mimeType: "application/hwp+zip", compression: "DEFLATE" });
  }

  /* ---------- TXT 생성 ---------- */
  function buildTxt(model) {
    var lines = [model.title, "", model.metaLine, ""];
    model.paras.forEach(function (p) { lines.push(p.text); });
    return lines.join("\n");
  }

  /* ---------- 파일명·저장 ---------- */
  function sanitizeName(s) {
    return String(s).replace(/[\\/:*?"<>|\s]+/g, "_").replace(/^_+|_+$/g, "").slice(0, 60) || "문서";
  }
  function docBaseName(model) {
    var prefix = model.isContract ? "계약서초안" : "문서초안";
    var d = new Date();
    var ymd = d.getFullYear() + String(d.getMonth() + 1).padStart(2, "0") + String(d.getDate()).padStart(2, "0");
    return prefix + "_" + sanitizeName(model.title) + "_" + ymd;
  }

  function saveDoc() {
    if (!state.selected) { toast("먼저 왼쪽에서 서식을 선택하세요."); return; }
    if (!state.filled) generate();
    var fmtSel = $("fmt").value;
    var model = state.model;
    var base = docBaseName(model);
    var DOCX_MIME = "application/vnd.openxmlformats-officedocument.wordprocessingml.document";
    if (fmtSel === "docx") {
      var D = docxLib;
      var doc = new D.Document({ sections: [{ properties: {}, children: buildDocxChildren(model) }] });
      D.Packer.toBlob(doc).then(function (blob) {
        deliver(blob, base + ".docx", DOCX_MIME,
          [{ description: "워드 문서", accept: { "application/vnd.openxmlformats-officedocument.wordprocessingml.document": [".docx"] } }]);
      }).catch(function (e) { toast("DOCX 생성 실패: " + e.message); });
    } else if (fmtSel === "hwpx") {
      makeHwpxBlob(model).then(function (blob) {
        deliver(blob, base + ".hwpx", "application/hwp+zip",
          [{ description: "한글 문서", accept: { "application/hwp+zip": [".hwpx"] } }]);
      }).catch(function (e) { toast("HWPX 생성 실패: " + e.message); });
    } else {
      var blob = new Blob([buildTxt(model)], { type: "text/plain;charset=utf-8" });
      deliver(blob, base + ".txt", "text/plain",
        [{ description: "텍스트", accept: { "text/plain": [".txt"] } }]);
    }
  }

  function deliver(blob, filename, mime, types) {
    if (dirHandle) {
      saveToDir(blob, filename).then(function (ok) {
        if (ok) toast("작업 폴더에 저장 완료: " + filename);
        else saveWithPicker(blob, filename, mime, types);
      });
    } else {
      saveWithPicker(blob, filename, mime, types);
    }
  }

  function saveWithPicker(blob, filename, mime, types) {
    if (window.showSaveFilePicker) {
      window.showSaveFilePicker({ suggestedName: filename, types: types })
        .then(function (handle) {
          return handle.createWritable().then(function (w) {
            return w.write(blob).then(function () { return w.close(); });
          });
        })
        .then(function () { toast("저장 완료: " + filename); })
        .catch(function (e) {
          if (e && e.name === "AbortError") return;
          downloadFallback(blob, filename);
        });
    } else {
      downloadFallback(blob, filename);
    }
  }

  function downloadFallback(blob, filename) {
    var url = URL.createObjectURL(blob);
    var a = document.createElement("a");
    a.href = url; a.download = filename;
    document.body.appendChild(a); a.click();
    document.body.removeChild(a);
    setTimeout(function () { URL.revokeObjectURL(url); }, 4000);
    toast("다운로드 완료: " + filename);
  }

  /* ---------- 복사·인쇄·토스트 ---------- */
  function docPlainText() {
    return buildTxt(state.model || buildDocModel(state.selected, state.skeleton, null, false));
  }
  function copyDraft() {
    var text = docPlainText();
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).then(function () { toast("클립보드에 복사했습니다."); },
        function () { fallbackCopy(text); });
    } else { fallbackCopy(text); }
  }
  function fallbackCopy(text) {
    var ta = document.createElement("textarea");
    ta.value = text; document.body.appendChild(ta); ta.select();
    try { document.execCommand("copy"); toast("클립보드에 복사했습니다."); }
    catch (e) { toast("복사에 실패했습니다."); }
    document.body.removeChild(ta);
  }
  var toastTimer = null;
  function toast(msg) {
    var el = $("toast");
    el.textContent = msg; el.classList.add("show");
    clearTimeout(toastTimer);
    toastTimer = setTimeout(function () { el.classList.remove("show"); }, 2600);
  }

  if (typeof document !== "undefined") init();

  /* node 테스트용 export */
  if (typeof module !== "undefined" && module.exports) {
    module.exports = {
      buildDocModel: buildDocModel, buildDocxChildren: buildDocxChildren,
      buildSectionXml: buildSectionXml, makeHwpxFiles: makeHwpxFiles,
      buildTxt: buildTxt, genericSections: genericSections,
      docBaseName: docBaseName, koDate: koDate,
      buildAiPrompt: buildAiPrompt, callLLM: callLLM, parseAiDraft: parseAiDraft
    };
  }
})();
