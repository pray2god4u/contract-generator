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
    $("btn-review-run").addEventListener("click", reviewContract);
    $("btn-instant-run").addEventListener("click", instantGenerate);
    $("btn-instant-open").addEventListener("click", openInstantInBrowse);
    $("btn-instant-save").addEventListener("click", function () { saveDoc($("fmt-instant").value); });
    restoreFolder();
    restoreAiCfg();
    initTabs();
    initGuide();
    initMyBox();
    render();
  }

  /* ---------- 사용법 가이드 (프롬프트·샘플 원클릭) ---------- */
  function initGuide() {
    /* 프롬프트: 「이 프롬프트로 생성 →」 / 「복사」 */
    Array.prototype.forEach.call(document.querySelectorAll(".prompt"), function (box) {
      var text = "";
      var span = box.querySelector(".ptext");
      if (span) text = span.textContent.trim();
      var use = box.querySelector(".use-btn");
      if (use) {
        use.addEventListener("click", function () {
          $("instant-input").value = text;
          var link = $("nav-instant");
          if (link) link.click();
          $("instant-input").focus();
          toast("프롬프트를 넣었습니다. 「✨ 바로 생성」을 누르세요. (사내망 필요)");
        });
      }
      var cp = box.querySelector(".copy-btn");
      if (cp) {
        cp.addEventListener("click", function () {
          copyToClipboard(text);
          cp.textContent = "복사됨 ✓";
          setTimeout(function () { cp.textContent = "복사"; }, 1500);
        });
      }
    });
    /* 샘플 계약서 → 검토 탭 */
    var sampleBtn = $("btn-guide-sample");
    if (sampleBtn) {
      sampleBtn.addEventListener("click", function () {
        var el = $("sample-contract-text");
        var text = el ? el.textContent.replace(/^\s*\n/, "").replace(/\s*$/, "") : "";
        $("review-input").value = text;
        var link = $("nav-review");
        if (link) link.click();
        $("review-input").focus();
        toast("샘플 계약서를 넣었습니다. 「🔍 AI 검토 시작」을 누르세요. (사내망 필요)");
      });
    }
  }

  function copyToClipboard(text) {
    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text);
      return;
    }
    var ta = document.createElement("textarea");
    ta.value = text;
    document.body.appendChild(ta);
    ta.select();
    try { document.execCommand("copy"); } catch (e) { /* 무시 */ }
    document.body.removeChild(ta);
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
    var g = (typeof globalThis !== "undefined" && globalThis.AI_CFG) || null; /* node 테스트용 */
    return {
      url: (($("ai-url") && $("ai-url").value) || (g && g.url) || AI_DEFAULTS.url).trim(),
      model: (($("ai-model") && $("ai-model").value) || (g && g.model) || AI_DEFAULTS.model).trim(),
      temp: parseFloat(($("ai-temp") && $("ai-temp").value) || (g && g.temp) || AI_DEFAULTS.temp) || 0.3,
      maxtok: parseInt(($("ai-maxtok") && $("ai-maxtok").value) || (g && g.maxtok) || AI_DEFAULTS.maxtok, 10) || 4096
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

  function callLLM(cfgOrSystem, messagesOrUser, onDeltaOrNothing, maybeOnDelta) {
    /* 두 시그니처 지원: (cfg, messages[], onDelta) | (cfg, system, user, onDelta) */
    var cfg, messages, onDelta;
    if (Array.isArray(messagesOrUser)) {
      cfg = cfgOrSystem; messages = messagesOrUser; onDelta = onDeltaOrNothing;
    } else {
      cfg = cfgOrSystem;
      messages = [
        { role: "system", content: messagesOrUser },
        { role: "user", content: onDeltaOrNothing }
      ];
      onDelta = maybeOnDelta;
    }
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

  /* ---------- 탭 전환 ---------- */
  function initTabs() {
    var tabs = [["browse", "nav-browse"], ["review", "nav-review"], ["instant", "nav-instant"], ["mybox", "nav-mybox"], ["guide", "nav-guide"]];
    tabs.forEach(function (t) {
      var link = $(t[1]);
      if (!link) return;
      link.addEventListener("click", function (ev) {
        ev.preventDefault();
        tabs.forEach(function (u) {
          var sec = $(u[0]);
          if (sec) sec.hidden = (u[0] !== t[0]);
          var nav = $(u[1]);
          if (nav) nav.classList.toggle("active", u[0] === t[0]);
        });
        if (t[0] === "mybox") renderMyBox();
        window.scrollTo(0, 0);
      });
    });
  }

  /* ---------- 나만의 서식함 ---------- */
  var BOX_KEY = "cg_mybox";
  function boxData() {
    try { return JSON.parse(localStorage.getItem(BOX_KEY) || "{}"); }
    catch (e) { return {}; }
  }
  function boxSave(d) {
    try { localStorage.setItem(BOX_KEY, JSON.stringify(d)); } catch (e) { /* 무시 */ }
  }
  function isFav(id) { return (boxData().favs || []).indexOf(id) > -1; }
  function toggleFav(id) {
    var d = boxData();
    d.favs = d.favs || [];
    var i = d.favs.indexOf(id);
    if (i > -1) d.favs.splice(i, 1); else d.favs.push(id);
    boxSave(d);
    toast(i > -1 ? "즐겨찾기에서 제거했습니다." : "즐겨찾기에 추가했습니다.");
    render(); /* 목록의 별 갱신 */
  }
  function addHistory(rec, fmt) {
    var d = boxData();
    d.hist = d.hist || [];
    d.hist.unshift({ id: rec.id, t: rec.t, mc: rec.mc, sc: rec.sc, ds: rec.ds, fmt: fmt, at: new Date().toISOString() });
    if (d.hist.length > 30) d.hist.length = 30;
    boxSave(d);
  }
  function recById(id) {
    for (var i = 0; i < IDX.records.length; i++) if (IDX.records[i].id === id) return IDX.records[i];
    return null;
  }
  function boxRow(r, extra) {
    var div = document.createElement("div");
    div.className = "result-item";
    div.innerHTML = "<div class='ri-main'><b>" + esc(r.t) + "</b><span class='ri-sub'>" +
      esc(r.mc) + " › " + esc(r.sc) + (extra ? " · " + extra : "") + "</span></div>";
    var btn = document.createElement("button");
    btn.textContent = "선택";
    btn.addEventListener("click", function () {
      select(r.id);
      document.querySelector("#browse").hidden = false;
      ["review", "instant", "mybox", "guide"].forEach(function (id) { $(id).hidden = true; });
      $("nav-browse").classList.add("active");
      ["nav-review", "nav-instant", "nav-mybox", "nav-guide"].forEach(function (id) { $(id).classList.remove("active"); });
      window.scrollTo(0, 0);
    });
    div.appendChild(btn);
    return div;
  }
  var DEFAULT_CLAUSES = [
    "분쟁 발생 시 서울중앙지방법원을 관할 법원으로 한다.",
    "본 계약과 관련한 비밀정보를 제3자에게 누설하지 아니한다.",
    "일방의 채무 불이행 시 상대방은 상당한 기간 최고 후 계약을 해제(해지)할 수 있다.",
    "계약 해제 시 이미 지급된 금원은 반환하지 아니한다.",
    "본 계약에 정하지 아니한 사항은 민법 등 관련 법령에 따른다."
  ];
  function renderMyBox() {
    var d = boxData();
    var favList = $("fav-list"), histList = $("hist-list");
    favList.innerHTML = ""; histList.innerHTML = "";
    var favs = d.favs || [];
    if (!favs.length) favList.innerHTML = "<p class='count'>아직 즐겨찾기가 없습니다 — 서식 탐색 목록에서 ⭐를 눌러보세요.</p>";
    favs.forEach(function (id) {
      var r = recById(id);
      if (r) favList.appendChild(boxRow(r));
    });
    var hist = d.hist || [];
    if (!hist.length) histList.innerHTML = "<p class='count'>아직 작성 이력이 없습니다.</p>";
    hist.forEach(function (h) {
      var r = recById(h.id) || { id: h.id, t: h.t, mc: h.mc, sc: h.sc };
      var dt = (h.at || "").slice(0, 10);
      histList.appendChild(boxRow(r, (h.fmt || "").toUpperCase() + " 저장 · " + dt));
    });
    var grid = $("clause-grid");
    grid.innerHTML = "";
    var clauses = d.clauses || DEFAULT_CLAUSES.slice();
    clauses.forEach(function (c, idx) {
      var chip = document.createElement("button");
      chip.className = "clause-chip";
      chip.textContent = c;
      chip.title = "클릭하면 특약사항 입력란에 추가됩니다";
      chip.addEventListener("click", function () {
        var ta = $("f-extra");
        if (ta) {
          ta.value = ta.value ? ta.value.replace(/\s*$/, "") + (ta.value ? "\n" : "") + c : c;
          toast("특약사항에 추가했습니다.");
        } else {
          toast("먼저 서식을 선택해 주세요.");
        }
      });
      var del = document.createElement("button");
      del.textContent = "×"; del.className = "clause-del"; del.title = "삭제";
      del.addEventListener("click", function () {
        var dd = boxData(); dd.clauses = dd.clauses || DEFAULT_CLAUSES.slice();
        dd.clauses.splice(idx, 1); boxSave(dd); renderMyBox();
      });
      var wrap = document.createElement("div");
      wrap.style.display = "flex"; wrap.style.gap = "6px"; wrap.style.alignItems = "flex-start";
      wrap.appendChild(chip); wrap.appendChild(del);
      grid.appendChild(wrap);
    });
  }
  function initMyBox() {
    $("btn-box-clear").addEventListener("click", function () {
      if (confirm("즐겨찾기·이력·특약을 모두 지울까요?")) { localStorage.removeItem(BOX_KEY); renderMyBox(); toast("서식함을 비웠습니다."); }
    });
    $("btn-clause-add").addEventListener("click", function () {
      var v = $("clause-input").value.trim();
      if (!v) return;
      var d = boxData(); d.clauses = d.clauses || DEFAULT_CLAUSES.slice();
      d.clauses.push(v); boxSave(d);
      $("clause-input").value = "";
      renderMyBox(); toast("특약 조항을 추가했습니다.");
    });
    $("btn-review-clear").addEventListener("click", function () { $("review-input").value = ""; });
    $("btn-instant-clear").addEventListener("click", function () { $("instant-input").value = ""; $("instant-output").hidden = true; });
  }

  /* ---------- AI 계약서 검토 ---------- */
  function buildReviewPrompt(text) {
    return {
      system: "당신은 대한민국 계약서 검토 전문가입니다. 주어진 계약서를 조항별로 분석하여 다음 형식으로만 출력합니다.\n\n형식:\n[요약] 전체 평가 2~3문장\n[리스크] 번호. 조항명 — 문제점 (불리함/모호함/법령소지 중 택1) — 수정 제안\n[누락] 번호. 누락된 필수 조항 — 왜 필요한가\n[종합] 종합 의견 1~2문장\n\n규칙: 리스크가 없으면 '없음', 누락도 없으면 '없음'. 실무 관점에서 구체적으로.",
      user: "다음 계약서를 검토해 주세요:\n\n" + text
    };
  }
  function reviewContract() {
    var text = $("review-input").value.trim();
    if (text.length < 50) { toast("계약서 본문을 50자 이상 붙여넣어 주세요."); return; }
    var btn = $("btn-review-run");
    btn.disabled = true;
    var out = $("review-output");
    out.hidden = false;
    out.innerHTML = "<p class='count'>AI가 계약서를 분석 중입니다…</p>";
    var cfg = aiCfg();
    var p = buildReviewPrompt(text.slice(0, 12000));
    callLLM(cfg, p.system, p.user, function (acc) {
      out.innerHTML = "<div class='ai-stream'>" + esc(acc) + "</div>";
      out.scrollTop = out.scrollHeight;
    }).then(function (full) {
      btn.disabled = false;
      out.innerHTML = "<div class='ai-stream'>" + esc(full) + "</div>";
      toast("검토 완료");
    }).catch(function (e) {
      btn.disabled = false;
      out.innerHTML = "<p class='hint'>검토 실패: " + esc(e.message) +
        "</p><p class='hint'>사내망에서 접속했는지 확인하세요.</p>";
    });
  }

  /* ---------- AI 바로 생성 ---------- */
  function buildInstantPrompt(request) {
    var catList = Object.keys(IDX.categories).map(function (mc) {
      var scs = Object.keys(IDX.categories[mc]).slice(0, 8).join("/");
      return mc + "(" + scs + ")";
    }).join(", ");
    return {
      system: "당신은 대한민국 법률 문서 전문가입니다. 사용자의 자연어 요청을 분석해 1단계로 아래 JSON만 출력합니다(다른 설명 금지):\n{\"title\":\"문서제목\",\"mc\":\"대분류\",\"sc\":\"중분류\",\"params\":{\"갑\":\"...\",\"을\":\"...\",\"계약금액\":\"...\",\"계약기간\":\"...\",\"특약사항\":\"...\"},\"clauses\":[\"제1조 (제목): 내용\", ...]}\n대분류 후보: " + catList + "\nclauses는 실제 계약서 조문 6~10개를 완성된 문장으로 작성하세요. 사용자가 안 준 정보는 합리적으로 가정하되 params에 표기.",
      user: request
    };
  }
  function parseInstantJson(full) {
    var m = full.match(/\{[\s\S]*\}/);
    if (!m) throw new Error("AI 응답에서 계획(JSON)을 찾지 못했습니다.");
    return JSON.parse(m[0]);
  }
  function instantGenerate() {
    var req = $("instant-input").value.trim();
    if (req.length < 10) { toast("요청을 10자 이상 입력해 주세요."); return; }
    var btn = $("btn-instant-run");
    btn.disabled = true;
    var out = $("instant-output");
    out.hidden = false;
    out.innerHTML = "<p class='count'>AI가 요청을 분석하고 초안을 작성합니다…</p>";
    var cfg = aiCfg();
    var p = buildInstantPrompt(req);
    callLLM(cfg, p.system, p.user, function (acc) {
      out.innerHTML = "<div class='ai-stream'>" + esc(acc.slice(0, 600)) + (acc.length > 600 ? " …" : "") + "</div>";
    }).then(function (full) {
      var plan;
      try { plan = parseInstantJson(full); }
      catch (e) { throw new Error("계획 파싱 실패 — 다시 시도해 주세요. (" + e.message + ")"); }
      var paras = [];
      paras.push({ kind: "title", text: plan.title || "계약서" });
      if (plan.params) {
        Object.keys(plan.params).forEach(function (k) {
          var v = plan.params[k];
          if (k === "특약사항") return;
          paras.push({ kind: "body", text: k + ": " + v });
        });
      }
      (plan.clauses || []).forEach(function (c) {
        var mm = c.match(/^(제\d+조)\s*\(([^)]*)\)\s*[:：]?\s*([\s\S]*)$/);
        if (mm) {
          paras.push({ kind: "article", text: mm[1] + " (" + mm[2] + ")" });
          var body = (mm[3] || "").trim();
          if (body) body.split(/\n+/).forEach(function (ln) {
            if (ln.trim()) paras.push({ kind: "body", text: ln.trim() });
          });
        } else {
          paras.push({ kind: "body", text: c });
        }
      });
      paras.push({ kind: "body", text: "\n이 계약의 성립을 증명하기 위하여 계약서 2부를 작성하고 갑·을이 각각 서명 날인 후 1부씩 보관한다." });
      paras.push({ kind: "body", text: koDate(new Date()) });
      paras.push({ kind: "body", text: "갑: " + ((plan.params || {}).갑 || "________________") + " (서명·인)" });
      paras.push({ kind: "body", text: "을: " + ((plan.params || {}).을 || "________________") + " (서명·인)" });
      paras.push({ kind: "body", text: "※ 본 초안은 사내 GenAI(dsm-120b)로 생성된 예시이며, 법적 효력 검토를 변호사 등 전문가와 하시기 바랍니다." });
      state.filled = true;
      state.model = {
        title: plan.title || "계약서",
        metaLine: "AI 바로 생성 · " + koDate(new Date()) + " · 출처: 사내 GenAI(dsm-120b)",
        paras: paras
      };
      state.selected = { t: plan.title || "계약서", mc: plan.mc || "", sc: plan.sc || "", ds: "AI" };
      renderModel();
      renderInstantPreview();
      out.innerHTML = "<p class='count'>✅ 초안 생성 완료 — 아래 「초안 미리보기」에서 확인하고 「💾 문서 파일 저장」으로 내려받으세요.</p>";
      btn.disabled = false;
      addHistory({ id: "instant-" + Date.now(), t: plan.title || "계약서", mc: plan.mc || "AI", sc: plan.sc || "", ds: "AI" }, "AI");
      toast("AI 초안이 생성되었습니다.");
    }).catch(function (e) {
      btn.disabled = false;
      out.innerHTML = "<p class='hint'>생성 실패: " + esc(e.message) + "</p>";
    });
  }

  /* ---------- 바로 생성 미리보기 ---------- */
  function renderInstantPreview() {
    var pv = $("instant-preview"), doc = $("instant-doc");
    if (!pv || !doc) return;
    var m = state.model;
    var html = "<h3>" + esc(m.title) + " (초안)</h3>";
    html += "<p class='doc-sub'>" + esc(m.metaLine) + "</p>";
    m.paras.forEach(function (p) {
      if (p.kind === "article") html += "<div class='art-title'>" + esc(p.text) + "</div>";
      else if (p.kind === "meta") html += "<p class='hint'>" + esc(p.text) + "</p>";
      else html += "<p class='sec'>" + esc(p.text) + "</p>";
    });
    doc.innerHTML = html;
    pv.hidden = false;
  }
  function openInstantInBrowse() {
    $("nav-browse").listeners && null; /* no-op guard */
    var tabs = [["browse", "nav-browse"], ["review", "nav-review"], ["instant", "nav-instant"], ["mybox", "nav-mybox"], ["guide", "nav-guide"]];
    tabs.forEach(function (t) {
      var sec = $(t[0]);
      if (sec) sec.hidden = (t[0] !== "browse");
      var nav = $(t[1]);
      if (nav) nav.classList.toggle("active", t[0] === "browse");
    });
    $("fields-panel").hidden = false;
    window.scrollTo(0, 0);
    toast("서식 탐색 탭의 문서 영역에서 편집·저장할 수 있습니다.");
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
    var prevSc = scSel.value; /* 사용자가 고른 중분류 보존 */
    scSel.innerHTML = '<option value="">전체</option>';
    if (mc) {
      Object.keys(IDX.categories[mc] || {}).sort().forEach(function (sc) {
        var o = document.createElement("option");
        o.value = sc; o.textContent = sc + " (" + IDX.categories[mc][sc] + ")";
        scSel.appendChild(o);
      });
    }
    if (prevSc && mc && (IDX.categories[mc] || {})[prevSc] != null) {
      scSel.value = prevSc; /* 대분류가 바뀌지 않았다면 선택 유지 */
    }
    var rows = filtered();
    $("count").textContent = "총 " + fmt(rows.length) + "건" + (rows.length > 100 ? " · 100건만 표시" : "");
    var list = rows.slice(0, 100).map(function (r) {
      var hasTpl = tplById[r.id] ? " · <span class='badge purple'>구조템플릿</span>" : "";
      var fav = isFav(r.id) ? "<span class='fav on' data-fav='" + esc(r.id) + "' title='즐겨찾기 해제'>★</span>" : "<span class='fav' data-fav='" + esc(r.id) + "' title='즐겨찾기'>☆</span>";
      return '<div class="result-item" data-id="' + esc(r.id) + '"><b>' + fav + esc(r.t) + "</b>" +
        "<div class='meta'><span class='badge'>" + esc(r.ds) + "</span>" +
        "<span>" + esc(r.mc) + " › " + esc(r.sc) + " › " + esc(r.dc) + "</span>" +
        "<span>" + esc(r.law || "법령 미명시") + " · " + r.nsec + "개 구획</span>" + hasTpl + "</div></div>";
    }).join("");
    $("results").innerHTML = list || "<p class='count'>검색 결과가 없습니다.</p>";
    Array.prototype.forEach.call($("results").children, function (el) {
      el.addEventListener("click", function (ev) {
        var favEl = ev.target && ev.target.getAttribute && ev.target.getAttribute("data-fav");
        if (favEl) { toggleFav(favEl); return; } /* 별 클릭은 선택하지 않음 */
        select(el.getAttribute("data-id"));
      });
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
    if (!state.selected) { toast("목록에서 찾을 수 없는 문서입니다."); return; }
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
        m.paras.push({ kind: "sign", text: buildSignBlock(v) });
        var att = buildAttachBlock(v);
        if (att) m.paras.push({ kind: "sign", text: att });
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

  /* ---------- 실무 서명·첨부 블록 ---------- */
  function buildSignBlock(v) {
    var L = [];
    L.push("─────────────────────────────");
    L.push("(갑) " + (v.a || "________________") + "  대표자: " + (v.ra || "________") + "   (인)");
    L.push("주소: " + (v.aaddr || "________________________"));
    L.push("");
    L.push("(을) " + (v.b || "________________") + "  대표자: " + (v.rb || "________") + "   (인)");
    L.push("주소: " + (v.baddr || "________________________"));
    return L.join("\n");
  }
  function buildAttachBlock(v) {
    var items = (v.attach || "").split(/[,\n]/).map(function (x) { return x.trim(); }).filter(Boolean);
    if (!items.length) return "";
    var L = ["─────────────────────────────", "[첨부목록]"];
    items.forEach(function (t, i) { L.push("  " + (i + 1) + ". " + t); });
    return L.join("\n");
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
      } else if (p.kind === "sign") {
        String(p.text).split("\n").forEach(function (ln) {
          out.push(para([run(ln || " ", { size: 20 })], { spacing: { after: 40 } }));
        });
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
    return String(p.text).split("\n").map(function (ln, i) {
      return '<hp:p paraPrIDRef="' + (p.kind === "article" ? "2" : "0") + '" styleIDRef="0">' +
        run + "<hp:t>" + xmlEsc(ln) + "</hp:t></hp:run></hp:p>";
    }).join("");
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

  function saveDoc(fmtOverride) {
    if (!state.selected) { toast("먼저 왼쪽에서 서식을 선택하세요."); return; }
    if (!state.filled) generate();
    var fmtSel = fmtOverride || $("fmt").value;
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
    if (state.selected) {
      addHistory({ id: state.selected.id, t: state.selected.t, mc: state.selected.mc, sc: state.selected.sc }, filename.split(".").pop());
    }
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
      buildAiPrompt: buildAiPrompt, callLLM: callLLM, parseAiDraft: parseAiDraft,
      buildReviewPrompt: buildReviewPrompt,
      buildInstantPrompt: buildInstantPrompt, parseInstantJson: parseInstantJson,
      buildSignBlock: buildSignBlock, buildAttachBlock: buildAttachBlock
    };
  }
})();
