/* 계약서 초안 생성기 앱 로직 */
(function () {
  "use strict";
  var IDX = window.CAT_INDEX;
  var TPL = window.CAT_TEMPLATES.templates || [];
  var tplById = {};
  TPL.forEach(function (t) { tplById[t.id] = t; });

  var $ = function (id) { return document.getElementById(id); };
  var state = { selected: null, draft: null };

  /* ---------- 초기화 ---------- */
  function init() {
    var s = IDX.stats || {};
    var tplCount = TPL.length;
    var mcCount = Object.keys(IDX.categories).length;
    var metrics = [
      { b: fmt(s.total || IDX.records.length), span: "색인된 서식 문서" },
      { b: mcCount, span: "대분류" },
      { b: tplCount, span: "구조 템플릿" },
      { b: "2", span: "AI Hub 데이터셋" }
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
    $("btn-copy").addEventListener("click", copyDraft);
    $("btn-print").addEventListener("click", function () { window.print(); });
    render();
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
    var r = state.selected;
    var tpl = tplById[r.id];
    var secs;
    if (tpl) {
      secs = tpl.sections;
    } else {
      secs = genericSections(r);
    }
    state.skeleton = secs;
    renderDoc(secs, false);
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

  /* ---------- 초안 생성 ---------- */
  function fieldVal(id) { var el = $(id); return el ? el.value.trim() : ""; }

  function generate() {
    if (!state.selected) return;
    var r = state.selected;
    var vals = {
      a: fieldVal("f-party-a") || "갑(상호)",
      ra: fieldVal("f-rep-a") || "(대표자)",
      b: fieldVal("f-party-b") || "을(상호)",
      rb: fieldVal("f-rep-b") || "(대표자)",
      amount: fieldVal("f-amount") || "금 ____원",
      period: fieldVal("f-period") || "____년 __월 __일부터 ____년 __월 __일까지",
      date: fieldVal("f-date") || "____년 __월 __일",
      place: fieldVal("f-place") || "________________",
      extra: fieldVal("f-extra")
    };
    state.vals = vals;
    renderDoc(state.skeleton, true, vals);
    toast("초안이 생성되었습니다. 복사·인쇄할 수 있어요.");
  }

  function renderDoc(secs, filled, vals) {
    var r = state.selected;
    var v = vals || {};
    var title = r.t + (filled ? " (초안)" : " — 구조 미리보기");
    var html = "<h3>" + esc(title) + "</h3>";
    html += "<p class='doc-sub'>" + esc(r.mc) + " › " + esc(r.sc) + " › " + esc(r.dc) +
      " · 관련법령: " + esc(r.law || "미명시") + " · 출처: AI 허브(" + esc(r.ds) + ")</p>";

    if (filled) {
      html += "<p class='sec'>계약서 번호: 제 ______ 호</p>";
      html += "<p class='sec'>" + esc(v.a) + "(대표자 " + esc(v.ra) + ", 이하 \"갑\")과 " +
        esc(v.b) + "(대표자 " + esc(v.rb) + ", 이하 \"을\")은 " + esc(v.date) + " " +
        esc(v.place) + "에서 아래와 같이 계약을 체결한다.</p>";
    }

    var artNo = 0, lastArt = null, itemNo = 0;
    var extraItems = (v.extra ? v.extra.split(/[,\n]/) : []).map(function (s) { return s.trim(); }).filter(Boolean);

    secs.forEach(function (s) {
      if (s.ft === "제목") return;
      if (s.art !== lastArt) {
        if (s.art != null) {
          artNo = s.art;
          var artTitle = artTitleOf(s);
          html += "<div class='art-title'>제" + artNo + "조 (" + esc(artTitle) + ")</div>";
        }
        lastArt = s.art; itemNo = 0;
      }
      itemNo++;
      var no = s.art != null ? "<span class='no'>" + circled(itemNo) + "</span>" : "";
      var body = filled ? fillText(s.hint, v, s) : esc(s.hint);
      html += "<p class='sec'>" + no + body + "</p>";
    });

    if (filled && extraItems.length) {
      artNo++;
      html += "<div class='art-title'>제" + artNo + "조 (특약사항)</div>";
      extraItems.forEach(function (t, i) {
        html += "<p class='sec'><span class='no'>" + circled(i + 1) + "</span>" + esc(t) + "</p>";
      });
    }

    if (filled) {
      html += "<p class='sec' style='margin-top:24px'>이 계약의 성립을 증명하기 위하여 계약서 2부를 작성하고 갑·을이 각각 서명 날인 후 1부씩 보관한다.</p>";
      html += "<p class='sec' style='margin-top:16px'>" + esc(v.date) + "</p>";
      html += "<div style='display:flex;gap:48px;margin-top:12px'>" +
        "<div><b>갑</b> " + esc(v.a) + "<br>대표자 " + esc(v.ra) + " (서명·인)</div>" +
        "<div><b>을</b> " + esc(v.b) + "<br>대표자 " + esc(v.rb) + " (서명·인)</div></div>";
      html += "<p class='hint' style='margin-top:24px'>※ 본 초안은 AI Hub 법률 문서 서식 데이터의 구조를 참고한 예시이며, 법적 효력 검토를 변호사 등 전문가와 하시기 바랍니다.</p>";
    } else {
      html += "<p class='hint' style='margin-top:24px'>※ 위 구획은 원문 데이터의 조문 구조를 요약한 안내입니다. 좌측 상단 「초안 생성」에서 기본 정보를 입력하면 계약서 초안으로 변환됩니다.</p>";
    }
    $("doc").innerHTML = html;
  }

  function artTitleOf(s) {
    var lb = (s.lb || []).filter(function (x) { return x && x !== "기타"; });
    if (lb.length) return lb[0];
    if (s.hint && s.hint.indexOf("기재 항목:") === 0) return s.hint.replace("기재 항목: ", "").split(",")[0];
    return "계약조건";
  }

  function circled(n) {
    var c = "①②③④⑤⑥⑦⑧⑨⑩⑪⑫⑬⑭⑮";
    return c[n - 1] || n + ".";
  }

  function fillText(hint, v, s) {
    var t = esc(hint);
    if (/계약금액|대금/.test(hint)) t += " " + esc(v.amount);
    if (/기한|기간/.test(hint)) t += " " + esc(v.period);
    return t + " ______________________";
  }

  /* ---------- 복사·인쇄·토스트 ---------- */
  function docPlainText() {
    return $("doc").innerText.replace(/\n{3,}/g, "\n\n");
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
    toastTimer = setTimeout(function () { el.classList.remove("show"); }, 2200);
  }

  init();
})();
