/* 엑셀 정리기 — 화면.
 *
 * 파일 → 읽기(SheetJS, 브라우저 안) → 진단 → 처방 선택(rules/*.json) → 확인 후 정리 → 원본과 정리본 비교 → 내려받기
 * 서버로 아무것도 보내지 않는다. fetch 는 «내 사이트의» rules/ 와 data/ 를 읽을 때만 쓴다.
 * 사용자 파일 · 규정 파일 · 정답 파일은 FileReader 로만 읽고, 어디로도 보내지 않는다.
 * 처방은 rules/*.json 에만 있다. 이 파일에는 처방을 적지 않는다.
 */
(function () {
  "use strict";
  const $ = s => document.querySelector(s);
  const $$ = s => Array.from(document.querySelectorAll(s));
  const F = Clean.FLAG;
  const S = {
    name: "", base: "", rows: [], sheets: [], cols: [],
    rules: null, rulesFrom: "", units: [], enabled: new Set(),
    pending: null,                      // 충돌을 확인받아야 하는 새 처방
    preview: null, result: null, idem: null,
    policy: null, policyName: "", policyRes: [],
    diag: null, diagKind: "전체", cmpMode: "touched",
  };
  const esc = v => String(v === undefined || v === null ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const fmt = n => (typeof n === "number" ? n.toLocaleString("ko-KR") : n);
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
  };
  const savedKey = () => `cleaner:rules:${S.base}`;

  // ── 읽기 ────────────────────────────────────────────
  function readWorkbook(buf, name) {
    const isCsv = /\.csv$/i.test(name);
    const wb = isCsv
      ? XLSX.read(new TextDecoder("utf-8").decode(buf).replace(/^﻿/, ""), { type: "string", raw: true })
      : XLSX.read(buf, { type: "array", cellDates: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: !isCsv });
    // 엑셀 날짜 칸은 Date 로 읽힌다 — 날짜 글자(YYYY-MM-DD)로 바꿔 둔다. 값은 그대로
    const pad = n => String(n).padStart(2, "0");
    rows.forEach(r => Object.keys(r).forEach(k => {
      const v = r[k];
      if (v instanceof Date && !isNaN(v)) r[k] = `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    }));
    return { rows, sheets: wb.SheetNames };
  }

  async function loadFile(buf, name) {
    const { rows, sheets } = readWorkbook(buf, name);
    Object.assign(S, { name, base: name.replace(/\.(xlsx|xls|csv)$/i, ""), rows, sheets, result: null, policyRes: [], pending: null });
    S.cols = rows.length ? Object.keys(rows[0]) : [];
    Object.freeze(S.rows);               // 원본은 바꾸지 않는다 — 정리본은 새로 만든다
    $("#file-info").hidden = false;
    $("#file-info").innerHTML = `${esc(name)} · ${fmt(rows.length)}줄 × ${S.cols.length}칸`
      + (sheets.length > 1 ? ` <span class="warn">· 시트 ${sheets.length}개 중 첫 시트(${esc(sheets[0])})만 읽었습니다</span>` : "");
    ["#s-result", "#s-down"].forEach(s => { $(s).hidden = true; });
    await loadRulesFor();
    renderDiag();
    $("#s-diag").hidden = false; $("#s-rules").hidden = false;
    markSteps(3);
    $("#s-diag").scrollIntoView();
  }

  function markSteps(upto) { $$(".steps a").forEach(a => a.classList.toggle("done", +a.dataset.step <= upto)); }

  // ── 처방 불러오기 ────────────────────────────────────
  async function loadRulesFor() {
    let site = null;
    try {
      const r = await fetch(`rules/${encodeURIComponent(S.base)}.json`);
      site = r.ok ? await r.json() : null;
    } catch (e) { site = null; }
    const saved = store.get(savedKey());
    S.rules = null; S.rulesFrom = "";
    if (site) useRules(site, `사이트의 rules/${S.base}.json`);
    $("#rules-saved").hidden = !saved;
    if (saved && !site) offerRules(saved.rules, "이 브라우저에 저장한 처방", saved.enabled);
    renderRules();
  }

  function useRules(rules, from, enabledIds) {
    S.rules = rules; S.rulesFrom = from; S.pending = null;
    S.units = Clean.ruleUnits(rules);
    const ids = S.units.map(u => u.id);
    S.enabled = new Set(enabledIds ? enabledIds.filter(id => ids.includes(id)) : ids);
  }

  // 지금 처방과 새 처방이 부딪히면 자동으로 바꾸지 않고 알린다
  function offerRules(rules, from, enabledIds) {
    const c = S.rules ? Clean.conflicts(S.rules, rules) : [];
    if (c.length) { S.pending = { rules, from, enabledIds, conflicts: c }; }
    else useRules(rules, from, enabledIds);
    renderRules();
  }

  // ── 2 진단 ──────────────────────────────────────────
  const KIND_CLASS = { "빈칸": "빈칸", "중복": "중복", "표기 불일치": "표기", "형식 불일치": "형식", "이상치 후보": "이상치", "칸 사이 모순": "모순" };
  const KINDS = ["빈칸", "중복", "표기 불일치", "형식 불일치", "이상치 후보", "칸 사이 모순"];
  const badge = k => `<span class="bd ${KIND_CLASS[k] || ""}">${esc(k)}</span>`;

  function renderDiag() {
    const d = S.diag = Clean.diagnose(S.rows, S.rules);
    const n = k => d.items.filter(x => x.kind === k).reduce((a, x) => a + x.count, 0);
    $("#diag-kpi").innerHTML = [
      ["줄 × 칸", `${fmt(d.n)} × ${d.cols}`, "머리줄 제외", "dark"],
      ["찾은 문제", fmt(d.items.length), "유형 · 칸별 항목 수", ""],
      ["영향 받는 줄", fmt(d.affectedRows), `전체 ${fmt(d.n)}줄 중`, "red"],
      ["빈칸", fmt(n("빈칸")), "칸 수", ""],
      ["완전히 같은 줄", fmt(d.duplicates), "앞 줄과 같은 줄", ""],
      ["이상치 후보", fmt(n("이상치 후보")), "오류라는 뜻은 아님", "amber"],
    ].map(kpi).join("");
    $("#diag-filter").innerHTML = ["전체", ...KINDS].map(k => {
      const c = k === "전체" ? d.items.length : d.items.filter(x => x.kind === k).length;
      return `<button data-k="${esc(k)}" aria-pressed="${S.diagKind === k}">${esc(k)}<span class="c">${c}</span></button>`;
    }).join("");
    drawDiagTable();
    const cats = Object.entries(d.columns).filter(([, i]) => i.values);
    $("#values").innerHTML = cats.map(([c, i]) =>
      `<p><b>${esc(c)}</b> (${Object.keys(i.values).length}종) — ${Object.entries(i.values).sort((a, b) => b[1] - a[1])
        .map(([v, k]) => `«${esc(v)}» ${k}`).join(" · ")}</p>`).join("") || `<p class="muted">종류가 30가지 이하인 글자 칸이 없습니다.</p>`;
  }

  function drawDiagTable() {
    const q = $("#diag-q").value.trim().toLowerCase();
    const items = S.diag.items.filter(x => (S.diagKind === "전체" || x.kind === S.diagKind) && (!q
      || [x.kind, x.col, x.note, ...x.examples.map(e => e.value)].some(v => String(v).toLowerCase().includes(q))
      || x.rows.some(r => String(r) === q)));
    $("#diag-table").innerHTML = `<thead><tr>${["유형", "칸", "건수", "영향 받는 행", "예 (행: 값)", "설명"].map(h => `<th><span class="rz">${h}</span></th>`).join("")}</tr></thead><tbody>`
      + (items.length ? items.map(x => `<tr><td>${badge(x.kind)}</td><td><b>${esc(x.col)}</b></td><td class="num">${fmt(x.count)}</td>
        <td class="wrap">${x.rows.slice(0, 14).join(", ")}${x.rows.length > 14 ? ` <span class="muted">외 ${x.rows.length - 14}</span>` : ""}</td>
        <td class="wrap">${x.examples.map(e => `<span class="muted">${e.row}:</span> ${esc(e.value === "" ? "(빈칸)" : e.value)}`).join(" · ")}</td>
        <td class="note">${esc(x.note)}</td></tr>`).join("")
        : `<tr><td colspan="6" class="empty">${S.diag.items.length ? "찾는 조건에 맞는 항목이 없습니다." : "규칙 없이 찾을 수 있는 증상이 없습니다. 그래도 값 목록은 눈으로 보십시오."}</td></tr>`)
      + "</tbody>";
  }

  // ── 3 처방 선택 ─────────────────────────────────────
  const GROUPS = [
    ["형식", "자동 정리 — 모양만 바꿈", "공백 · 쉼표 · 단위 · 날짜 표기처럼 뜻은 그대로 두고 모양을 맞춥니다."],
    ["값", "사람이 정한 값 바꾸기", "처방 파일에 사람이 적은 표대로만 값을 바꿉니다 (예: «Sales» → «영업»). 뜻이 걸린 일이라 꼭 확인하세요."],
    ["지움", "줄 지우기", "지운 줄은 «제외된 줄»과 변경 기록에 남깁니다."],
    ["표시", "검토 표시 — 고치지 않음", "빈칸 · 범위 밖 · 칸 사이 모순은 값을 바꾸지 않고 «_표시» 칸에 까닭을 적습니다. 진짜 값은 사람이 확인합니다."],
  ];

  function renderRules() {
    const has = !!S.rules;
    $("#rules-missing").hidden = has || !!S.pending;
    $("#rules-src").innerHTML = has
      ? `처방: <b>${esc(S.rules.title || "처방")}</b> · 정한 사람 <b>${esc(S.rules.owner || "(적혀 있지 않음)")}</b> · ${esc(S.rulesFrom)}`
        + (S.rules.order_note ? `<br>순서 — ${esc(S.rules.order_note)}` : "")
      : "";
    const cf = $("#conflict");
    if (S.pending) {
      cf.hidden = false;
      cf.innerHTML = `<b>지금 처방(${esc(S.rulesFrom || "없음")})과 새 처방(${esc(S.pending.from)})이 부딪힙니다.</b> 자동으로 고르지 않았습니다. 어느 쪽을 쓸지 정하세요.
        <ul>${S.pending.conflicts.map(c => `<li>${esc(c)}</li>`).join("")}</ul>
        <div class="toolbar" style="margin:10px 0 0"><button class="btn" id="cf-keep">지금 처방 그대로</button><button class="btn primary" id="cf-take">새 처방으로 바꾸기</button></div>`;
      $("#cf-keep").onclick = () => { S.pending = null; renderRules(); };
      $("#cf-take").onclick = () => { const p = S.pending; useRules(p.rules, p.from, p.enabledIds); renderDiag(); renderRules(); };
    } else cf.hidden = true;

    if (!has) { $("#rules-groups").innerHTML = ""; $("#policy-box").hidden = true; updatePreview(); return; }
    $("#rules-groups").innerHTML = GROUPS.map(([g, title, desc]) => {
      const us = S.units.filter(u => u.group === g);
      if (!us.length) return "";
      return `<div class="rgroup"><div class="rgroup-h"><h3>${title}</h3><p>${desc}</p>
        <button class="btn" data-all="${g}" data-on="1">모두 켜기</button><button class="btn" data-all="${g}" data-on="0">모두 끄기</button></div>`
        + us.map(u => `<label class="rule ${S.enabled.has(u.id) ? "" : "off"}" data-id="${esc(u.id)}">
          <input type="checkbox" data-unit="${esc(u.id)}" ${S.enabled.has(u.id) ? "checked" : ""}>
          <div class="col">${esc(u.col)}</div>
          <div><div class="how">${esc(u.how)}</div><div class="why">이유 — ${esc(u.why || "(적혀 있지 않음)")}</div></div>
          <div class="imp" data-imp="${esc(u.id)}"></div></label>`).join("") + "</div>";
    }).join("");
    renderPolicyBox();
    updatePreview();
  }

  // 켠 처방으로 미리 돌려 본다 — 원본은 그대로, 결과는 화면에 쓰지 않는다
  function updatePreview() {
    const btn = $("#btn-apply");
    if (!S.rules) { $("#confirm-text").innerHTML = `<b>처방이 없어 정리할 수 없습니다.</b>`; btn.disabled = true; return; }
    const p = S.preview = Clean.apply(S.rows, S.rules, { enabled: [...S.enabled] });
    const s = p.summary;
    S.units.forEach(u => {
      const el = document.querySelector(`[data-imp="${CSS.escape(u.id)}"]`);
      if (!el) return;
      const n = s.byRule[u.id] || 0;
      const unit = u.group === "표시" ? "줄 표시" : u.group === "지움" ? "줄 지움" : "칸 바뀜";
      el.innerHTML = S.enabled.has(u.id) ? `예상 <b>${fmt(n)}</b> ${unit}` : `<span class="muted">꺼짐</span>`;
    });
    $("#confirm-text").innerHTML = `켠 처방 <b>${S.enabled.size}</b> / ${S.units.length}개 · 예상: 값 <b>${fmt(s.changes)}칸</b>(${fmt(s.changedRows)}줄) 바뀜 ·
      <b>${fmt(s.dropped)}줄</b> 지움 · <b>${fmt(s.flagged)}줄</b> 검토 표시<br><span class="muted">원본은 그대로 두고 새 정리본을 만듭니다.</span>`;
    btn.disabled = !!S.pending;
  }

  // ── 규정 확인 (처방 파일에 policy_fields 가 있을 때만) ──────
  const NEEDS = [
    ["출장 국가 · 도시", f => f.place], ["주 또는 지역", f => f.region], ["출장 시작일", f => f.date], ["출장 종료일", f => f.end],
    ["숙박일(박수)", f => f.nights], ["부서", f => f.dept], ["직무", f => f.job], ["비용항목", f => f.items && Object.keys(f.items).join(" · ")],
    ["원래 결제 통화", f => f.currency || ($("#currency").value ? `화면에서 정함: ${$("#currency").value}` : "")],
    ["적용 환율 · 환산 기준일", () => S.policy && (S.policy.fx || []).length ? "규정 파일" : ""],
    ["회사 출장비 규정 · 유효기간", () => S.policy ? `${S.policyName} (${S.policy.valid_from || "?"} ~ ${S.policy.valid_to || "?"})` : ""],
  ];
  function renderPolicyBox() {
    const f = S.rules && S.rules.policy_fields;
    $("#policy-box").hidden = !f;
    if (!f) return;
    $("#policy-fields").innerHTML = `<div class="fieldlist">${NEEDS.map(([k, get]) => {
      const v = get(f);
      return `<span class="bd ${v ? "ok" : "warn"}" title="${esc(v || "없음")}">${v ? "있음" : "없음"} · ${esc(k)}</span>`;
    }).join("")}</div><p class="muted">없는 정보 때문에 판단할 수 없는 줄은 «판단 불가»로 검토 필요 목록에 남깁니다.</p>`;
    const p = S.policy;
    $("#policy-info").innerHTML = p
      ? `<b>${esc(S.policyName)}</b> · ${esc(p.name || "")} ${p.version ? "· 버전 " + esc(p.version) : ""} · 회사 규정 ${(p.items || []).length}개 ·
         채택한 외부 기준 ${(p.external || []).filter(e => e.adopted === true).length}개 (등록 ${(p.external || []).length}개)`
      : "아직 고르지 않았습니다. 규정 없이는 금액이 적정한지 판단하지 않습니다.";
  }

  // ── 4 정리 실행과 비교 ───────────────────────────────
  function runApply() {
    if (!S.rules || S.pending) return;
    const opts = { enabled: [...S.enabled] };
    const r1 = Clean.apply(S.rows, S.rules, opts);
    const r2 = Clean.apply(r1.rows, S.rules, opts);           // 한 번 더 — 두 번 돌려도 같은가
    S.result = r1;
    S.idem = Clean.same(r1.rows, r2.rows) && r2.summary.changes === 0 && r2.summary.dropped === 0;
    const pf = S.rules.policy_fields;
    // 규정 파일이 없으면 비용 적정성은 판단하지 않는다 — 금액 크기만으로 판단하지 않는다
    S.policyRes = pf && S.policy ? Clean.checkPolicy(r1.rows, r1.src, S.policy, { fields: pf, currency: $("#currency").value }) : [];
    S.ranAt = new Date();
    buildMaps();
    $("#s-result").hidden = false; $("#s-down").hidden = false;   // 먼저 보이게 해야 줄 높이를 잴 수 있다
    renderResult();
    markSteps(5);
    $("#s-result").scrollIntoView();
  }

  let M = {};
  function buildMaps() {
    const r = S.result;
    M = { chg: new Map(), flagRow: new Map(), flagCell: new Map(), drop: new Map(), clean: new Map() };
    r.cells.forEach(l => M.chg.set(`${l.row}|${l.col}`, l));
    r.flags.forEach(f => {
      (M.flagRow.get(f.row) || M.flagRow.set(f.row, []).get(f.row)).push(f);
      f.col.split(" · ").forEach(c => (M.flagCell.get(`${f.row}|${c}`) || M.flagCell.set(`${f.row}|${c}`, []).get(`${f.row}|${c}`)).push(f));
    });
    r.dropped.forEach(d => M.drop.set(d.row, d));
    r.rows.forEach((row, i) => M.clean.set(r.src[i], row));
    S.policyRes.filter(p => p.status !== "한도 안").forEach(p => { if (!M.flagRow.has(p.row)) M.flagRow.set(p.row, []); });
  }

  function kpi([k, v, s, cls]) { return `<div class="kpi ${cls || ""}"><div class="k">${k}</div><div class="v">${v}</div><div class="s">${s || ""}</div></div>`; }

  function reviewItems() {
    const r = S.result;
    const a = r.flags.map(f => ({ 구분: "처방 표시", 행: f.row, 키: f.key, 칸: f.col, 사유: f.reason, 지금값: f.col.split(" · ").map(c => `${c}=${(M.clean.get(f.row) || {})[c] ?? ""}`).join(", "), 근거: unitText(f.rule) }));
    const b = S.policyRes.filter(p => p.status !== "한도 안").map(p => ({ 구분: `규정 확인 · ${p.status}`, 행: p.row, 키: p.key, 칸: p.item, 사유: p.reason, 지금값: p.amount, 근거: p.basis || "" }));
    return a.concat(b).sort((x, y) => x.행 - y.행);
  }
  const unitText = id => { const u = S.units.find(x => x.id === id); return u ? `${u.col}: ${u.how}` : id; };

  function renderResult() {
    const r = S.result, s = r.summary;
    const rev = reviewItems();
    const revRows = new Set(rev.map(x => x.행)).size;
    const pol = S.policyRes;
    $("#res-kpi").innerHTML = [
      ["줄 (전 → 후)", `${fmt(s.before)} → ${fmt(s.after)}`, "원본은 그대로", "dark"],
      ["바뀐 칸", fmt(s.changes), "값이 달라진 칸", "red"],
      ["바뀐 줄", fmt(s.changedRows), "한 칸이라도 바뀐 줄", "red"],
      ["지운 줄", fmt(s.dropped), "«제외된 줄»에 기록", ""],
      ["검토 필요 줄", fmt(revRows), `처방 표시 ${fmt(s.flagged)}줄${pol.length ? ` · 규정 확인 ${fmt(new Set(pol.filter(p => p.status !== "한도 안").map(p => p.row)).size)}줄` : ""}`, "amber"],
    ].map(kpi).join("");
    $("#idem").innerHTML = S.idem
      ? `<b class="pass">✓ 두 번 돌려도 같은 결과</b> — 정리본에 같은 처방을 한 번 더 적용해도 바뀌는 것이 없습니다.`
      : `<b class="fail">✕ 두 번 돌리면 결과가 달라집니다</b> — 처방이 자기 결과를 또 바꿉니다. 처방을 다시 보십시오.`;
    if (S.rules.policy_fields) {
      const pc = st => pol.filter(p => p.status === st).length;
      $("#idem").innerHTML += S.policy
        ? `<br><b>비용 적정성 (${esc(S.policyName)})</b> — 한도 안 ${pc("한도 안")} · 한도 초과 검토 ${pc("한도 초과 — 검토 필요")} · 예외 조건 ${pc("예외 승인 조건 해당")} · 판단 불가 ${pc("판단 불가")} (항목 수)`
        : `<br><b>비용 적정성은 확인하지 않았습니다</b> — 회사 규정 파일이 없습니다. 금액이 크다는 이유만으로 규정 위반으로 보지 않습니다.`;
    }
    const counts = { log: s.changes, review: rev.length, drop: s.dropped };
    $$(".tabs button").forEach(b => {
      if (!b.dataset.label) b.dataset.label = b.textContent.trim();
      const c = counts[b.dataset.tab];
      b.innerHTML = esc(b.dataset.label) + (c !== undefined ? `<span class="c">${fmt(c)}</span>` : "");
    });
    drawCompare(); drawTable("#t-orig", S.rows.map((row, i) => ({ rn: i + 2, row })), S.cols, false);
    drawTable("#t-clean", r.rows.map((row, i) => ({ rn: r.src[i], row })), S.cols.concat(F), true);
    drawLog(); drawReview(); drawDrop();
    $("#compare").innerHTML = ""; $("#cell-detail").hidden = true;
  }

  // 값 하나를 칸에 — 숫자는 오른쪽, 날짜는 고정폭, 앞뒤 공백은 ␣ 로 보이게
  function cell(v, extraCls, attrs) {
    let cls = extraCls || "", text;
    if (v === "" || v === null || v === undefined) { cls += " blank"; text = "(빈칸)"; }
    else if (typeof v === "number") { cls += " num"; text = String(v); }
    else {
      const s = String(v);
      if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(s.trim())) cls += " date";
      text = s !== s.trim() ? s.replace(/^\s+|\s+$/g, m => "␣".repeat(m.length)) : s;
    }
    return `<td class="${cls.trim()}" ${attrs || ""}>${esc(text)}</td>`;
  }
  const head = cols => `<thead><tr><th class="rn">행</th>${cols.map(c => `<th><span class="rz">${esc(c)}</span></th>`).join("")}</tr></thead>`;

  function drawTable(sel, list, cols, isClean) {
    const max = 3000;
    $(sel).innerHTML = `<table class="grid">${head(cols)}<tbody>${list.slice(0, max).map(({ rn, row }) =>
      `<tr><td class="rn">${rn}</td>${cols.map(c => c === F ? `<td class="note">${esc(row[c])}</td>`
        : cell(row[c], isClean && M.chg.has(`${rn}|${c}`) ? "chg" : isClean && M.flagCell.has(`${rn}|${c}`) ? "flg" : "")).join("")}</tr>`).join("")}</tbody></table>`
      + (list.length > max ? `<p class="empty">앞 ${fmt(max)}줄만 보여 줍니다. 전체는 내려받아 보세요.</p>` : "");
  }

  function drawCompare() {
    const q = $("#cmp-q").value.trim().toLowerCase();
    const cols = S.cols;
    const list = S.rows.map((row, i) => i + 2).filter(rn => {
      const changed = cols.some(c => M.chg.has(`${rn}|${c}`)), flagged = M.flagRow.has(rn), dropped = M.drop.has(rn);
      const ok = S.cmpMode === "all" || (S.cmpMode === "changed" && changed) || (S.cmpMode === "flagged" && flagged)
        || (S.cmpMode === "touched" && (changed || flagged || dropped));
      if (!ok) return false;
      if (!q) return true;
      const o = S.rows[rn - 2], c = M.clean.get(rn) || {};
      return String(rn) === q || cols.some(k => String(o[k]).toLowerCase().includes(q) || String(c[k] ?? "").toLowerCase().includes(q));
    });
    const max = 2000, shown = list.slice(0, max);
    const statusL = rn => {
      const b = [];
      if (cols.some(c => M.chg.has(`${rn}|${c}`))) b.push(`<span class="bd bad">바뀜</span>`);
      if (M.flagRow.has(rn)) b.push(`<span class="bd warn">검토</span>`);
      if (M.drop.has(rn)) b.push(`<span class="bd 중복">지움</span>`);
      return b.join(" ");
    };
    const left = `<table class="grid">${head(cols.concat("상태"))}<tbody>${shown.map(rn => {
      const o = S.rows[rn - 2];
      return `<tr data-rn="${rn}"><td class="rn">${rn}</td>${cols.map(c => cell(o[c], M.chg.has(`${rn}|${c}`) ? "chg" : "", `data-c="${esc(c)}"`)).join("")}<td class="note">${statusL(rn)}</td></tr>`;
    }).join("")}</tbody></table>`;
    const right = `<table class="grid">${head(cols.concat("검토 사유"))}<tbody>${shown.map(rn => {
      const d = M.drop.get(rn);
      if (d) return `<tr class="del" data-rn="${rn}"><td class="rn">${rn}</td>${cols.map(c => cell(S.rows[rn - 2][c])).join("")}<td class="note">지움 — 행 ${d.sameAs}과 같음</td></tr>`;
      const c = M.clean.get(rn) || {};
      const why = (M.flagRow.get(rn) || []).map(f => f.reason).concat(S.policyRes.filter(p => p.row === rn && p.status !== "한도 안").map(p => `${p.item}: ${p.status}`));
      return `<tr data-rn="${rn}"><td class="rn">${rn}</td>${cols.map(k => cell(c[k], M.chg.has(`${rn}|${k}`) ? "chg" : M.flagCell.has(`${rn}|${k}`) ? "flg" : "", `data-c="${esc(k)}"`)).join("")}<td class="note">${esc(why.join(" · "))}</td></tr>`;
    }).join("")}</tbody></table>`;
    const empty = `<p class="empty">이 보기에 해당하는 줄이 없습니다.</p>`;
    $("#cmp-left").innerHTML = shown.length ? left : empty;
    $("#cmp-right").innerHTML = shown.length ? right : empty;
    if (list.length > max) $("#cmp-right").insertAdjacentHTML("beforeend", `<p class="empty">앞 ${fmt(max)}줄만 보여 줍니다. 찾기로 좁혀 보세요.</p>`);
    alignRows();
  }

  // 원본과 정리본의 같은 줄이 같은 높이에 오게 맞춘다 (먼저 모두 재고, 그다음 한꺼번에 쓴다)
  function alignRows() {
    const L = $$("#cmp-left tr"), R = $$("#cmp-right tr");
    if (L.length !== R.length) return;
    L.concat(R).forEach(t => { t.style.height = ""; });
    const h = L.map((t, i) => Math.max(t.getBoundingClientRect().height, R[i].getBoundingClientRect().height));
    L.forEach((t, i) => { t.style.height = h[i] + "px"; R[i].style.height = h[i] + "px"; });
  }

  function showDetail(rn, col) {
    const l = M.chg.get(`${rn}|${col}`), fl = M.flagCell.get(`${rn}|${col}`) || [];
    if (!l && !fl.length) return;
    $$(".split tr.sel").forEach(t => t.classList.remove("sel"));
    $$(`.split tr[data-rn="${rn}"]`).forEach(t => t.classList.add("sel"));
    const d = $("#cell-detail");
    d.hidden = false;
    d.innerHTML = `<dl><dt>행 · 칸</dt><dd>${rn} · <b>${esc(col)}</b></dd>`
      + (l ? `<dt>변경 전</dt><dd class="from">${esc(l.from === "" ? "(빈칸)" : l.from)}</dd><dt>변경 후</dt><dd class="to">${esc(l.to === "" ? "(빈칸)" : l.to)}</dd>
        <dt>적용 처방</dt><dd>${l.rules.map(id => esc(unitText(id))).join("<br>")}</dd><dt>변경 사유</dt><dd>${esc(l.rule)}</dd>` : "")
      + (fl.length ? `<dt>검토 필요</dt><dd>${fl.map(f => `${esc(f.reason)} <span class="muted">(${esc(unitText(f.rule))})</span>`).join("<br>")} — 값은 고치지 않았습니다</dd>` : "")
      + "</dl>";
  }

  function simpleTable(sel, heads, rows, q) {
    const ql = (q || "").trim().toLowerCase();
    const list = ql ? rows.filter(r => r.some(v => String(v).toLowerCase().includes(ql))) : rows;
    $(sel).innerHTML = list.length
      ? `<table class="grid"><thead><tr>${heads.map(h => `<th><span class="rz">${esc(h)}</span></th>`).join("")}</tr></thead><tbody>${list.slice(0, 3000).map(r =>
        `<tr>${r.map((v, i) => i === 0 ? `<td class="rn">${esc(v)}</td>` : (heads[i] === "이유" || heads[i] === "사유" || heads[i] === "근거" || heads[i] === "처방") ? `<td class="note">${esc(v)}</td>` : cell(v)).join("")}</tr>`).join("")}</tbody></table>`
      : `<p class="empty">${rows.length ? "찾는 조건에 맞는 것이 없습니다." : "없습니다."}</p>`;
  }
  const logRows = () => S.result.cells.map(l => [l.row, l.key, l.col, l.from, l.to, l.rules.map(unitText).join(" / "), l.rule]);
  function drawLog() { simpleTable("#t-log", ["행", "키", "칸", "변경 전", "변경 후", "처방", "이유"], logRows(), $("#log-q").value); }
  function drawReview() { simpleTable("#t-review", ["행", "구분", "키", "칸", "사유", "지금 값", "근거"], reviewItems().map(x => [x.행, x.구분, x.키, x.칸, x.사유, x.지금값, x.근거]), $("#rev-q").value); }
  function drawDrop() { simpleTable("#t-drop", ["행", "같은 줄", ...S.cols], S.result.dropped.map(d => [d.row, `행 ${d.sameAs}`, ...d.values])); }

  // ── 5 내려받기 ──────────────────────────────────────
  const stamp = () => S.ranAt ? S.ranAt.toLocaleString("ko-KR") : "";
  function saveXlsx(name, sheets) {
    const wb = XLSX.utils.book_new();
    sheets.forEach(([title, rows]) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(rows.length ? rows : [{ 내용: "없음" }]), title));
    XLSX.writeFile(wb, name);
  }
  function saveCsv(name, rows) {
    const csv = "﻿" + XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(rows.length ? rows : [{ 내용: "없음" }]));  // 엑셀에서 한글이 깨지지 않게 BOM
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    a.download = name; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  const logObjs = () => S.result.cells.map(l => ({ 행: l.row, 키: l.key, 칸: l.col, 변경전: l.from, 변경후: l.to, 처방: l.rules.map(unitText).join(" / "), 이유: l.rule }));
  const dropObjs = () => S.result.dropped.map(d => Object.assign({ 행: d.row, 같은줄: d.sameAs }, Object.fromEntries(S.cols.map((c, i) => [c, d.values[i]]))));
  const unitObjs = () => S.units.map(u => ({ 처방: u.id, 켬: S.enabled.has(u.id) ? "켬" : "끔", 묶음: u.group, 칸: u.col, 방식: u.how, 이유: u.why, 영향: S.result.summary.byRule[u.id] || 0 }));
  const infoObjs = () => [{ 항목: "파일", 값: S.name }, { 항목: "처방", 값: `${S.rules.title || ""} (${S.rulesFrom})` }, { 항목: "정한 사람", 값: S.rules.owner || "" },
    { 항목: "실행 시각", 값: stamp() }, { 항목: "줄", 값: `${S.result.summary.before} → ${S.result.summary.after}` },
    { 항목: "두 번 돌려도 같은가", 값: S.idem ? "같음" : "다름" }, { 항목: "규정 파일", 값: S.policy ? S.policyName : "없음" }];
  const reviewObjs = () => reviewItems().map(x => ({ 행: x.행, 구분: x.구분, 키: x.키, 칸: x.칸, 사유: x.사유, 지금값: x.지금값, 근거: x.근거 }));

  $("#dl-clean").addEventListener("click", () => saveXlsx(`${S.base}_정리본.xlsx`, [["정리본", S.result.rows]]));
  $("#dl-log").addEventListener("click", () => saveXlsx(`${S.base}_변경기록.xlsx`, [["바뀐 칸", logObjs()], ["지운 줄", dropObjs()], ["적용한 처방", unitObjs()], ["정보", infoObjs()]]));
  $("#dl-review").addEventListener("click", () => saveXlsx(`${S.base}_검토필요목록.xlsx`, [["검토 필요", reviewObjs()]]));
  $("#dl-log-csv").addEventListener("click", () => saveCsv(`${S.base}_변경기록.csv`, logObjs().concat(S.result.dropped.map(d => ({ 행: d.row, 키: d.key, 칸: "(행 전체)", 변경전: "", 변경후: "지움", 처방: "중복", 이유: `행 ${d.sameAs}과 같음` })))));
  $("#dl-review-csv").addEventListener("click", () => saveCsv(`${S.base}_검토필요목록.csv`, reviewObjs()));

  // 켠 처방만 남긴 처방 파일 — 다음 파일에도 다시 쓸 수 있다
  const STEP_KEYS = { trim: ["trim"], map: ["map"], blank_to: ["blank_to"], strip_units: ["strip_units"], number: ["number"], date: ["date"], rescale: ["rescale"], missing: ["missing"], unparsed: ["unparsed"], range: ["range"] };
  function prunedRules() {
    const R = JSON.parse(JSON.stringify(S.rules));
    Object.entries(R.columns || {}).forEach(([col, r]) => {
      Object.entries(STEP_KEYS).forEach(([k, keys]) => { if (!S.enabled.has(`${col}|${k}`)) keys.forEach(x => { if (x in r && !(k === "date" && S.enabled.has(`${col}|unparsed`))) delete r[x]; }); });
      if (!Object.keys(r).some(k => k !== "why")) delete R.columns[col];
    });
    R.cross_checks = (R.cross_checks || []).filter((c, i) => S.enabled.has(`cc|${i}`));
    if (!S.enabled.has("dup")) delete R.duplicates;
    R.saved_note = `엑셀 정리기에서 켠 처방만 남겨 저장 (${new Date().toLocaleString("ko-KR")})`;
    return R;
  }
  $("#rules-save").addEventListener("click", () => {
    if (!S.rules) return;
    const ok = store.set(savedKey(), { rules: S.rules, enabled: [...S.enabled], saved_at: new Date().toISOString() });
    $("#rules-save").textContent = ok ? "이 브라우저에 저장했습니다 ✓" : "저장하지 못했습니다 (브라우저 설정)";
    $("#rules-saved").hidden = !ok;
  });
  $("#rules-export").addEventListener("click", () => {
    if (!S.rules) return;
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([JSON.stringify(prunedRules(), null, 2)], { type: "application/json" }));
    a.download = `${S.base}.json`; a.click();
  });
  $("#rules-saved").addEventListener("click", () => {
    const saved = store.get(savedKey());
    if (saved) { offerRules(saved.rules, "이 브라우저에 저장한 처방", saved.enabled); renderDiag(); }
  });
  $("#rules-file").addEventListener("change", async e => {
    const f = e.target.files[0]; e.target.value = "";
    if (!f) return;
    try { offerRules(JSON.parse(await f.text()), `내가 고른 파일 ${f.name}`); renderDiag(); }
    catch (err) { $("#rules-src").textContent = "처방 파일을 읽지 못했습니다 — JSON 형식을 확인하십시오"; }
  });
  $("#rules-groups").addEventListener("change", e => {
    const id = e.target.dataset.unit;
    if (!id) return;
    if (e.target.checked) S.enabled.add(id); else S.enabled.delete(id);
    e.target.closest(".rule").classList.toggle("off", !e.target.checked);
    updatePreview();
  });
  $("#rules-groups").addEventListener("click", e => {
    const g = e.target.dataset.all;
    if (!g) return;
    S.units.filter(u => u.group === g).forEach(u => { if (e.target.dataset.on === "1") S.enabled.add(u.id); else S.enabled.delete(u.id); });
    renderRules();
  });

  // 규정 파일 — 사용자 파일과 같은 방식으로 이 브라우저 안에서만 읽는다. 외부 기준을 인터넷에서 가져오지 않는다
  $("#policy-file").addEventListener("change", async e => {
    const f = e.target.files[0]; e.target.value = "";
    if (!f) return;
    try {
      const p = JSON.parse(await f.text());
      if (p.kind !== "travel_policy") throw new Error("kind");
      S.policy = p; S.policyName = f.name;
    } catch (err) { S.policy = null; $("#policy-info").textContent = "규정 파일을 읽지 못했습니다 — 양식(kind: travel_policy)을 확인하십시오"; return; }
    renderPolicyBox();
  });
  $("#currency").addEventListener("change", renderPolicyBox);
  $("#policy-template").addEventListener("click", () => {
    const t = {
      kind: "travel_policy", name: "회사 출장비 규정 (회사가 채워 넣음)", version: "", source: "규정 문서 이름 · 조항",
      valid_from: "YYYY-MM-DD", valid_to: "YYYY-MM-DD", currency: "KRW",
      items: [{ item: "숙박비", match: { place: "출장지 값 그대로" }, per: "night", max: null, currency: "KRW", valid_from: "", valid_to: "",
        exceptions: [{ label: "사전 승인", when: [["승인여부", "==", "Y"]] }], note: "max 가 비어 있으면 판단하지 않습니다. 부서 · 직무 차이는 규정에 실제로 있을 때만 match 에 넣습니다" }],
      external: [{ name: "GSA Per Diem", adopted: false, adopted_by: "회사 규정의 어느 조항이 채택했는지", source_url: "https://www.gsa.gov/travel/plan-book/per-diem-rates",
        year: "FY", region: "미국 본토(CONUS)", currency: "USD", rates: [],
        note: "회사 규정이 채택했을 때만 adopted:true. 한국 민간기업의 의무 기준이 아닙니다. 요율은 공식 출처에서 확인해 지역 · 기간별로 rates 에 적습니다" }],
      fx: [{ from: "USD", to: "KRW", rate: null, date: "YYYY-MM-DD", source: "환율 출처" }],
    };
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([JSON.stringify(t, null, 2)], { type: "application/json" }));
    a.download = "출장비규정_양식.json"; a.click();
  });

  // ── 정답 대조 ────────────────────────────────────────
  $("#answer-file").addEventListener("change", async e => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f || !S.result) return;
    // 정답 파일은 내 컴퓨터에서 고른다 — 사이트에 두지 않고, 어디로도 보내지 않는다
    const { rows } = readWorkbook(await f.arrayBuffer(), f.name);
    const mine = S.cols;
    const theirs = rows.length ? Object.keys(rows[0]) : [];
    const A = S.rules.answer || {};
    // 처방 파일에 대조 기준이 없으면: 첫 칸으로 같은 줄을 찾고, 두 파일에 다 있는 나머지 칸을 비교한다
    const key = A.key || mine.slice(0, 1);
    const cols = A.compare || mine.filter(c => !key.includes(c) && theirs.includes(c));
    const c = Clean.compare(S.result.rows, rows, key, cols);
    const diff = Object.entries(c.diff).map(([k, n]) => `${esc(k)} ${n}줄`).join(" · ") || "없음";
    $("#compare").innerHTML = `<p class="small muted">같은 줄 찾는 칸: ${esc(key.join(" + "))} · 비교한 칸: ${esc(cols.join(" · "))}</p>
      <div class="kpis">${[["정답 줄", fmt(c.answerRows)], ["일치", fmt(c.match), "", "dark"], ["정답에 없는 줄", fmt(c.notInAnswer)], ["어긋난 칸", diff, "", "red"]].map(kpi).join("")}</div>`
      + (c.examples.length ? `<p class="small">어긋난 예 — ${c.examples.slice(0, 6).map(x => `${esc(x.key.split("|")[0])} ${esc(x.col)}: ${esc(x.got)} (정답 ${esc(x.want)})`).join(" / ")}</p>` : "");
  });

  // ── 연결 ────────────────────────────────────────────
  $("#btn-apply").addEventListener("click", runApply);
  $("#file").addEventListener("change", async e => { const f = e.target.files[0]; e.target.value = ""; if (f) loadFile(await f.arrayBuffer(), f.name); });
  const drop = $("#drop");
  ["dragenter", "dragover"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", async e => { const f = e.dataTransfer.files[0]; if (f) loadFile(await f.arrayBuffer(), f.name); });
  $$("[data-sample]").forEach(b => b.addEventListener("click", async () => {
    const name = b.dataset.sample;
    const buf = await (await fetch("data/" + encodeURIComponent(name))).arrayBuffer();
    loadFile(buf, name);
  }));
  $("#diag-filter").addEventListener("click", e => { const b = e.target.closest("button"); if (!b) return; S.diagKind = b.dataset.k; renderDiag(); });
  $("#diag-q").addEventListener("input", drawDiagTable);
  $(".tabs").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    $$(".tabs button").forEach(x => x.setAttribute("aria-selected", String(x === b)));
    $$(".tabpane").forEach(p => { p.hidden = p.dataset.pane !== b.dataset.tab; });
    if (b.dataset.tab === "cmp") alignRows();
  });
  $("#cmp-filter").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    S.cmpMode = b.dataset.f;
    $$("#cmp-filter button").forEach(x => x.setAttribute("aria-pressed", String(x === b)));
    drawCompare();
  });
  $("#cmp-q").addEventListener("input", drawCompare);
  $("#log-q").addEventListener("input", drawLog);
  $("#rev-q").addEventListener("input", drawReview);
  $(".split").addEventListener("click", e => {
    const td = e.target.closest("td[data-c]"), tr = e.target.closest("tr[data-rn]");
    if (td && tr) showDetail(+tr.dataset.rn, td.dataset.c);
  });
  // 원본과 정리본 표를 같이 내려간다
  let syncing = false;
  ["#cmp-left", "#cmp-right"].forEach((s, i, all) => $(s).addEventListener("scroll", () => {
    if (syncing) return; syncing = true;
    $(all[1 - i]).scrollTop = $(s).scrollTop;
    requestAnimationFrame(() => { syncing = false; });
  }));
})();
