/* 엑셀 정리기 — 화면.
 *
 * 파일 → 읽기(SheetJS, 브라우저 안) → 합계 행 분리 → 진단 → 처방 선택(rules/*.json) → 확인 후 정리
 *      → 핵심 수치 · 금액 대사 → 원본과 정리본 비교 · 검토 · 중복 제외 → 자동 검증 → 내려받기
 * 서버로 아무것도 보내지 않는다. fetch 는 «내 사이트의» rules/ 와 data/ 를 읽을 때만 쓴다.
 * 사용자 파일 · 규정 파일 · 정답 파일은 이 브라우저 안에서만 읽고, 어디로도 보내지 않는다.
 * 처방은 rules/*.json 에만 있다. 이 파일에는 처방을 적지 않는다. 숫자는 모두 이 자리에서 다시 계산한다.
 */
(function () {
  "use strict";
  const $ = s => document.querySelector(s);
  const $$ = s => Array.from(document.querySelectorAll(s));
  const F = Clean.FLAG;
  const S = {
    name: "", base: "", rows: [], rowNos: [], summary: [], snapshot: "", sheets: [], cols: [],
    rules: null, rulesFrom: "", units: [], enabled: new Set(),
    pending: null,                      // 충돌을 확인받아야 하는 새 처방
    preview: null, result: null, idem: null, recon: [],
    policy: null, policyName: "", policyRes: [],
    review: new Map(),                  // 원본 행 → 처리 상태 (이 화면 안에서만, 저장하지 않음)
    diag: null, diagKind: "전체", cmpMode: "all", revMode: "all",
  };
  const esc = v => String(v === undefined || v === null ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;").replace(/"/g, "&quot;");
  const fmt = n => (typeof n === "number" ? n.toLocaleString("ko-KR") : n);
  const store = {
    get(k) { try { return JSON.parse(localStorage.getItem(k)); } catch (e) { return null; } },
    set(k, v) { try { localStorage.setItem(k, JSON.stringify(v)); return true; } catch (e) { return false; } },
  };
  const savedKey = () => `cleaner:rules:${S.base}`;
  const unitLabel = () => ($("#currency") && $("#currency").value) || "단위 미확인";

  // ── 읽기 ────────────────────────────────────────────
  // 엑셀의 각 줄이 원본 몇 행인지(__rowNum__)와 수식(SUM 등)을 함께 읽는다. 합계 행을 데이터와 나누기 위해서다
  function readWorkbook(buf, name) {
    const isCsv = /\.csv$/i.test(name);
    const wb = isCsv
      ? XLSX.read(new TextDecoder("utf-8").decode(buf).replace(/^﻿/, ""), { type: "string", raw: true })
      : XLSX.read(buf, { type: "array", cellDates: true, cellFormula: true });
    const ws = wb.Sheets[wb.SheetNames[0]];
    const rows = XLSX.utils.sheet_to_json(ws, { defval: "", raw: !isCsv });
    const rowNos = rows.map((r, i) => (r.__rowNum__ !== undefined ? r.__rowNum__ + 1 : i + 2));
    const pad = n => String(n).padStart(2, "0");
    rows.forEach(r => Object.keys(r).forEach(k => {
      const v = r[k];
      // 엑셀 날짜 칸은 Date 로 읽힌다 — 날짜 글자(YYYY-MM-DD)로 바꿔 둔다. 값은 그대로
      if (v instanceof Date && !isNaN(v)) r[k] = `${v.getFullYear()}-${pad(v.getMonth() + 1)}-${pad(v.getDate())}`;
    }));
    const formulas = {};
    if (!isCsv && ws["!ref"]) {
      const rg = XLSX.utils.decode_range(ws["!ref"]);
      const head = {};
      for (let c = rg.s.c; c <= rg.e.c; c++) { const h = ws[XLSX.utils.encode_cell({ r: rg.s.r, c })]; head[c] = h ? String(h.v) : ""; }
      const idxOfRow = new Map(rowNos.map((n, i) => [n, i]));
      Object.keys(ws).filter(k => k[0] !== "!").forEach(addr => {
        const cell = ws[addr];
        if (!cell || !cell.f) return;
        const a = XLSX.utils.decode_cell(addr);
        const i = idxOfRow.get(a.r + 1);
        if (i === undefined) return;
        (formulas[i] = formulas[i] || []).push({ col: head[a.c], addr, f: cell.f, v: cell.v });
      });
    }
    return { rows, rowNos, formulas, sheets: wb.SheetNames };
  }

  async function loadFile(buf, name) {
    const { rows, rowNos, formulas, sheets } = readWorkbook(buf, name);
    const sp = Clean.splitSummary(rows, formulas, rowNos);
    Object.assign(S, {
      name, base: name.replace(/\.(xlsx|xls|csv)$/i, ""), sheets, result: null, policyRes: [], pending: null,
      rows: sp.data, rowNos: sp.rowNos, summary: sp.summary, review: new Map(),
    });
    S.cols = rows.length ? Object.keys(rows[0]) : [];
    S.snapshot = JSON.stringify(S.rows);
    Object.freeze(S.rows); S.rows.forEach(r => Object.freeze(r));   // 원본은 바꾸지 않는다 — 정리본은 새로 만든다
    $("#file-info").hidden = false;
    $("#file-info").innerHTML = `${esc(name)} · 데이터 <b>${fmt(S.rows.length)}건</b> × ${S.cols.length}칸`
      + (S.summary.length ? ` · <span class="warn">합계 행 ${S.summary.length}줄은 데이터에서 빼고 따로 둡니다 — ${S.summary.map(s => `행 ${s.row}: ${esc(s.reason)}`).join(" / ")}</span>` : "")
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
    const d = S.diag = Clean.diagnose(S.rows, S.rules, { rowNos: S.rowNos });
    const n = k => d.items.filter(x => x.kind === k).reduce((a, x) => a + x.count, 0);
    $("#diag-kpi").innerHTML = [
      ["데이터 × 칸", `${fmt(d.n)} × ${d.cols}`, S.summary.length ? `합계 행 ${S.summary.length}줄 제외` : "머리줄 제외", "dark"],
      ["찾은 문제", fmt(d.items.length), "유형 · 칸별 항목 수", ""],
      ["영향 받는 줄", fmt(d.affectedRows), `데이터 ${fmt(d.n)}건 중`, "accent"],
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
    $("#diag-table").innerHTML = `<thead><tr>${["유형", "칸", "건수", "영향 받는 행 (원본 행 번호)", "예 (행: 값)", "설명"].map(h => `<th><span class="rz">${h}</span></th>`).join("")}</tr></thead><tbody>`
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
    ["값", "사용자가 승인한 값 바꾸기", "처방 파일에 사람이 적은 매핑표대로만 값을 바꿉니다 (예: «Sales» → «영업»). 표에 없는 값은 그대로 둡니다."],
    ["지움", "줄 지우기 — 금액에 영향", "지울 줄과 금액을 아래에 미리 보여 줍니다. 지운 줄은 «중복 제외»와 제외행목록에 남깁니다."],
    ["표시", "검토 표시 — 고치지 않음", "값을 바꾸지 않고 «검토 사유»를 적습니다. «탐지 기준»은 사용자가 정한 기준이지 회사 규정이 아닙니다."],
  ];

  function renderRules() {
    const has = !!S.rules;
    $("#rules-missing").hidden = has || !!S.pending;
    $("#rules-src").innerHTML = has
      ? `처방: <b>${esc(S.rules.title || "처방")}</b> · 정한 사람(승인) <b>${esc(S.rules.owner || "(적혀 있지 않음)")}</b> · ${esc(S.rulesFrom)}`
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
          <div class="col">${esc(u.col)}${u.severity === "무결성" ? `<div><span class="bd p1">우선 1 · 무결성</span></div>` : ""}${u.basis ? `<div><span class="bd basis" title="${esc(u.basis)}">탐지 기준 · 규정 아님</span></div>` : ""}</div>
          <div><div class="how">${esc(u.how)}</div><div class="why">이유 — ${esc(u.why || "(적혀 있지 않음)")}</div></div>
          <div class="imp" data-imp="${esc(u.id)}"></div></label>`).join("") + "</div>";
    }).join("");
    renderPolicyBox();
    updatePreview();
  }

  // 켠 처방으로 미리 돌려 본다 — 원본은 그대로, 결과는 화면에 쓰지 않는다
  function updatePreview() {
    const btn = $("#btn-apply");
    if (!S.rules) { $("#confirm-text").innerHTML = `<b>처방이 없어 정리할 수 없습니다.</b>`; btn.disabled = true; $("#drop-preview").hidden = true; return; }
    const p = S.preview = Clean.apply(S.rows, S.rules, { enabled: [...S.enabled], rowNos: S.rowNos });
    const s = p.summary;
    S.units.forEach(u => {
      const el = document.querySelector(`[data-imp="${CSS.escape(u.id)}"]`);
      if (!el) return;
      const n = s.byRule[u.id] || 0;
      const unit = u.group === "표시" ? "건 표시" : u.group === "지움" ? "줄 지움" : "칸 바뀜";
      el.innerHTML = S.enabled.has(u.id) ? `예상 <b>${fmt(n)}</b> ${unit}` : `<span class="muted">꺼짐</span>`;
    });
    // 지울 줄을 실행 전에 보여 준다 — 신청번호와 금액
    const rc = reconCols();
    $("#drop-preview").hidden = !p.dropped.length;
    if (p.dropped.length) {
      const tot = c => p.dropped.reduce((a, d) => a + (typeof d.cleaned[c] === "number" ? d.cleaned[c] : 0), 0);
      $("#drop-preview").innerHTML = `<h3>실행하면 정리본에서 뺄 줄 ${p.dropped.length}건 <span class="tag">처방: 중복</span></h3>
        <p class="small muted">기준: 처방을 적용한 뒤 모든 칸이 앞 줄과 같은 줄. 금액 합계에 영향을 줍니다. 같은 신청이 아니라면 «줄 지우기» 처방을 끄세요.</p>
        <div class="tablebox"><table class="grid"><thead><tr><th class="rn">원본 행</th><th>${esc(S.cols[0])}</th><th>같은 줄</th>${rc.map(c => `<th class="r">${esc(c)}</th>`).join("")}</tr></thead><tbody>`
        + p.dropped.map(d => `<tr><td class="rn">${d.row}</td><td>${esc(d.key)}</td><td>행 ${d.sameAs}</td>${rc.map(c => `<td class="num">${fmt(d.cleaned[c])}</td>`).join("")}</tr>`).join("")
        + `<tr class="total"><td class="rn"></td><td colspan="2">빠지는 금액 합계 (${esc(unitLabel())})</td>${rc.map(c => `<td class="num">${fmt(tot(c))}</td>`).join("")}</tr></tbody></table></div>`;
    }
    $("#confirm-text").innerHTML = `켠 처방 <b>${S.enabled.size}</b> / ${S.units.length}개 · 예상: 데이터 ${fmt(s.before)}건 → <b>${fmt(s.after)}건</b> ·
      값 <b>${fmt(s.changes)}셀</b>(${fmt(s.changedRows)}행) 바뀜 · <b>${fmt(s.dropped)}건</b> 제외 · <b>${fmt(s.flagged)}건</b> 검토 표시<br><span class="muted">원본은 그대로 두고 새 정리본을 만듭니다. 실행은 업무 승인이 아닙니다.</span>`;
    btn.disabled = !!S.pending;
  }

  // 금액 대사에 쓸 칸 — 합계 행에 숫자가 있던 칸, 합계 행이 없으면 처방에서 숫자로 바꾸는 칸
  function reconCols() {
    if (!S.rules) return [];
    const fromSummary = S.cols.filter(c => S.summary.some(s => typeof s.values[c] === "number"));
    if (fromSummary.length) return fromSummary;
    return Object.entries(S.rules.columns || {}).filter(([c, r]) => r.number && S.cols.includes(c)).map(([c]) => c);
  }

  // ── 규정 확인 (처방 파일에 policy_fields 가 있을 때만) ──────
  const NEEDS = [
    ["출장 국가 · 도시", f => f.place], ["주 또는 지역", f => f.region], ["출장 시작일", f => f.date], ["출장 종료일 · 숙박일별 날짜", f => f.end],
    ["숙박일(박수)", f => f.nights], ["부서", f => f.dept], ["직무", f => f.job], ["비용항목", f => f.items && Object.keys(f.items).join(" · ")],
    ["교통 수단 구분(항공 · 철도 · 택시)", f => f.transport], ["숙박비 세금 포함 여부", f => f.tax],
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
      return `<span class="bd ${v ? "ok" : "warn"}" title="${esc(v || "없음")}">${v ? "✓ 있음" : "– 없음"} · ${esc(k)}</span>`;
    }).join("")}</div><p class="muted">없는 칸은 추정하지 않습니다. 판단에 필요한데 없으면 «판단 불가»로 남기고 어떤 정보가 없는지 적습니다.</p>`;
    const p = S.policy;
    $("#policy-info").innerHTML = p
      ? `<b>${esc(S.policyName)}</b> · ${esc(p.name || "")} ${p.version ? "· 버전 " + esc(p.version) : ""} · 회사 규정 ${(p.items || []).length}개 ·
         채택한 외부 기준 ${(p.external || []).filter(e => e.adopted === true).length}개 (등록 ${(p.external || []).length}개)`
      : "아직 고르지 않았습니다. 규정 없이는 금액이 적정한지 판단하지 않습니다 (규정 미설정).";
  }

  // ── 4 정리 실행 ─────────────────────────────────────
  function runApply() {
    if (!S.rules || S.pending) return;
    const opts = { enabled: [...S.enabled], rowNos: S.rowNos };
    const r1 = Clean.apply(S.rows, S.rules, opts);
    const r2 = Clean.apply(r1.rows, S.rules, { enabled: opts.enabled, rowNos: r1.src });   // 한 번 더 — 두 번 돌려도 같은가
    S.result = r1;
    S.idem = Clean.same(r1.rows, r2.rows) && r2.summary.changes === 0 && r2.summary.dropped === 0;
    S.recon = Clean.reconcile(S.rows, r1, reconCols(), S.summary);
    const pf = S.rules.policy_fields;
    // 규정 파일이 없으면 비용 적정성은 판단하지 않는다 — 금액 크기만으로 판단하지 않는다
    S.policyRes = pf && S.policy ? Clean.checkPolicy(r1.rows, r1.src, S.policy, { fields: pf, currency: $("#currency").value }) : [];
    S.ranAt = new Date();
    S.review = new Map();
    buildMaps();
    $("#s-result").hidden = false; $("#s-down").hidden = false;   // 먼저 보이게 해야 줄 높이를 잴 수 있다
    renderResult();
    renderDown();
    markSteps(5);
    $("#s-result").scrollIntoView();
  }

  let M = {};
  function buildMaps() {
    const r = S.result;
    M = { chg: new Map(), flagRow: new Map(), flagCell: new Map(), drop: new Map(), clean: new Map(), orig: new Map() };
    r.cells.forEach(l => M.chg.set(`${l.row}|${l.col}`, l));
    r.flags.forEach(f => {
      (M.flagRow.get(f.row) || M.flagRow.set(f.row, []).get(f.row)).push(f);
      f.col.split(" · ").forEach(c => (M.flagCell.get(`${f.row}|${c}`) || M.flagCell.set(`${f.row}|${c}`, []).get(`${f.row}|${c}`)).push(f));
    });
    r.dropped.forEach(d => M.drop.set(d.row, d));
    r.rows.forEach((row, i) => M.clean.set(r.src[i], row));
    S.rows.forEach((row, i) => M.orig.set(S.rowNos[i], row));
  }

  function kpi([k, v, s, cls, go]) {
    const tag = go ? "button" : "div";
    return `<${tag} class="kpi ${cls || ""}${go ? " link" : ""}" ${go ? `data-go="${go}"` : ""}><div class="k">${k}</div><div class="v">${v}</div><div class="s">${s || ""}</div></${tag}>`;
  }

  // 검토 목록 — 실제 데이터 줄만, 한 줄에 하나. 우선순위 1 = 금액 · 데이터 무결성, 2 = 담당자 확인
  function reviewItems() {
    const out = new Map();
    const add = (row, o) => {
      const cur = out.get(row) || { 행: row, 신청번호: o.key, 우선: 2, 유형: [], 현재값: [], 사유: [], 근거: [], 기준: [] };
      if (o.severity === "무결성") cur.우선 = 1;
      cur.유형.push(o.kind); cur.사유.push(o.reason); if (o.need && !cur.근거.includes(o.need)) cur.근거.push(o.need);
      if (o.value && !cur.현재값.includes(o.value)) cur.현재값.push(o.value);
      if (o.basis && !cur.기준.includes(o.basis)) cur.기준.push(o.basis);
      out.set(row, cur);
    };
    const cur = (rn, col) => col.split(" · ").map(c => `${c}=${fmtCell((M.clean.get(rn) || {})[c])}`).join(", ");
    S.result.flags.forEach(f => {
      const u = S.units.find(x => x.id === f.rule) || {};
      add(f.row, { key: f.key, severity: f.severity, kind: u.basis ? "탐지 기준(사용자 설정)" : u.step === "cross" ? "칸 사이 모순" : u.step === "missing" ? "빈칸" : u.step === "range" ? "범위 밖" : u.step === "unparsed" ? "날짜 못 읽음" : "검토", reason: f.reason, need: f.need, value: cur(f.row, f.col), basis: u.basis });
    });
    S.policyRes.filter(p => p.status !== "한도 안").forEach(p => add(p.row, {
      key: p.key, severity: "확인", kind: p.status === "판단 불가" ? "규정 미판정" : p.status === "예외 승인 조건 해당" ? "규정 예외 조건" : "규정 한도 초과(확인 필요)",
      reason: `${p.item}: ${p.reason}`, need: p.status === "판단 불가" ? "부족한 정보를 채우거나 규정 확인" : "승인 · 증빙 확인", value: `${p.item}=${fmtCell(p.amount)}`, basis: p.basis,
    }));
    return [...out.values()].sort((a, b) => a.우선 - b.우선 || a.행 - b.행).map(x => Object.assign(x, { 상태: S.review.get(x.행) || "검토 대기" }));
  }
  const unitText = id => { const u = S.units.find(x => x.id === id); return u ? `${u.col}: ${u.how}` : id; };
  const fmtCell = v => (v === "" || v === null || v === undefined ? "(빈칸)" : typeof v === "number" ? fmt(v) : v);

  function renderResult() {
    const r = S.result, s = r.summary;
    const rev = reviewItems();
    const p1 = rev.filter(x => x.우선 === 1).length;
    const reconOk = S.recon.every(x => x.ok);
    const done = rev.filter(x => x.상태 === "확인 완료").length;
    const pol = S.policyRes;
    // 상태 — 기계적 정리와 업무 승인은 다르다
    $("#status").innerHTML = [
      `<span class="st ok">✓ 정리 실행 완료</span>`,
      rev.length ? `<span class="st warn">⚠ 업무 검토 필요 ${fmt(rev.length)}건${p1 ? ` (무결성 ${p1})` : ""}</span>` : `<span class="st ok">검토 필요 없음</span>`,
      S.rules.policy_fields ? (S.policy ? `<span class="st info">규정 확인: ${esc(S.policyName)}</span>` : `<span class="st grey">규정 적합성 미판정 (규정 미설정)</span>`) : "",
      `<span class="st grey">승인 대기 — 확인 ${done}/${rev.length}</span>`,
      S.recon.length ? `<span class="st ${reconOk ? "ok" : "bad"}">${reconOk ? "✓ 합계 검증 통과" : "✕ 합계 검증 실패"}</span>` : "",
    ].join("");
    $("#res-kpi").innerHTML = [
      ["원본 데이터", `${fmt(s.before)}건`, S.summary.length ? `합계 행 ${S.summary.length}줄 별도` : "머리줄 제외", "dark"],
      ["정리 후", `${fmt(s.after)}건`, "정리본 데이터", ""],
      ["변경", `${fmt(s.changes)}셀`, `${fmt(s.changedRows)}행에서`, "", "log"],
      ["중복 제외", `${fmt(s.dropped)}건`, "목록과 금액 보기", "", "drop"],
      ["검토 필요", `${fmt(rev.length)}건`, `실제 데이터 줄 · 처방 표시 ${fmt(s.flagged)}${pol.length ? ` + 규정` : ""}`, "amber", "review"],
      ["합계 검증", S.recon.length ? (reconOk ? "통과" : "실패") : "대상 없음", "금액 대사 보기", reconOk ? "okc" : "bad", "recon"],
    ].map(kpi).join("");
    renderRecon();
    $("#idem").innerHTML = (S.idem
      ? `<b class="pass">✓ 두 번 돌려도 같은 결과</b> — 정리본에 같은 처방을 한 번 더 적용해도 바뀌는 것이 없습니다.`
      : `<b class="fail">✕ 두 번 돌리면 결과가 달라집니다</b> — 처방이 자기 결과를 또 바꿉니다. 처방을 다시 보십시오.`)
      + (S.rules.policy_fields ? (S.policy
        ? `<br><b>비용 적정성 (${esc(S.policyName)})</b> — ${["한도 안", "한도 초과 — 검토 필요", "예외 승인 조건 해당", "판단 불가"].map(st => `${st} ${pol.filter(p => p.status === st).length}`).join(" · ")} (항목 수)`
        : `<br><b>비용 적정성은 판단하지 않았습니다</b> — 회사 규정 파일이 없습니다. 금액이 크다는 이유만으로 규정 위반으로 보지 않습니다.`) : "");
    const counts = { log: s.changes, review: rev.length, drop: s.dropped };
    $$("#res-tabs button").forEach(b => {
      if (!b.dataset.label) b.dataset.label = b.textContent.trim();
      const c = counts[b.dataset.tab];
      b.innerHTML = esc(b.dataset.label) + (c !== undefined ? `<span class="c">${fmt(c)}</span>` : "");
    });
    drawCompare();
    drawTable("#t-orig", S.rows.map((row, i) => ({ rn: S.rowNos[i], row })), S.cols, false);
    drawTable("#t-clean", r.rows.map((row, i) => ({ rn: r.src[i], row })), S.cols.concat(F), true);
    drawLog(); drawReview(); drawDrop();
    $("#compare").innerHTML = ""; $("#cell-detail").hidden = true;
  }

  // 금액 대사표 — 원본 합계 행 → 숫자 칸만 → 처방 적용 후 → 중복 제외 → 정리 후
  function renderRecon() {
    const R = S.recon;
    if (!R.length) { $("#recon").innerHTML = `<h3>금액 대사</h3><p class="muted small">숫자로 다루는 칸이 없어 금액 대사를 하지 않았습니다.</p>`; return; }
    const line = (label, get, cls, note) => `<tr class="${cls || ""}"><td>${label}${note ? `<div class="small muted">${note}</div>` : ""}</td>${R.map(x => `<td class="num">${get(x)}</td>`).join("")}</tr>`;
    const sgn = n => (n > 0 ? "+" : n < 0 ? "−" : "") + fmt(Math.abs(n));
    $("#recon").innerHTML = `<h3 id="recon-h">금액 대사 <span class="tag">단위: ${esc(unitLabel())}</span></h3>
      <p class="small muted">파일에 통화 칸이 없으면 단위를 정하지 않습니다. 모든 숫자는 데이터 줄에서 다시 계산했습니다 — 원본 합계 행 값은 비교용으로만 씁니다.</p>
      <div class="tablebox"><table class="grid recon"><thead><tr><th>단계</th>${R.map(x => `<th class="r">${esc(x.col)}</th>`).join("")}</tr></thead><tbody>`
      + (S.summary.length ? line("① 원본 파일의 합계 행 값", x => x.fileTotal === null ? "—" : fmt(x.fileTotal), "", `행 ${S.summary.map(s => s.row).join(", ")} · ${esc(S.summary[0].reason)}`) : "")
      + line("② 원본 데이터 중 숫자로 저장된 값만 더한 합계", x => fmt(x.numbersOnly), "", "엑셀 SUM 이 보는 값")
      + (S.summary.length ? line("&nbsp;&nbsp;①과 ②", x => x.fileTotalMatchesNumbers === null ? "—" : x.fileTotalMatchesNumbers ? "같음" : `<b class="fail">다름</b>`, "sub") : "")
      + line("③ 처방으로 더해진 금액 (글자 금액 → 숫자 등)", x => sgn(x.formatDelta) + (x.textInOriginal ? `<div class="small muted">글자 금액 ${x.textInOriginal}건</div>` : ""), "sub")
      + line("④ 처방 적용 후 전체 합계 (② + ③, 지우기 전)", x => fmt(x.afterRules), "strong")
      + line("⑤ 중복으로 뺀 줄의 금액", x => "−" + fmt(x.droppedAmt), "sub", `${S.result.dropped.length}건`)
      + line("⑥ 정리 후 합계 (정리본 데이터 줄에서 다시 계산)", x => fmt(x.cleaned), "strong total")
      + line("검증: ④ − ⑤ = ⑥", x => (x.ok ? `<b class="pass">✓ 통과</b>` : `<b class="fail">✕ 실패</b>`), "sub")
      + `</tbody></table></div>`
      + (S.summary.length && R.some(x => x.fileTotal !== null && x.fileTotal !== x.cleaned)
        ? `<p class="small">원본 합계 행 값(①)은 정리본에 옮기지 않습니다. 정리본.xlsx 의 합계는 정리된 데이터 줄을 더하는 <b>SUM 수식으로 새로 씁니다</b>(⑥).</p>` : "");
  }

  // 값 하나를 칸에 — 숫자는 쉼표 · 오른쪽, 날짜는 고정폭, 숫자처럼 보이는 글자는 «글자» 표시, 앞뒤 공백은 ␣
  function cell(v, extraCls, attrs, icon) {
    let cls = extraCls || "", text, tag = "";
    if (v === "" || v === null || v === undefined) { cls += " blank"; text = "(빈칸)"; }
    else if (typeof v === "number") { cls += " num"; text = fmt(v); }
    else {
      const s = String(v);
      if (/^\d{4}[-/.]\d{1,2}[-/.]\d{1,2}$/.test(s.trim())) cls += " date";
      else if (/^\s*-?[\d,]+(\.\d+)?\s*(원|개|박)?\s*$/.test(s)) { cls += " num"; tag = `<span class="txt">글자</span>`; }
      text = s !== s.trim() ? s.replace(/^\s+|\s+$/g, m => "␣".repeat(m.length)) : s;
    }
    return `<td class="${cls.trim()}" ${attrs || ""}>${icon ? `<span class="ico" title="검토 필요">⚠</span>` : ""}${tag}${esc(text)}</td>`;
  }
  const head = (cols, labels) => `<thead><tr><th class="rn">원본 행</th>${cols.map((c, i) => `<th class="${i === 0 ? "key" : ""}"><span class="rz">${esc(labels ? labels[i] : c)}</span></th>`).join("")}</tr></thead>`;
  const keyCell = (v, cls, attrs, icon) => cell(v, (cls || "") + " key", attrs, icon);

  function drawTable(sel, list, cols, isClean) {
    const max = 3000;
    const labels = cols.map(c => (c === F ? "검토 사유" : c));
    $(sel).innerHTML = `<table class="grid">${head(cols, labels)}<tbody>${list.slice(0, max).map(({ rn, row }) =>
      `<tr><td class="rn">${rn}</td>${cols.map((c, i) => c === F ? `<td class="note">${esc(row[c])}</td>`
        : (i === 0 ? keyCell : cell)(row[c], isClean && M.chg.has(`${rn}|${c}`) ? "chg" : "", "", isClean && M.flagCell.has(`${rn}|${c}`))).join("")}</tr>`).join("")}</tbody></table>`
      + (list.length > max ? `<p class="empty">앞 ${fmt(max)}줄만 보여 줍니다. 전체는 내려받아 보세요.</p>` : "");
  }

  // 원본과 정리본을 «원본 행 번호(+신청번호)»로 맞춰 나란히 — 중복을 지워도 줄이 밀리지 않는다
  function drawCompare() {
    const q = $("#cmp-q").value.trim().toLowerCase();
    const cols = S.cols;
    const list = S.rowNos.filter(rn => {
      const changed = cols.some(c => M.chg.has(`${rn}|${c}`)), flagged = M.flagRow.has(rn), dropped = M.drop.has(rn);
      const ok = S.cmpMode === "all" || (S.cmpMode === "changed" && changed) || (S.cmpMode === "flagged" && flagged) || (S.cmpMode === "dropped" && dropped);
      if (!ok) return false;
      if (!q) return true;
      const o = M.orig.get(rn), c = M.clean.get(rn) || {};
      return String(rn) === q || cols.some(k => String(o[k]).toLowerCase().includes(q) || String(c[k] ?? "").toLowerCase().includes(q));
    });
    const max = 2000, shown = list.slice(0, max);
    const left = `<table class="grid">${head(cols)}<tbody>${shown.map(rn => {
      const o = M.orig.get(rn), d = M.drop.has(rn);
      return `<tr data-rn="${rn}" class="${d ? "del" : ""}"><td class="rn">${rn}</td>${cols.map((c, i) => (i === 0 ? keyCell : cell)(o[c], M.chg.has(`${rn}|${c}`) ? "chg" : "", `data-c="${esc(c)}"`)).join("")}</tr>`;
    }).join("")}</tbody></table>`;
    const right = `<table class="grid">${head(cols.concat("상태 · 검토 사유"))}<tbody>${shown.map(rn => {
      const d = M.drop.get(rn);
      if (d) return `<tr class="del" data-rn="${rn}"><td class="rn">${rn}</td>${cols.map((c, i) => (i === 0 ? keyCell : cell)(M.orig.get(rn)[c])).join("")}<td class="note"><span class="bd 중복">중복 제외</span> 행 ${d.sameAs}과 같음</td></tr>`;
      const c = M.clean.get(rn) || {};
      const fl = M.flagRow.get(rn) || [];
      const pol = S.policyRes.filter(p => p.row === rn && p.status !== "한도 안");
      const why = fl.map(f => f.reason).concat(pol.map(p => `${p.item}: ${p.status}`));
      const ch = cols.some(k => M.chg.has(`${rn}|${k}`));
      const st = (ch ? `<span class="bd chgb">변경</span> ` : "") + (why.length ? `<span class="bd warn">⚠ 검토</span> ${esc(why.join(" · "))}` : "");
      return `<tr data-rn="${rn}"><td class="rn">${rn}</td>${cols.map((k, i) => (i === 0 ? keyCell : cell)(c[k], M.chg.has(`${rn}|${k}`) ? "chg" : "", `data-c="${esc(k)}"`, M.flagCell.has(`${rn}|${k}`))).join("")}<td class="note">${st}</td></tr>`;
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
    d.innerHTML = `<dl><dt>원본 행 · 칸</dt><dd>${rn} · ${esc((M.orig.get(rn) || {})[S.cols[0]])} · <b>${esc(col)}</b></dd>`
      + (l ? `<dt>원본 값</dt><dd class="from">${esc(fmtCell(l.from))}${typeof l.from === "string" ? " <span class='muted'>(글자)</span>" : ""}</dd><dt>정리 값</dt><dd class="to">${esc(fmtCell(l.to))}${typeof l.to === "number" ? " <span class='muted'>(숫자)</span>" : ""}</dd>
        <dt>적용 처방</dt><dd>${l.rules.map(id => `<code>${esc(id)}</code> ${esc(unitText(id))}`).join("<br>")}</dd><dt>변경 이유</dt><dd>${esc(l.rule)}</dd>` : "")
      + (fl.length ? `<dt>검토 필요</dt><dd>${fl.map(f => `${esc(f.reason)} <span class="muted">— ${esc(f.need || "")}</span>`).join("<br>")}<br><b>값은 고치지 않았습니다.</b></dd>` : "")
      + "</dl>";
  }

  function simpleTable(sel, heads, rows, q, opts) {
    const ql = (q || "").trim().toLowerCase();
    const list = ql ? rows.filter(r => r.some(v => String(v).toLowerCase().includes(ql))) : rows;
    const noteCols = new Set(["이유", "사유", "근거", "처방", "필요한 근거", "기준"]);
    $(sel).innerHTML = list.length
      ? `<table class="grid"><thead><tr>${heads.map((h, i) => `<th class="${i === 0 ? "rn" : ""}"><span class="rz">${esc(h)}</span></th>`).join("")}</tr></thead><tbody>${list.slice(0, 3000).map(r =>
        `<tr>${r.map((v, i) => i === 0 ? `<td class="rn">${esc(v)}</td>` : (opts && opts.html && opts.html.includes(i)) ? `<td>${v}</td>` : noteCols.has(heads[i]) ? `<td class="note">${esc(v)}</td>` : cell(v)).join("")}</tr>`).join("")}</tbody>${opts && opts.foot ? opts.foot : ""}</table>`
      : `<p class="empty">${rows.length ? "찾는 조건에 맞는 것이 없습니다." : "없습니다."}</p>`;
  }
  const logRows = () => S.result.cells.map(l => [l.row, l.key, l.col, l.from, l.to, l.rules.join(", "), l.rule]);
  function drawLog() { simpleTable("#t-log", ["원본 행", "신청번호(첫 칸)", "칸", "원본 값", "정리 값", "처방", "이유"], logRows(), $("#log-q").value); }
  function drawReview() {
    const all = reviewItems();
    const list = all.filter(x => S.revMode === "all" || String(x.우선) === S.revMode);
    const sel = x => `<select data-rv="${x.행}">${["검토 대기", "확인 완료", "보류"].map(o => `<option ${o === x.상태 ? "selected" : ""}>${o}</option>`).join("")}</select>`;
    simpleTable("#t-review", ["원본 행", "우선순위", "신청번호", "문제 유형", "현재 값", "사유", "필요한 근거", "처리 상태"],
      list.map(x => [x.행, `<span class="bd ${x.우선 === 1 ? "p1" : "p2"}">${x.우선 === 1 ? "1 · 무결성" : "2 · 담당자 확인"}</span>`, x.신청번호, x.유형.join(" · "), x.현재값.join(" / "),
        x.사유.join(" · ") + (x.기준.length ? " — 탐지 기준(사용자 설정, 규정 아님)" : ""), x.근거.join(" · "), sel(x)]),
      $("#rev-q").value, { html: [1, 7] });
    const done = all.filter(x => x.상태 === "확인 완료").length;
    $("#rev-progress").textContent = `확인 완료 ${done} / ${all.length}건 · 처리 상태는 이 화면에만 있고 저장되지 않습니다. 검토필요목록.xlsx 에 함께 내려받습니다.`;
  }
  function drawDrop() {
    const rc = reconCols();
    const d = S.result.dropped;
    const foot = d.length ? `<tfoot><tr class="total"><td class="rn"></td><td colspan="2">빠진 금액 합계 (${esc(unitLabel())})</td>${rc.map(c => `<td class="num">${fmt(d.reduce((a, x) => a + (typeof x.cleaned[c] === "number" ? x.cleaned[c] : 0), 0))}</td>`).join("")}<td colspan="${S.cols.length}"></td></tr></tfoot>` : "";
    simpleTable("#t-drop", ["원본 행", S.cols[0], "같은 줄", ...rc.map(c => `${c} (정리 값)`), ...S.cols.map(c => `원본 ${c}`)],
      d.map(x => [x.row, x.key, `행 ${x.sameAs}`, ...rc.map(c => x.cleaned[c]), ...x.values]), "", { foot });
  }

  // ── 5 자동 검증과 내려받기 ───────────────────────────
  // 원본과 정리본을 직접 다시 비교해 셀 수를 센다 — 엔진의 기록과 따로 센 값이 같아야 한다
  function verify() {
    const r = S.result, s = r.summary;
    let recount = 0;
    r.rows.forEach((row, i) => { const o = M.orig.get(r.src[i]); S.cols.forEach(c => { if (String(o[c]) !== String(row[c]) && !(Clean.isBlank(o[c]) && Clean.isBlank(row[c]))) recount++; }); });
    const rev = reviewItems();
    const flaggedRows = new Set(r.flags.map(f => f.row));
    const summaryRowNos = new Set(S.summary.map(x => x.row));
    const sheetTotals = reconCols().map(c => [c, r.rows.reduce((a, x) => a + (typeof x[c] === "number" ? x[c] : 0), 0)]);
    const checks = [
      ["행 수 대사", `원본 데이터 ${fmt(s.before)} = 정리본 ${fmt(s.after)} + 제외 ${fmt(s.dropped)}`, s.before === s.after + s.dropped && r.rows.length === s.after],
      ["변경 기록 대사", `변경기록 ${fmt(r.cells.length)}셀 = 원본과 정리본을 다시 비교해 센 ${fmt(recount)}셀`, r.cells.length === recount],
      ["제외 목록 대사", `제외행목록 ${fmt(r.dropped.length)}건 = 줄 수 차이 ${fmt(s.before - s.after)}`, r.dropped.length === s.before - s.after],
      ["검토 목록 대사", `검토필요목록 ${fmt(rev.length)}건 ⊇ 처방 표시 ${fmt(flaggedRows.size)}건, 합계 행 포함 0건`, [...flaggedRows].every(x => rev.some(y => y.행 === x)) && !rev.some(y => summaryRowNos.has(y.행))],
      ["금액 대사", S.recon.length ? S.recon.map(x => `${x.col} ${fmt(x.afterRules)} − ${fmt(x.droppedAmt)} = ${fmt(x.cleaned)}`).join(" · ") : "금액 칸 없음", S.recon.every(x => x.ok)],
      ["정리본 합계 = 화면 합계", sheetTotals.map(([c, t]) => `${c} ${fmt(t)}`).join(" · ") || "금액 칸 없음", sheetTotals.every(([c, t]) => { const x = S.recon.find(y => y.col === c); return !x || x.cleaned === t; })],
      ["원본 불변", "읽은 원본 데이터가 처음과 같음", JSON.stringify(S.rows) === S.snapshot],
    ];
    return checks;
  }
  function renderDown() {
    const checks = S.checks = verify();
    const ok = checks.every(c => c[2]);
    $("#verify").innerHTML = `<div class="vhead ${ok ? "ok" : "bad"}">${ok ? `✓ 내려받기 전 자동 검증 통과 (${checks.length}/${checks.length})` : `✕ 자동 검증 실패 — 내려받은 파일을 최종 자료로 쓰지 마세요`}
      <span class="muted"> · 검증 통과는 «숫자가 서로 맞는다»는 뜻이지 업무 승인이 아닙니다 (승인 대기)</span></div>
      <ul>${checks.map(([k, v, pass]) => `<li class="${pass ? "pass" : "fail"}"><b>${pass ? "✓" : "✕"} ${esc(k)}</b> — ${esc(v)}</li>`).join("")}</ul>`;
    $$(".dcard").forEach(b => b.classList.toggle("failed", !ok));
  }

  const stamp = () => (S.ranAt ? S.ranAt.toLocaleString("ko-KR") : "");
  const excelSerial = iso => { const [y, m, d] = iso.split("-").map(Number); return (Date.UTC(y, m - 1, d) - Date.UTC(1899, 11, 30)) / 86400000; };
  const ISO = /^\d{4}-\d{2}-\d{2}$/;
  const validISO = s => { if (!ISO.test(s)) return false; const [y, m, d] = s.split("-").map(Number); const t = new Date(Date.UTC(y, m - 1, d)); return t.getUTCMonth() === m - 1 && t.getUTCDate() === d; };

  /* 형식이 있는 시트를 만든다: 숫자 #,##0 · 날짜는 진짜 날짜형(yyyy-mm-dd) · 필터 · 열 너비 · (필요하면) SUM 합계 줄
   * opts.dateCols: 날짜형으로 쓸 칸 이름 · opts.totalCols: SUM 줄을 쓸 칸 이름 */
  function makeSheet(headers, rows, opts) {
    opts = opts || {};
    const aoa = [headers].concat(rows.map(r => headers.map(h => (r[h] === undefined ? "" : r[h]))));
    const ws = XLSX.utils.aoa_to_sheet(aoa);
    const n = rows.length;
    headers.forEach((h, c) => {
      for (let r = 1; r <= n; r++) {
        const a = XLSX.utils.encode_cell({ r, c }), cellObj = ws[a];
        if (!cellObj) continue;
        if (cellObj.t === "n") cellObj.z = Number.isInteger(cellObj.v) ? "#,##0" : "#,##0.##";
        else if (opts.dateCols && opts.dateCols.includes(h) && cellObj.t === "s" && validISO(cellObj.v)) {
          ws[a] = { t: "n", v: excelSerial(cellObj.v), z: "yyyy-mm-dd" };
        }
      }
    });
    if (opts.totalCols && opts.totalCols.length && n) {
      const tr = n + 1;
      ws[XLSX.utils.encode_cell({ r: tr, c: 0 })] = { t: "s", v: "합계" };
      opts.totalCols.forEach(h => {
        const c = headers.indexOf(h); if (c < 0) return;
        const col = XLSX.utils.encode_col(c);
        const v = rows.reduce((a, r) => a + (typeof r[h] === "number" ? r[h] : 0), 0);
        ws[XLSX.utils.encode_cell({ r: tr, c })] = { t: "n", f: `SUM(${col}2:${col}${n + 1})`, v, z: "#,##0" };
      });
      ws["!ref"] = XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: tr, c: headers.length - 1 } });
    }
    ws["!autofilter"] = { ref: XLSX.utils.encode_range({ s: { r: 0, c: 0 }, e: { r: Math.max(n, 1), c: headers.length - 1 } }) };
    ws["!cols"] = headers.map((h, c) => {
      const w = Math.max(String(h).length * 2, ...rows.slice(0, 500).map(r => { const v = r[h]; return typeof v === "number" ? fmt(v).length + 2 : String(v ?? "").length * 1.6; }));
      return { wch: Math.min(Math.max(Math.ceil(w) + 2, 8), 60) };
    });
    return ws;
  }

  // SheetJS 무료판은 틀 고정을 쓰지 못한다 — 만든 xlsx(zip) 안의 시트 XML 에 «1행 고정»을 넣는다. 모두 이 브라우저 안에서
  function freezeHeader(buf, sheetCount) {
    try {
      const cfb = XLSX.CFB.read(new Uint8Array(buf), { type: "array" });
      for (let i = 1; i <= sheetCount; i++) {
        const e = XLSX.CFB.find(cfb, `/xl/worksheets/sheet${i}.xml`);
        if (!e) continue;
        let xml = new TextDecoder().decode(e.content);
        const pane = `<sheetView workbookViewId="0"><pane ySplit="1" topLeftCell="A2" activePane="bottomLeft" state="frozen"/><selection pane="bottomLeft"/></sheetView>`;
        if (/<sheetViews>[\s\S]*?<\/sheetViews>/.test(xml)) xml = xml.replace(/<sheetViews>[\s\S]*?<\/sheetViews>/, `<sheetViews>${pane}</sheetViews>`);
        else xml = xml.replace(/(<dimension[^>]*\/>)/, `$1<sheetViews>${pane}</sheetViews>`);
        e.content = new TextEncoder().encode(xml); e.size = e.content.length;
      }
      return XLSX.CFB.write(cfb, { fileType: "zip", type: "array" });
    } catch (err) { return buf; }   // 실패하면 고정 없이 내보낸다 (데이터는 같다)
  }
  function saveBook(name, sheets) {
    const wb = XLSX.utils.book_new();
    sheets.forEach(([title, ws]) => XLSX.utils.book_append_sheet(wb, ws, title));
    const out = freezeHeader(XLSX.write(wb, { bookType: "xlsx", type: "array" }), sheets.length);
    S.lastBook = { name, wb };   // 시험용 — 무엇을 썼는지 확인
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([out], { type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" }));
    a.download = name; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 2000);
  }
  function saveCsv(name, rows) {
    const csv = "﻿" + XLSX.utils.sheet_to_csv(XLSX.utils.json_to_sheet(rows.length ? rows : [{ 내용: "없음" }]));  // 엑셀에서 한글이 깨지지 않게 BOM
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    a.download = name; a.click();
    setTimeout(() => URL.revokeObjectURL(a.href), 1000);
  }
  const dateCols = () => Object.entries(S.rules.columns || {}).filter(([, r]) => r.date).map(([c]) => c);
  const reconRows = () => S.recon.flatMap(x => [
    { 칸: x.col, 단계: "① 원본 파일의 합계 행 값", 금액: x.fileTotal === null ? "" : x.fileTotal },
    { 칸: x.col, 단계: "② 원본 숫자 칸만 합계", 금액: x.numbersOnly },
    { 칸: x.col, 단계: "③ 처방으로 더해진 금액", 금액: x.formatDelta },
    { 칸: x.col, 단계: "④ 처방 적용 후 전체 합계", 금액: x.afterRules },
    { 칸: x.col, 단계: "⑤ 중복으로 뺀 금액", 금액: -x.droppedAmt },
    { 칸: x.col, 단계: "⑥ 정리 후 합계", 금액: x.cleaned },
    { 칸: x.col, 단계: "검증 ④−⑤=⑥", 금액: x.ok ? "통과" : "실패" },
  ]).map(o => Object.assign(o, { 단위: unitLabel() }));
  const reconSheet = () => makeSheet(["칸", "단계", "금액", "단위"], reconRows());
  const verifySheet = () => makeSheet(["검증", "내용", "결과"], (S.checks || verify()).map(([k, v, p]) => ({ 검증: k, 내용: v, 결과: p ? "통과" : "실패" })));
  const logObjs = () => S.result.cells.map(l => ({ 신청번호: l.key, 원본행: l.row, 원본칸: l.col, 원본값: l.from, 정리값: l.to, 처방ID: l.rules.join(", "), 처방: l.rules.map(unitText).join(" / "), 변경이유: l.rule, 처리시각: stamp(), 행제거: "아니오" }))
    .concat(S.result.dropped.map(d => ({ 신청번호: d.key, 원본행: d.row, 원본칸: "(행 전체)", 원본값: "", 정리값: "", 처방ID: "dup", 처방: "중복 — 모든 칸이 같은 줄", 변경이유: `행 ${d.sameAs}과 같음. ${(S.rules.duplicates || {}).why || ""}`, 처리시각: stamp(), 행제거: "예" })));
  const unitObjs = () => S.units.map(u => ({ 처방ID: u.id, 켬: S.enabled.has(u.id) ? "켬" : "끔", 묶음: u.group, 칸: u.col, 방식: u.how, 이유: u.why, 기준: u.basis || "", 영향: S.result.summary.byRule[u.id] || 0 }));
  const infoObjs = () => [["파일", S.name], ["처방", `${S.rules.title || ""} (${S.rulesFrom})`], ["처방을 정한 사람", S.rules.owner || ""], ["실행 시각", stamp()],
    ["원본 데이터", `${S.result.summary.before}건 (합계 행 ${S.summary.length}줄 별도)`], ["정리 후", `${S.result.summary.after}건`], ["두 번 돌려도 같은가", S.idem ? "같음" : "다름"],
    ["금액 단위", unitLabel()], ["규정 파일", S.policy ? S.policyName : "없음 — 규정 적합성 미판정"], ["상태", "정리 실행 완료 · 업무 승인 대기"]].map(([항목, 값]) => ({ 항목, 값 }));

  $("#dl-clean").addEventListener("click", () => {
    const headers = S.cols.concat("검토 사유");
    const rows = S.result.rows.map(r => Object.assign({}, r, { "검토 사유": r[F] }));
    saveBook(`${S.base}_정리본.xlsx`, [["정리본", makeSheet(headers, rows, { dateCols: dateCols(), totalCols: reconCols() })], ["금액 대사", reconSheet()], ["자동 검증", verifySheet()]]);
  });
  $("#dl-log").addEventListener("click", () => {
    const lh = ["신청번호", "원본행", "원본칸", "원본값", "정리값", "처방ID", "처방", "변경이유", "처리시각", "행제거"];
    saveBook(`${S.base}_변경기록.xlsx`, [["변경 기록", makeSheet(lh, logObjs())], ["적용한 처방", makeSheet(["처방ID", "켬", "묶음", "칸", "방식", "이유", "기준", "영향"], unitObjs())],
      ["금액 대사", reconSheet()], ["정보", makeSheet(["항목", "값"], infoObjs())]]);
  });
  const reviewObjs = () => reviewItems().map(x => ({ 우선순위: x.우선 === 1 ? "1 · 금액/데이터 무결성" : "2 · 담당자 확인", 신청번호: x.신청번호, 원본행: x.행, 문제유형: x.유형.join(" · "),
    현재값: x.현재값.join(" / "), 사유: x.사유.join(" · "), 필요한근거: x.근거.join(" · "), 기준: x.기준.join(" · "), 처리상태: x.상태 }));
  $("#dl-review").addEventListener("click", () => saveBook(`${S.base}_검토필요목록.xlsx`,
    [["검토 필요", makeSheet(["우선순위", "신청번호", "원본행", "문제유형", "현재값", "사유", "필요한근거", "기준", "처리상태"], reviewObjs())]]));
  $("#dl-drop").addEventListener("click", () => {
    const rc = reconCols();
    const rows = S.result.dropped.map(d => Object.assign({ 원본행: d.row, 같은줄: d.sameAs }, Object.fromEntries(S.cols.map((c, i) => [c, d.values[i]]))));
    saveBook(`${S.base}_제외행목록.xlsx`, [["제외 행", makeSheet(["원본행", "같은줄", ...S.cols], rows, { dateCols: [], totalCols: rc })]]);
  });
  $("#dl-log-csv").addEventListener("click", () => saveCsv(`${S.base}_변경기록.csv`, logObjs()));
  $("#dl-review-csv").addEventListener("click", () => saveCsv(`${S.base}_검토필요목록.csv`, reviewObjs()));

  // 켠 처방만 남긴 처방 파일 — 다음 파일에도 다시 쓸 수 있다
  const STEP_KEYS = { trim: ["trim"], map: ["map"], blank_to: ["blank_to"], strip_units: ["strip_units"], number: ["number"], date: ["date"], rescale: ["rescale"], missing: ["missing"], unparsed: ["unparsed"], range: ["range"] };
  function prunedRules() {
    const R = JSON.parse(JSON.stringify(S.rules));
    Object.entries(R.columns || {}).forEach(([col, r]) => {
      Object.entries(STEP_KEYS).forEach(([k, keys]) => { if (!S.enabled.has(`${col}|${k}`)) keys.forEach(x => { if (x in r && !(k === "date" && S.enabled.has(`${col}|unparsed`))) delete r[x]; }); });
      if (!Object.keys(r).some(k => !["why", "need", "severity", "basis"].includes(k))) delete R.columns[col];
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
  $("#currency").addEventListener("change", () => { renderPolicyBox(); if (S.rules) updatePreview(); if (S.result) { renderRecon(); } });
  $("#policy-template").addEventListener("click", () => {
    const t = {
      kind: "travel_policy", name: "회사 출장비 규정 (회사가 채워 넣음)", version: "", source: "규정 문서 이름 · 조항",
      valid_from: "YYYY-MM-DD", valid_to: "YYYY-MM-DD", currency: "KRW",
      items: [{ item: "숙박비", match: { place: "출장지 값 그대로" }, per: "night", max: null, currency: "KRW", valid_from: "", valid_to: "",
        exceptions: [{ label: "사전 승인", when: [["승인여부", "==", "Y"]] }], note: "max 가 비어 있으면 판단하지 않습니다. 부서 · 직무 차이는 규정에 실제로 있을 때만 match 에 넣습니다" }],
      external: [{ name: "GSA Per Diem", adopted: false, adopted_by: "회사 규정의 어느 조항이 채택했는지", source_url: "https://www.gsa.gov/travel/plan-book/per-diem-rates",
        year: "FY", region: "미국 본토(CONUS)", currency: "USD", rates: [],
        note: "회사 규정이 채택했을 때만 adopted:true. 한국 회사에 자동 적용되는 법적 상한이 아닙니다. 요율은 공식 출처에서 확인해 지역 · 기간별로 rates 에 적습니다" }],
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
    const rd = readWorkbook(await f.arrayBuffer(), f.name);
    const rows = Clean.splitSummary(rd.rows, rd.formulas, rd.rowNos).data;
    const mine = S.cols;
    const theirs = rows.length ? Object.keys(rows[0]) : [];
    const A = S.rules.answer || {};
    // 처방 파일에 대조 기준이 없으면: 첫 칸으로 같은 줄을 찾고, 두 파일에 다 있는 나머지 칸을 비교한다
    const key = A.key || mine.slice(0, 1);
    const cols = A.compare || mine.filter(c => !key.includes(c) && theirs.includes(c));
    const c = Clean.compare(S.result.rows, rows, key, cols);
    const diff = Object.entries(c.diff).map(([k, n]) => `${esc(k)} ${n}줄`).join(" · ") || "없음";
    $("#compare").innerHTML = `<p class="small muted">같은 줄 찾는 칸: ${esc(key.join(" + "))} · 비교한 칸: ${esc(cols.join(" · "))}</p>
      <div class="kpis">${[["정답 줄", fmt(c.answerRows)], ["일치", fmt(c.match), "", "dark"], ["정답에 없는 줄", fmt(c.notInAnswer)], ["어긋난 칸", diff, "", "amber"]].map(kpi).join("")}</div>`
      + (c.examples.length ? `<p class="small">어긋난 예 — ${c.examples.slice(0, 6).map(x => `${esc(x.key.split("|")[0])} ${esc(x.col)}: ${esc(x.got)} (정답 ${esc(x.want)})`).join(" / ")}</p>` : "");
  });

  // ── 연결 ────────────────────────────────────────────
  function openTab(name) {
    const b = $(`#res-tabs button[data-tab="${name}"]`); if (!b) return;
    $$("#res-tabs button").forEach(x => x.setAttribute("aria-selected", String(x === b)));
    $$(".tabpane").forEach(p => { p.hidden = p.dataset.pane !== name; });
    if (name === "cmp") alignRows();
  }
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
  $("#res-tabs").addEventListener("click", e => { const b = e.target.closest("button"); if (b) openTab(b.dataset.tab); });
  $("#res-kpi").addEventListener("click", e => {
    const b = e.target.closest("[data-go]"); if (!b) return;
    if (b.dataset.go === "recon") { $("#recon").scrollIntoView({ block: "start" }); return; }
    openTab(b.dataset.go); $("#res-tabs").scrollIntoView({ block: "start" });
  });
  $("#cmp-filter").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    S.cmpMode = b.dataset.f;
    $$("#cmp-filter button").forEach(x => x.setAttribute("aria-pressed", String(x === b)));
    drawCompare();
  });
  $("#pane-switch").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    $$("#pane-switch button").forEach(x => x.setAttribute("aria-pressed", String(x === b)));
    $(".split").classList.toggle("show-left", b.dataset.p === "left");
    $(".split").classList.toggle("show-right", b.dataset.p === "right");
  });
  $("#rev-filter").addEventListener("click", e => {
    const b = e.target.closest("button"); if (!b) return;
    S.revMode = b.dataset.r;
    $$("#rev-filter button").forEach(x => x.setAttribute("aria-pressed", String(x === b)));
    drawReview();
  });
  $("#t-review").addEventListener("change", e => {
    const rv = e.target.dataset.rv; if (!rv) return;
    S.review.set(+rv, e.target.value);
    const rev = reviewItems();
    $("#rev-progress").textContent = `확인 완료 ${rev.filter(x => x.상태 === "확인 완료").length} / ${rev.length}건 · 처리 상태는 이 화면에만 있고 저장되지 않습니다. 검토필요목록.xlsx 에 함께 내려받습니다.`;
    renderResultStatusOnly();
  });
  function renderResultStatusOnly() {
    const rev = reviewItems(), done = rev.filter(x => x.상태 === "확인 완료").length;
    const el = $$("#status .st").find(x => x.textContent.startsWith("승인 대기"));
    if (el) el.textContent = `승인 대기 — 확인 ${done}/${rev.length}`;
  }
  $("#cmp-q").addEventListener("input", drawCompare);
  $("#log-q").addEventListener("input", drawLog);
  $("#rev-q").addEventListener("input", drawReview);
  $(".split").addEventListener("click", e => {
    const td = e.target.closest("td[data-c]"), tr = e.target.closest("tr[data-rn]");
    if (td && tr) showDetail(+tr.dataset.rn, td.dataset.c);
  });
  // 원본과 정리본 표를 같이 움직인다 (세로 · 가로)
  let syncing = false;
  ["#cmp-left", "#cmp-right"].forEach((s, i, all) => $(s).addEventListener("scroll", () => {
    if (syncing) return; syncing = true;
    const a = $(s), b = $(all[1 - i]);
    b.scrollTop = a.scrollTop; b.scrollLeft = a.scrollLeft;
    requestAnimationFrame(() => { syncing = false; });
  }));
})();
