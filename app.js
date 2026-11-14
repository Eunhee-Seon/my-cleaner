/* 엑셀 정리기 — 화면.
 *
 * 파일 → 읽기(SheetJS, 브라우저 안) → 진단(clean.js) → 규칙(rules/*.json) → 정리 → 기록·내려받기
 * 서버로 아무것도 보내지 않는다. fetch 는 예시 파일과 규칙 파일을 «내 사이트에서» 읽을 때만 쓴다.
 */
(function () {
  "use strict";
  const $ = s => document.querySelector(s);
  const S = { name: "", rows: [], rules: null, result: null };
  const esc = v => String(v === undefined || v === null ? "" : v)
    .replace(/&/g, "&amp;").replace(/</g, "&lt;").replace(/>/g, "&gt;");

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
    S.name = name; S.rows = rows;
    $("#file-info").hidden = false;
    $("#file-info").textContent = `${name} · ${rows.length}줄 · ${rows.length ? Object.keys(rows[0]).length : 0}칸`
      + (sheets.length > 1 ? ` · ⚠ 시트 ${sheets.length}개 중 첫 시트만 읽었습니다` : "");
    preview(rows);
    diagnose(rows);
    await loadRulesFor(name);
    $("#sec-result").hidden = true;
  }

  function preview(rows) {
    const cols = rows.length ? Object.keys(rows[0]) : [];
    $("#preview").hidden = false;
    $("#preview").innerHTML = `<table><thead><tr>${cols.map(c => `<th>${esc(c)}</th>`).join("")}</tr></thead><tbody>`
      + rows.slice(0, 8).map(r => `<tr>${cols.map(c => `<td>${esc(r[c])}</td>`).join("")}</tr>`).join("")
      + `</tbody></table>`;
  }

  // ── 진단 ────────────────────────────────────────────
  function diagnose(rows) {
    const d = Clean.diagnose(rows);
    $("#sec-diag").hidden = false;
    $("#symptoms").innerHTML = d.symptoms.length
      ? d.symptoms.map(s => `<div class="sym"><span class="kind ${s.kind}">${s.kind}</span><span>${esc(s.text)}</span></div>`).join("")
      : `<p class="muted">규칙 없이 찾을 수 있는 증상은 없습니다. 그래도 값 목록은 눈으로 보십시오.</p>`;
    const cats = Object.entries(d.columns).filter(([, i]) => i.values);
    $("#values").innerHTML = cats.map(([c, i]) =>
      `<p><b>${esc(c)}</b> — ${Object.entries(i.values).sort((a, b) => b[1] - a[1])
        .map(([v, n]) => `«${esc(v)}» ${n}`).join(" · ")}</p>`).join("") || `<p class="muted">종류가 적은 글자 열이 없습니다.</p>`;
  }

  // ── 규칙 ────────────────────────────────────────────
  async function loadRulesFor(name) {
    const base = name.replace(/\.(xlsx|xls|csv)$/i, "");
    try {
      const r = await fetch(`rules/${encodeURIComponent(base)}.json`);
      S.rules = r.ok ? await r.json() : null;
    } catch (e) { S.rules = null; }
    showRules();
  }

  // 처방 파일의 내용을 사람이 읽는 말로 — 처방 자체는 rules/*.json 에만 있다
  function describe(r) {
    const out = [];
    if (r.trim) out.push("앞뒤 공백 지움");
    if (r.map) {
      const pairs = Object.entries(r.map);
      out.push(pairs.length ? pairs.map(([a, b]) => `«${a}»→«${b}»`).join(" · ") : "바꿀 이름 표가 비어 있음");
    }
    if (r.blank_to !== undefined) out.push(`빈칸은 «${r.blank_to}»`);
    if (r.strip_units) out.push(r.strip_units.map(u => `«${u}»`).join(" ") + " 뗌");
    if (r.number) out.push("숫자로");
    if (r.date) out.push("날짜를 YYYY-MM-DD 하나로");
    if (r.unparsed === "flag") out.push("못 읽은 날짜는 고치지 않고 표시");
    if (r.rescale) out.push(`${r.rescale.over} 넘으면 ${r.rescale.divide}로 나눔`);
    if (r.missing === "flag") out.push("빈칸은 채우지 않고 표시");
    if (r.range) out.push(`${r.range.min}~${r.range.max} 밖이면 표시`);
    return out.join(" · ");
  }
  const cond = w => `${w[0]} ${w[1]} ${typeof w[2] === "number" ? w[2].toLocaleString("ko-KR") : `«${w[2]}»`}`;

  function showRules() {
    $("#sec-rules").hidden = false;
    const R = S.rules;
    if (!R) {
      $("#rules-owner").hidden = false;
      $("#rules-owner").textContent = "이 파일에 맞는 처방이 없습니다. 처방 파일(.json)을 골라 올리십시오 — 처방은 사람이 정합니다.";
      $("#rules").innerHTML = ""; $("#rules-order").textContent = "";
      $("#btn-apply").disabled = true;
      return;
    }
    $("#btn-apply").disabled = false;
    $("#rules-owner").hidden = false;
    $("#rules-owner").textContent = `${R.title || "처방"} — 정한 사람: ${R.owner || "(적혀 있지 않음)"}`;
    const card = (c, what, why) => `<div class="rule"><div class="c">${esc(c)}</div>`
      + (what ? `<div class="w"><b>처방</b> ${esc(what)}</div>` : "")
      + `<div class="w"><b>이유</b> ${esc(why || "(적혀 있지 않음)")}</div></div>`;
    const parts = Object.entries(R.columns || {}).map(([c, r]) => card(c, describe(r), r.why));
    (R.cross_checks || []).forEach(cc => parts.push(card(`두 칸을 같이 — ${cc.label}`, cc.when.map(cond).join(" 이고 ") + " 이면 표시", cc.why)));
    if (R.duplicates) parts.push(card("중복", R.duplicates.action === "drop" ? "완전히 같은 줄은 하나만 남김" : "", R.duplicates.why));
    $("#rules").innerHTML = parts.join("");
    $("#rules-order").textContent = R.order_note ? "순서 — " + R.order_note : "";
  }

  // ── 적용 ────────────────────────────────────────────
  function applyRules() {
    const r1 = Clean.apply(S.rows, S.rules);
    const r2 = Clean.apply(r1.rows, S.rules);        // 한 번 더 — 두 번 돌려도 같은가
    S.result = r1;
    const s = r1.summary;
    $("#sec-result").hidden = false;
    $("#stats").innerHTML = [
      ["줄 (전 → 후)", `${s.before} → ${s.after}`],
      ["지운 줄 (완전 중복)", s.dropped],
      ["고친 값", s.changes],
      ["표시한 줄 (지우지 않음)", s.flagged],
    ].map(([k, v]) => `<div><div class="k">${k}</div><div class="v">${v}</div></div>`).join("");
    const ok = Clean.same(r1.rows, r2.rows) && r2.summary.changes === 0 && r2.summary.dropped === 0;
    $("#idem").innerHTML = ok
      ? `<b class="pass">✓ 두 번 돌려도 같은 결과</b> — 정리본에 규칙을 한 번 더 적용해도 바뀌는 것이 없습니다`
      : `<b class="fail">✕ 두 번 돌리면 결과가 달라집니다</b> — 규칙이 자기 결과를 또 바꿉니다. 규칙을 다시 보십시오`;
    const shown = r1.log.slice(0, 12);
    $("#log").innerHTML = `<table><thead><tr><th>행</th><th>열</th><th>전</th><th>후</th><th>규칙</th></tr></thead><tbody>`
      + shown.map(l => `<tr><td>${l.row}</td><td>${esc(l.col)}</td><td>${esc(l.from)}</td><td>${esc(l.to)}</td><td>${esc(l.rule).slice(0, 30)}…</td></tr>`).join("")
      + (r1.log.length > 12 ? `<tr><td colspan="5">… 외 ${r1.log.length - 12}줄 — «변경 기록 내려받기»</td></tr>` : "")
      + `</tbody></table>`;
    $("#compare").innerHTML = "";
  }

  // ── 내려받기 ─────────────────────────────────────────
  function download(name, rows) {
    const ws = XLSX.utils.json_to_sheet(rows);
    const csv = "﻿" + XLSX.utils.sheet_to_csv(ws);           // 엑셀에서 한글이 깨지지 않게 BOM
    const a = document.createElement("a");
    a.href = URL.createObjectURL(new Blob([csv], { type: "text/csv;charset=utf-8" }));
    a.download = name; a.click();
  }
  const base = () => S.name.replace(/\.(xlsx|xls|csv)$/i, "");
  $("#dl-clean").addEventListener("click", () => download(`${base()}_정리본.csv`, S.result.rows));
  $("#dl-log").addEventListener("click", () => download(`${base()}_변경기록.csv`,
    S.result.log.map(l => ({ 행: l.row, 열: l.col, 전: l.from, 후: l.to, 규칙: l.rule }))));

  // ── 정답 대조 ────────────────────────────────────────
  $("#answer-file").addEventListener("change", async e => {
    const f = e.target.files[0];
    e.target.value = "";
    if (!f) return;
    if (!S.result) { $("#compare").innerHTML = `<p class="small">먼저 «규칙대로 정리»를 누르십시오.</p>`; return; }
    // 정답 파일은 내 컴퓨터에서 고른다 — 사이트에 두지 않고, 어디로도 보내지 않는다
    const { rows } = readWorkbook(await f.arrayBuffer(), f.name);
    const mine = S.result.rows.length ? Object.keys(S.result.rows[0]).filter(c => c !== Clean.FLAG) : [];
    const theirs = rows.length ? Object.keys(rows[0]) : [];
    const A = S.rules.answer || {};
    // 처방 파일에 대조 기준이 없으면: 첫 칸으로 같은 줄을 찾고, 두 파일에 다 있는 나머지 칸을 비교한다
    const key = A.key || mine.slice(0, 1);
    const cols = A.compare || mine.filter(c => !key.includes(c) && theirs.includes(c));
    const c = Clean.compare(S.result.rows, rows, key, cols);
    const diff = Object.entries(c.diff).map(([k, n]) => `${esc(k)} ${n}줄`).join(" · ") || "없음";
    $("#compare").innerHTML = `<p class="small muted">같은 줄 찾는 칸: ${esc(key.join(" + "))} · 비교한 칸: ${esc(cols.join(" · "))}</p>
      <div class="stat">
        <div><div class="k">정답 줄</div><div class="v">${c.answerRows}</div></div>
        <div><div class="k">일치</div><div class="v">${c.match}</div></div>
        <div><div class="k">정답에 없는 줄</div><div class="v">${c.notInAnswer}</div></div>
        <div><div class="k">어긋난 칸</div><div class="v" style="font-size:14px">${diff}</div></div></div>`
      + (c.examples.length ? `<p class="small">어긋난 예 — ${c.examples.slice(0, 5).map(x => `${esc(x.key.split("|")[0])} ${esc(x.col)}: ${esc(x.got)} (정답 ${esc(x.want)})`).join(" / ")}</p>` : "");
  });

  // ── 연결 ────────────────────────────────────────────
  $("#file").addEventListener("change", async e => {
    const f = e.target.files[0];
    if (f) loadFile(await f.arrayBuffer(), f.name);
  });
  const drop = $("#drop");
  ["dragenter", "dragover"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.add("over"); }));
  ["dragleave", "drop"].forEach(t => drop.addEventListener(t, e => { e.preventDefault(); drop.classList.remove("over"); }));
  drop.addEventListener("drop", async e => {
    const f = e.dataTransfer.files[0];
    if (f) loadFile(await f.arrayBuffer(), f.name);
  });
  document.querySelectorAll("[data-sample]").forEach(b => b.addEventListener("click", async () => {
    const name = b.dataset.sample;
    const buf = await (await fetch("data/" + encodeURIComponent(name))).arrayBuffer();
    loadFile(buf, name);
  }));
  $("#rules-file").addEventListener("change", async e => {
    const f = e.target.files[0];
    if (!f) return;
    e.target.value = "";
    try { S.rules = JSON.parse(await f.text()); showRules(); $("#sec-result").hidden = true; }
    catch (err) { $("#rules-owner").textContent = "처방 파일을 읽지 못했습니다 — JSON 형식을 확인하십시오"; }
  });
  $("#btn-apply").addEventListener("click", applyRules);
})();
