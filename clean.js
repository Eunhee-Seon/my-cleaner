/* 정리 엔진 — 진단 · 처방 적용 · 규정 확인. 화면과 상관없이 계산만 한다.
 *
 *   diagnose(rows, rules?)         원본을 고치지 않고 증상을 찾는다 (빈칸 · 중복 · 표기 · 형식 · 이상치 후보 · 칸 사이 모순)
 *   ruleUnits(rules)               처방 파일을 «하나씩 켜고 끌 수 있는 처방»으로 나눈다
 *   apply(rows, rules, {enabled})  켠 처방만 적용해 «새 정리본»을 만든다. 원본 배열은 건드리지 않는다
 *   same(a, b)                     두 결과가 같은가 — «두 번 돌려도 같은 결과» 확인용
 *   compare(rows, answer, key, cols)  정답과 대조
 *   conflicts(a, b)                두 처방 파일이 같은 칸을 다르게 다루는지
 *   checkPolicy(rows, src, policy, ctx)  회사 규정(1순위) → 회사가 채택한 외부 기준(2순위) → 판단 불가(3순위)
 *
 * 처방은 rules/*.json 에만 있다. 이 파일에는 처방을 적지 않는다.
 * 브라우저(window.Clean)와 node(require) 양쪽에서 쓴다. 파일은 밖으로 나가지 않는다.
 */
(function (root) {
  "use strict";

  const FLAG = "_표시";
  const isBlank = v => v === null || v === undefined || String(v).trim() === "";
  const numLike = v => typeof v === "number" || /^\s*-?[\d,]+(\.\d+)?\s*$/.test(String(v));
  const toNum = v => (typeof v === "number" ? v : Number(String(v).replace(/,/g, "").trim()));
  const columnsOf = rows => (rows.length ? Object.keys(rows[0]).filter(c => c !== FLAG) : []);

  const DATE_FORMATS = [
    // 연도가 네 자리인 것만 읽는다. «25/7/23»처럼 연도가 두 자리면 뜻이 갈리므로 읽지 않는다
    { name: "YYYY-MM-DD", re: /^(\d{4})-(\d{1,2})-(\d{1,2})$/, y: 1, m: 2, d: 3 },
    { name: "YYYY/MM/DD", re: /^(\d{4})\/(\d{1,2})\/(\d{1,2})$/, y: 1, m: 2, d: 3 },
    { name: "DD-MM-YYYY", re: /^(\d{2})-(\d{2})-(\d{4})$/, y: 3, m: 2, d: 1 },
    { name: "YYYY.MM.DD", re: /^(\d{4})\.(\d{1,2})\.(\d{1,2})$/, y: 1, m: 2, d: 3 },
  ];
  const ISO = /^\d{4}-\d{2}-\d{2}$/;
  const dateFormatOf = v => (DATE_FORMATS.find(f => f.re.test(String(v).trim())) || {}).name || null;

  function quantile(sorted, q) {
    const pos = (sorted.length - 1) * q, lo = Math.floor(pos), hi = Math.ceil(pos);
    return sorted[lo] + (sorted[hi] - sorted[lo]) * (pos - lo);
  }
  const fmtN = n => (typeof n === "number" ? n.toLocaleString("ko-KR") : String(n));

  // ── 진단 ─────────────────────────────────────────────────────
  // items: {kind, col, count, rows:[엑셀 행], examples:[{row, value}], note}
  // 이상치는 «후보»일 뿐 오류라는 뜻이 아니다. 진단은 아무것도 고치지 않는다.
  function diagnose(rows, rules) {
    const cols = columnsOf(rows);
    const out = { n: rows.length, cols: cols.length, columns: {}, symptoms: [], items: [] };
    const R = i => i + 2;                                  // 엑셀 행 번호 (머리줄이 1행)
    const item = (kind, col, idx, note, valueOf) => {
      if (!idx.length) return;
      out.items.push({
        kind, col, count: idx.length, rows: idx.map(R),
        examples: idx.slice(0, 5).map(i => ({ row: R(i), value: valueOf ? valueOf(i) : rows[i][col] })),
        note: note || "",
      });
    };

    for (const c of cols) {
      const vals = rows.map(r => r[c]);
      const filledIdx = vals.map((v, i) => i).filter(i => !isBlank(vals[i]));
      const filled = filledIdx.map(i => vals[i]);
      const info = { empty: vals.length - filled.length };

      item("빈칸", c, vals.map((v, i) => i).filter(i => isBlank(vals[i])), "비었거나 공백뿐인 칸", () => "(빈칸)");

      // 형식 불일치 — 대부분 숫자인데 글자가 섞였다 («3개», «1,234 원», 글자로 저장된 «2,440,000»)
      const nums = filled.filter(numLike);
      info.numericShare = filled.length ? nums.length / filled.length : 0;
      if (info.numericShare >= 0.6) {
        // CSV 는 모든 값을 글자로 읽는다 — «12»는 문제가 아니다. 숫자가 아니거나 쉼표 · 공백이 낀 글자만 센다
        const textIdx = filledIdx.filter(i => !numLike(vals[i]) || (typeof vals[i] === "string" && /[,\s]/.test(vals[i])));
        if (nums.length < filled.length) {
          info.notNumber = [...new Set(filled.filter(v => !numLike(v)).map(String))].slice(0, 6);
          info.notNumberCount = filled.length - nums.length;
        }
        item("형식 불일치", c, textIdx, "숫자 칸인데 글자로 저장된 값 (쉼표 · 단위 포함)");
      }

      // 이상치 후보 — 숫자 칸에서 사분위 범위의 3배 밖 · 음수. 오류라는 뜻은 아니다
      if (info.numericShare >= 0.6 && nums.length >= 8) {
        const s = nums.map(toNum).sort((a, b) => a - b);
        const q1 = quantile(s, 0.25), q3 = quantile(s, 0.75), iqr = q3 - q1;
        const lo = q1 - 3 * iqr, hi = q3 + 3 * iqr;
        const odd = s.filter(v => v < lo || v > hi);
        info.stats = { min: s[0], max: s[s.length - 1], median: quantile(s, 0.5), q1, q3, lo, hi };
        if (odd.length) info.outliers = { count: odd.length, examples: [...new Set(odd)].slice(-5) };
        const neg = s.filter(v => v < 0).length;
        if (neg) info.negative = neg;
        item("이상치 후보", c, filledIdx.filter(i => numLike(vals[i]) && (toNum(vals[i]) < lo || toNum(vals[i]) > hi)),
          `사분위 범위 3배 밖 (${fmtN(Math.round(lo))} ~ ${fmtN(Math.round(hi))}). 크다는 이유만으로 오류는 아니다`);
        item("이상치 후보", c, filledIdx.filter(i => numLike(vals[i]) && toNum(vals[i]) < 0), "음수");
      }

      // 날짜 표기 여러 가지
      const fmts = {};
      filled.forEach(v => { const f = dateFormatOf(v); if (f) fmts[f] = (fmts[f] || 0) + 1; });
      if (Object.keys(fmts).length && Object.values(fmts).reduce((a, b) => a + b, 0) >= filled.length * 0.8) {
        info.dateFormats = fmts;
        const main = Object.entries(fmts).sort((a, b) => b[1] - a[1])[0][0];
        item("형식 불일치", c, filledIdx.filter(i => dateFormatOf(vals[i]) !== main),
          `날짜 표기 ${Object.keys(fmts).length}가지 — 가장 많은 것은 ${main}`);
      }

      // 표기 불일치 — 앞뒤 공백 · 대소문자/띄어쓰기만 다른 값. 종류가 적은 글자 칸은 값 목록 전부
      const distinct = {};
      filled.forEach(v => { const k = String(v); distinct[k] = (distinct[k] || 0) + 1; });
      const kinds = Object.keys(distinct).length;
      item("표기 불일치", c, filledIdx.filter(i => typeof vals[i] === "string" && vals[i] !== vals[i].trim()), "앞뒤 공백");
      if (info.numericShare < 0.6 && !info.dateFormats && kinds <= 30) {
        info.values = distinct;
        const groups = {};
        Object.keys(distinct).forEach(v => {
          const k = v.toLowerCase().replace(/\s+/g, "");
          (groups[k] = groups[k] || []).push(v);
        });
        info.lookAlike = Object.values(groups).filter(g => g.length > 1);
        const alike = new Set(info.lookAlike.flat());
        item("표기 불일치", c, filledIdx.filter(i => alike.has(String(vals[i]))), "대소문자 · 띄어쓰기만 다른 값");
      }
      out.columns[c] = info;
    }

    // 중복 — 모든 칸이 같은 줄, 첫 칸(줄을 알아보는 값)이 두 번 이상
    const seen = new Map();
    const dupIdx = [];
    rows.forEach((r, i) => {
      const k = JSON.stringify(cols.map(c => r[c]));
      if (seen.has(k)) dupIdx.push(i); else seen.set(k, i);
    });
    out.duplicates = dupIdx.length;
    item("중복", "(줄 전체)", dupIdx, "모든 칸이 앞의 줄과 같음",
      i => `행 ${R(seen.get(JSON.stringify(cols.map(c => rows[i][c]))))}과 같음`);
    if (cols.length) {
      const first = cols[0], cnt = {};
      rows.forEach(r => { if (!isBlank(r[first])) cnt[r[first]] = (cnt[r[first]] || 0) + 1; });
      item("중복", first, rows.map((r, i) => i).filter(i => !isBlank(rows[i][first]) && cnt[rows[i][first]] > 1),
        "줄을 알아보는 값이 두 번 이상");
    }

    // 칸 사이 모순 — 처방 파일에 사람이 정한 규칙이 있을 때만 (규칙을 지어내지 않는다)
    (rules && rules.cross_checks || []).filter(cc => cc.when.length > 1).forEach(cc => {
      item("칸 사이 모순", cc.when.map(w => w[0]).join(" · "),
        rows.map((r, i) => i).filter(i => cc.when.every(w => test(rows[i], w))),
        `${cc.label} — 처방 파일에 정한 규칙`, i => cc.when.map(w => `${w[0]}=${rows[i][w[0]]}`).join(", "));
    });

    // 증상 카드 — 사람이 읽을 요약 (예전 화면과 같은 모양)
    for (const [c, i] of Object.entries(out.columns)) {
      if (i.empty) out.symptoms.push({ kind: "빈칸", col: c, text: `${c} — 빈칸 ${i.empty}개` });
      if (i.notNumberCount) out.symptoms.push({ kind: "섞인 형식", col: c, text: `${c} — 숫자 열에 글자 ${i.notNumberCount}개 (예: ${i.notNumber.slice(0, 3).join(", ")})` });
      if (i.outliers) out.symptoms.push({ kind: "튀는 값", col: c, text: `${c} — 튀는 값 ${i.outliers.count}개 (예: ${i.outliers.examples.join(", ")})` });
      if (i.negative) out.symptoms.push({ kind: "튀는 값", col: c, text: `${c} — 음수 ${i.negative}개` });
      if (i.dateFormats && Object.keys(i.dateFormats).length > 1) out.symptoms.push({ kind: "섞인 형식", col: c, text: `${c} — 날짜 표기 ${Object.keys(i.dateFormats).length}가지 (${Object.keys(i.dateFormats).join(" · ")})` });
      if (i.values && Object.keys(i.values).length > 1 && (i.lookAlike.length || Object.keys(i.values).length > 6)) {
        out.symptoms.push({ kind: "표기", col: c, text: `${c} — ${Object.keys(i.values).length}가지 표기${i.lookAlike.length ? ` (비슷한 것: ${i.lookAlike.map(g => g.map(x => `«${x}»`).join("/")).join(", ")})` : ""}` });
      }
    }
    if (dupIdx.length) out.symptoms.push({ kind: "중복", col: "(행 전체)", text: `완전히 같은 행 ${dupIdx.length}개 — 날짜 표기를 먼저 맞추면 더 나올 수 있습니다` });
    out.affectedRows = new Set(out.items.flatMap(x => x.rows)).size;
    return out;
  }

  // ── 처방을 하나씩 켜고 끌 수 있는 단위로 ──────────────────────
  // group: 형식(모양만 바꿈) · 값(사람이 정한 표로 값을 바꿈) · 지움 · 표시(고치지 않고 표시)
  const STEP = {
    trim: { group: "형식", how: () => "앞뒤 공백을 지운다" },
    map: { group: "값", how: r => { const p = Object.entries(r.map); return p.length ? p.map(([a, b]) => `«${a}» → «${b}»`).join(" · ") : "바꿀 이름 표가 비어 있다"; } },
    blank_to: { group: "값", how: r => `빈칸을 «${r.blank_to}»로 채운다` },
    strip_units: { group: "형식", how: r => r.strip_units.map(u => `«${u}»`).join(" ") + "를 뗀다" },
    number: { group: "형식", how: () => "숫자로 바꾼다 (쉼표를 뗀다)" },
    date: { group: "형식", how: () => "날짜를 YYYY-MM-DD 하나로 맞춘다 (연도 네 자리만)" },
    rescale: { group: "값", how: r => `${fmtN(r.rescale.over)}를 넘으면 ${r.rescale.divide}로 나눈다` },
    missing: { group: "표시", how: () => "빈칸은 채우지 않고 표시한다" },
    unparsed: { group: "표시", how: () => "읽지 못한 날짜는 고치지 않고 표시한다" },
    range: { group: "표시", how: r => `${fmtN(r.range.min)} ~ ${fmtN(r.range.max)} 밖이면 표시한다` },
  };
  const STEP_ORDER = ["trim", "map", "blank_to", "strip_units", "number", "date", "rescale"];
  const condText = w => `${w[0]} ${w[1]} ${typeof w[2] === "number" ? fmtN(w[2]) : `«${w[2]}»`}`;

  function ruleUnits(rules) {
    const units = [];
    for (const [col, r] of Object.entries((rules && rules.columns) || {})) {
      for (const k of Object.keys(STEP)) {
        const present = k === "missing" ? r.missing === "flag" : k === "unparsed" ? r.date && r.unparsed === "flag"
          : k === "range" ? r.range && r.range.action === "flag" : k === "blank_to" ? r.blank_to !== undefined : !!r[k];
        if (present) units.push({ id: `${col}|${k}`, col, step: k, group: STEP[k].group, how: STEP[k].how(r), why: r.why || "" });
      }
    }
    ((rules && rules.cross_checks) || []).forEach((cc, i) => units.push({
      id: `cc|${i}`, col: cc.when.map(w => w[0]).join(" · "), step: "cross", group: "표시",
      how: cc.when.map(condText).join(" 이고 ") + ` 이면 «${cc.label}» 표시`, why: cc.why || "", label: cc.label,
    }));
    if (rules && rules.duplicates && rules.duplicates.action === "drop") {
      units.push({ id: "dup", col: "(줄 전체)", step: "dup", group: "지움", how: "모든 칸이 같은 줄은 처음 것만 남기고 지운다", why: rules.duplicates.why || "" });
    }
    return units;
  }

  // ── 처방 적용 ─────────────────────────────────────────────────
  function toISODate(v) {
    const s = String(v).trim();
    for (const f of DATE_FORMATS) {
      const m = s.match(f.re);
      if (m) return `${m[f.y]}-${m[f.m].padStart(2, "0")}-${m[f.d].padStart(2, "0")}`;
    }
    return v;
  }

  function test(row, [col, op, value]) {
    const v = row[col];
    if (op === "blank") return isBlank(v);
    if (isBlank(v)) return false;          // 빈칸은 0이 아니다 — 빈칸을 0으로 읽으면 없는 사실이 생긴다
    const n = typeof value === "number" ? toNum(v) : String(v);
    switch (op) {
      case "==": return n === value;
      case "!=": return n !== value;
      case ">": return toNum(v) > value;
      case "<": return toNum(v) < value;
      default: return false;
    }
  }

  /* 켠 처방만 적용한다. opts.enabled 가 없으면 처방 파일 전부.
   * 돌려주는 것:
   *   rows     새 정리본 (원본 배열과 다른 새 객체)
   *   src      정리본 각 줄이 원본 몇 행에서 왔는지
   *   log      바뀐 칸 하나마다 {row, col, from, to, rule(이유), rules:[처방 id], key} + 지운 줄
   *   cells    바뀐 칸만 (log 에서 지운 줄을 뺀 것)
   *   flags    표시 하나마다 {row, col, reason, rule, key}
   *   dropped  지운 줄 {row, key, sameAs, values}
   *   summary  {before, after, dropped, flagged, changes(바뀐 칸), changedRows(바뀐 줄), byColumn, byRule}
   */
  function apply(input, rules, opts) {
    const enabled = opts && opts.enabled ? new Set(opts.enabled) : null;
    const on = id => !enabled || enabled.has(id);
    const log = [], flags = [];
    const cols = columnsOf(input);
    const rows = input.map(r => { const o = {}; cols.forEach(c => { o[c] = r[c]; }); o[FLAG] = []; return o; });
    const keyOf = row => (cols.length ? row[cols[0]] : "");
    const flag = (i, col, reason, rule) => { rows[i][FLAG].push(reason); flags.push({ row: i + 2, col, reason, rule, key: keyOf(rows[i]) }); };

    for (const [col, r] of Object.entries((rules && rules.columns) || {})) {
      if (!cols.includes(col)) continue;
      const id = k => `${col}|${k}`;
      rows.forEach((row, i) => {
        let v = row[col];
        const before = v, used = [];
        const step = (k, fn) => { if (!on(id(k))) return; const b = v; v = fn(v); if (String(v) !== String(b)) used.push(id(k)); };
        if (r.trim) step("trim", x => (typeof x === "string" ? x.trim() : x));
        if (r.map) step("map", x => (Object.prototype.hasOwnProperty.call(r.map, String(x)) ? r.map[String(x)] : x));
        if (r.blank_to !== undefined) step("blank_to", x => (isBlank(x) ? r.blank_to : x));
        if (r.strip_units) step("strip_units", x => { if (typeof x !== "string") return x; r.strip_units.forEach(u => { x = x.split(u).join("").trim(); }); return x; });
        if (r.number) step("number", x => (!isBlank(x) && numLike(x) ? toNum(x) : x));
        if (r.date) step("date", x => (!isBlank(x) ? toISODate(x) : x));
        if (r.rescale) step("rescale", x => (typeof x === "number" && x > r.rescale.over ? x / r.rescale.divide : x));
        // 글자 «12»가 숫자 12가 된 것처럼 겉보기가 같으면 기록하지 않는다. «2,440,000»→2440000 처럼 모양이 바뀌면 기록한다
        if (String(v) !== String(before) && !(isBlank(v) && isBlank(before))) {
          log.push({ row: i + 2, col, from: before, to: v, rule: r.why || "", rules: used, key: keyOf(row) });
        }
        row[col] = v;

        if (r.missing === "flag" && on(id("missing")) && isBlank(v)) flag(i, col, `${col} 빈칸`, id("missing"));
        if (r.date && r.unparsed === "flag" && on(id("unparsed")) && !isBlank(v) && !ISO.test(String(v))) flag(i, col, `${col} 못 읽은 날짜(${v})`, id("unparsed"));
        if (r.range && r.range.action === "flag" && on(id("range")) && !isBlank(v) && typeof v === "number" && (v < r.range.min || v > r.range.max)) {
          flag(i, col, `${col} 범위 밖(${v})`, id("range"));
        }
      });
    }

    ((rules && rules.cross_checks) || []).forEach((cc, k) => {
      if (!on(`cc|${k}`)) return;
      rows.forEach((row, i) => { if (cc.when.every(w => test(row, w))) flag(i, cc.when.map(w => w[0]).join(" · "), cc.label, `cc|${k}`); });
    });

    let kept = rows, src = rows.map((r, i) => i + 2);
    const dropped = [];
    if (rules && rules.duplicates && rules.duplicates.action === "drop" && on("dup")) {
      const seen = new Map();
      kept = []; src = [];
      rows.forEach((row, i) => {
        const k = JSON.stringify(cols.map(c => row[c]));
        if (seen.has(k)) {
          dropped.push({ row: i + 2, key: keyOf(row), sameAs: seen.get(k), values: cols.map(c => row[c]) });
          log.push({ row: i + 2, col: "(행 전체)", from: "", to: "지움", rule: rules.duplicates.why || "완전히 같은 행", rules: ["dup"], key: keyOf(row) });
        } else { seen.set(k, i + 2); kept.push(row); src.push(i + 2); }
      });
    }
    const droppedRows = new Set(dropped.map(d => d.row));
    const keptFlags = flags.filter(f => !droppedRows.has(f.row));

    kept.forEach(row => { row[FLAG] = row[FLAG].join(" · "); });
    // 바뀐 칸은 정리본에 남은 줄만 센다 — 지운 줄은 dropped 에 따로 남는다
    const cells = log.filter(l => l.col !== "(행 전체)" && !droppedRows.has(l.row));
    const byColumn = {}, byRule = {};
    cells.forEach(l => { byColumn[l.col] = (byColumn[l.col] || 0) + 1; l.rules.forEach(id => { byRule[id] = (byRule[id] || 0) + 1; }); });
    keptFlags.forEach(f => { byRule[f.rule] = (byRule[f.rule] || 0) + 1; });
    if (dropped.length) byRule.dup = dropped.length;
    return {
      rows: kept, src, log, cells, flags: keptFlags, dropped,
      summary: {
        before: input.length, after: kept.length, dropped: dropped.length,
        flagged: kept.filter(r => r[FLAG]).length,
        changes: cells.length, changedRows: new Set(cells.map(l => l.row)).size,
        byColumn, byRule,
      },
    };
  }

  function same(a, b) {
    if (a.length !== b.length) return false;
    return a.every((r, i) => JSON.stringify(r) === JSON.stringify(b[i]));
  }

  /* 정답과 대조 — 정답은 «완성 기준»이 아니라 «내 규칙이 무엇을 놓쳤나»를 보는 거울이다.
   * key 로 같은 행을 찾고, cols 의 값을 비교한다. 숫자는 값으로 비교한다. */
  function compare(rows, answer, key, cols) {
    const k = r => key.map(c => String(r[c])).join("|");
    const want = new Map(answer.map(r => [k(r), r]));
    let match = 0, notInAnswer = 0;
    const diff = {}, examples = [];
    rows.forEach(r => {
      const a = want.get(k(r));
      if (!a) { notInAnswer += 1; return; }
      let ok = true;
      for (const c of cols) {
        const x = r[c], y = a[c];
        const eq = String(x) === String(y) || (numLike(x) && numLike(y) && !isBlank(x) && !isBlank(y) && toNum(x) === toNum(y));
        if (!eq) {
          ok = false; diff[c] = (diff[c] || 0) + 1;
          if (examples.length < 8) examples.push({ key: k(r), col: c, got: x, want: y });
        }
      }
      if (ok) match += 1;
    });
    return { answerRows: answer.length, match, notInAnswer, diff, examples };
  }

  // ── 두 처방 파일이 부딪히는가 ─────────────────────────────────
  // 같은 칸 같은 값을 다른 값으로 바꾸거나, 같은 칸의 기준이 다르면 알린다. 어느 쪽이 맞는지는 정하지 않는다.
  function conflicts(a, b) {
    const out = [];
    if (!a || !b) return out;
    const A = a.columns || {}, B = b.columns || {};
    for (const col of Object.keys(A)) {
      if (!B[col]) continue;
      const x = A[col], y = B[col];
      Object.keys(x.map || {}).forEach(k => {
        if (y.map && Object.prototype.hasOwnProperty.call(y.map, k) && String(y.map[k]) !== String(x.map[k])) {
          out.push(`${col}: «${k}»를 한쪽은 «${x.map[k]}», 다른 쪽은 «${y.map[k]}»로 바꿉니다`);
        }
      });
      if (x.blank_to !== undefined && y.blank_to !== undefined && String(x.blank_to) !== String(y.blank_to)) out.push(`${col}: 빈칸을 «${x.blank_to}» / «${y.blank_to}»로 채웁니다`);
      if (x.range && y.range && (x.range.min !== y.range.min || x.range.max !== y.range.max)) out.push(`${col}: 범위가 ${x.range.min}~${x.range.max} / ${y.range.min}~${y.range.max}로 다릅니다`);
      if (x.rescale && y.rescale && JSON.stringify(x.rescale) !== JSON.stringify(y.rescale)) out.push(`${col}: 단위 바꾸기 기준이 다릅니다`);
    }
    // 같은 이름의 검사가 여러 개일 수 있다 — 이름이 같은데 조건 묶음이 다를 때만 알린다
    const whens = list => list.reduce((m, c) => { (m[c.label] = m[c.label] || new Set()).add(JSON.stringify(c.when)); return m; }, {});
    const wa = whens(a.cross_checks || []), wb = whens(b.cross_checks || []);
    Object.keys(wb).forEach(l => {
      if (wa[l] && (wa[l].size !== wb[l].size || [...wb[l]].some(x => !wa[l].has(x)))) out.push(`«${l}»: 조건이 다릅니다`);
    });
    if (a.duplicates && b.duplicates && a.duplicates.action !== b.duplicates.action) out.push("중복 처리 방법이 다릅니다");
    return out;
  }

  // ── 규정 확인 (출장비 등) ─────────────────────────────────────
  /* 우선순위: 1순위 회사 규정(policy.items) → 2순위 회사가 채택한 외부 기준(policy.external, adopted:true 이고 rates 가 있을 때만)
   *          → 3순위 판단 불가(검토 필요). 기준 금액을 추정하지 않고, 원본 금액을 바꾸지 않는다.
   * ctx.fields: 처방 파일의 policy_fields — 어느 칸이 날짜 · 장소 · 박수 · 비용항목 · 통화인지
   * ctx.currency: 파일에 통화 칸이 없을 때 사람이 화면에서 정한 통화 (없으면 판단 불가)
   * 결과 status: «한도 안» · «한도 초과 — 검토 필요» · «예외 승인 조건 해당» · «판단 불가» */
  function checkPolicy(rows, src, policy, ctx) {
    const F = (ctx && ctx.fields) || {};
    const out = [];
    const inRange = (d, from, to) => (!from || d >= from) && (!to || d <= to);
    const matches = (row, m) => Object.entries(m || {}).every(([c, v]) => {
      const col = c === "place" ? F.place : c;
      return col && !isBlank(row[col]) && (Array.isArray(v) ? v.map(String).includes(String(row[col])) : String(row[col]) === String(v));
    });
    const basisOf = (p, r, kind) => [kind, p.name || "", r.source || p.source || "", p.version ? `버전 ${p.version}` : "",
      r.year ? `기준 연도 ${r.year}` : "", (r.valid_from || p.valid_from) ? `적용 ${r.valid_from || p.valid_from} ~ ${r.valid_to || p.valid_to || ""}` : ""].filter(Boolean).join(" · ");

    rows.forEach((row, i) => {
      const rowNo = src ? src[i] : i + 2;
      const key = Object.keys(row)[0] ? row[Object.keys(row)[0]] : "";
      Object.entries(F.items || {}).forEach(([item, col]) => {
        const res = { row: rowNo, key, item, amount: row[col], status: "판단 불가", basis: "", reason: "" };
        out.push(res);
        const missing = [];
        if (isBlank(row[col]) || typeof row[col] !== "number") missing.push(`${item} 금액(숫자)`);
        const date = F.date ? row[F.date] : null;
        if (!F.date || isBlank(date) || !ISO.test(String(date))) missing.push("출장일");
        const cur = (F.currency && !isBlank(row[F.currency])) ? String(row[F.currency]).trim() : (ctx && ctx.currency) || "";
        if (!cur) missing.push("통화 (화면에서 정하거나 통화 칸이 있어야 함)");
        if (!policy) missing.push("회사 규정 파일");
        if (missing.length) { res.reason = `부족한 정보: ${missing.join(" · ")}`; return; }
        if (!inRange(date, policy.valid_from, policy.valid_to)) { res.reason = `출장일 ${date}이 규정 유효기간(${policy.valid_from || ""} ~ ${policy.valid_to || ""}) 밖`; return; }

        // 1순위 회사 규정 — 가장 구체적인(맞춘 조건이 가장 많은) 것. 같은 수준에서 한도가 다르면 판단하지 않는다
        let cands = (policy.items || []).filter(r => r.item === item && matches(row, r.match) && inRange(date, r.valid_from, r.valid_to))
          .map(r => ({ r, kind: "1순위 회사 규정", spec: Object.keys(r.match || {}).length }));
        // 2순위 회사가 채택한 외부 기준 — 회사 규정에 맞는 것이 없을 때만
        if (!cands.length) {
          (policy.external || []).filter(e => e.adopted === true).forEach(e => {
            (e.rates || []).filter(r => r.item === item && matches(row, r.match) && inRange(date, r.valid_from, r.valid_to))
              .forEach(r => cands.push({ r: Object.assign({ source: e.source_url || e.name, year: e.year }, r), kind: `2순위 외부 기준(${e.name})`, spec: Object.keys(r.match || {}).length, p: e }));
          });
        }
        if (!cands.length) { res.reason = `이 ${F.place ? row[F.place] + " · " : ""}${item}에 맞는 회사 규정이나 회사가 채택한 외부 기준이 없음`; return; }
        const top = Math.max(...cands.map(c => c.spec));
        cands = cands.filter(c => c.spec === top);
        if (new Set(cands.map(c => `${c.r.max}|${c.r.currency || policy.currency}|${c.r.per}`)).size > 1) { res.reason = "맞는 규정이 둘 이상이고 한도가 서로 다름"; return; }
        const { r, kind, p } = cands[0];
        let limit = r.max;
        if (typeof limit !== "number") { res.reason = "규정에 한도 금액이 적혀 있지 않음"; return; }
        if (r.per === "night") {
          const n = F.nights ? row[F.nights] : null;
          if (isBlank(n) || typeof n !== "number" || n < 0) { res.reason = "1박당 기준인데 박수를 알 수 없음"; return; }
          limit = limit * n;
        }
        const rc = r.currency || (p && p.currency) || policy.currency;
        if (!rc) { res.reason = "규정의 통화가 적혀 있지 않음"; return; }
        let fxNote = "";
        if (rc !== cur) {
          const fx = (policy.fx || []).find(x => x.from === rc && x.to === cur && (!x.date || x.date <= date));
          if (!fx) { res.reason = `규정은 ${rc}, 금액은 ${cur} — 적용 환율과 환산 기준일이 없음`; return; }
          limit = limit * fx.rate; fxNote = ` (환율 ${fx.rate}, 기준일 ${fx.date || "적혀 있지 않음"})`;
        }
        res.basis = basisOf(p || policy, r, kind) + fxNote;
        res.limit = Math.round(limit);
        if (row[col] <= limit) { res.status = "한도 안"; return; }
        const ex = (r.exceptions || policy.exceptions || []).find(e => (e.when || []).every(w => test(row, w)));
        if (ex) { res.status = "예외 승인 조건 해당"; res.reason = `${ex.label} — 한도 ${fmtN(res.limit)} 초과지만 규정의 예외 조건에 해당`; return; }
        res.status = "한도 초과 — 검토 필요";
        res.reason = `한도 ${fmtN(res.limit)} ${cur}를 넘음. 규정 위반으로 단정하지 않는다 — 승인 여부를 확인`;
      });
    });
    return out;
  }

  const api = { diagnose, ruleUnits, apply, same, compare, conflicts, checkPolicy, FLAG, isBlank, numLike };
  if (typeof module !== "undefined" && module.exports) module.exports = api;
  else root.Clean = api;
})(this);
