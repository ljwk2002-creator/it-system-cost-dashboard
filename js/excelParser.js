/**
 * excelParser.js
 * SheetJS Workbook → 정규화된 항목 목록. 계산·화면 로직은 없다.
 * - Sheet명 / Header 위치 / Column명을 고정하지 않고 탐색한다 (Alias Mapping).
 * - 값을 보정하지 않는다. 해석할 수 없는 값은 issues에 기록하고 원본을 그대로 보존한다.
 */

/* ---------- 날짜 유틸 (UTC 일련번호 = "day number") ---------- */
const DAY = 86400000;
export const dim = (y, m) => new Date(Date.UTC(y, m, 0)).getUTCDate();
export const dn = ({ y, m, d }) => Date.UTC(y, m - 1, d) / DAY;
export const ymd = (n) => {
  const t = new Date(n * DAY);
  return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() };
};
export const addMonths = (n, k) => {
  const { y, m, d } = ymd(n);
  const t = y * 12 + m - 1 + k;
  const Y = Math.floor(t / 12), M = (t % 12) + 1;
  return dn({ y: Y, m: M, d: Math.min(d, dim(Y, M)) });
};

/* ---------- Header Alias ---------- */
// 비교용 정규화: 소문자, 괄호 안 내용·공백·구두점 제거. 예) "총 금액(원)" → "총금액"
export const norm = (s) =>
  String(s ?? '').toLowerCase().replace(/\(.*?\)|\[.*?\]/g, '').replace(/[\s_.:·/\\-]/g, '');

const ALIASES = {
  itemName: ['항목명', '항목', '자산명', '시스템', '시스템명', 'system', 'systemname', '품명'],
  amount: ['총금액', '취득금액', '계약금액', '금액', 'amount'],
  // 전산용역비 Sheet에서 '감가상각 기간'으로 적혀 있어도 Sheet 문맥상 계약기간으로 해석된다.
  period: ['감가상각기간', '상각기간', '계약기간', '사용기간', '기간', 'period'],
  monthly: ['매월감가상각액', '월감가상각액', '월상각액', '매월상각액'],
  nextAmount: ['예상계약금액', '차기계약금액', '연장계약금액'],
  spend: ['지출월', '지출시기', '지급월', '지급시기'],
  remark: ['비고', 'remark', 'remarks', 'note', 'notes'],
};
const FIELD = new Map(Object.entries(ALIASES).flatMap(([f, names]) => names.map((n) => [norm(n), f])));
const FIELD_LABEL = { itemName: '항목명', amount: '금액', period: '기간' };
const REQUIRED = ['itemName', 'amount', 'period'];

const SHEETS = {
  depreciation: { label: '감가상각액', keywords: ['감가상각', '무형자산', '상각'] },
  service: { label: '전산용역비', keywords: ['전산용역', '용역'] },
};

// 전산용역비 비고에 관리자가 명시한 '집계 제외' 사유. 오류가 아니라 알려진 상태이므로 ⚠ 대신 제외로 표시한다.
const EXCLUDE = [
  [/금액\s*미정/, '금액 미정'],
  [/일정\s*미정/, '일정 미정'],
];

// 비고의 Project 비용처리 지시. 예) "2026년~2027년 LNIC Project에서 비용처리" · "2026.12~2028.11 LNIC Project 귀속"
const PROJECT = /([A-Za-z0-9&-]+)\s*(?:project|프로젝트|pjt)(?:\s*에서)?\s*(?:귀속|비용\s*처리)/i;
const PROJECT_RANGE = /(\d{4})\s*(?:[.\-/년]\s*(\d{1,2})(?:\s*[.\-/월]\s*\d{1,2})?)?[^~〜∼\d]{0,3}[~〜∼–-]\s*(\d{4})\s*(?:[.\-/년]\s*(\d{1,2}))?/;

// 비고의 '리스계약 중' 지시: 그 기간 비용은 0원. 예) "2026년~2027년 LNIC Project에서 205백만원 리스계약 중"
const LEASE = /리스\s*계약\s*중/;
// 비고 · 지출 월의 '일시납': 계약기간 월할 안분 대신 지급월에 전액 비용 인식
const LUMP_SUM = /일시납/;

const SINGLE = /(\d{4})\s*(?:년|[.\-/]\s*(\d{1,2}))/; // 기간 대신 한 해(2027년) 또는 한 달(2027.12)
// 특정 연도 화면에만 보이는 비고 줄. 예) "2026년: 첫해 유지보수비 0원" ('0원'이 있으면 그해 계약금액·비용 0원)
const YEAR_NOTE = /^\s*(\d{4})\s*년\s*:\s*(.+)$/;

/** 기간 → { start, end } (월 단위: 시작월 1일 ~ 종료월 말일, 연도만 있으면 1월~12월). 없거나 해석 불가면 전체 기간(null) */
function parseRange(text) {
  const r = text.match(PROJECT_RANGE);
  if (r) {
    const [m1, m2] = [+(r[2] ?? 1), +(r[4] ?? 12)];
    if (m1 < 1 || m1 > 12 || m2 < 1 || m2 > 12) return { start: null, end: null };
    return { start: { y: +r[1], m: m1, d: 1 }, end: { y: +r[3], m: m2, d: dim(+r[3], m2) } };
  }
  const s = text.match(SINGLE);
  const [y, m] = s ? [+s[1], s[2] ? +s[2] : null] : [];
  if (!s || (m != null && (m < 1 || m > 12))) return { start: null, end: null };
  return { start: { y, m: m ?? 1, d: 1 }, end: { y, m: m ?? 12, d: dim(y, m ?? 12) } };
}

const lineOf = (remark, re) => String(remark ?? '').split('\n').find((l) => re.test(l));

/** 비고 → { name, start, end }. 비용은 그대로 집계하고, 해당 기간분을 Project 비용으로 구분 표시한다 (지시가 적힌 줄의 기간). */
export function parseProject(remark) {
  const line = lineOf(remark, PROJECT);
  return line ? { name: line.match(PROJECT)[1], ...parseRange(line) } : null;
}

/** 비고 → 리스계약 기간 { start, end } (그 기간 비용 0원) 또는 null */
export function parseLease(remark) {
  const line = lineOf(remark, LEASE);
  return line ? parseRange(line) : null;
}

/** 비고 → { base: 연도 지정 줄을 뺀 비고, notes: { 연도: 그해에만 보일 비고 }, zeroYears: [계약금액·비용 0원 연도] } */
export function parseYearNotes(remark) {
  const lines = String(remark ?? '').split('\n');
  const notes = {};
  for (const l of lines) { const m = l.match(YEAR_NOTE); if (m) notes[+m[1]] = m[2].trim(); }
  return {
    base: lines.filter((l) => !YEAR_NOTE.test(l)).join('\n').trim(),
    notes,
    zeroYears: Object.keys(notes).filter((y) => /(?<![\d,])0\s*원/.test(notes[y])).map(Number),
  };
}

const TOTAL = /^(sub|grand)?total$|^(소|총|총합)?계$|^합계$/;
const SUBTOTAL = /^subtotal$|^소계$/;
const MAX_ROWS = 3000, MAX_COLS = 60; // 서식만 넓게 잡힌 Sheet 방어

/* ---------- 기간 Parsing ---------- */
// 2024.02 / 2024-02-01 / 2024/2/1 / 2024년 2월 1일
const DATE_TOKEN = /(\d{4})\s*[.\-/년]\s*(\d{1,2})(?:\s*[.\-/월]\s*(\d{1,2}))?\s*[일월]?/g;

/**
 * 기간 문자열 → { status: 'ok'|'invalid'|'unparsed'|'missing', text, start, end, note }
 * 존재하지 않는 날짜(예: 6월 31일)는 JS Date로 넘기기 전에 판정하여 자동 보정을 막는다.
 */
export function parsePeriod(value) {
  const text = String(value ?? '').trim();
  if (!text) return { status: 'missing', text };
  const tokens = [...text.matchAll(DATE_TOKEN)];
  if (tokens.length !== 2) return { status: 'unparsed', text };
  const [a, b] = tokens.map(([raw, y, m, d]) => ({ raw: raw.trim(), y: +y, m: +m, d: d == null ? null : +d }));
  const bad = [a, b].filter((t) => t.y < 1900 || t.y > 2199 || t.m < 1 || t.m > 12 ||
    (t.d != null && (t.d < 1 || t.d > dim(t.y, t.m))));
  if (bad.length) return { status: 'invalid', text, reason: `유효하지 않은 날짜 ${bad.map((t) => t.raw).join(', ')}` };
  const start = { y: a.y, m: a.m, d: a.d ?? 1 };
  const end = { y: b.y, m: b.m, d: b.d ?? dim(b.y, b.m) };
  if (dn(start) > dn(end)) return { status: 'invalid', text, reason: '시작일이 종료일보다 늦음' };
  const note = text.replace(DATE_TOKEN, '').replace(/[~〜∼–—-]/g, '').trim();
  return { status: 'ok', text, start, end, note };
}

/* ---------- Cell helpers ---------- */
const text = (c) => (c == null || c.v == null ? '' : String(typeof c.v === 'string' ? c.v : c.w ?? c.v).trim());

function parseAmount(c) {
  const raw = text(c);
  if (!raw) return { value: null, issue: '금액 공란' };
  if (c.t === 'e') return { value: null, issue: `금액 셀 오류값 ${raw}` };
  const n = typeof c.v === 'number' ? c.v : Number(raw.replace(/[₩,\s원]/g, ''));
  if (!Number.isFinite(n)) return { value: null, issue: `숫자가 아닌 금액 "${raw}"` };
  if (n === 0) return { value: null, issue: '금액 0원 — 확정 금액인지 확인 필요' };
  if (n < 0) return { value: null, issue: `음수 금액 ${raw}` };
  return { value: n };
}

const optionalNumber = (c) => (c && typeof c.v === 'number' && c.t !== 'e' ? c.v : null);

function readGrid(ws, utils) {
  if (!ws?.['!ref']) return { grid: [], r0: 0, c0: 0 };
  const { s, e } = utils.decode_range(ws['!ref']);
  const grid = [];
  for (let r = s.r; r <= Math.min(e.r, s.r + MAX_ROWS); r++) {
    const row = [];
    for (let c = s.c; c <= Math.min(e.c, s.c + MAX_COLS); c++) row.push(ws[utils.encode_cell({ r, c })] ?? null);
    grid.push(row);
  }
  return { grid, r0: s.r, c0: s.c };
}

function findHeader(grid) {
  for (let r = 0; r < Math.min(grid.length, 30); r++) {
    const cols = {};
    grid[r].forEach((cell, c) => {
      const f = FIELD.get(norm(text(cell)));
      if (f && cols[f] == null) cols[f] = c;
    });
    if (cols.itemName != null && Object.keys(cols).length >= 3) return { r, cols };
  }
  return null;
}

function findSheet(names, keywords) {
  for (const k of keywords) {
    const hit = names.find((n) => norm(n).includes(k));
    if (hit) return hit;
  }
  return null;
}

/* ---------- Sheet → items ---------- */
function parseSheet(wb, key, utils) {
  const spec = SHEETS[key];
  const out = { key, label: spec.label, sheet: null, items: [], warnings: [], excelTotal: null, error: null };
  out.sheet = findSheet(wb.SheetNames, spec.keywords);
  if (!out.sheet) {
    out.error = `'${spec.label}' Sheet를 찾을 수 없습니다. (Workbook Sheet: ${wb.SheetNames.join(', ')})`;
    return out;
  }
  const { grid, r0, c0 } = readGrid(wb.Sheets[out.sheet], utils);
  const header = findHeader(grid);
  if (!header) {
    out.error = `'${out.sheet}' Sheet 상단에서 Header 행(항목명 · 금액 · 기간)을 찾을 수 없습니다.`;
    return out;
  }
  const missing = REQUIRED.filter((f) => header.cols[f] == null);
  if (missing.length) {
    out.error = `'${out.sheet}' Sheet에 필수 Column이 없습니다: ${missing.map((f) => FIELD_LABEL[f]).join(', ')}`;
    return out;
  }

  const labels = Object.entries(header.cols).sort((a, b) => a[1] - b[1]).map(([f, c]) => ({ f, c, label: text(grid[header.r][c]) }));
  let afterTotal = false;
  for (let r = header.r + 1; r < grid.length; r++) {
    const row = grid[r], rowNo = r0 + r + 1;
    const cell = (f) => (header.cols[f] == null ? null : row[header.cols[f]]);
    const name = text(cell('itemName'));
    if (!name) {
      if (!afterTotal && labels.some(({ f }) => f !== 'itemName' && text(cell(f))))
        out.warnings.push(`${out.sheet} ${rowNo}행: 항목명 없이 값이 입력되어 집계에서 제외되었습니다.`);
      continue;
    }
    const k = norm(name);
    if (TOTAL.test(k)) {
      if (!SUBTOTAL.test(k) && !afterTotal) {
        out.excelTotal = { rowNo, amount: optionalNumber(cell('amount')), monthly: optionalNumber(cell('monthly')) };
        afterTotal = true;
      }
      continue;
    }
    if (afterTotal) {
      out.warnings.push(`${out.sheet} ${rowNo}행 '${name}': TOTAL 행 아래에 있어 집계에서 제외되었습니다.`);
      continue;
    }

    const amt = parseAmount(cell('amount'));
    const item = {
      id: `${key}-${rowNo}`, kind: key, sheet: out.sheet, rowNo, name,
      amount: amt.value, amountText: text(cell('amount')),
      period: parsePeriod(text(cell('period'))),
      remark: text(cell('remark')),
      ...(({ base, notes, zeroYears }) => ({
        remarkBase: base, yearNotes: notes, zeroYears, project: parseProject(base), lease: parseLease(base),
      }))(parseYearNotes(text(cell('remark')))),
      lumpSum: LUMP_SUM.test(`${text(cell('remark'))} ${text(cell('spend'))}`),
      spendText: text(cell('spend')),
      monthlyExcel: null, nextAmount: optionalNumber(cell('nextAmount')),
      raw: labels.map(({ c, label }) => ({ label, addr: utils.encode_cell({ r: r0 + r, c: c0 + c }), text: text(row[c]) })),
      issues: [],
    };

    if (amt.issue) item.issues.push({ type: 'amount', msg: amt.issue });
    else if (/확인\s*필요/.test(item.remark)) item.issues.push({ type: 'amount', msg: "비고에 '확인 필요' 기재" });

    const p = item.period;
    if (p.status === 'invalid') item.issues.push({ type: 'date', msg: `${p.reason} (원본: ${p.text})` });
    else if (p.status === 'missing') item.issues.push({ type: 'period', msg: '기간 공란' });
    else if (p.status === 'unparsed') item.issues.push({ type: 'period', msg: `날짜로 변환할 수 없는 기간 (원본: ${p.text})` });
    else if (p.note) item.issues.push({ type: 'data', msg: `기간 원문의 부가 정보 "${p.note}" 확인 필요 — 날짜 범위만 계산에 사용` });

    const mc = cell('monthly');
    if (text(mc)) {
      item.monthlyExcel = optionalNumber(mc);
      if (item.monthlyExcel == null) item.issues.push({ type: 'data', msg: `월 상각액 값을 읽을 수 없음 "${text(mc)}"` });
    }

    const ex = key === 'service' && EXCLUDE.find(([re]) => re.test(item.remark));
    if (ex) {
      item.excluded = ex[1];
      item.issues = item.issues.filter((i) => i.type === 'date'); // 금액·기간은 사유로 설명됨. 날짜 오타만 계속 표시
    }
    item.amountOk = !item.excluded && item.amount != null && !item.issues.some((i) => i.type === 'amount');
    out.items.push(item);
  }
  return out;
}

export function parseWorkbook(wb, utils) {
  return {
    sheetNames: wb.SheetNames,
    depreciation: parseSheet(wb, 'depreciation', utils),
    service: parseSheet(wb, 'service', utils),
  };
}
