/**
 * contractTimeline.js
 * 전산용역 계약 금액 산정. 기본 가정: 현재 계약 종료 후 동일 주기로 연장되며, 갱신마다 연 3~5% 물가인상
 * (백만원 단위로 딱 떨어지는 금액). 관리자가 '차기계약금액'을 입력하면 첫 갱신에는 그 금액을 사용.
 * 비고에 계약 시작 전부터의 Project 비용처리 기간이 있으면 그 기간에도 같은 금액·주기로 비용이 발생한다고 본다.
 * 금액은 계약기간에 걸쳐 월할 안분한다 (각 월의 포함 일수 ÷ 그 달 일수 기준).
 */
import { dn, ymd, dim, addMonths } from './excelParser.js';

const EXPLICIT = /(?<![\d.])(\d{1,2})\s*(년|개월)/;
const pad = (n) => String(n).padStart(2, '0');
export const fmtDn = (n) => { const { y, m, d } = ymd(n); return `${y}.${pad(m)}.${pad(d)}`; };
export const cycleLabel = (c) => (c.unit === 'days' ? `${c.n}일` : c.n % 12 ? `${c.n}개월` : `${c.n / 12}년`);

/** a~b(포함) 구간의 월 환산 길이 = Σ(해당 월 포함 일수 ÷ 그 달 일수). 04.01~03.31 → 12.0 */
export function fracMonths(a, b) {
  let total = 0;
  for (let d = a; d <= b;) {
    const { y, m } = ymd(d);
    const days = dim(y, m);
    const e = Math.min(b, dn({ y, m, d: days }));
    total += (e - d + 1) / days;
    d = e + 1;
  }
  return total;
}

/**
 * 계약 주기 산정 우선순위
 * 1. 항목명·기간 원문에 '1년', '3년', 'N개월' 명시
 * 2. 시작일 + N개월 - 1일 = 종료일이 성립하는 달력 월 주기
 * 3. 실제 일수
 * (기간 자체를 해석할 수 없으면 prepareContracts에서 cur=null → '확인 필요')
 */
function detectCycle(it) {
  const hit = `${it.name} ${it.period.text}`.match(EXPLICIT);
  if (hit && +hit[1] > 0) return { unit: 'months', n: +hit[1] * (hit[2] === '년' ? 12 : 1), basis: `'${hit[0]}' 명시` };
  const { s, e } = it.cur, { start, end } = it.period;
  const approx = (end.y - start.y) * 12 + (end.m - start.m);
  for (const n of [approx, approx + 1]) {
    if (n > 0 && addMonths(s, n) - 1 === e) return { unit: 'months', n, basis: '시작·종료일 기준 월 단위' };
  }
  return { unit: 'days', n: e - s + 1, basis: '실제 일수 기준' };
}

export function prepareContracts(items) {
  for (const it of items) {
    it.cur = null;
    if (it.period.status !== 'ok') continue;
    it.cur = { s: dn(it.period.start), e: dn(it.period.end) };
    it.cycle = detectCycle(it);
    const actual = fracMonths(it.cur.s, it.cur.e);
    if (it.cycle.unit === 'months' && Math.abs(actual - it.cycle.n) > 1) {
      it.issues.push({ type: 'period', msg: `명시된 계약 주기(${cycleLabel(it.cycle)})와 현재 계약기간(약 ${actual.toFixed(1)}개월)이 다름 — 명시 주기로 연장 가정` });
    }
    const rule = parseSpend(it.spendText);
    if (!rule || !it.amountOk) continue;
    if (rule.unparsed) {
      it.issues.push({ type: 'data', msg: `지출 월 "${rule.text}"을 해석할 수 없음 — 월별 금액은 월할 안분으로 표시` });
      continue;
    }
    const paid = payments({ ...it.cur, amount: it.amount }, rule, 1).reduce((a, x) => a + x.amount, 0);
    if (Math.abs(paid - it.amount) > Math.max(1, it.amount * 0.005)) {
      it.issues.push({ type: 'data', msg: `지출 월 금액 합계 ₩${Math.round(paid).toLocaleString('ko-KR')} ≠ 계약금액 ₩${it.amount.toLocaleString('ko-KR')}` });
    }
  }
  return items;
}

/* ---------- 지출 월 ---------- */
const num = (s) => Number(String(s).replace(/,/g, ''));

/**
 * '지출 월' 문구 해석. 예) "매년 4월 1회" · "매월 6,500,000원씩" · "매월 5,500,000원씩 + 매년 28,000,000원씩" · "매월 약 11,000,000원씩"
 * → { text, monthly: 월 금액 | 'even'(계약금액 균등 분할) | null, stated: 기재된 월 금액, lumps: [{ month: 1~12|null, amount|null }] }
 *   해석할 수 없으면 { text, unparsed: true }, 비어 있으면 null
 */
export function parseSpend(value) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  if (!text) return null;
  const m = text.match(/매월\s*(약\s*)?(?:([\d,]+)\s*원)?/);
  const lumps = [...text.matchAll(/매년\s*(?:(\d{1,2})\s*월)?\s*(?:([\d,]+)\s*원)?/g)].map(([, mo, amt]) => ({
    month: mo && +mo >= 1 && +mo <= 12 ? +mo : null,
    amount: amt ? num(amt) : null,
  }));
  if (!m && !lumps.length) {
    // "1월 4월 7월 10월 년 4회" → 적힌 달마다 계약금액 균등 분할
    const listed = /[년연]\s*\d+\s*회/.test(text)
      ? [...text.matchAll(/(?<!\d)(\d{1,2})\s*월/g)].map(([, mo]) => +mo).filter((mo) => mo >= 1 && mo <= 12) : [];
    if (!listed.length) return { text, unparsed: true };
    return { text, stated: null, monthly: null, lumps: listed.map((month) => ({ month, amount: null })) };
  }
  const stated = m?.[2] ? num(m[2]) : null;
  // '약' 또는 금액 없는 '매월' → 계약금액을 균등 분할 (월 합계가 계약금액과 일치하도록)
  return { text, stated, monthly: m ? (stated != null && !m[1] ? stated : 'even') : null, lumps };
}

/** 한 계약기간의 지출 일정 [{ mi: 월 인덱스(y*12+m-1), amount }]. factor = 이 기간 금액 ÷ 현재 계약금액 */
function payments(p, rule, factor) {
  const { y, m } = ymd(p.s);
  const s = y * 12 + m - 1;
  const n = Math.max(1, Math.round(fracMonths(p.s, p.e)));
  const known = rule.lumps.reduce((a, l) => a + (l.amount ?? 0) * factor, 0);
  const open = rule.lumps.filter((l) => l.amount == null).length;
  const monthly = rule.monthly === 'even' ? (p.amount - known) / n : (rule.monthly ?? 0) * factor;
  const rest = p.amount - known - monthly * n; // 금액이 적히지 않은 '매년 N월' 몫
  const out = [];
  if (rule.monthly != null) for (let k = 0; k < n; k++) out.push({ mi: s + k, amount: monthly });
  for (const l of rule.lumps) {
    let mi = s; // 지급월이 없으면 계약 시작월로 가정
    if (l.month) for (let k = 0; k < 12; k++) if ((s + k) % 12 === l.month - 1) { mi = s + k; break; }
    out.push({ mi, amount: l.amount != null ? l.amount * factor : open ? rest / open : 0 });
  }
  return out;
}

/**
 * 선택 연도의 월별 지출액 (지출 월 기준, 계약 전 Project 기간 · 연장 가정 포함).
 * → { months: [12], kind: [12]('pre'|'ext'|null — 가정분 지출이면 그 종류), rule } / 지출 월이 없거나 해석 불가면 null
 */
export function spendInYear(it, year) {
  const rule = parseSpend(it.spendText);
  if (!it.cur || !it.amountOk || !rule || rule.unparsed) return null;
  const months = Array(12).fill(0), kind = Array(12).fill(null);
  const zero = zeroSpans(it);
  for (const p of allPeriods(it, dn({ y: year, m: 12, d: 31 }))) {
    for (const x of payments(p, rule, p.amount / it.amount)) {
      if (Math.floor(x.mi / 12) !== year || !x.amount) continue;
      const d = dn({ y: year, m: (x.mi % 12) + 1, d: 1 });
      if (zero.some(([a, b]) => a <= d && d <= b)) continue; // 리스계약 기간 · 0원 해
      months[x.mi % 12] += x.amount;
      if (p.kind !== 'cur') kind[x.mi % 12] = p.kind;
    }
  }
  return { months, kind, rule };
}

/** 현재 계약 종료 다음 날부터 동일 주기로 반복한 연장 가정 기간 (until 이전에 시작하는 것까지) */
export function renewals(it, until) {
  const out = [];
  for (let s = it.cur.e + 1; s <= until && out.length < 600;) {
    const e = it.cycle.unit === 'months' ? addMonths(s, it.cycle.n) - 1 : s + it.cycle.n - 1;
    out.push({ s, e });
    s = e + 1;
  }
  return out;
}

/* ---------- 물가인상 (연장 가정 금액) ---------- */
export const INFLATION = { min: 0.03, max: 0.05, target: 0.04 };
const MIL = 1e6;

/**
 * 전년 금액 → 다음 해 금액. 3~5% 인상 범위에서 백만원 단위로 딱 떨어지는 금액 중 4%에 가장 가까운 것.
 * 범위 안에 백만원 단위 금액이 없으면(소액 계약) 4% 인상액을 백만원 단위로 반올림.
 */
export function escalate(amount) {
  const lo = Math.ceil((amount * (1 + INFLATION.min)) / MIL - 1e-9);
  const hi = Math.floor((amount * (1 + INFLATION.max)) / MIL + 1e-9);
  const target = Math.round((amount * (1 + INFLATION.target)) / MIL);
  return (lo <= hi ? Math.min(hi, Math.max(lo, target)) : target) * MIL;
}

/**
 * 연장 가정 기간과 금액. 갱신마다 계약 연수만큼 물가인상 반영 (첫 갱신에 '차기계약금액'이 있으면 그 금액)
 * flat = true: 물가인상 없이 (전년 대비 변동 비교용)
 */
export function renewalPeriods(it, until, flat = false) {
  const years = Math.max(1, Math.round((it.cycle.unit === 'months' ? it.cycle.n : it.cycle.n / 30.44) / 12));
  let amount = it.amount;
  return renewals(it, until).map((r, i) => {
    if (i === 0 && it.nextAmount != null) amount = it.nextAmount;
    else if (!flat) for (let k = 0; k < years; k++) amount = escalate(amount);
    return { ...r, amount };
  });
}

/**
 * 계약 시작 전이 비고의 Project 비용처리 기간에 걸리면, 그 기간에도 같은 금액·같은 주기로 비용이 발생한다고 본다.
 * 예) 계약 2028.12.01~ + 비고 "2026.12~2028.11 LNIC Project 귀속" → 2026.12.01~2027.11.30, 2027.12.01~2028.11.30
 */
export function priorPeriods(it) {
  const ps = it.project?.start ? dn(it.project.start) : null;
  const out = [];
  // Project 시작일 이후에 시작하는 주기만 (예: 비고 2026년~ + 12월 시작 계약 → 2026.12부터)
  for (let s = it.cur.s; ps != null && out.length < 600;) {
    const ns = it.cycle.unit === 'months' ? addMonths(s, -it.cycle.n) : s - it.cycle.n;
    if (ns < ps) break;
    out.unshift({ s: ns, e: s - 1, amount: it.amount });
    s = ns;
  }
  return out;
}

const mOf = (n) => { const { y, m } = ymd(n); return y * 12 + m - 1; };
/** 비고의 리스계약 기간(비용 0원) [시작, 종료] day number. 기간 미기재 = 전체 */
/** 비용 0원 구간 [시작, 종료] day number 목록: 비고의 리스계약 기간 + 연도 지정 비고의 '0원' 해 */
const zeroSpans = (it) => [
  ...(it.lease ? [[it.lease.start ? dn(it.lease.start) : -Infinity, it.lease.end ? dn(it.lease.end) : Infinity]] : []),
  ...(it.zeroYears ?? []).map((y) => [dn({ y, m: 1, d: 1 }), dn({ y, m: 12, d: 31 })]),
];
/** [s, e]에서 spans를 뺀 구간들 */
const minus = (s, e, spans) => spans.reduce((parts, [x, y]) =>
  parts.flatMap(([a, b]) => [[a, Math.min(b, x - 1)], [Math.max(a, y + 1), b]]).filter(([a, b]) => a <= b), [[s, e]]);

/**
 * 비용 기간 전체: 계약 전 Project 기간(pre) → 현재 계약(cur) → 연장 가정(ext, until 이전 시작분까지)
 * cmp = true (연간 계약금액 비교용): 물가인상 제외 + 첫 기간이 연중에 시작하면 그해 1월부터 시작 전까지를 같은 금액으로 채움
 */
function allPeriods(it, until, cmp = false) {
  const out = [
    ...priorPeriods(it).map((p) => ({ ...p, kind: 'pre' })),
    { ...it.cur, amount: it.amount, kind: 'cur' },
    ...renewalPeriods(it, until, cmp).map((p) => ({ ...p, kind: 'ext' })),
  ];
  const first = out[0], jan1 = dn({ y: ymd(first.s).y, m: 1, d: 1 });
  if (cmp && first.s > jan1) {
    const s = it.cycle.unit === 'months' ? addMonths(first.s, -it.cycle.n) : first.s - it.cycle.n;
    out.unshift({ s, e: first.s - 1, from: jan1, amount: first.amount, kind: 'cur' });
  }
  return out;
}

/**
 * [a, b] 구간에 귀속되는 금액 { pre: 계약 전 Project 기간분, cur: 현재 계약분, ext: 연장 가정분, total }. 금액·기간을 확정할 수 없으면 null
 * 기본은 월할 안분. 비고 '일시납'이면 지급월(지출 월)에 전액. 비고 '리스계약 중' 기간은 0원. cmp: 연간 계약금액 비교용 (allPeriods)
 */
export function accrued(it, a, b, cmp = false) {
  if (!it.cur || !it.amountOk) return null;
  const out = { pre: 0, cur: 0, ext: 0 };
  const rule = it.lumpSum ? parseSpend(it.spendText) : null;
  const zero = zeroSpans(it);
  const add = (p, s, e, f) => { // [s, e] 중 0원 구간을 뺀 부분만 f(s, e)로 더함
    for (const [x, y] of minus(s, e, zero)) out[p.kind] += f(x, y);
  };
  for (const p of allPeriods(it, b, cmp)) {
    const from = Math.max(a, p.from ?? a);
    if (rule && !rule.unparsed) {
      const pays = payments(p, rule, p.amount / it.amount);
      add(p, from, b, (x, y) => pays.filter((q) => q.mi >= mOf(x) && q.mi <= mOf(y)).reduce((s, q) => s + q.amount, 0));
    } else {
      add(p, Math.max(p.s, from), Math.min(p.e, b), (x, y) => (p.amount * fracMonths(x, y)) / fracMonths(p.s, p.e));
    }
  }
  out.total = out.pre + out.cur + out.ext;
  return out;
}

const yearBounds = (year) => [dn({ y: year, m: 1, d: 1 }), dn({ y: year, m: 12, d: 31 })];

export const yearTotal = (it, year, cmp = false) => accrued(it, ...yearBounds(year), cmp)?.total ?? 0;

/** 선택 연도 금액 중 비고의 Project 비용처리 기간에 해당하는 금액 (기간 미기재 = 전체) */
export function projectInYear(it, year, cmp = false) {
  const p = it.project;
  if (!p) return 0;
  const [a, b] = yearBounds(year);
  const s = Math.max(a, p.start ? dn(p.start) : a), e = Math.min(b, p.end ? dn(p.end) : b);
  return s <= e ? accrued(it, s, e, cmp)?.total ?? 0 : 0;
}

/** 선택 연도 금액 + 1~12월 월별 금액. 집계 불가 항목은 null */
export function yearView(it, year) {
  const total = accrued(it, ...yearBounds(year));
  if (!total) return null;
  const months = [];
  for (let m = 1; m <= 12; m++) {
    const ms = dn({ y: year, m, d: 1 });
    months.push(accrued(it, ms, ms + dim(year, m) - 1).total);
  }
  return { ...total, months };
}
