/**
 * contractTimeline.js — 전산용역비 연도별 금액 · 지출 월 (2026-10-01 지시)
 * 연도 금액
 *   1. Excel에 'YYYY년 금액' 열이 있는 해 → 그 금액 (예: 2026년)
 *   2. 기준 연도(마지막 'YYYY년 금액' 열의 다음 해)까지 → Excel 총 금액 (예: 2027년)
 *   3. 이후 → 매년 3~5% 물가인상 (백만원 단위로 딱 떨어지는 금액)
 *   - 계약 시작 연도 전에는 비용 없음. 단 비고의 Project 비용처리 기간인 해는 비용 발생 (계약 전 Project 비용)
 *   - 비고의 리스계약 기간 · 'YYYY년: …0원' 해는 0원
 * 지출: 그해 '지출 월'(그해 열이 있으면 그 문구)대로 그해 금액을 나눈다 (총 금액 ÷ 지출 횟수).
 */
import { dn, ymd } from './excelParser.js';

const pad = (n) => String(n).padStart(2, '0');
export const fmtDn = (n) => { const { y, m, d } = ymd(n); return `${y}.${pad(m)}.${pad(d)}`; };
const won = (v) => `₩${Math.round(v).toLocaleString('ko-KR')}`;

/* ---------- 지출 월 ---------- */
const num = (s) => Number(String(s).replace(/,/g, ''));
const month = (mo) => (mo && +mo >= 1 && +mo <= 12 ? +mo : null);

/**
 * '지출 월' 문구 해석. 예) "매년 4월 1회" · "2026년 4월 1회" · "매월 6,500,000원씩" · "매월 약 11,000,000원씩"
 *   · "매월 5,500,000원씩 + 6월 28,000,000원" · "매월 5,500,000원씩 + 매년 28,000,000원씩" · "1월 4월 7월 10월 년 4회"
 * → { text, monthly: 월 금액 | 'even'(그해 금액 균등 분할) | null, stated: 기재된 월 금액, lumps: [{ month: 1~12|null, amount|null }] }
 *   해석할 수 없으면 { text, unparsed: true }, 비어 있으면 null
 */
export function parseSpend(value) {
  const text = String(value ?? '').replace(/\s+/g, ' ').trim();
  if (!text) return null;
  const m = text.match(/매월\s*(약\s*)?(?:([\d,]+)\s*원)?/);
  const lumps = [
    ...[...text.matchAll(/매년\s*(?:(\d{1,2})\s*월)?\s*(?:([\d,]+)\s*원)?/g)].map(([, mo, amt]) => ({ month: month(mo), amount: amt ? num(amt) : null })),
    ...[...text.matchAll(/(?<!매년\s*)(?<!\d)(\d{1,2})\s*월\s*([\d,]+)\s*원/g)].map(([, mo, amt]) => ({ month: month(mo), amount: num(amt) })),
  ];
  if (!m && !lumps.length) {
    // "2026년 4월 1회" · "1월 4월 7월 10월 년 4회" → 적힌 달마다 그해 금액 균등 분할
    const listed = /\d+\s*회/.test(text) ? [...text.matchAll(/(?<!\d)(\d{1,2})\s*월/g)].map(([, mo]) => month(mo)).filter(Boolean) : [];
    if (!listed.length) return { text, unparsed: true };
    return { text, stated: null, monthly: null, lumps: listed.map((mo) => ({ month: mo, amount: null })) };
  }
  const stated = m?.[2] ? num(m[2]) : null;
  // '약' 또는 금액 없는 '매월' → 그해 금액 균등 분할 (월 합계가 그해 금액과 일치하도록)
  return { text, stated, monthly: m ? (stated != null && !m[1] ? stated : 'even') : null, lumps };
}

/** 금액이 모두 적힌 문구의 연간 합계 (균등 분할 · 금액 없는 지출이 있으면 null) */
const statedTotal = (rule) => (rule.monthly === 'even' || rule.lumps.some((l) => l.amount == null)
  ? null : (rule.monthly ?? 0) * 12 + rule.lumps.reduce((s, l) => s + l.amount, 0));

/**
 * 그해 금액을 지출 월에 나눔 → { pays: [12], monthly: 월 지출 | null, lumps: [{ m, amount }] }
 * 금액이 모두 적혀 있으면 그해 금액에 맞게 비례 조정 (물가인상 연도). 지급월이 없는 연 1회분은 계약 시작월.
 */
function schedule(rule, amount, startMonth) {
  const pays = Array(12).fill(0);
  if (!rule || rule.unparsed) return { pays: pays.fill(amount / 12), monthly: amount / 12, lumps: [] };
  const lumps = rule.lumps.map((l) => ({ m: l.month ?? startMonth, amount: l.amount }));
  let monthly = rule.monthly === 'even' ? null : rule.monthly;
  const face = statedTotal(rule);
  if (face) {
    const f = amount / face;
    if (monthly != null) monthly *= f;
    lumps.forEach((l) => { l.amount *= f; });
  } else {
    const known = lumps.reduce((s, l) => s + (l.amount ?? 0), 0);
    if (rule.monthly === 'even') monthly = (amount - known) / 12;
    const open = lumps.filter((l) => l.amount == null);
    open.forEach((l) => { l.amount = (amount - (monthly ?? 0) * 12 - known) / open.length; });
  }
  if (monthly) pays.forEach((_, i) => { pays[i] += monthly; });
  lumps.forEach((l) => { pays[l.m - 1] += l.amount; });
  return { pays, monthly: monthly || null, lumps };
}

/** 지출 월 해석 · 금액 합계 점검 (Excel 지출 월 + 'YYYY년 지출 월' 열). baseYear: Excel 총 금액을 쓰는 마지막 해 */
export function prepareContracts(items, baseYear = new Date().getFullYear()) {
  for (const it of items) {
    it.baseYear = baseYear;
    it.cur = it.period.status === 'ok' ? { s: dn(it.period.start), e: dn(it.period.end) } : null;
    if (!it.amountOk) continue;
    const texts = [['지출 월', it.spendText, it.amount],
      ...Object.entries(it.yearSpend ?? {}).map(([y, t]) => [`${y}년 지출 월`, t, it.yearAmounts?.[y] ?? it.amount])];
    for (const [label, text, amount] of texts) {
      const rule = parseSpend(text);
      if (!rule) continue;
      if (rule.unparsed) {
        it.issues.push({ type: 'data', msg: `${label} "${rule.text}"을 해석할 수 없음 — 월별 금액은 매월 균등으로 표시` });
        continue;
      }
      const face = statedTotal(rule);
      if (face != null && Math.abs(face - amount) > Math.max(1, amount * 0.005)) {
        it.issues.push({ type: 'data', msg: `${label} 금액 합계 ${won(face)} ≠ 금액 ${won(amount)} — 금액에 맞춰 비례 표시` });
      }
    }
  }
  return items;
}

/* ---------- 물가인상 ---------- */
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

/* ---------- 연도별 금액 ---------- */
const mi = ({ y, m }) => y * 12 + m - 1;
const within = (r, Y, m) => (!r.start || mi(r.start) <= Y * 12 + m - 1) && (!r.end || Y * 12 + m - 1 <= mi(r.end));
const covers = (r, Y) => !!r && (!r.start || r.start.y <= Y) && (!r.end || r.end.y >= Y);
/** 그 달이 0원인지: 비고의 리스계약 기간 · 'YYYY년: …0원' 해 */
const isZero = (it, Y, m) => !!it.zeroYears?.includes(Y) || (!!it.lease && within(it.lease, Y, m));

/**
 * 선택 연도 금액 · 지출. 금액 · 기간을 확정할 수 없으면 null
 * → { amount: 그해 비용, gross: 0원 처리 전 금액, basis: 'year'|'excel'|'inflation'|'none', pre: 계약 전 Project 비용,
 *     pays: [12] 월별 지출, monthly, lumps: [{ m, amount }], spendText }
 * cmp = true: 전년 대비 변동 비교용 — 연간 계약금액(Excel 총 금액) 기준. 물가인상 · 'YYYY년 금액' 열(실적)과의 차이는 변동으로 보지 않음
 */
export function yearPlan(it, Y, cmp = false) {
  if (!it.cur || !it.amountOk) return null;
  const explicit = cmp ? null : it.yearAmounts?.[Y];
  const startY = ymd(it.cur.s).y;
  const pre = explicit == null && Y < startY && covers(it.project, Y);
  let gross = 0, basis = 'none';
  if (explicit != null) [gross, basis] = [explicit, 'year'];
  else if (Y >= startY || pre) {
    [gross, basis] = [it.amount, Y <= it.baseYear ? 'excel' : 'inflation'];
    if (!cmp) for (let y = it.baseYear; y < Y; y++) gross = escalate(gross);
  }
  const spendText = (!cmp && it.yearSpend?.[Y]) || it.spendText;
  const sch = schedule(parseSpend(spendText), gross, it.period.start.m);
  const pays = sch.pays.map((a, i) => (isZero(it, Y, i + 1) ? 0 : a));
  const amount = pays.reduce((s, a) => s + a, 0);
  return { amount, gross, basis, pre: pre && amount > 0, pays, monthly: sch.monthly, lumps: sch.lumps, spendText };
}

export const yearTotal = (it, Y, cmp = false) => yearPlan(it, Y, cmp)?.amount ?? 0;

/** 선택 연도 금액 중 비고의 Project 비용처리 기간(월 단위)에 지출되는 금액 */
export function projectInYear(it, Y, cmp = false) {
  const p = yearPlan(it, Y, cmp);
  return p && it.project ? p.pays.reduce((s, a, i) => s + (within(it.project, Y, i + 1) ? a : 0), 0) : 0;
}
