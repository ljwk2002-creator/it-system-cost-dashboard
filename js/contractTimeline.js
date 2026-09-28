/**
 * contractTimeline.js
 * 전산용역 계약 금액 산정. 기본 가정: 현재 계약 종료 후 동일 주기 · 동일 금액으로 연장된다.
 * (관리자가 '차기계약금액'을 입력하면 연장 구간에는 그 금액을 사용)
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
  }
  return items;
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

/** [a, b] 구간에 귀속되는 금액 { cur: 현재 계약분, ext: 연장 가정분, total }. 금액·기간을 확정할 수 없으면 null */
export function accrued(it, a, b) {
  if (!it.cur || !it.amountOk) return null;
  const periods = [
    { ...it.cur, kind: 'cur', amount: it.amount },
    ...renewals(it, b).map((r) => ({ ...r, kind: 'ext', amount: it.nextAmount ?? it.amount })),
  ];
  const out = { cur: 0, ext: 0 };
  for (const p of periods) {
    const f = fracMonths(Math.max(p.s, a), Math.min(p.e, b));
    if (f) out[p.kind] += (p.amount * f) / fracMonths(p.s, p.e);
  }
  out.total = out.cur + out.ext;
  return out;
}

const yearBounds = (year) => [dn({ y: year, m: 1, d: 1 }), dn({ y: year, m: 12, d: 31 })];

export const yearTotal = (it, year) => accrued(it, ...yearBounds(year))?.total ?? 0;

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
