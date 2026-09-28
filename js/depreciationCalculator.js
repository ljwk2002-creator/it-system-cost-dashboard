/**
 * depreciationCalculator.js
 * 무형자산 월 상각액 · 연도별 상각비. 월 단위, 시작월·종료월 포함.
 *   총 상각개월 = 종료월 - 시작월 + 1   (2024.02~2029.01 → 60개월)
 *   연도 상각비 = 월 상각액 × 해당 연도와 겹치는 개월 수
 */

const mIdx = ({ y, m }) => y * 12 + m - 1;
const won = (v) => `₩${Math.round(v).toLocaleString('ko-KR')}`;

export function prepareDepreciation(items) {
  for (const it of items) {
    it.monthly = null;
    if (it.period.status !== 'ok') continue;
    it.startMi = mIdx(it.period.start);
    it.endMi = mIdx(it.period.end);
    it.months = it.endMi - it.startMi + 1;
    it.monthlyCalc = it.amount == null ? null : it.amount / it.months;
    // Excel 값이 숫자로 읽히면 우선 사용하되, 계산값과 0.1% 넘게 다르면 숨기지 않고 표시한다.
    it.monthly = it.monthlyExcel ?? it.monthlyCalc;
    if (it.monthlyExcel != null && it.monthlyCalc != null &&
        Math.abs(it.monthlyExcel - it.monthlyCalc) > Math.max(1, it.monthlyCalc * 0.001)) {
      it.issues.push({ type: 'data', msg: `월 상각액 불일치 — Excel ${won(it.monthlyExcel)} / 계산 ${won(it.monthlyCalc)} (총 금액 ÷ ${it.months}개월)` });
    }
  }
  return items;
}

export const monthsInYear = (it, year) =>
  it.monthly == null ? 0 : Math.max(0, Math.min(it.endMi, year * 12 + 11) - Math.max(it.startMi, year * 12) + 1);

export const amountInYear = (it, year) => (it.monthly ?? 0) * monthsInYear(it, year);

export const totalInYear = (items, year) => items.reduce((s, it) => s + amountInYear(it, year), 0);

/** 선택 연도 말 기준 누적 상각개월 */
export const elapsedAtYearEnd = (it, year) => Math.min(it.months, Math.max(0, year * 12 + 11 - it.startMi + 1));

export function statusInYear(it, year) {
  if (it.monthly == null) return 'check';
  if (it.startMi > year * 12 + 11) return 'planned';
  if (it.endMi < year * 12) return 'done';
  return 'active';
}

/** 상각 스케줄이 존재하는 [최소 연도, 최대 연도] */
export function yearSpan(items) {
  const ok = items.filter((it) => it.monthly != null);
  if (!ok.length) return null;
  return [Math.floor(Math.min(...ok.map((it) => it.startMi)) / 12), Math.floor(Math.max(...ok.map((it) => it.endMi)) / 12)];
}
