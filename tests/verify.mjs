// 계산 Logic 검증 — 실행: node tests/verify.mjs [xlsx 경로]
// 1) 합성 입력 단위 검증: 항상 통과해야 한다.
// 2) 예시 Excel 검증(명세 §38): 기대값은 "제공된 예시 Excel" 기준의 검증용 값이며 대시보드 코드에는 없다.
//    data/current.xlsx를 다른 Excel로 교체한 뒤에는 예시 파일 경로를 인자로 주거나 2)의 실패를 무시한다.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { parseWorkbook, parsePeriod } from '../js/excelParser.js';
import * as Dep from '../js/depreciationCalculator.js';
import * as Svc from '../js/contractTimeline.js';

const XLSX = createRequire(import.meta.url)('../js/vendor/xlsx.core.min.js');
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.01, `${msg}: ${a} ≠ ${b}`);

/* ---------- 1) 단위 검증 ---------- */
const bad = parsePeriod('2026.07.01~2027.06.31');
assert.equal(bad.status, 'invalid', '6월 31일은 유효하지 않은 날짜');
assert.match(bad.reason, /2027\.06\.31/);
assert.equal(bad.text, '2026.07.01~2027.06.31', '원본 보존');
for (const t of ['매월', '분기', '개발완료 후 유지보수 계약 예정', '확인 필요']) assert.equal(parsePeriod(t).status, 'unparsed', t);
assert.equal(parsePeriod('').status, 'missing');

const dep = Dep.prepareDepreciation([{ amount: 600, period: parsePeriod('2027.10~2032.09'), monthlyExcel: null, issues: [] }])[0];
assert.equal(dep.months, 60);
assert.equal(Dep.monthsInYear(dep, 2027), 3, '2027.10~12 → 3개월');
assert.equal(Dep.amountInYear(dep, 2027), 30);
assert.equal(Dep.statusInYear(dep, 2026), 'planned');
assert.equal(Dep.statusInYear(dep, 2033), 'done');
const mismatch = Dep.prepareDepreciation([{ amount: 600, period: parsePeriod('2027.10~2032.09'), monthlyExcel: 12, issues: [] }])[0];
assert.ok(mismatch.issues.some((i) => i.type === 'data'), '월 상각액 불일치 감지');

const mk = (name, period) => Svc.prepareContracts([{ name, period: parsePeriod(period), amount: 120, amountOk: true, nextAmount: null, issues: [] }])[0];
const until = 99999;
const r1 = mk('A', '2026.04.01~2027.03.31');
assert.deepEqual([r1.cycle.unit, r1.cycle.n], ['months', 12]);
assert.deepEqual(Svc.renewals(r1, until)[0], { s: r1.cur.e + 1, e: mk('x', '2027.04.01~2028.03.31').cur.e });
const r3 = mk('B', '2025.06.01~2028.05.31');
assert.equal(r3.cycle.n, 36);
assert.equal(Svc.fmtDn(Svc.renewals(r3, until)[0].e), '2031.05.31', '3년 계약 → 2028.06.01~2031.05.31');
assert.equal(mk('C (1년)', '2026.09.26~2027.09.25').cycle.basis, "'1년' 명시");
const rd = mk('D', '2026.02.01~2027.01.30');
assert.deepEqual([rd.cycle.unit, rd.cycle.n], ['days', 364], '월 경계가 아니면 실제 일수');
assert.ok(mk('E (1년)', '2025.01.01~2027.12.31').issues.some((i) => i.type === 'period'), '명시 주기와 실제 기간 불일치 감지');

// Workbook 구조 변화 대응: Header 위치 이동 · Alias · Sheet/Column 누락 · TOTAL 아래 행
const book = (sheets) => {
  const wb = XLSX.utils.book_new();
  for (const [name, aoa] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
  return parseWorkbook(wb, XLSX.utils);
};
const alt = book({
  '무형자산 상각': [[], ['제목'], [], ['자산명', '취득금액', '상각기간', '월 상각액', 'Remarks'], ['X', 1200, '2027.01~2027.12', 100, ''], ['합계', 1200]],
  '전산용역비': [['System', '계약금액', '감가상각 기간'], ['Y', 0, '2027.06.31~2028.06.30'], ['TOTAL', 0], ['Z', 5, '2027.01~2027.12']],
});
assert.equal(alt.depreciation.items.length, 1, 'Header 위치·Alias 탐색');
assert.equal(alt.depreciation.items[0].monthlyExcel, 100);
assert.equal(alt.depreciation.excelTotal.amount, 1200, 'TOTAL 행은 항목이 아닌 검증용');
assert.equal(alt.service.items[0].period.status, 'invalid', "전산용역비의 '감가상각 기간' 헤더 = 계약기간");
assert.ok(alt.service.items[0].issues.some((i) => i.type === 'amount'), '0원 → 금액 확인 필요');
assert.equal(alt.service.warnings.length, 1, 'TOTAL 아래 행은 조용히 버리지 않고 경고');
const ex = book({ '전산용역비': [['항목명', '금액', '계약기간', '비고'], ['P', 0, '분기', 'ABC Project 귀속'], ['Q', 10, '', '금액 미정'], ['R', 0, '', '확인 필요']] }).service.items;
assert.deepEqual(ex.map((i) => i.excluded ?? null), ['Project 귀속', '금액 미정', null], '비고 키워드 → 집계 제외');
assert.ok(ex[0].issues.length === 0 && ex[2].issues.length > 0, '제외 사유가 없으면 여전히 ⚠');
const broken = book({ '감가상각액': [['항목명', '비고', '금액']], Sheet2: [['a']] });
assert.match(broken.depreciation.error, /필수 Column.*기간/, '필수 Column 누락 오류');
assert.match(broken.service.error, /Sheet를 찾을 수 없습니다/, '필수 Sheet 누락 오류');
console.log('✓ 단위 검증 통과');

/* ---------- 2) 예시 Excel 검증 ---------- */
const file = process.argv[2] ?? new URL('../data/current.xlsx', import.meta.url);
const data = parseWorkbook(XLSX.read(readFileSync(file), { type: 'buffer', cellFormula: false }), XLSX.utils);
const D = Dep.prepareDepreciation(data.depreciation.items);
const S = Svc.prepareContracts(data.service.items);
const sum = (a, f) => a.reduce((s, x) => s + (f(x) ?? 0), 0);
const won = (v) => `₩${Math.round(v).toLocaleString('ko-KR')}`;

console.log(`\nSheet: ${data.depreciation.sheet} (${D.length}건) / ${data.service.sheet} (${S.length}건)`);
const totalAmt = sum(D, (i) => i.amount), totalMonthly = sum(D, (i) => i.monthly);
console.log(`무형자산 총 금액 ${won(totalAmt)} | 월 상각액 합계 ${totalMonthly}`);
const [y0, y1] = Dep.yearSpan(D);
for (let y = y0; y <= y1; y++) console.log(`  ${y}년 상각비 ${won(Dep.totalInYear(D, y)).padStart(16)}`);

const Y = 2027;
console.log(`\n${Y}년 전산용역비 (연장 가정 포함 · 월할 안분)            현재분         연장분`);
for (const it of S) {
  const v = Svc.yearView(it, Y);
  const cols = v ? [v.total, v.cur, v.ext].map((x) => won(x).padStart(14)).join(' ') : (it.excluded ? `집계 제외(${it.excluded})` : '⚠ 미집계').padStart(14);
  console.log(`  ${cols}  ${it.name}${it.issues.length ? `  ⚠ ${it.issues.map((i) => i.msg).join(' / ')}` : ''}`);
}
const views = S.map((i) => Svc.yearView(i, Y)).filter(Boolean);
console.log(`  ${Y}년 전산용역비 합계 ${won(sum(views, (v) => v.total))} (현재 계약분 ${won(sum(views, (v) => v.cur))} + 연장 가정분 ${won(sum(views, (v) => v.ext))})`);
console.log(`  월별: ${Array.from({ length: 12 }, (_, m) => (sum(views, (v) => v.months[m]) / 1e6).toFixed(1)).join(' / ')} (백만원)`);

near(totalAmt, 4165396425, '총 무형자산 금액');
near(totalMonthly, 69423273.75, '월 감가상각액 합계');
near(Dep.totalInYear(D, 2027), 833079285, '2027년 상각비 = 69,423,273.75 × 12');
near(totalAmt, data.depreciation.excelTotal.amount, 'Excel TOTAL 행과 일치');

const pds = S.find((i) => i.name.startsWith('PDS'));
const pv = Svc.yearView(pds, 2027);
const [next] = Svc.renewals(pds, pds.cur.e + 1);
assert.equal(`${Svc.fmtDn(next.s)}~${Svc.fmtDn(next.e)}`, '2027.04.01~2028.03.31', 'PDS 연장 가정 기간');
near(pv.cur, 40000000, 'PDS 2027 현재 계약분 = 160,000,000 × 3/12');
near(pv.ext, 120000000, 'PDS 2027 연장 가정분 = 160,000,000 × 9/12');
pv.months.forEach((a, i) => near(a, 160000000 / 12, `PDS ${i + 1}월 = 160,000,000 ÷ 12`));
near(sum(pv.months, (a) => a), pv.total, '월별 합 = 연간 합');

// 2026-09-28 사용자 지시 반영분 (data/current.xlsx 비고·기간)
const find = (n) => S.find((i) => i.name.startsWith(n));
near(Svc.yearView(find('배관설계관리시스템(S-GEN)'), 2027).total, 23000000, 'S-GEN 자동 연장: 2027년 23,000,000');
near(Svc.yearView(find('Agentic AI(유지보수)'), 2027).total, 100000000, 'Agentic AI 2026.10.30~ 1년 계약 + 연장 → 2027년 100,000,000');
assert.deepEqual(
  S.filter((i) => i.excluded).map((i) => [i.name.split(' ')[0], i.excluded]),
  [['ePMCS', '일정 미정'], ['MIDAS', '금액 미정'], ['벤틀리', 'Project 귀속']], '비고 기반 집계 제외');
assert.ok(S.filter((i) => i.excluded).every((i) => !i.issues.length && Svc.yearView(i, 2027) === null), '집계 제외는 오류가 아니며 합계에서 빠짐');
assert.equal(S.filter((i) => !i.excluded && !Svc.yearView(i, 2027)).length, 0, '미집계(오류) 0건');
console.log('\n✓ 예시 Excel 검증 통과 (§38)');
