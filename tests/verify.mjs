// 계산 Logic 검증 — 실행: node tests/verify.mjs [xlsx 경로]
// 1) 합성 입력 단위 검증: 항상 통과해야 한다.
// 2) 예시 Excel 검증(명세 §38): 기대값은 "제공된 예시 Excel" 기준의 검증용 값이며 대시보드 코드에는 없다.
//    data/current.xlsx를 다른 Excel로 교체한 뒤에는 예시 파일 경로를 인자로 주거나 2)의 실패를 무시한다.
import assert from 'node:assert/strict';
import { readFileSync } from 'node:fs';
import { createRequire } from 'node:module';
import { parseWorkbook, parsePeriod, parseProject, parseLease, parseYearNotes } from '../js/excelParser.js';
import * as Dep from '../js/depreciationCalculator.js';
import * as Svc from '../js/contractTimeline.js';

const XLSX = createRequire(import.meta.url)('../js/vendor/xlsx.core.min.js');
const near = (a, b, msg) => assert.ok(Math.abs(a - b) < 0.01, `${msg}: ${a} ≠ ${b}`);
const M = (v) => Math.round(v / 1e5) / 10; // 백만원, 소수점 1자리

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

// Workbook 구조 변화 대응: Header 위치 이동 · Alias · Sheet/Column 누락 · TOTAL 아래 행 · 연도별 열
const book = (sheets) => {
  const wb = XLSX.utils.book_new();
  for (const [name, aoa] of Object.entries(sheets)) XLSX.utils.book_append_sheet(wb, XLSX.utils.aoa_to_sheet(aoa), name);
  return parseWorkbook(wb, XLSX.utils);
};
const alt = book({
  '무형자산 상각': [[], ['제목'], [], ['자산명', '취득금액', '상각기간', '월 상각액', 'Remarks'], ['X', 1200, '2027.01~2027.12', 100, ''], ['합계', 1200]],
  '전산용역비': [['System', '계약금액', '감가상각 기간', '2026년 금액', '2026년 지출 월'], ['Y', 0, '2027.06.31~2028.06.30', '-', ''], ['W', 5, '2027.01~2027.12', 4, '2026년 3월 1회'], ['TOTAL', 0], ['Z', 5, '2027.01~2027.12']],
});
assert.equal(alt.depreciation.items.length, 1, 'Header 위치·Alias 탐색');
assert.equal(alt.depreciation.items[0].monthlyExcel, 100);
assert.equal(alt.depreciation.excelTotal.amount, 1200, 'TOTAL 행은 항목이 아닌 검증용');
assert.equal(alt.service.items[0].period.status, 'invalid', "전산용역비의 '감가상각 기간' 헤더 = 계약기간");
assert.ok(alt.service.items[0].issues.some((i) => i.type === 'amount'), '0원 → 금액 확인 필요');
assert.equal(alt.service.warnings.length, 1, 'TOTAL 아래 행은 조용히 버리지 않고 경고');
assert.deepEqual(alt.service.yearCols, [2026], "'2026년 금액' 열 인식");
assert.deepEqual([alt.service.items[0].yearAmounts, alt.service.items[1].yearAmounts, alt.service.items[1].yearSpend],
  [{}, { 2026: 4 }, { 2026: '2026년 3월 1회' }], "연도별 금액 · 지출 월 ('-'는 없음)");
const ex = book({ '전산용역비': [['항목명', '금액', '계약기간', '비고'], ['P', 10, '2027.01~2027.12', 'ABC Project 귀속'], ['Q', 10, '', '금액 미정'], ['R', 0, '', '확인 필요']] }).service.items;
assert.deepEqual(ex.map((i) => i.excluded ?? null), [null, '금액 미정', null], '비고 키워드 → 집계 제외 (Project 귀속은 집계 + 구분)');
assert.equal(ex[0].project?.name, 'ABC', 'Project 귀속 → project');
assert.ok(ex[1].issues.length === 0 && ex[2].issues.length > 0, '제외 사유가 없으면 여전히 ⚠');
const broken = book({ '감가상각액': [['항목명', '비고', '금액']], Sheet2: [['a']] });
assert.match(broken.depreciation.error, /필수 Column.*기간/, '필수 Column 누락 오류');
assert.match(broken.service.error, /Sheet를 찾을 수 없습니다/, '필수 Sheet 누락 오류');

// 비고의 Project 비용처리 기간 · 리스계약 · 연도 지정 비고
assert.deepEqual(parseProject('2026년~2027년 LNIC Project에서 비용처리'), { name: 'LNIC', start: { y: 2026, m: 1, d: 1 }, end: { y: 2027, m: 12, d: 31 } });
assert.deepEqual(parseProject('AVEVA E3D\n2026.12~2028.11 LNIC Project 귀속'), { name: 'LNIC', start: { y: 2026, m: 12, d: 1 }, end: { y: 2028, m: 11, d: 30 } });
assert.deepEqual(parseProject('2026.12.01~2028.11.30 LNIC Project 귀속').end, { y: 2028, m: 11, d: 30 }, '일 단위 표기 → 월 단위');
assert.deepEqual(parseProject('LNIC PDH-PP Project 귀속'), { name: 'PDH-PP', start: null, end: null }, '기간 없음 → 전체 기간');
assert.equal(parseProject('AVEVA E3D'), null);
assert.deepEqual(parseProject('AVEVA E3D\n2027년 LNIC Project 귀속 (12월 일시납)'), { name: 'LNIC', start: { y: 2027, m: 1, d: 1 }, end: { y: 2027, m: 12, d: 31 } }, '한 해만');
assert.deepEqual(parseLease('2026년~2027년 LNIC Project에서 205백만원 리스계약 중'), { start: { y: 2026, m: 1, d: 1 }, end: { y: 2027, m: 12, d: 31 } });
assert.equal(parseProject('2026년~2027년 LNIC Project에서 205백만원 리스계약 중'), null, '리스계약은 Project 비용 아님');
assert.deepEqual(parseYearNotes('2026년: 첫해 유지보수비 0원\nAVEVA E3D'), { base: 'AVEVA E3D', notes: { 2026: '첫해 유지보수비 0원' }, zeroYears: [2026] });
assert.deepEqual(parseYearNotes('2026년: 100,000원 확인').zeroYears, [], "'100,000원'은 0원 아님");
const dp = Dep.prepareDepreciation([{ amount: 600, period: parsePeriod('2026.12~2031.11'), monthlyExcel: null, project: parseProject('2026.12~2028.12 X Project 귀속'), issues: [] }])[0];
assert.deepEqual([2026, 2027, 2028, 2029].map((y) => Dep.projectInYear(dp, y)), [10, 120, 120, 0], '상각 Project 기간분');

// 지출 월 해석
assert.deepEqual(Svc.parseSpend('매년 4월 1회').lumps, [{ month: 4, amount: null }]);
assert.deepEqual(Svc.parseSpend('2026년 4월 1회').lumps, [{ month: 4, amount: null }], '연도가 붙은 1회 지출');
assert.equal(Svc.parseSpend('매월 6,500,000원씩').monthly, 6500000);
assert.equal(Svc.parseSpend('매월 약 11,000,000원씩').monthly, 'even', "'약' → 그해 금액 균등 분할");
assert.deepEqual(Svc.parseSpend('매월 5,500,000원씩 +\n매년 28,000,000원씩').lumps, [{ month: null, amount: 28000000 }]);
assert.deepEqual(Svc.parseSpend('매월 5,500,000원씩 +\n6월 28,000,000원').lumps, [{ month: 6, amount: 28000000 }], '달 + 금액');
assert.equal(Svc.parseSpend('매년 6월 28,000,000원').lumps.length, 1, '매년 N월 금액은 한 번만');
assert.ok(Svc.parseSpend('2028년 6월 재계약 예정').unparsed);
assert.deepEqual(Svc.parseSpend('1월 4월 7월 10월\n년 4회').lumps, [1, 4, 7, 10].map((month) => ({ month, amount: null })), '적힌 달마다 균등');

// 전산용역비 연도별 금액: 'YYYY년 금액' 열 → 기준 연도(2027)까지 Excel 총 금액 → 이후 매년 3~5% 인상
const mk = (o) => Svc.prepareContracts([{ name: 'A', period: parsePeriod('2026.04.01~2027.03.31'), amount: 120e6, amountOk: true,
  spendText: '매년 4월 1회', issues: [], ...o }], 2027)[0];
const a = mk({ yearAmounts: { 2026: 100e6 }, yearSpend: { 2026: '2026년 4월 1회' } });
assert.deepEqual([2025, 2026, 2027, 2028, 2029].map((y) => M(Svc.yearTotal(a, y))), [0, 100, 120, 125, 130], '계약 전 0 · 2026 열 · 2027 Excel · 2028부터 인상');
assert.deepEqual([2026, 2027, 2028].map((y) => M(Svc.yearTotal(a, y, true))), [120, 120, 120], "변동 비교용: 연간 계약금액 기준 (물가인상 · '2026년 금액' 실적 차이 제외)");
assert.deepEqual(Svc.yearPlan(a, 2026).pays.map(M), [0, 0, 0, 100, 0, 0, 0, 0, 0, 0, 0, 0], '2026년 4월 1회');
const doc = mk({ amount: 94e6, period: parsePeriod('2026.06.01~2027.05.31'), spendText: '매월 5,500,000원씩 + 매년 28,000,000원씩',
  yearAmounts: { 2026: 94e6 }, yearSpend: { 2026: '매월 5,500,000원씩 +\n6월 28,000,000원' } });
near(Svc.yearPlan(doc, 2026).pays[5], 33.5e6, '6월 = 매월 5.5 + 28');
near(Svc.yearPlan(doc, 2027).pays[5], 33.5e6, '지급월 없는 연 1회분 → 계약 시작월(6월)');
near(Svc.yearPlan(doc, 2028).monthly, 5.5e6 * 98 / 94, '물가인상 연도: 적힌 금액을 비례 조정 (94 → 98)');
const nav = mk({ amount: 138e6, spendText: '매월 약 11,000,000원씩', yearAmounts: { 2026: 132e6 }, yearSpend: { 2026: '매월 약 11,000,000원씩' } });
assert.deepEqual([2026, 2027].map((y) => M(Svc.yearPlan(nav, y).monthly)), [11, 11.5], "'매월 약' → 그해 금액 ÷ 12");
const j = mk({ period: parsePeriod('2028.12.01~2029.11.30'), amount: 102e6, spendText: '매년 12월 1회',
  project: parseProject('2027년 X Project 귀속'), zeroYears: [2026], yearAmounts: { 2026: 102e6 } });
assert.deepEqual([2026, 2027, 2028, 2029].map((y) => M(Svc.yearTotal(j, y))), [0, 102, 106, 110], "2026 '0원' 비고 · 2027 계약 전 Project 비용 · 2028부터 인상");
assert.deepEqual([2026, 2027, 2028].map((y) => M(Svc.projectInYear(j, y))), [0, 102, 0], 'Project 기간 지출분');
assert.ok(Svc.yearPlan(j, 2027).pre && !Svc.yearPlan(j, 2028).pre, '계약 전 Project 비용 표시');
const l = mk({ period: parsePeriod('2026.01.01~2026.12.31'), amount: 64e6, spendText: '1월 4월 7월 10월 년 4회', lease: parseLease('2026년~2027년 리스계약 중') });
assert.deepEqual([2026, 2027, 2028].map((y) => M(Svc.yearTotal(l, y))), [0, 0, 67], '리스계약 기간 0원 · 2028 64 → 67');
assert.deepEqual(Svc.yearPlan(l, 2028).pays.map((x) => x / 1e6), [16.75, 0, 0, 16.75, 0, 0, 16.75, 0, 0, 16.75, 0, 0], '1·4·7·10월 균등');
assert.ok(mk({ spendText: '매월 1원씩' }).issues.some((i) => i.type === 'data'), '지출 금액 합계 ≠ 금액 감지');
assert.equal(Svc.yearPlan(mk({ period: parsePeriod('확인 필요') }), 2027), null, '기간 오류 → 미집계');

// 물가인상: 3~5% 범위에서 백만원 단위로 딱 떨어지는 금액 (4%에 가장 가까운 것)
assert.equal(Svc.escalate(160e6), 166e6, '160 → 166');
assert.equal(Svc.escalate(200e6), 208e6, '200 → 208 (정확히 4%)');
assert.equal(Svc.escalate(40e6), 42e6, '40 → 42 (범위 안 유일한 값 = 5%)');
assert.equal(Svc.escalate(37e6), 38e6, '37: 3~5% 안에 백만원 단위 없음 → 4% 인상액 반올림');
for (const v of [23e6, 94e6, 115e6, 145e6, 1e9]) {
  const r = Svc.escalate(v) / v - 1;
  assert.ok(Svc.escalate(v) % 1e6 === 0 && r >= 0.03 - 1e-9 && r <= 0.05 + 1e-9, `${v / 1e6}: 3~5% · 백만원 단위`);
}
console.log('✓ 단위 검증 통과');

/* ---------- 2) 예시 Excel 검증 ---------- */
const file = process.argv[2] ?? new URL('../data/current.xlsx', import.meta.url);
const data = parseWorkbook(XLSX.read(readFileSync(file), { type: 'buffer', cellFormula: false }), XLSX.utils);
const D = Dep.prepareDepreciation(data.depreciation.items);
const base = data.service.yearCols.length ? Math.max(...data.service.yearCols) + 1 : new Date().getFullYear();
const S = Svc.prepareContracts(data.service.items, base);
const sum = (arr, f) => arr.reduce((s, x) => s + (f(x) ?? 0), 0);
const won = (v) => `₩${Math.round(v).toLocaleString('ko-KR')}`;

console.log(`\nSheet: ${data.depreciation.sheet} (${D.length}건) / ${data.service.sheet} (${S.length}건) · 연도별 열 ${data.service.yearCols.join(', ')} · Excel 총 금액 ~${base}년`);
const totalAmt = sum(D, (i) => i.amount), totalMonthly = sum(D, (i) => i.monthly);
console.log(`무형자산 총 금액 ${won(totalAmt)} | 월 상각액 합계 ${totalMonthly}`);
const [y0, y1] = Dep.yearSpan(D);
for (let y = y0; y <= y1; y++) console.log(`  ${y}년 상각비 ${won(Dep.totalInYear(D, y)).padStart(16)}`);

const YS = [2026, 2027, 2028];
console.log(`\n전산용역비 (백만원)        ${YS.join('     ')}`);
for (const it of S) {
  const cols = YS.map((y) => { const p = Svc.yearPlan(it, y); return (p ? M(p.amount).toFixed(1) : it.excluded ? '제외' : '⚠').padStart(8); }).join('');
  console.log(`  ${cols}  ${it.name}${it.issues.length ? `  ⚠ ${it.issues.map((i) => i.msg).join(' / ')}` : ''}`);
}
console.log(`  ${YS.map((y) => M(sum(S, (i) => Svc.yearTotal(i, y))).toFixed(1).padStart(8)).join('')}  합계`);

// 2026-10-01 base 갱신: 3JCNS CAS, E&I(760,000,000 · 2026.12~2031.11) 추가
near(totalAmt, 4925396425, '총 무형자산 금액');
near(totalMonthly, 82089940.41666667, '월 감가상각액 합계');
near(Dep.totalInYear(D, 2027), 985079285, '2027년 상각비 = 82,089,940.42 × 12');
near(totalAmt, data.depreciation.excelTotal.amount, 'Excel TOTAL 행과 일치');

// 2026-10-01 지시: 2026년 = 사용자 제공 표('2026년 금액' 열) · 2027년 = Excel 총 금액 · 2028년부터 물가인상
const find = (n) => S.find((i) => i.name.startsWith(n));
assert.equal(base, 2027, "'2026년 금액' 열 → Excel 총 금액은 2027년");
assert.deepEqual(YS.map((y) => M(sum(S, (i) => Svc.yearTotal(i, y)))), [1144.6, 1325, 1447], '전산용역비 합계 2026 · 2027 · 2028');
assert.deepEqual(YS.map((y) => M(Svc.yearTotal(find('PDS'), y))), [121.6, 160, 166], 'PDS: 2026 표 · 2027 Excel · 2028 +3.75%');
assert.deepEqual(Svc.yearPlan(find('PDS'), 2026).pays.map(M), [0, 0, 0, 121.6, 0, 0, 0, 0, 0, 0, 0, 0], 'PDS 2026년 4월 1회');
near(Svc.yearPlan(find('문서관리시스템 유지보수'), 2026).pays[5], 33.5e6, '문서관리 유지보수 2026: 매월 5.5 + 6월 28');
near(Svc.yearPlan(find('Navisworks'), 2026).monthly, 11e6, 'Navisworks 2026: 132 ÷ 12');
assert.deepEqual(
  S.filter((i) => i.excluded).map((i) => [i.name.split(' ')[0], i.excluded]),
  [['MIDAS', '금액 미정']], '비고 기반 집계 제외');
assert.ok(S.filter((i) => i.excluded).every((i) => !i.issues.length && Svc.yearPlan(i, 2027) === null), '집계 제외는 오류가 아니며 합계에서 빠짐');
assert.equal(S.filter((i) => !i.excluded && !Svc.yearPlan(i, 2027)).length, 0, '미집계(오류) 0건');
assert.equal(S.filter((i) => i.issues.length).length, 0, '지출 월 해석 · 합계 오류 0건');

// LNIC Project (사본 비고: 전산용역 23행 '2026년: 첫해 유지보수비 0원' · 2027년 LNIC / 26행 2026년~2027년 리스계약 중 / 감가상각 23행 2026.12~2028.12)
const jcns = find('3JCNS'), bentley = find('벤틀리');
const dj = D.find((i) => i.name.startsWith('3JCNS'));
const lnic = (y) => Dep.projectTotalInYear(D, y) + sum(S, (i) => Svc.projectInYear(i, y));
const yrs = (f) => [2026, 2027, 2028, 2029].map((y) => M(f(y)));
assert.deepEqual(yrs((y) => Svc.yearTotal(jcns, y)), [0, 102, 106, 110], '3JCNS 유지보수: 2026 첫해 0원 · 2027 12월 102 · 2028부터 인상');
assert.deepEqual(yrs((y) => Svc.projectInYear(jcns, y)), [0, 102, 0, 0], '3JCNS 유지보수: 2027년분만 LNIC, 2028년은 판관비');
assert.deepEqual([jcns.zeroYears, jcns.yearNotes[2026], jcns.remarkBase], [[2026], '첫해 유지보수비 0원', 'AVEVA E3D\n2027년 LNIC Project 귀속 (12월 일시납)'], '연도 지정 비고');
assert.deepEqual(yrs((y) => Svc.yearTotal(bentley, y)), [0, 0, 67, 70], '벤틀리: 2026~2027 리스계약 기간 0원, 2028년 64 → 67');
assert.deepEqual(yrs((y) => Dep.projectInYear(dj, y)), [12.7, 152, 152, 0], '3JCNS 상각: 2026.12~2028.12 (25개월) LNIC');
near(lnic(2026), 760e6 / 60, '2026 LNIC 행별 계산 12.7 (그래프는 확정 처리액 405)');
// 상세 표 빨간 테두리(전산용역비): app.js와 같은 기준 — 계약 전 Project 비용 · 연간 계약금액 · Project 비용 변동
const changed = (Y, C) => S.filter((i) => Svc.yearPlan(i, Y)?.pre
  || Math.abs(Svc.yearTotal(i, Y, true) - Svc.yearTotal(i, C, true)) > 0.05e6
  || Math.abs(Svc.projectInYear(i, Y, true) - Svc.projectInYear(i, C, true)) > 0.05e6).map((i) => i.name.split(' ')[0]);
assert.deepEqual([changed(2026, 2027), changed(2027, 2026), changed(2028, 2027)], [['3JCNS'], ['3JCNS'], ['3JCNS', '벤틀리']], '2026·2027은 3JCNS 유지보수만, 2028은 3JCNS · 벤틀리');
console.log(`  LNIC Project 행별 계산: ${[2026, 2027, 2028, 2029].map((y) => `${y} ${M(lnic(y)).toFixed(1)}`).join(' / ')} (백만원)`);
console.log('\n✓ 예시 Excel 검증 통과 (§38)');
