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
const ex = book({ '전산용역비': [['항목명', '금액', '계약기간', '비고'], ['P', 10, '2027.01~2027.12', 'ABC Project 귀속'], ['Q', 10, '', '금액 미정'], ['R', 0, '', '확인 필요']] }).service.items;
assert.deepEqual(ex.map((i) => i.excluded ?? null), [null, '금액 미정', null], '비고 키워드 → 집계 제외 (Project 귀속은 집계 + 구분)');
assert.equal(ex[0].project?.name, 'ABC', 'Project 귀속 → project');
assert.ok(ex[1].issues.length === 0 && ex[2].issues.length > 0, '제외 사유가 없으면 여전히 ⚠');
// 비고의 Project 비용처리 기간
assert.deepEqual(parseProject('2026년~2027년 LNIC Project에서 비용처리'), { name: 'LNIC', start: { y: 2026, m: 1, d: 1 }, end: { y: 2027, m: 12, d: 31 } });
assert.deepEqual(parseProject('AVEVA E3D\n2026.12~2028.11 LNIC Project 귀속'), { name: 'LNIC', start: { y: 2026, m: 12, d: 1 }, end: { y: 2028, m: 11, d: 30 } });
assert.deepEqual(parseProject('2026.12.01~2028.11.30 LNIC Project 귀속').end, { y: 2028, m: 11, d: 30 }, '일 단위 표기 → 월 단위');
assert.deepEqual(parseProject('LNIC PDH-PP Project 귀속'), { name: 'PDH-PP', start: null, end: null }, '기간 없음 → 전체 기간');
assert.equal(parseProject('AVEVA E3D'), null);
assert.deepEqual(parseProject('AVEVA E3D\n2027년 LNIC Project 귀속 (12월 일시납)'), { name: 'LNIC', start: { y: 2027, m: 1, d: 1 }, end: { y: 2027, m: 12, d: 31 } }, '한 해만');
assert.deepEqual(parseYearNotes('2026년: 첫해 유지보수비 0원\nAVEVA E3D'), { base: 'AVEVA E3D', notes: { 2026: '첫해 유지보수비 0원' }, zeroYears: [2026] });
assert.deepEqual(parseYearNotes('2026년: 100,000원 확인').zeroYears, [], "'100,000원'은 0원 아님");
const dp = Dep.prepareDepreciation([{ amount: 600, period: parsePeriod('2026.12~2031.11'), monthlyExcel: null, project: parseProject('2026.12~2028.12 X Project 귀속'), issues: [] }])[0];
assert.deepEqual([2026, 2027, 2028, 2029].map((y) => Dep.projectInYear(dp, y)), [10, 120, 120, 0], '상각 Project 기간분');
const pre = Svc.prepareContracts([{ name: 'M', period: parsePeriod('2028.12.01~2029.11.30'), amount: 120e6, amountOk: true, nextAmount: null,
  spendText: '매년 12월 1회', project: parseProject('2026.12~2028.11 X Project 귀속'), issues: [] }])[0];
assert.deepEqual(Svc.priorPeriods(pre).map((p) => `${Svc.fmtDn(p.s)}~${Svc.fmtDn(p.e)}`), ['2026.12.01~2027.11.30', '2027.12.01~2028.11.30'], '계약 전 Project 기간 → 같은 주기로 선행');
near(Svc.yearView(pre, 2027).pre, 120e6, '2027년 전체가 선행 기간분');
near(Svc.yearView(pre, 2028).total, 120e6, '2028 = 선행 110 + 현재 계약 10');
near(Svc.projectInYear(pre, 2028), 110e6, '2028.01~11만 Project 기간분');
near(Svc.spendInYear(pre, 2026).months[11], 120e6, '선행 기간에도 지출 월(12월) 적용');
// 일시납 · 리스계약 중
const lump = Svc.prepareContracts([{ ...pre, lumpSum: true, cur: null }])[0];
assert.deepEqual([2026, 2027, 2028].map((y) => Svc.yearTotal(lump, y) / 1e6), [120, 120, 120], '일시납: 지급월(12월)에 전액');
near(Svc.yearView(lump, 2027).months[11], 120e6, '일시납: 12월에만');
assert.deepEqual(parseLease('2026년~2027년 LNIC Project에서 205백만원 리스계약 중'), { start: { y: 2026, m: 1, d: 1 }, end: { y: 2027, m: 12, d: 31 } });
assert.equal(parseProject('2026년~2027년 LNIC Project에서 205백만원 리스계약 중'), null, '리스계약은 Project 비용 아님');
const lease = Svc.prepareContracts([{ name: 'L', period: parsePeriod('2026.01.01~2026.12.31'), amount: 120e6, amountOk: true, nextAmount: null, lease: parseLease('2026.07~2027.06 리스계약 중'), issues: [] }])[0];
assert.deepEqual([2026, 2027, 2028].map((y) => Math.round(Svc.yearTotal(lease, y) / 1e6)), [60, 63, 130], '리스 기간 0원: 2026 1~6월분 · 2027 7~12월분만');
const broken = book({ '감가상각액': [['항목명', '비고', '금액']], Sheet2: [['a']] });
assert.match(broken.depreciation.error, /필수 Column.*기간/, '필수 Column 누락 오류');
assert.match(broken.service.error, /Sheet를 찾을 수 없습니다/, '필수 Sheet 누락 오류');
// 지출 월 해석
assert.deepEqual(Svc.parseSpend('매년 4월 1회').lumps, [{ month: 4, amount: null }]);
assert.equal(Svc.parseSpend('매월 6,500,000원씩').monthly, 6500000);
assert.equal(Svc.parseSpend('매월 약 11,000,000원씩').monthly, 'even', "'약' → 계약금액 균등 분할");
assert.deepEqual(Svc.parseSpend('매월 5,500,000원씩 +\n매년 28,000,000원씩').lumps, [{ month: null, amount: 28000000 }]);
assert.ok(Svc.parseSpend('2028년 6월 재계약 예정').unparsed);
assert.deepEqual(Svc.parseSpend('1월 4월 7월 10월\n년 4회').lumps, [1, 4, 7, 10].map((month) => ({ month, amount: null })), '적힌 달마다 균등');
const sp = (name, period, spendText, amount = 120) =>
  Svc.prepareContracts([{ name, period: parsePeriod(period), amount, amountOk: true, nextAmount: null, spendText, issues: [] }])[0];
const apr = sp('A', '2026.04.01~2027.03.31', '매년 4월 1회', 120e6);
assert.deepEqual(Svc.spendInYear(apr, 2027).months.map((a) => a / 1e6), [0, 0, 0, 125, 0, 0, 0, 0, 0, 0, 0, 0], '연 1회 → 4월에만 지출 (갱신 120 → 125)');
assert.deepEqual([2026, 2027, 2028].map((y) => Svc.yearTotal(apr, y, true) / 1e6), [120, 120, 120], '변동 비교용: 물가인상 제외 + 첫해 계약 시작 전 개월을 같은 금액으로 채움');
assert.ok(sp('B', '2026.01.01~2026.12.31', '매월 1원씩').issues.some((i) => i.type === 'data'), '지출 합계 ≠ 계약금액 감지');

// 물가인상: 3~5% 범위에서 백만원 단위로 딱 떨어지는 금액 (4%에 가장 가까운 것)
assert.equal(Svc.escalate(160e6), 166e6, '160 → 166');
assert.equal(Svc.escalate(200e6), 208e6, '200 → 208 (정확히 4%)');
assert.equal(Svc.escalate(40e6), 42e6, '40 → 42 (범위 안 유일한 값 = 5%)');
assert.equal(Svc.escalate(37e6), 38e6, '37: 3~5% 안에 백만원 단위 없음 → 4% 인상액 반올림');
for (const a of [23e6, 94e6, 115e6, 145e6, 1e9]) {
  const r = Svc.escalate(a) / a - 1;
  assert.ok(Svc.escalate(a) % 1e6 === 0 && r >= 0.03 - 1e-9 && r <= 0.05 + 1e-9, `${a / 1e6}: 3~5% · 백만원 단위`);
}
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
console.log(`\n${Y}년 전산용역비 (연장 가정 포함 · 월할 안분)            현재분         연장분   계약 전 Project`);
for (const it of S) {
  const v = Svc.yearView(it, Y);
  const cols = v ? [v.total, v.cur, v.ext, v.pre].map((x) => won(x).padStart(14)).join(' ') : (it.excluded ? `집계 제외(${it.excluded})` : '⚠ 미집계').padStart(14);
  console.log(`  ${cols}  ${it.name}${it.issues.length ? `  ⚠ ${it.issues.map((i) => i.msg).join(' / ')}` : ''}`);
}
const views = S.map((i) => Svc.yearView(i, Y)).filter(Boolean);
console.log(`  ${Y}년 전산용역비 합계 ${won(sum(views, (v) => v.total))} (현재 계약분 ${won(sum(views, (v) => v.cur))} + 연장 가정분 ${won(sum(views, (v) => v.ext))} + 계약 전 Project 기간분 ${won(sum(views, (v) => v.pre))})`);
console.log(`  월별: ${Array.from({ length: 12 }, (_, m) => (sum(views, (v) => v.months[m]) / 1e6).toFixed(1)).join(' / ')} (백만원)`);

// 2026-10-01 base 갱신: 3JCNS CAS, E&I(760,000,000 · 2026.12~2031.11) 추가
near(totalAmt, 4925396425, '총 무형자산 금액');
near(totalMonthly, 82089940.41666667, '월 감가상각액 합계');
near(Dep.totalInYear(D, 2027), 985079285, '2027년 상각비 = 82,089,940.42 × 12');
near(totalAmt, data.depreciation.excelTotal.amount, 'Excel TOTAL 행과 일치');

const find = (n) => S.find((i) => i.name.startsWith(n));
const pds = find('PDS');
const pv = Svc.yearView(pds, 2027);
const [next] = Svc.renewalPeriods(pds, pds.cur.e + 1);
assert.equal(`${Svc.fmtDn(next.s)}~${Svc.fmtDn(next.e)}`, '2027.04.01~2028.03.31', 'PDS 연장 가정 기간');
assert.equal(next.amount, 166e6, 'PDS 갱신 160 → 166 (+3.75%, 백만원 단위)');
near(pv.cur, 40e6, 'PDS 2027 현재 계약분 = 160 × 3/12');
near(pv.ext, 124.5e6, 'PDS 2027 연장 가정분 = 166 × 9/12');
near(sum(pv.months, (a) => a), pv.total, '월별 합 = 연간 합');

// 2026-09-28 사용자 지시 반영분 (data/current.xlsx 비고·기간) + 2026-09-29 물가인상
near(Svc.yearView(find('배관설계관리시스템(S-GEN)'), 2027).total, 23e6 * 6 / 12 + 24e6 * 6 / 12, 'S-GEN: 1~6월 현재 23 + 7~12월 갱신 24');
near(Svc.yearView(find('Agentic AI(유지보수)'), 2027).total, 100e6 * (9 + 29 / 31) / 12 + 104e6 * (2 + 2 / 31) / 12, 'Agentic AI: 갱신 100 → 104');
assert.deepEqual(
  S.filter((i) => i.excluded).map((i) => [i.name.split(' ')[0], i.excluded]),
  [['MIDAS', '금액 미정']], '비고 기반 집계 제외');
assert.ok(S.filter((i) => i.excluded).every((i) => !i.issues.length && Svc.yearView(i, 2027) === null), '집계 제외는 오류가 아니며 합계에서 빠짐');
assert.equal(S.filter((i) => !i.excluded && !Svc.yearView(i, 2027)).length, 0, '미집계(오류) 0건');

// 2026-09-29 '지출 월' 열
const spend = (n) => Svc.spendInYear(find(n), 2027).months;
near(spend('PDS')[3], 166e6, 'PDS: 4월에 갱신 금액 166');
assert.equal(spend('PDS').filter(Boolean).length, 1, 'PDS: 매년 4월 1회 → 나머지 달 지출 없음');
spend('문서관리시스템 운영 관리').forEach((a) => near(a, 81e6 / 12, '운영 관리: 갱신 78 → 81, 매월 6.5 × 81/78'));
near(spend('문서관리시스템 유지보수')[5], (5.5e6 + 28e6) * 98 / 94, '유지보수 6월: (매월 5.5 + 연 1회 28, 시작월 가정) × 갱신 98/94');
near(spend('Navisworks')[0], 138e6 / 12, "Navisworks 1월: '매월 약' → 현재 계약 138 ÷ 12");
near(spend('Navisworks')[11], 144e6 / 12, 'Navisworks 12월: 갱신 144 ÷ 12');

// 2026-10-01 LNIC Project (사본 비고: 전산용역 23행 2026년~2027년 · 12월 일시납 / 26행 2026년~2027년 리스계약 중 / 감가상각 23행 2026.12~2028.12)
const jcns = find('3JCNS'), bentley = find('벤틀리');
const dj = D.find((i) => i.name.startsWith('3JCNS'));
const lnic = (y) => Dep.projectTotalInYear(D, y) + sum(S, (i) => Svc.projectInYear(i, y));
const yrs = (f) => [2026, 2027, 2028, 2029].map((y) => Math.round(f(y) / 1e5) / 10);
assert.deepEqual(yrs((y) => Svc.yearTotal(jcns, y)), [0, 102, 102, 106], '3JCNS 유지보수: 2026 첫해 0원, 매년 12월 일시납 (계약 2028.12~ 전 2027년 12월 포함)');
assert.deepEqual(yrs((y) => Svc.projectInYear(jcns, y)), [0, 102, 0, 0], '3JCNS 유지보수: 2027년 12월분만 LNIC, 2028년은 판관비');
assert.deepEqual(Svc.spendInYear(jcns, 2027).months.map((a) => a / 1e6), [0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 0, 102], '12월 일시납');
assert.deepEqual([jcns.zeroYears, jcns.yearNotes[2026], jcns.remarkBase], [[2026], '첫해 유지보수비 0원', 'AVEVA E3D\n2027년 LNIC Project 귀속 (12월 일시납)'], '연도 지정 비고');
assert.deepEqual(yrs((y) => Svc.yearTotal(bentley, y)), [0, 0, 70, 73], '벤틀리: 2026~2027 리스계약 기간 0원, 2028년부터 64 → 67 → 70');
assert.ok(Svc.spendInYear(bentley, 2027).months.every((a) => a === 0), '벤틀리 리스 기간 지출 없음');
assert.deepEqual(Svc.spendInYear(bentley, 2028).months.map((a) => a / 1e6), [17.5, 0, 0, 17.5, 0, 0, 17.5, 0, 0, 17.5, 0, 0], '벤틀리 1·4·7·10월 70 ÷ 4');
assert.deepEqual(yrs((y) => Dep.projectInYear(dj, y)), [12.7, 152, 152, 0], '3JCNS 상각: 2026.12~2028.12 (25개월) LNIC');
near(lnic(2026), 760e6 / 60, '2026 LNIC 행별 계산 12.7 (그래프는 확정 처리액 405)');
console.log(`  LNIC Project 행별 계산: ${[2026, 2027, 2028, 2029].map((y) => `${y} ${(lnic(y) / 1e6).toFixed(1)}`).join(' / ')} (백만원)`);
console.log('\n✓ 예시 Excel 검증 통과 (§38)');
