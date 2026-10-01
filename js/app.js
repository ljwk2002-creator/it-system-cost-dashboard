/**
 * app.js — 데이터 로드 · 상태 · 화면 렌더링.
 * 화면의 모든 금액·항목·기간은 data/current.xlsx에서 읽어 계산한다 (코드에 데이터 없음 — 예외: Excel에 없는 확정 처리액 PROJECT_FIXED).
 * Excel 원문 문자열은 반드시 esc()를 거쳐 HTML에 들어간다.
 * 화면 구성: 연간 합계 카드 + 연도별 추이 2개 → 연도 선택 시 상세 내역(무형자산 · 전산용역 표)
 */
import { parseWorkbook } from './excelParser.js';
import * as Dep from './depreciationCalculator.js';
import * as Svc from './contractTimeline.js';
import { columnChart } from './charts.js';

const DATA_URL = 'data/current.xlsx'; // 관리자가 교체하는 파일 (README 참조)
const MONTHS = Array.from({ length: 12 }, (_, i) => `${i + 1}월`);
const DEP_STATUS = { active: '상각 중', planned: '상각 예정', done: '상각 완료', check: '확인 필요' };
const ISSUE = { date: '날짜 확인 필요', amount: '금액 확인 필요', period: '기간 확인 필요', data: '데이터 확인 필요' };
// 연도별 Project 확정 처리액: 그해 Project 비용(무형자산+전산용역)은 비고 기준 행별 계산 대신 이 금액으로 표기.
// Excel에 없는 경영 확정값이라 코드에 둔다 (2026-10-01 지시).
const PROJECT_FIXED = { 2026: { amount: 405e6, note: '판관비로 처리될 비용 중 405백만원 LNIC Project로 처리' } };

const $ = (s) => document.querySelector(s);
const esc = (s) => String(s ?? '').replace(/[&<>"']/g, (c) => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' })[c]);
const pad = (n) => String(n).padStart(2, '0');
const sum = (arr, f) => arr.reduce((s, x) => s + (f(x) ?? 0), 0);
const clamp = (v, lo, hi) => Math.min(hi, Math.max(lo, v));
const span = (a, b) => Array.from({ length: b - a + 1 }, (_, i) => a + i);

const state = { data: null, source: null, year: null, years: [], q: '', dep: 'all', svc: 'all', detail: false, checkOpen: false };
let unit = { div: 1e6, label: '백만원' };
const chartArgs = new Map(); // chart element id → columnChart 인자 (폭이 바뀌면 다시 그림)
let lastFocus = null;

/* ---------- 표시 형식: 모든 금액은 단위와 함께 (표시만 반올림, 원본 값은 그대로) ---------- */
const won = (v) => `₩${v.toLocaleString('ko-KR', { maximumFractionDigits: 2 })}`;
const inUnit = (v) => (v / unit.div).toLocaleString('ko-KR', { minimumFractionDigits: 1, maximumFractionDigits: 1 }); // 백만원, 소수점 1자리
const plain = (v) => `${inUnit(v)} ${unit.label}`;
const money = (v) => `<span class="num" data-tip="${esc(won(v))}">${inUnit(v)}<span class="u">${unit.label}</span></span>`;
const bigMoney = (v, cls) => `<p class="${cls}" data-tip="${esc(won(v))}"><span class="cur">₩</span>${inUnit(v)}<span class="unit">${unit.label}</span></p>`;
const fmtDate = (d) => `${d.getFullYear()}.${pad(d.getMonth() + 1)}.${pad(d.getDate())}`;
const fmtDateTime = (d) => `${fmtDate(d)} ${pad(d.getHours())}:${pad(d.getMinutes())}`;
const fmtYM = ({ y, m }) => `${y}.${pad(m)}`;
const validDate = (d) => (d instanceof Date && !Number.isNaN(d.getTime()) ? d : null);
/** " · 전년 대비 ▲x%" — 전년이 화면 표시 범위에 있을 때만 (지난 연도는 데이터가 불완전) */
const yoy = (cur, prev, year) => (prev > 0 && state.years.includes(year - 1)
  ? ` · 전년 대비 <b class="${cur >= prev ? 'up' : 'down'}">${cur >= prev ? '▲' : '▼'} ${Math.abs(((cur - prev) / prev) * 100).toFixed(1)}%</b>` : '');

/* ---------- 로드 ---------- */
async function loadDefault() {
  if (location.protocol === 'file:') {
    throw new Error('index.html을 파일(file://)로 직접 열면 Excel을 읽을 수 없습니다. README의 "실행 방법"대로 웹서버를 통해 열어 주세요.');
  }
  let res;
  try {
    res = await fetch(DATA_URL, { cache: 'no-cache' });
  } catch {
    throw new Error(`Excel 파일을 불러오지 못했습니다 (${DATA_URL}). 네트워크 상태를 확인하세요.`);
  }
  if (!res.ok) throw new Error(`Excel 파일을 찾을 수 없습니다: ${DATA_URL} (HTTP ${res.status})`);
  const lm = res.headers.get('Last-Modified');
  ingest(await res.arrayBuffer(), { name: DATA_URL.split('/').pop(), fileTime: lm ? new Date(lm) : null });
}

function ingest(buf, source) {
  if (!window.XLSX) throw new Error('Excel Parser(js/vendor/xlsx.core.min.js)를 불러오지 못했습니다.');
  let wb;
  try {
    wb = window.XLSX.read(buf, { type: 'array', cellFormula: false, cellHTML: false });
  } catch (e) {
    throw new Error(`Excel Parsing 실패 — 올바른 .xlsx 파일인지 확인하세요. (${e.message})`);
  }
  const data = parseWorkbook(wb, window.XLSX.utils);
  const { depreciation: D, service: S } = data;
  if (D.error && S.error) throw new Error(`${D.error}\n${S.error}`);
  Dep.prepareDepreciation(D.items);
  // Excel 총 금액을 쓰는 마지막 해 = 마지막 'YYYY년 금액' 열의 다음 해 (없으면 올해). 이후는 물가인상
  Svc.prepareContracts(S.items, S.yearCols?.length ? Math.max(...S.yearCols) + 1 : new Date().getFullYear());

  state.data = data;
  state.source = { ...source, saved: validDate(wb.Props?.ModifiedDate) };
  const names = [...new Set([...D.items, ...S.items].map((i) => i.project?.name).filter(Boolean))];
  state.projLabel = `${names.join('·')} Project`.trim();
  const max = Math.max(sum(D.items, (i) => i.amount), sum(S.items, (i) => i.amount));
  unit = max >= 1e10 ? { div: 1e8, label: '억원' } : { div: 1e6, label: '백만원' };
  initYears();
  render();
}

function initYears() {
  const ys = [...state.data.depreciation.items, ...state.data.service.items]
    .filter((i) => i.period.status === 'ok')
    .flatMap((i) => [i.period.start.y, i.period.end.y]);
  const now = new Date().getFullYear();
  const lo = ys.length ? Math.min(...ys) : now, hi = ys.length ? Math.max(...ys) : now;
  const from = clamp(now, lo, hi); // 지난 연도는 표시하지 않음: 올해부터 데이터 마지막 연도까지
  state.years = span(from, hi);
  const params = new URLSearchParams(location.search);
  state.year = clamp(Number(params.get('year')) || state.year || now, from, hi);
  state.detail ||= params.get('detail') === '1';
}

function syncUrl() {
  try {
    const url = new URL(location.href);
    url.searchParams.set('year', state.year);
    if (state.detail) url.searchParams.set('detail', '1'); else url.searchParams.delete('detail');
    history.replaceState(null, '', url);
  } catch { /* URL 공유 기능만 영향 */ }
}

/** 상세 내역 열기. target: 스크롤할 요소 선택자 */
function openDetail(target = '#detail') {
  state.detail = true;
  syncUrl();
  render();
  $(target)?.scrollIntoView({ behavior: 'smooth', block: 'start' });
}

function closeDetail() {
  state.detail = false;
  syncUrl();
  render();
  window.scrollTo({ top: 0, behavior: 'smooth' });
}

const svcYear = (y) => sum(state.data.service.items, (it) => Svc.yearTotal(it, y));
// Project 비용(비고의 비용처리 기간분). 합계는 확정 처리액이 있으면 그 금액
const projDep = (y) => Dep.projectTotalInYear(state.data.depreciation.items, y);
const projSvc = (y) => sum(state.data.service.items, (it) => Svc.projectInYear(it, y));
const projYear = (y) => PROJECT_FIXED[y]?.amount ?? projDep(y) + projSvc(y);
const projDepShown = (y) => PROJECT_FIXED[y]?.amount ?? projDep(y); // 확정 처리액은 무형자산 그래프에도 같은 금액으로 표기 (2026-10-01 지시)
/** 선택 연도가 비고의 리스계약 기간(비용 0원)에 걸리는지 */
const leasedIn = (it, Y) => !!it.lease && (!it.lease.start || (it.lease.start.y <= Y && it.lease.end.y >= Y));
/** 선택 연도 비고: 'YYYY년:' 줄이 있으면 그해에는 그 내용만, 아니면 나머지 비고 */
const remarkIn = (it, Y) => it.yearNotes?.[Y] ?? it.remarkBase ?? it.remark;
const zeroYear = (it, Y) => !!it.zeroYears?.includes(Y); // 비고 'YYYY년: … 0원' → 그해 비용 0원

/* ----- 전산용역비 금액 기준 ----- */
const svcBase = () => state.data.service.items[0]?.baseYear ?? new Date().getFullYear();
/** 선택 연도 금액 기준 (표 · 그래프 공통) */
const svcBasis = (Y) => (state.data.service.yearCols?.includes(Y) ? `Excel '${Y}년 금액' 열`
  : Y <= svcBase() ? 'Excel 총 금액' : `Excel 총 금액 + ${svcBase() + 1}년부터 매년 3~5% 인상`);
const BASIS = { year: (Y) => `${Y}년 금액 열`, excel: () => 'Excel 총 금액', inflation: () => '물가인상 반영', none: () => '계약 시작 전' };
const zeroLabel = (it, Y) => (leasedIn(it, Y) ? '리스계약 기간' : zeroYear(it, Y) ? '비고' : '');
/** 지출 월 + 1회 금액 (그해 금액 ÷ 지출 횟수). 예) "4월 121.6백만원" · "매월 5.5백만원 + 6월 28.0백만원" · "1·4·7·10월 각 16.0백만원" */
function spendCell(p) {
  if (!p?.amount) return '<span class="dim">–</span>';
  const u = `<span class="u">${unit.label}</span>`;
  const parts = p.monthly && p.pays.some((a, i) => a > 0 && !p.lumps.some((l) => l.m === i + 1)) ? [`매월 ${inUnit(p.monthly)}${u}`] : [];
  const lumps = p.lumps.filter((l) => l.amount > 0 && p.pays[l.m - 1] > 0);
  if (lumps.length > 1 && lumps.every((l) => Math.abs(l.amount - lumps[0].amount) < 1)) {
    parts.push(`${lumps.map((l) => l.m).join('·')}월 각 ${inUnit(lumps[0].amount)}${u}`);
  } else lumps.forEach((l) => parts.push(`${l.m}월 ${inUnit(l.amount)}${u}`));
  return `<span class="spend">${parts.join(' + ')}</span>`;
}
const projTag = (v) => (v > 0 ? `<small class="proj-tag" data-tip="${esc(`${state.projLabel} 비용처리 ${won(v)}`)}">${esc(state.projLabel)} ${inUnit(v)}<span class="u">${unit.label}</span></small>` : '');
const projRange = (p) => (p.start ? `${fmtYM(p.start)} ~ ${fmtYM(p.end)}` : '전체 기간');

/* ---------- 비교 연도 대비 변동 (상세 표 빨간 테두리) ---------- */
// 비교 연도: 전년. 단 표시 첫해(올해)는 전년이 없으므로 다음 해와 비교 → 2026 ↔ 2027 양쪽 모두 표시
const cmpYear = (Y) => (state.years.includes(Y - 1) || !state.years.includes(Y + 1) ? Y - 1 : Y + 1);
// 표시 단위 반올림 · 일 단위 계약의 연간 1% 미만 흔들림은 변동으로 보지 않음
const differs = (a, b) => Math.abs(a - b) > Math.max(0.05 * unit.div, Math.max(Math.abs(a), Math.abs(b)) * 0.01);
const chgNote = (C, v, proj, leased) => `<small class="chg-note" data-tip="${esc(`${C}년 대비 변동 (전산용역비는 연간 계약금액 기준 · 물가인상분과 연도별 실적 금액 차이 제외)`)}">${C}년 ${plain(v)}${proj > 0 ? ` · ${esc(state.projLabel)}` : leased ? ' · 리스계약' : ''}</small>`;
const chgLegend = (C, n) => `<br><span class="chg-legend"><i></i>빨간 테두리: ${C}년 대비 변동 ${n}건</span>`;

/* ---------- 렌더링 ---------- */
function render() {
  const { depreciation: D, service: S } = state.data;
  const Y = state.year;
  const C = cmpYear(Y);
  const dep = D.items.map((it) => ({ it, st: Dep.statusInYear(it, Y), months: Dep.monthsInYear(it, Y), amount: Dep.amountInYear(it, Y), proj: Dep.projectInYear(it, Y),
    chg: differs(Dep.amountInYear(it, Y), Dep.amountInYear(it, C)) || differs(Dep.projectInYear(it, Y), Dep.projectInYear(it, C)) }));
  const svc = S.items.map((it) => {
    const p = Svc.yearPlan(it, Y);
    // Excel 계약 시작 전 Project 기간 비용 = 계약에 없던 신규 비용 → 그해 변동으로 표시 (예: 3JCNS 유지보수 2027)
    const pre = !!p?.pre;
    return { it, p, proj: Svc.projectInYear(it, Y), pre,
      chg: pre || differs(Svc.yearTotal(it, Y, true), Svc.yearTotal(it, C, true)) || differs(Svc.projectInYear(it, Y, true), Svc.projectInYear(it, C, true)) };
  });
  const ctx = { Y, C, D, S, dep, svc, depTotal: sum(dep, (r) => r.amount), svcTotal: sum(svc, (r) => r.p?.amount), check: checkSummary(D, S) };
  renderHeader(ctx);
  renderHero(ctx);
  renderAlert(ctx);
  renderCharts(ctx);
  $('#detail').hidden = !state.detail;
  if (state.detail) {
    $('#detailTitle').textContent = `${Y}년 상세 내역`;
    renderDep(ctx);
    renderSvc(ctx);
    renderCheck(ctx);
  }
}

function renderHeader({ Y, D, S }) {
  $('#printYear').textContent = `· ${Y}년`;

  const { saved, fileTime, name, preview } = state.source;
  const shown = saved ?? fileTime;
  $('#updated').textContent = shown ? fmtDateTime(shown) : '확인 불가';
  $('#updated').dataset.tip = [
    saved && `Excel 저장 시각: ${fmtDateTime(saved)}`,
    fileTime && `${preview ? '파일 수정' : '서버 반영'} 시각: ${fmtDateTime(fileTime)}`,
  ].filter(Boolean).join('\n') || '파일 시각 정보 없음';
  $('#meta').textContent = `조회 기준일 ${fmtDate(new Date())} · 원본 ${name} (${[D.sheet, S.sheet].filter(Boolean).join(' / ')} Sheet)`;
}

function renderHero({ Y, D, S, depTotal, svcTotal }) {
  const total = depTotal + svcTotal;
  const prev = (D.error ? 0 : Dep.totalInYear(D.items, Y - 1)) + (S.error ? 0 : svcYear(Y - 1));
  const pct = (v) => (total ? (v / total) * 100 : 0);
  const proj = projYear(Y), fixed = PROJECT_FIXED[Y];
  const part = (cls, label, v) => `<button type="button" class="split__item" data-open-detail="#${cls}Sub">
      <span class="split__top"><i class="dot dot--${cls}"></i><span class="split__name">${label}</span><span class="split__pct">${pct(v).toFixed(0)}%</span></span>
      <b data-tip="${esc(won(v))}">${inUnit(v)}<span class="u">${unit.label}</span></b></button>`;
  const missing = [D.error && '무형자산', S.error && '전산용역비'].filter(Boolean);
  $('#hero').innerHTML = `
    <div class="hero__main">
      <p class="hero__label">${Y}년 IT System 비용 합계 <span>무형자산상각비 + 전산용역비</span></p>
      ${bigMoney(total, 'hero__value')}
      <p class="hero__sub">월평균 ${plain(total / 12)}${yoy(total, prev, Y)}
        ${missing.length ? ` · <span class="warn">⚠ ${missing.join('·')} 데이터 오류로 제외</span>` : ''}</p>
      ${proj > 0 ? `<p class="hero__proj"><i class="dot dot--proj"></i>그중 ${esc(state.projLabel)} 비용
        <b data-tip="${esc(won(proj))}">${inUnit(proj)}<span class="u">${unit.label}</span></b>${fixed ? `<small>비고: ${esc(fixed.note)}</small>` : ''}</p>` : ''}
      <button type="button" class="btn btn--primary no-print" data-open-detail="#detail">${Y}년 상세 내역 보기 →</button>
    </div>
    <div class="hero__split">
      <div class="split" role="img" aria-label="무형자산상각비 ${pct(depTotal).toFixed(0)}%, 전산용역비 ${pct(svcTotal).toFixed(0)}%">
        <i class="split__seg split__seg--dep" style="width:${pct(depTotal).toFixed(2)}%"></i><i class="split__seg split__seg--svc" style="width:${pct(svcTotal).toFixed(2)}%"></i>
      </div>
      <div class="split__legend">${part('dep', '무형자산상각비', depTotal)}${part('svc', '전산용역비', svcTotal)}</div>
    </div>`;
}

/* ----- DATA CHECK ----- */
function totalCheck(sec, field, label, get) {
  const excel = sec.excelTotal?.[field];
  if (sec.error || excel == null) return null;
  const parsed = sum(sec.items, get);
  const bad = Math.abs(parsed - excel) >= 1;
  return {
    bad,
    msg: bad
      ? `${sec.sheet} TOTAL ${label} 불일치 — 원본 TOTAL ${won(excel)} / 항목 합계 ${won(parsed)}`
      : `${sec.sheet} TOTAL ${label} 일치 (${won(excel)})`,
  };
}

function checkSummary(D, S) {
  const items = [...D.items, ...S.items];
  const flagged = items.filter((i) => i.issues.length);
  const excluded = S.items.filter((i) => i.excluded);
  const totals = [
    totalCheck(D, 'amount', '총 금액', (i) => i.amount),
    totalCheck(D, 'monthly', '월 상각액', (i) => i.monthlyExcel),
    totalCheck(S, 'amount', '총 금액', (i) => i.amount),
  ].filter(Boolean);
  const notes = [
    ...[D, S].filter((x) => x.error).map((x) => ({ bad: true, msg: x.error })),
    ...totals,
    ...[...D.warnings, ...S.warnings].map((msg) => ({ bad: true, msg })),
    ...excluded.map((it) => ({ info: true, msg: `${it.sheet} ${it.rowNo}행 · ${it.name} — 집계 제외 (${it.excluded}) · 비고: ${it.remark}` })),
  ];
  return {
    items, flagged, excluded, notes,
    normal: items.filter((i) => !i.issues.length && !i.excluded).length,
    counts: Object.entries(ISSUE).map(([t, label]) => [label, items.filter((i) => i.issues.some((x) => x.type === t)).length]).filter(([, n]) => n),
    badTotals: totals.filter((t) => t.bad).length,
    hasTotals: totals.length > 0,
    bad: flagged.length + notes.filter((n) => n.bad).length,
  };
}

function renderAlert({ check }) {
  $('#alert').innerHTML = check.bad
    ? `<button type="button" class="alert" data-open-check>⚠ 데이터 확인 필요 ${check.bad}건 — 상세 내역의 DATA CHECK에서 확인하세요</button>` : '';
}

function renderCheck({ check: c }) {
  $('#check').innerHTML = `<details class="check${c.bad ? ' check--warn' : ''}"${state.checkOpen ? ' open' : ''}>
    <summary>
      <span class="check__title">DATA CHECK</span>
      <span class="pill pill--ok">✓ 정상 항목 ${c.normal}</span>
      ${c.counts.map(([l, n]) => `<span class="pill pill--warn">⚠ ${l} ${n}</span>`).join('')}
      ${c.excluded.length ? `<span class="pill pill--info">ⓘ 집계 제외 ${c.excluded.length}</span>` : ''}
      ${c.hasTotals ? `<span class="pill ${c.badTotals ? 'pill--warn' : 'pill--ok'}">${c.badTotals ? `⚠ 원본 TOTAL 불일치 ${c.badTotals}` : '✓ 원본 TOTAL 일치'}</span>` : ''}
      <span class="check__more no-print">상세</span>
    </summary>
    <div class="check__body">
      ${c.flagged.length ? `<ul class="check__list">${c.flagged.map((it) => `<li>
        <button type="button" data-goto="${esc(it.id)}">${esc(it.sheet)} ${it.rowNo}행 · ${esc(it.name)}</button>
        <span>${it.issues.map((x) => `<em>${ISSUE[x.type]}</em> ${esc(x.msg)}`).join('<br>')}</span></li>`).join('')}</ul>`
      : '<p>확인이 필요한 항목이 없습니다.</p>'}
      <ul class="check__notes">${c.notes.map((n) => `<li class="${n.bad ? 'is-bad' : n.info ? 'is-info' : ''}">${n.bad ? '⚠' : n.info ? 'ⓘ' : '✓'} ${esc(n.msg)}</li>`).join('')}</ul>
    </div>
  </details>`;
  $('#check details').addEventListener('toggle', (e) => { state.checkOpen = e.target.open; });
}

/* ----- 연도별 추이 ----- */
function setChart(id, args) {
  chartArgs.set(id, args);
  drawChart(document.getElementById(id));
}
function drawChart(el) {
  const args = chartArgs.get(el.id);
  if (args) el.innerHTML = columnChart({ ...args, width: Math.max(260, el.clientWidth) });
}
function renderCharts({ Y, D, S }) {
  const P = state.projLabel;
  const common = { selected: Y, unit: unit.label, label: (d) => inUnit(d.value), projLabel: (d) => inUnit(d.proj) };
  const depOf = (y) => (D.error ? 0 : Dep.totalInYear(D.items, y));
  const svcOf = (y) => (S.error ? 0 : svcYear(y));
  const projTip = (d) => (d.proj > 0 ? `\n▨ 그중 ${P} 비용 ${won(d.proj)}${PROJECT_FIXED[d.key] ? ' (확정 처리액)' : ''}` : '');
  setChart('totalYearChart', {
    ...common,
    data: state.years.map((y) => {
      const parts = [{ cls: 'seg--dep', value: depOf(y) }, { cls: 'seg--svc', value: svcOf(y) }];
      return { key: y, name: String(y), value: parts[0].value + parts[1].value, parts, proj: projYear(y) };
    }),
    tip: (d) => `${d.key}년 IT System 비용 합계 ${won(d.value)}\n무형자산상각비 ${won(d.parts[0].value)}\n전산용역비 ${won(d.parts[1].value)} (${svcBasis(d.key)})${projTip(d)}\n선택하면 상세 내역`,
  });
  if (D.error) {
    chartArgs.delete('depYearChart');
    $('#depYearChart').innerHTML = sectionError(D.error);
  } else {
    setChart('depYearChart', {
      ...common,
      data: state.years.map((y) => ({ key: y, name: String(y), value: depOf(y), proj: projDepShown(y) })),
      tip: (d) => `${d.key}년 무형자산상각비 ${won(d.value)}${projTip(d)}\n선택하면 상세 내역`,
    });
  }

  // 우측 상단: Project 비용 합계 (표시 연도 기준) + 확정 처리액 비고
  const pYears = state.years.filter((y) => projYear(y) > 0);
  const pTotal = sum(pYears, projYear);
  const range = pYears.length > 1 ? `${pYears[0]}~${pYears.at(-1)}년` : `${pYears[0]}년`;
  $('#projSum').innerHTML = pYears.length ? `<p class="proj-sum__head"><i class="dot dot--proj"></i>${esc(P)} 비용 합계 <small>${range}</small></p>
    <p class="proj-sum__value" data-tip="${esc(pYears.map((y) => `${y}년 ${won(projYear(y))}`).join('\n'))}">${inUnit(pTotal)}<span class="u">${unit.label}</span></p>` : '';
  $('#projNote').innerHTML = pYears.filter((y) => PROJECT_FIXED[y])
    .map((y) => `▨ ${y}년 ${esc(P)} 비용 ${plain(PROJECT_FIXED[y].amount)} — 비고: ${esc(PROJECT_FIXED[y].note)}`).join('<br>');
  document.querySelectorAll('.legend__proj').forEach((el) => {
    el.hidden = !(el.closest('.chart-card--dep') ? state.years.some((y) => projDepShown(y) > 0) : pYears.length);
    el.lastChild.textContent = P;
  });
}

/* ----- 공통 셀 ----- */
const matches = (it) => !state.q || it.name.toLowerCase().includes(state.q.toLowerCase());
const chips = (group, defs, count) => defs.map(([k, l]) =>
  `<button type="button" class="chip" data-group="${group}" data-filter="${k}" aria-pressed="${state[group] === k}">${l} <b>${count(k)}</b></button>`).join('');
const sectionError = (msg) => `<div class="section-error" role="alert">⚠ ${esc(msg)}<small>원본 Excel을 확인해 주세요.</small></div>`;
const emptyRow = (n) => `<tr class="empty"><td colspan="${n}">조건에 맞는 항목이 없습니다.</td></tr>`;
const issueIcon = (it) => (it.issues.length
  ? ` <span class="warn" data-tip="${esc(it.issues.map((x) => `${ISSUE[x.type]}: ${x.msg}`).join('\n'))}" aria-label="확인 필요">⚠</span>` : '');
const amountWarn = (it) => `<span class="warn" data-tip="${esc(`원본 값: ${it.amountText || '(공란)'}\n${it.issues.filter((x) => x.type === 'amount').map((x) => x.msg).join('\n')}`)}">⚠ 금액 확인 필요</span>`;
const share = (v, total) => (v > 0 && total > 0
  ? `<div class="share"><span class="share__bar"><i style="width:${((v / total) * 100).toFixed(2)}%"></i></span><span class="share__pct">${((v / total) * 100).toFixed(1)}%</span></div>`
  : '<span class="dim">–</span>');

function periodTip(it) {
  const p = it.period;
  if (p.status === 'invalid') return `${p.reason} — 자동 보정하지 않았습니다.\n원본: ${p.text}`;
  return `${it.kind === 'service' ? '계약기간' : '상각기간'}을 날짜로 변환할 수 없습니다.\n원본 Excel을 확인하십시오.\n원본: ${p.text || '(공란)'}`;
}
const periodLabel = (it) => (it.period.status === 'invalid' ? '⚠ 날짜 확인 필요' : `⚠ ${it.kind === 'service' ? '계약기간' : '상각기간'} 확인 필요`);
const periodWarn = (it) => `<span class="warn" data-tip="${esc(periodTip(it))}">${periodLabel(it)}</span><small class="raw">${esc(it.period.text || '(공란)')}</small>`;

/** 상세 표 제목 아래: 그중 Project 비용 (행별 계산. 확정 처리액이 있는 연도는 그 사실을 함께 표시) */
const projNote = (v, Y) => (v > 0
  ? `<br><span class="proj-note"><i class="dot dot--proj"></i>그중 ${esc(state.projLabel)} 비용처리 ${plain(v)}${PROJECT_FIXED[Y] ? ` (행별 계산 · 그래프는 ${Y}년 확정 처리액 ${plain(PROJECT_FIXED[Y].amount)})` : ''}</span>` : '');

function subHead({ eyebrow, title, total, note, filter }) {
  return `<div class="sub__title"><p class="eyebrow">${eyebrow}</p><h3>${title}</h3></div>
    ${bigMoney(total, 'sub__total')}
    <p class="sub__note">${note}</p>
    <div class="chips no-print">${filter}</div>`;
}

/* ----- 무형자산 상세 ----- */
function renderDep({ Y, C, D, dep, depTotal }) {
  if (D.error) {
    $('#depHead').innerHTML = `<div class="sub__title"><h3>무형자산상각비</h3></div>${sectionError(D.error)}`;
    $('#depTable').innerHTML = '';
    return;
  }
  const is = {
    all: () => true, active: (r) => r.st === 'active', planned: (r) => r.st === 'planned',
    done: (r) => r.st === 'done', check: (r) => r.it.issues.length > 0, chg: (r) => r.chg,
  };
  const searched = dep.filter((r) => matches(r.it));
  $('#depHead').innerHTML = subHead({
    eyebrow: 'INTANGIBLE ASSETS · 무형자산 상세', title: `${Y}년 무형자산상각비`, total: depTotal,
    note: `상각 중 ${dep.filter((r) => r.st === 'active').length}건 · 월평균 ${plain(depTotal / 12)}${yoy(depTotal, Dep.totalInYear(D.items, Y - 1), Y)}
      · 총 취득금액 ${plain(sum(D.items, (i) => i.amount))}${projNote(sum(dep, (r) => r.proj), Y)}${chgLegend(C, dep.filter((r) => r.chg).length)}`,
    filter: chips('dep', [['all', '전체'], ['chg', `${C}년 대비 변동`], ['active', '상각 중'], ['planned', '상각 예정'], ['done', '상각 완료'], ['check', '⚠ 확인 필요']],
      (k) => searched.filter(is[k]).length),
  });
  const rows = searched.filter(is[state.dep]).sort((a, b) => b.amount - a.amount || (b.it.amount ?? 0) - (a.it.amount ?? 0));
  $('#depTable').innerHTML = `
    <thead><tr>
      <th scope="col">항목</th><th scope="col" class="r">총 취득금액</th><th scope="col">상각기간</th><th scope="col" class="r">월 상각액</th>
      <th scope="col" class="r em">${Y}년 상각액</th><th scope="col" class="share-col">비중</th><th scope="col">상태</th><th scope="col">비고</th>
    </tr></thead>
    <tbody>${rows.map((r) => depRow(r, depTotal, C)).join('') || emptyRow(8)}</tbody>
    <tfoot><tr>
      <th scope="row">합계 ${rows.length}건</th>
      <td class="r">${money(sum(rows, (r) => r.it.amount))}</td><td></td>
      <td class="r">${money(sum(rows, (r) => r.it.monthly))}</td>
      <td class="r em">${money(sum(rows, (r) => r.amount))}</td>
      <td>${share(sum(rows, (r) => r.amount), depTotal)}</td><td colspan="2"></td>
    </tr></tfoot>`;
}

function depRow({ it, st, months, amount, proj, chg }, total, C) {
  const p = it.period;
  const mismatch = it.issues.find((x) => x.type === 'data' && x.msg.startsWith('월 상각액'));
  return `<tr data-id="${esc(it.id)}" tabindex="0"${chg ? ' class="is-chg"' : ''}>
    <th scope="row" class="name">${esc(it.name)}${issueIcon(it)}</th>
    <td class="r">${it.amount != null ? money(it.amount) : amountWarn(it)}</td>
    <td class="nowrap">${p.status === 'ok' ? `${fmtYM(p.start)} ~ ${fmtYM(p.end)}<small>${it.months}개월</small>` : periodWarn(it)}</td>
    <td class="r">${it.monthly != null ? money(it.monthly) : '<span class="dim">–</span>'}${mismatch ? ` <span class="warn" data-tip="${esc(mismatch.msg)}">⚠</span>` : ''}</td>
    <td class="r em">${months ? `${money(amount)}<small>${months}개월</small>${projTag(proj)}` : '<span class="dim">–</span>'}${chg ? chgNote(C, Dep.amountInYear(it, C), Dep.projectInYear(it, C)) : ''}</td>
    <td>${share(amount, total)}</td>
    <td><span class="st st--${st}">${DEP_STATUS[st]}</span></td>
    <td class="remark">${esc(remarkIn(it, state.year))}</td>
  </tr>`;
}

/* ----- 전산용역 계약 현황 ----- */
function renderSvc({ Y, C, S, svc, svcTotal }) {
  if (S.error) {
    $('#svcHead').innerHTML = `<div class="sub__title"><h3>전산용역비</h3></div>${sectionError(S.error)}`;
    $('#svcTable').innerHTML = '';
    return;
  }
  const counted = svc.filter((r) => r.p);
  const excluded = svc.filter((r) => r.it.excluded).length;
  const failed = svc.length - counted.length - excluded;
  const is = { all: () => true, chg: (r) => r.chg, check: (r) => r.it.issues.length > 0, excluded: (r) => !!r.it.excluded };
  const searched = svc.filter((r) => matches(r.it));
  $('#svcHead').innerHTML = subHead({
    eyebrow: 'IT SERVICE CONTRACTS · 전산용역 계약 현황', title: `${Y}년 전산용역비`, total: svcTotal,
    note: `집계 ${counted.filter((r) => r.p.amount > 0).length}건 · 집계 제외 ${excluded}건${failed ? ` · <span class="warn">확인 필요 ${failed}건</span>` : ''}
${yoy(svcTotal, svcYear(Y - 1), Y)}<br>${Y}년 금액 기준: ${esc(svcBasis(Y))}${projNote(sum(svc, (r) => r.proj), Y)}${chgLegend(C, svc.filter((r) => r.chg).length)} <small>(연간 계약금액 기준 · 물가인상분과 연도별 실적 금액 차이 제외)</small>`,
    filter: chips('svc', [['all', '전체'], ['chg', `${C}년 대비 변동`], ['check', '⚠ 확인 필요'], ['excluded', '집계 제외']], (k) => searched.filter(is[k]).length),
  });
  const rank = (r) => (r.p ? r.p.amount : r.it.excluded ? -2 : -1); // 집계 → 확인 필요 → 집계 제외
  const rows = searched.filter(is[state.svc]).sort((a, b) => rank(b) - rank(a) || (b.it.amount ?? 0) - (a.it.amount ?? 0));
  $('#svcTable').innerHTML = `
    <thead><tr>
      <th scope="col">System</th><th scope="col">지출 월 <small>· 1회 금액</small></th><th scope="col" class="r em">${Y}년 금액</th>
      <th scope="col" class="share-col">비중</th><th scope="col">비고</th>
    </tr></thead>
    <tbody>${rows.map((r) => svcRow(r, Y, svcTotal, C)).join('') || emptyRow(5)}</tbody>
    <tfoot><tr>
      <th scope="row">합계 ${rows.length}건</th><td></td>
      <td class="r em">${money(sum(rows, (r) => r.p?.amount))}</td>
      <td>${share(sum(rows, (r) => r.p?.amount), svcTotal)}</td>
      <td class="unit-note">집계 제외 · 확인 필요 항목은 합계에서 제외</td>
    </tr></tfoot>`;
}

function svcRow({ it, p, proj, pre, chg }, Y, total, C) {
  const why = !it.cur ? periodTip(it) : it.issues.filter((x) => x.type === 'amount').map((x) => x.msg).join('\n');
  const zero = p && !p.amount && zeroLabel(it, Y);
  const yearCell = it.excluded ? `<span class="tag" data-tip="${esc(`집계 제외 — 비고: ${it.remark}`)}">${esc(it.excluded)}</span>`
    : !p ? `<span class="warn" data-tip="${esc(`집계 제외\n${why}`)}">⚠ 미집계</span>`
      : p.amount || zero ? `<span class="num" data-tip="${esc(`${Y}년 ${won(p.amount)} · ${BASIS[p.basis](Y)}`)}">${inUnit(p.amount)}<span class="u">${unit.label}</span></span>${projTag(proj)}`
        + (zero ? `<small class="proj-tag" data-tip="${esc(`비고: ${remarkIn(it, Y)}`)}">${zero} · 0원</small>` : '')
        : '<span class="dim">–</span>';
  const spend = it.excluded || !p ? `<small class="dim">${esc((it.yearSpend?.[Y] ?? it.spendText).replace(/\s+/g, ' '))}</small>` : spendCell(p);
  return `<tr data-id="${esc(it.id)}" tabindex="0"${chg ? ' class="is-chg"' : ''}>
    <th scope="row" class="name">${esc(it.name)}${issueIcon(it)}</th>
    <td class="spend-col">${spend}</td>
    <td class="r em">${yearCell}${pre ? `<small class="chg-note" data-tip="${esc(`Excel 계약(${Svc.fmtDn(it.cur.s)}~) 시작 전 비고의 Project 비용처리 기간에 새로 발생하는 비용`)}">신규 · 계약 전 ${esc(state.projLabel)} 비용</small>`
      : chg ? chgNote(C, Svc.yearTotal(it, C), Svc.projectInYear(it, C), leasedIn(it, C)) : ''}</td>
    <td>${share(p?.amount ?? 0, total)}</td>
    <td class="remark">${esc(remarkIn(it, Y))}</td>
  </tr>`;
}

/* ---------- Drill-down (행 클릭) ---------- */
/* 팝업: 핵심 금액(KPI) → 요약 타일 → 막대 그래프 → 계약·산정 정보 → 원본 Excel 값(접힘) */
const kvList = (rows) => `<dl class="kv">${rows.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join('')}</dl>`;
const dwHead = (it, kind, tags) => `<p class="eyebrow">${kind} · ${state.year}년</p><h2 id="drawerTitle">${esc(it.name)}</h2>
  ${tags.length ? `<div class="dw-tags">${tags.join('')}</div>` : ''}
  ${it.issues.length ? `<ul class="issues">${it.issues.map((x) => `<li><b>⚠ ${ISSUE[x.type]}</b> ${esc(x.msg)}</li>`).join('')}</ul>` : ''}`;
const dwKpi = (label, v, sub) => `<div class="dw-kpi"><p class="dw-kpi__label">${label}</p>${bigMoney(v, 'dw-kpi__value')}${sub ? `<p class="dw-kpi__sub">${sub}</p>` : ''}</div>`;
const dwKpiNa = (label, text, sub) => `<div class="dw-kpi dw-kpi--na"><p class="dw-kpi__label">${label}</p><p class="dw-kpi__na">${text}</p>${sub ? `<p class="dw-kpi__sub">${sub}</p>` : ''}</div>`;
const dwTiles = (rows) => `<div class="dw-tiles">${rows.map(([k, v]) => `<div class="dw-tile"><span>${k}</span><b>${v}</b></div>`).join('')}</div>`;
const dwSec = (title, note, body) => `<section class="dw-sec"><h3>${title}${note ? ` <small>${note}</small>` : ''}</h3>${body}</section>`;
const dwNotes = (notes) => notes.filter(Boolean).map((n) => `<p class="dw-note">${n}</p>`).join('');
const dwRaw = (it) => `<details class="dw-raw"><summary>원본 Excel 값 (${esc(it.sheet)} ${it.rowNo}행)</summary><table class="mini src"><tbody>${it.raw.map((r) =>
  `<tr><th scope="row">${esc(r.label)}</th><td class="addr">${esc(r.addr)}</td><td>${esc(r.text) || '<span class="dim">(공란)</span>'}</td></tr>`).join('')}</tbody></table></details>`;
const moneyText = (v) => `${inUnit(v)} <small>${unit.label}</small>`;
const dwProjTag = (it) => (it.project ? [`<span class="tag tag--proj">${esc(it.project.name)} Project 비용처리 · ${projRange(it.project)}</span>`] : []);
const dwProjSub = (it, v) => (v > 0 ? ` · 그중 ${esc(it.project.name)} Project ${plain(v)}` : '');

/** 작은 막대 그래프. 값이 0인 칸은 금액·막대를 표시하지 않음 */
function barStrip(items, tone, selected) {
  const max = Math.max(...items.map((d) => d.value), 1);
  return `<div class="mbars mbars--${tone}">${items.map((d) => {
    const has = d.value > 0;
    return `<div class="mbar${has ? ' has' : ''}${d.key === selected ? ' is-sel' : ''}"${has ? ` data-tip="${esc(d.tip)}"` : ''}>
      <span class="mbar__v">${has ? `${inUnit(d.value)}<small>${unit.label}</small>` : ''}</span>
      <span class="mbar__track">${has ? `<i style="height:${Math.max(4, (d.value / max) * 100).toFixed(1)}%"></i>` : ''}</span>
      <span class="mbar__m">${d.name}</span></div>`;
  }).join('')}</div>`;
}

function depDetail(it) {
  const Y = state.year, p = it.period, ok = it.monthly != null;
  const st = Dep.statusInYear(it, Y);
  const done = ok ? Dep.elapsedAtYearEnd(it, Y) : 0;
  const pct = ok ? ((done / it.months) * 100).toFixed(1) : 0;
  const monthlyBasis = it.monthlyExcel != null
    ? `Excel 입력값${it.monthlyCalc != null ? ` (계산값 ${plain(it.monthlyCalc)})` : ''}` : '총 취득금액 ÷ 총 상각개월';
  return dwHead(it, '무형자산', [`<span class="st st--${st}">${DEP_STATUS[st]}</span>`, ...dwProjTag(it)])
    + (ok ? dwKpi(`${Y}년 상각액`, Dep.amountInYear(it, Y), `월 ${plain(it.monthly)} × ${Dep.monthsInYear(it, Y)}개월${dwProjSub(it, Dep.projectInYear(it, Y))}`)
      : dwKpiNa(`${Y}년 상각액`, '⚠ 계산 불가', '금액 또는 상각기간 확인 필요'))
    + dwTiles([
      ['총 취득금액', it.amount != null ? moneyText(it.amount) : amountWarn(it)],
      ['월 상각액', ok ? moneyText(it.monthly) : '–'],
      ['상각기간', p.status === 'ok' ? `${fmtYM(p.start)} ~ ${fmtYM(p.end)}` : periodWarn(it)],
    ])
    + (ok ? `<div class="dw-progress"><div class="dw-progress__top"><span>상각 진행 (${Y}년 말 기준)</span><b>${done} / ${it.months}개월 · ${pct}%</b></div>
        <div class="dw-progress__bar"><i style="width:${pct}%"></i></div></div>` : '')
    + (ok ? dwSec('연도별 상각액', `총 ${plain(it.monthly * it.months)}`, barStrip(span(p.start.y, p.end.y).map((y) => ({
      key: y, name: String(y), value: Dep.amountInYear(it, y), tip: `${y}년 ${won(Dep.amountInYear(it, y))} · ${Dep.monthsInYear(it, y)}개월`,
    })), 'dep', Y)) : '')
    + dwSec('산정 정보', '', kvList([
      ['월 상각액 기준', monthlyBasis],
      ...(it.project ? [['Project 비용처리', `${esc(it.project.name)} Project · ${projRange(it.project)}<small>상각액은 그대로 집계하고 이 기간분을 Project 비용으로 구분</small>`]] : []),
      ['비고', esc(remarkIn(it, Y)) || '–'],
    ]))
    + dwRaw(it);
}

function svcDetail(it) {
  const Y = state.year, p = Svc.yearPlan(it, Y), next = Svc.yearPlan(it, Y + 1);
  const leaseRange = it.lease && projRange(it.lease);
  const zero = p && !p.amount && zeroLabel(it, Y);
  const tags = dwProjTag(it);
  if (it.excluded) tags.push(`<span class="tag">집계 제외 · ${esc(it.excluded)}</span>`);
  if (it.lease) tags.push(`<span class="tag tag--proj">리스계약 기간 · ${leaseRange} · 비용 0원</span>`);
  if (p) tags.push(`<span class="tag tag--line">${Y}년 금액 기준: ${esc(BASIS[p.basis](Y))}</span>`);

  const kpi = it.excluded ? dwKpiNa(`${Y}년 금액`, `집계 제외 — ${esc(it.excluded)}`, esc(it.remark))
    : p ? dwKpi(`${Y}년 금액`, p.amount, zero ? `${leasedIn(it, Y) ? `리스계약 기간(${leaseRange}) — ` : ''}비고: ${esc(remarkIn(it, Y))}`
      : `${esc(BASIS[p.basis](Y))}${dwProjSub(it, Svc.projectInYear(it, Y))}`)
      : dwKpiNa(`${Y}년 금액`, '⚠ 미집계', '금액 또는 계약기간 확인 필요');

  let body = '';
  if (p) {
    const paid = p.pays.filter((a) => a > 0).length;
    const rule = Svc.parseSpend(p.spendText);
    const rate = next?.amount && p.amount ? ((next.amount / p.amount - 1) * 100).toFixed(1) : null;
    body = dwTiles([
      ['지출 월', esc(p.spendText.replace(/\s+/g, ' ')) || '<span class="dim">미기재</span>'],
      ['지출 1회 금액', spendCell(p)],
      [`${Y + 1}년 금액`, next ? `${moneyText(next.amount)}${rate != null ? `<small> · ${rate >= 0 ? '+' : ''}${rate}%</small>` : ''}` : '–'],
    ])
      + dwSec(`${Y}년 월별 지출`, '지출이 없는 달은 비워 둠', barStrip(p.pays.map((a, i) => ({
        key: i + 1, name: MONTHS[i], value: a, tip: `${Y}.${pad(i + 1)} 지출 ${won(a)}`,
      })), 'svc'))
      + `<p class="dw-sum">지출 합계 <b>${moneyText(p.amount)}</b> · ${paid}회 지출</p>`
      + dwNotes([
        zero && `비고에 따라 ${Y}년 비용은 0원입니다${leasedIn(it, Y) ? ` (리스계약 기간 ${leaseRange})` : ''}.`,
        p.pre && `Excel 계약(${Svc.fmtDn(it.cur.s)}~) 시작 전이지만 비고의 ${esc(it.project.name)} Project 비용처리 기간(${projRange(it.project)})이라 비용이 발생합니다.`,
        rule?.lumps.some((l) => !l.month) && `지급월이 적혀 있지 않은 연 1회분은 계약 시작월(${it.period.start.m}월)로 가정했습니다.`,
        rule?.monthly === 'even' && rule.stated != null
          && `"매월 약 ${rule.stated.toLocaleString('ko-KR')}원" → 그해 금액을 12등분해 표시했습니다.`,
        p.basis === 'inflation' && `Excel 총 금액 ${plain(it.amount)}에서 ${svcBase() + 1}년부터 매년 3~5% 인상(백만원 단위)한 금액입니다.`,
      ])
      + dwSec('연도별 금액', esc(svcBasis(Y)), barStrip(state.years.map((y) => {
        const q = Svc.yearPlan(it, y);
        return { key: y, name: String(y), value: q?.amount ?? 0, tip: `${y}년 ${won(q?.amount ?? 0)} · ${BASIS[q?.basis ?? 'none'](y)}` };
      }), 'svcy', Y));
  }

  const yc = state.data.service.yearCols ?? [];
  const info = [
    ['금액 기준', `${yc.map((y) => `${y}년: '${y}년 금액' 열`).join(' · ')}${yc.length ? ' · ' : ''}${svcBase()}년까지: Excel 총 금액${it.amount != null ? ` ${plain(it.amount)}` : ''}<small>${svcBase() + 1}년부터 매년 3~5% 인상 (백만원 단위)</small>`],
    ['Excel 계약기간', it.cur ? `${Svc.fmtDn(it.cur.s)} ~ ${Svc.fmtDn(it.cur.e)}` : periodWarn(it)],
  ];
  if (it.lease) info.push(['리스계약 기간', `${leaseRange}<small>이 기간 비용 0원 (비고 기준)</small>`]);
  if (it.project) info.push(['Project 비용처리', `${esc(it.project.name)} Project · ${projRange(it.project)}<small>이 기간에 지출되는 금액을 Project 비용으로 구분</small>`]);
  info.push(['비고', esc(remarkIn(it, Y)) || '–']);

  return dwHead(it, '전산용역비', tags) + kpi + body + dwSec('산정 정보', '', kvList(info)) + dwRaw(it);
}

function openDrawer(id) {
  const it = [...state.data.depreciation.items, ...state.data.service.items].find((x) => x.id === id);
  if (!it) return;
  $('#drawerBody').innerHTML = it.kind === 'depreciation' ? depDetail(it) : svcDetail(it);
  lastFocus = document.activeElement;
  $('#drawer').hidden = false;
  $('.drawer__close').focus();
}

function closeDrawer() {
  if ($('#drawer').hidden) return false;
  $('#drawer').hidden = true;
  lastFocus?.focus?.();
  return true;
}

function goto(id) {
  if (id.startsWith('depreciation')) state.dep = 'all'; else state.svc = 'all';
  state.q = '';
  $('#search').value = '';
  state.detail = true;
  render();
  const tr = document.querySelector(`tr[data-id="${CSS.escape(id)}"]`);
  if (!tr) return;
  tr.scrollIntoView({ block: 'center', behavior: 'smooth' });
  tr.classList.remove('flash');
  void tr.offsetWidth; // 애니메이션 재시작
  tr.classList.add('flash');
  tr.focus({ preventScroll: true });
}

/* ---------- 오류 · 미리보기 ---------- */
const errorCard = (title, msg) => `<div class="error-card" role="alert"><h2>${esc(title)}</h2><p>${esc(msg)}</p></div>`;

function fatal(msg) {
  document.body.classList.add('is-error');
  $('#banner').innerHTML = errorCard('데이터를 표시할 수 없습니다', msg);
  $('#meta').textContent = '데이터 로드 실패';
  $('.foot').open = true;
}

async function onLocalFile(e) {
  const f = e.target.files?.[0];
  if (!f) return;
  try {
    ingest(await f.arrayBuffer(), { name: f.name, fileTime: new Date(f.lastModified), preview: true });
    document.body.classList.remove('is-error');
    $('#banner').innerHTML = `<div class="banner" role="status"><b>미리보기 중</b> ${esc(f.name)} — 이 브라우저에서만 표시되며 서버 데이터(${DATA_URL})는 변경되지 않았습니다.
      <button type="button" class="btn btn--ghost" data-reload>원래 데이터로 돌아가기</button></div>`;
    window.scrollTo({ top: 0, behavior: 'smooth' });
  } catch (err) {
    $('#banner').innerHTML = errorCard('미리보기 실패', err.message);
  }
}

/* ---------- 이벤트 ---------- */
function bindEvents() {
  $('#search').addEventListener('input', (e) => { state.q = e.target.value.trim(); if (state.data) render(); });
  $('#printBtn').addEventListener('click', () => window.print());
  $('#localFile').addEventListener('change', onLocalFile);
  window.addEventListener('beforeprint', closeDrawer);

  document.addEventListener('click', (e) => {
    const t = e.target;
    const chip = t.closest('[data-filter]');
    if (chip) {
      state[chip.dataset.group] = chip.dataset.filter;
      render();
      document.querySelector(`[data-group="${chip.dataset.group}"][data-filter="${chip.dataset.filter}"]`)?.focus();
      return;
    }
    const bar = t.closest('.chart [data-key]');
    if (bar) {
      state.year = clamp(Number(bar.dataset.key), state.years[0], state.years.at(-1));
      return openDetail('#detail');
    }
    const open = t.closest('[data-open-detail]');
    if (open) return openDetail(open.dataset.openDetail || '#detail');
    if (t.closest('[data-open-check]')) {
      state.checkOpen = true;
      return openDetail('#check');
    }
    if (t.closest('[data-close-detail]')) return closeDetail();
    const go = t.closest('[data-goto]');
    if (go) return goto(go.dataset.goto);
    if (t.closest('[data-close]')) return closeDrawer();
    if (t.closest('[data-reload]')) return location.reload();
    const tr = t.closest('tr[data-id]');
    if (tr) openDrawer(tr.dataset.id);
  });

  document.addEventListener('keydown', (e) => {
    if (e.key === 'Escape') { if (!closeDrawer() && state.detail) closeDetail(); return; }
    if (e.key !== 'Enter' && e.key !== ' ') return;
    const el = document.activeElement;
    if (el?.matches?.('tr[data-id]')) {
      e.preventDefault();
      openDrawer(el.dataset.id);
    } else if (el?.matches?.('.chart [data-key]')) {
      e.preventDefault();
      el.dispatchEvent(new MouseEvent('click', { bubbles: true }));
    }
  });
}

function bindTips() {
  const tip = $('#tip');
  const show = (e) => {
    const el = e.target.closest?.('[data-tip]');
    if (!el?.dataset.tip) { tip.hidden = true; return; }
    tip.textContent = el.dataset.tip;
    tip.hidden = false;
    const r = el.getBoundingClientRect(), w = tip.offsetWidth, h = tip.offsetHeight;
    tip.style.left = `${clamp(r.left + r.width / 2 - w / 2, 8, innerWidth - w - 8)}px`;
    tip.style.top = `${r.top - h - 8 < 8 ? r.bottom + 8 : r.top - h - 8}px`;
  };
  document.addEventListener('mouseover', show);
  document.addEventListener('focusin', show);
  document.addEventListener('scroll', () => { tip.hidden = true; }, true);
}

bindEvents();
bindTips();
const resize = new ResizeObserver((entries) => entries.forEach((en) => drawChart(en.target)));
document.querySelectorAll('.chart').forEach((el) => resize.observe(el));
loadDefault().catch((e) => fatal(e.message));
