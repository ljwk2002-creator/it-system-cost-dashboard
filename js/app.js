/**
 * app.js — 데이터 로드 · 상태 · 화면 렌더링.
 * 화면의 모든 금액·항목·기간은 data/current.xlsx에서 읽어 계산한다 (코드에 데이터 없음).
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
const inUnit = (v, digits = 1) => (v / unit.div).toLocaleString('ko-KR', { minimumFractionDigits: digits, maximumFractionDigits: digits });
const plain = (v, digits) => `${inUnit(v, digits)} ${unit.label}`;
const money = (v, digits) => `<span class="num" data-tip="${esc(won(v))}">${inUnit(v, digits)}<span class="u">${unit.label}</span></span>`;
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
  Svc.prepareContracts(S.items);

  state.data = data;
  state.source = { ...source, saved: validDate(wb.Props?.ModifiedDate) };
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

/* ---------- 렌더링 ---------- */
function render() {
  const { depreciation: D, service: S } = state.data;
  const Y = state.year;
  const dep = D.items.map((it) => ({ it, st: Dep.statusInYear(it, Y), months: Dep.monthsInYear(it, Y), amount: Dep.amountInYear(it, Y) }));
  const svc = S.items.map((it) => ({ it, v: Svc.yearView(it, Y) }));
  const ctx = { Y, D, S, dep, svc, depTotal: sum(dep, (r) => r.amount), svcTotal: sum(svc, (r) => r.v?.total), check: checkSummary(D, S) };
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
  const common = { selected: Y, unit: unit.label, label: (d, digits) => inUnit(d.value, digits) };
  const depOf = (y) => (D.error ? 0 : Dep.totalInYear(D.items, y));
  const svcOf = (y) => (S.error ? 0 : svcYear(y));
  setChart('totalYearChart', {
    ...common,
    data: state.years.map((y) => {
      const parts = [{ cls: 'seg--dep', value: depOf(y) }, { cls: 'seg--svc', value: svcOf(y) }];
      return { key: y, name: String(y), value: parts[0].value + parts[1].value, parts };
    }),
    tip: (d) => `${d.key}년 IT System 비용 합계 ${won(d.value)}\n무형자산상각비 ${won(d.parts[0].value)}\n전산용역비 ${won(d.parts[1].value)} (연장 가정 포함)\n선택하면 상세 내역`,
  });
  if (D.error) {
    chartArgs.delete('depYearChart');
    $('#depYearChart').innerHTML = sectionError(D.error);
  } else {
    setChart('depYearChart', {
      ...common,
      data: state.years.map((y) => ({ key: y, name: String(y), value: depOf(y) })),
      tip: (d) => `${d.key}년 무형자산상각비 ${won(d.value)}\n선택하면 상세 내역`,
    });
  }
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

function subHead({ eyebrow, title, total, note, filter }) {
  return `<div class="sub__title"><p class="eyebrow">${eyebrow}</p><h3>${title}</h3></div>
    ${bigMoney(total, 'sub__total')}
    <p class="sub__note">${note}</p>
    <div class="chips no-print">${filter}</div>`;
}

/* ----- 무형자산 상세 ----- */
function renderDep({ Y, D, dep, depTotal }) {
  if (D.error) {
    $('#depHead').innerHTML = `<div class="sub__title"><h3>무형자산상각비</h3></div>${sectionError(D.error)}`;
    $('#depTable').innerHTML = '';
    return;
  }
  const is = {
    all: () => true, active: (r) => r.st === 'active', planned: (r) => r.st === 'planned',
    done: (r) => r.st === 'done', check: (r) => r.it.issues.length > 0,
  };
  const searched = dep.filter((r) => matches(r.it));
  $('#depHead').innerHTML = subHead({
    eyebrow: 'INTANGIBLE ASSETS · 무형자산 상세', title: `${Y}년 무형자산상각비`, total: depTotal,
    note: `상각 중 ${dep.filter((r) => r.st === 'active').length}건 · 월평균 ${plain(depTotal / 12)}${yoy(depTotal, Dep.totalInYear(D.items, Y - 1), Y)}
      · 총 취득금액 ${plain(sum(D.items, (i) => i.amount))}`,
    filter: chips('dep', [['all', '전체'], ['active', '상각 중'], ['planned', '상각 예정'], ['done', '상각 완료'], ['check', '⚠ 확인 필요']],
      (k) => searched.filter(is[k]).length),
  });
  const rows = searched.filter(is[state.dep]).sort((a, b) => b.amount - a.amount || (b.it.amount ?? 0) - (a.it.amount ?? 0));
  $('#depTable').innerHTML = `
    <thead><tr>
      <th scope="col">항목</th><th scope="col" class="r">총 취득금액</th><th scope="col">상각기간</th><th scope="col" class="r">월 상각액</th>
      <th scope="col" class="r em">${Y}년 상각액</th><th scope="col" class="share-col">비중</th><th scope="col">상태</th><th scope="col">비고</th>
    </tr></thead>
    <tbody>${rows.map((r) => depRow(r, depTotal)).join('') || emptyRow(8)}</tbody>
    <tfoot><tr>
      <th scope="row">합계 ${rows.length}건</th>
      <td class="r">${money(sum(rows, (r) => r.it.amount))}</td><td></td>
      <td class="r">${money(sum(rows, (r) => r.it.monthly), 2)}</td>
      <td class="r em">${money(sum(rows, (r) => r.amount))}</td>
      <td>${share(sum(rows, (r) => r.amount), depTotal)}</td><td colspan="2"></td>
    </tr></tfoot>`;
}

function depRow({ it, st, months, amount }, total) {
  const p = it.period;
  const mismatch = it.issues.find((x) => x.type === 'data' && x.msg.startsWith('월 상각액'));
  return `<tr data-id="${esc(it.id)}" tabindex="0">
    <th scope="row" class="name">${esc(it.name)}${issueIcon(it)}</th>
    <td class="r">${it.amount != null ? money(it.amount) : amountWarn(it)}</td>
    <td class="nowrap">${p.status === 'ok' ? `${fmtYM(p.start)} ~ ${fmtYM(p.end)}<small>${it.months}개월</small>` : periodWarn(it)}</td>
    <td class="r">${it.monthly != null ? money(it.monthly, 2) : '<span class="dim">–</span>'}${mismatch ? ` <span class="warn" data-tip="${esc(mismatch.msg)}">⚠</span>` : ''}</td>
    <td class="r em">${months ? `${money(amount)}<small>${months}개월</small>` : '<span class="dim">–</span>'}</td>
    <td>${share(amount, total)}</td>
    <td><span class="st st--${st}">${DEP_STATUS[st]}</span></td>
    <td class="remark">${esc(it.remark)}</td>
  </tr>`;
}

/* ----- 전산용역 계약 현황 ----- */
function renderSvc({ Y, S, svc, svcTotal }) {
  if (S.error) {
    $('#svcHead').innerHTML = `<div class="sub__title"><h3>전산용역비</h3></div>${sectionError(S.error)}`;
    $('#svcTable').innerHTML = '';
    return;
  }
  const counted = svc.filter((r) => r.v);
  const excluded = svc.filter((r) => r.it.excluded).length;
  const failed = svc.length - counted.length - excluded;
  const is = { all: () => true, check: (r) => r.it.issues.length > 0, excluded: (r) => !!r.it.excluded };
  const searched = svc.filter((r) => matches(r.it));
  $('#svcHead').innerHTML = subHead({
    eyebrow: 'IT SERVICE CONTRACTS · 전산용역 계약 현황', title: `${Y}년 전산용역비`, total: svcTotal,
    note: `집계 ${counted.filter((r) => r.v.total > 0).length}건 · 집계 제외 ${excluded}건${failed ? ` · <span class="warn">확인 필요 ${failed}건</span>` : ''}
${yoy(svcTotal, svcYear(Y - 1), Y)}<br>동일 주기·동일 금액 연장 가정 포함 — 현재 계약분 ${plain(sum(counted, (r) => r.v.cur))} + 연장 가정분 ${plain(sum(counted, (r) => r.v.ext))}`,
    filter: chips('svc', [['all', '전체'], ['check', '⚠ 확인 필요'], ['excluded', '집계 제외']], (k) => searched.filter(is[k]).length),
  });
  const rank = (r) => (r.v ? r.v.total : r.it.excluded ? -2 : -1); // 집계 → 확인 필요 → 집계 제외
  const rows = searched.filter(is[state.svc]).sort((a, b) => rank(b) - rank(a) || (b.it.amount ?? 0) - (a.it.amount ?? 0));
  $('#svcTable').innerHTML = `
    <thead><tr>
      <th scope="col">System</th><th scope="col" class="r">계약금액</th><th scope="col" class="r em">${Y}년 금액</th>
      <th scope="col" class="share-col">비중</th><th scope="col">비고</th>
    </tr></thead>
    <tbody>${rows.map((r) => svcRow(r, Y, svcTotal)).join('') || emptyRow(5)}</tbody>
    <tfoot><tr>
      <th scope="row">합계 ${rows.length}건</th>
      <td class="r">${money(sum(rows, (r) => (r.it.amountOk ? r.it.amount : 0)))}</td>
      <td class="r em">${money(sum(rows, (r) => r.v?.total))}</td>
      <td>${share(sum(rows, (r) => r.v?.total), svcTotal)}</td>
      <td class="unit-note">집계 제외 · 확인 필요 항목은 합계에서 제외</td>
    </tr></tfoot>`;
}

function svcRow({ it, v }, Y, total) {
  const cycle = it.cycle?.unit === 'months' && it.cycle.n !== 12 ? `<small>${Svc.cycleLabel(it.cycle)} 계약</small>` : '';
  const amountCell = it.excluded ? (it.amount != null ? money(it.amount) : '<span class="dim">–</span>')
    : it.amountOk ? money(it.amount) + cycle
      : it.amount != null ? `${money(it.amount)} ${issueIcon({ issues: it.issues.filter((x) => x.type === 'amount') })}`
        : amountWarn(it);
  const why = !it.cur ? periodTip(it) : it.issues.filter((x) => x.type === 'amount').map((x) => x.msg).join('\n');
  const yearCell = it.excluded ? `<span class="tag" data-tip="${esc(`집계 제외 — 비고: ${it.remark}`)}">${esc(it.excluded)}</span>`
    : !v ? `<span class="warn" data-tip="${esc(`집계 제외\n${why}`)}">⚠ 미집계</span>`
      : v.total ? `<span class="num" data-tip="${esc(`${Y}년 ${won(v.total)}\n현재 계약분 ${won(v.cur)}\n연장 가정분 ${won(v.ext)}`)}">${inUnit(v.total)}<span class="u">${unit.label}</span></span>`
        : '<span class="dim">–</span>';
  return `<tr data-id="${esc(it.id)}" tabindex="0">
    <th scope="row" class="name">${esc(it.name)}${issueIcon(it)}</th>
    <td class="r">${amountCell}</td>
    <td class="r em">${yearCell}</td>
    <td>${share(v?.total ?? 0, total)}</td>
    <td class="remark">${esc(it.remark)}</td>
  </tr>`;
}

/* ---------- Drill-down (행 클릭) ---------- */
const kvList = (rows) => `<dl class="kv">${rows.map(([k, v]) => `<dt>${esc(k)}</dt><dd>${v}</dd>`).join('')}</dl>`;
const drawerHead = (it, kind) => `<p class="eyebrow">${kind} · ${esc(it.sheet)} ${it.rowNo}행</p><h2 id="drawerTitle">${esc(it.name)}</h2>`
  + (it.issues.length ? `<ul class="issues">${it.issues.map((x) => `<li><b>⚠ ${ISSUE[x.type]}</b> ${esc(x.msg)}</li>`).join('')}</ul>` : '');
const rawTable = (it) => `<h3>원본 Excel 값</h3><table class="mini src"><tbody>${it.raw.map((r) =>
  `<tr><th scope="row">${esc(r.label)}</th><td class="addr">${esc(r.addr)}</td><td>${esc(r.text) || '<span class="dim">(공란)</span>'}</td></tr>`).join('')}</tbody></table>`;
const wonUnit = (v) => `${won(v)} <small>${plain(v)}</small>`;
const miniTable = (head, rows, foot, selected) => `<table class="mini"><thead><tr>${head.map((h, i) => `<th${i ? ' class="r"' : ''}>${h}</th>`).join('')}</tr></thead>
  <tbody>${rows.map((r) => `<tr${r[0] === selected ? ' class="is-sel"' : ''}>${r.map((c, i) => `<td${i ? ' class="r"' : ''}>${c}</td>`).join('')}</tr>`).join('')}</tbody>
  ${foot ? `<tfoot><tr>${foot.map((c, i) => `<td${i ? ' class="r"' : ''}>${c}</td>`).join('')}</tr></tfoot>` : ''}</table>`;

function depDetail(it) {
  const Y = state.year, p = it.period;
  const ok = it.monthly != null;
  const monthlyNote = it.monthlyExcel != null
    ? `Excel 입력값${it.monthlyCalc != null ? ` · 계산값 ${won(it.monthlyCalc)}` : ''}` : '계산값 (총 금액 ÷ 총 상각개월)';
  const yearly = ok ? `<h3>연도별 상각액</h3>${miniTable(['연도', '개월', '상각액', '원'],
    span(p.start.y, p.end.y).map((y) => [`${y}년`, `${Dep.monthsInYear(it, y)}개월`, plain(Dep.amountInYear(it, y)), won(Dep.amountInYear(it, y))]),
    ['합계', `${it.months}개월`, plain(it.monthly * it.months), won(it.monthly * it.months)], `${Y}년`)}` : '';
  return drawerHead(it, '무형자산') + kvList([
    ['총 취득금액', it.amount != null ? wonUnit(it.amount) : amountWarn(it)],
    ['상각기간', p.status === 'ok' ? `${fmtYM(p.start)} ~ ${fmtYM(p.end)} <small>${it.months}개월</small>` : periodWarn(it)],
    ['월 상각액', ok ? `${won(it.monthly)}<small>${plain(it.monthly, 2)} · ${monthlyNote}</small>` : '–'],
    [`${Y}년 상각액`, ok ? `${wonUnit(Dep.amountInYear(it, Y))}<small>${Dep.monthsInYear(it, Y)}개월 · 누적 ${Dep.elapsedAtYearEnd(it, Y)}/${it.months}개월</small>` : '–'],
    ['상태', DEP_STATUS[Dep.statusInYear(it, Y)]],
    ['비고', esc(it.remark) || '–'],
  ]) + yearly + rawTable(it);
}

function svcDetail(it) {
  const Y = state.year, v = Svc.yearView(it, Y);
  const rows = [
    ['계약금액', it.amountOk ? wonUnit(it.amount) : it.amount != null ? `${wonUnit(it.amount)} <span class="warn">⚠ 확인 필요</span>` : amountWarn(it)],
    ['현재 계약기간', it.cur ? `${Svc.fmtDn(it.cur.s)} ~ ${Svc.fmtDn(it.cur.e)}` : periodWarn(it)],
  ];
  if (it.cur) {
    const next = Svc.renewals(it, it.cur.e + 1)[0];
    rows.push(
      ['계약 주기', `${Svc.cycleLabel(it.cycle)} <small>${esc(it.cycle.basis)}</small>`],
      ['다음 연장 가정', `${Svc.fmtDn(next.s)} ~ ${Svc.fmtDn(next.e)}<small>${it.nextAmount != null ? `차기계약금액 ${plain(it.nextAmount)}` : '동일 금액 가정'}</small>`],
    );
  }
  rows.push(
    [`${Y}년 금액`, it.excluded ? `집계 제외 <span class="tag">${esc(it.excluded)}</span>`
      : v ? `${wonUnit(v.total)}<small>현재 계약분 ${plain(v.cur)} + 연장 가정분 ${plain(v.ext)}</small>`
        : '<span class="warn">⚠ 미집계</span> <small>금액 또는 기간 확인 필요</small>'],
    ['비고', esc(it.remark) || '–'],
  );
  const monthly = v ? `<h3>${Y}년 월별 금액</h3>${miniTable(['월', '금액', '원'],
    v.months.map((a, i) => [MONTHS[i], plain(a, 2), won(a)]), ['합계', plain(v.total, 2), won(v.total)])}` : '';
  return drawerHead(it, '전산용역비') + kvList(rows) + monthly + rawTable(it);
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
