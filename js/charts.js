/**
 * charts.js — 연도별 Column Chart (의존성 없는 SVG)
 * 입력은 숫자와 호출자가 만든 짧은 문자열(연도·금액·단위)뿐이다 (Excel 원문 텍스트는 받지 않음).
 * data: [{ key, name, value, parts?, proj? }]  selected: 강조할 key
 *   parts: [{ cls, value }] 가 있으면 아래부터 쌓은 누적 막대 (value = 합계)
 *   proj: 그중 Project 비용 → 막대 위쪽에 반투명 해치 박스(index.html #projHatch) + 박스 우측 상단에 금액(projLabel)
 * 데이터 레이블은 2줄: 금액(굵게) / 단위
 */
export function columnChart({ data, selected = null, width, height = 270, label, unit, tip, projLabel }) {
  const pad = { t: 44, r: 4, b: 28, l: 4 };
  const ih = height - pad.t - pad.b;
  const max = Math.max(...data.map((d) => d.value), 1);
  const slot = (width - pad.l - pad.r) / data.length;
  const bw = Math.min(64, slot * 0.6);
  const base = pad.t + ih;
  const compact = slot < 60; // 좁은 화면: 작은 글씨

  const projLabels = []; // 옆 막대에 가리지 않도록 모든 막대 뒤(맨 위 레이어)에 그림
  const bars = data.map((d, i) => {
    const sx = pad.l + slot * i, cx = sx + slot / 2;
    const h = d.value > 0 ? Math.max(2, (d.value / max) * ih) : 0;
    const sel = d.key === selected;
    let top = base;
    const cols = (d.parts ?? [{ cls: '', value: d.value }]).map((p) => {
      const ph = d.parts ? (p.value / max) * ih : h;
      top -= ph;
      return `<rect class="col ${p.cls}" x="${cx - bw / 2}" y="${top}" width="${bw}" height="${ph}" rx="${d.parts ? 0 : 3}"/>`;
    }).join('');
    const ph = d.proj > 0 ? Math.min(h, Math.max(3, (d.proj / max) * ih)) : 0;
    if (ph) {
      const inside = compact && ph >= 22; // 좁은 화면: 박스 안 가운데 (막대 사이 공간이 부족)
      const lx = inside ? cx : cx + bw / 2 + 3;
      projLabels.push(`<text class="proj__v${inside ? ' is-in' : ''}" x="${lx}" y="${base - h + (inside ? 11 : 9)}">${projLabel(d)}<tspan class="proj__u" x="${lx}" dy="${inside ? 9 : 11}">${unit}</tspan></text>`);
    }
    return `<g class="bar${sel ? ' is-sel' : ''}" data-key="${d.key}" data-tip="${tip(d)}" tabindex="0" role="button" aria-label="${d.name} ${label(d)} ${unit}">
      <rect class="hit" x="${sx}" y="0" width="${slot}" height="${height}"/>
      ${cols}${ph ? `<rect class="proj" x="${cx - bw / 2}" y="${base - h}" width="${bw}" height="${ph}"/>` : ''}
      <text class="val" x="${cx}" y="${base - h - 21}">${label(d)}<tspan class="val__u" x="${cx}" dy="14">${unit}</tspan></text>
      <text class="xl" x="${cx}" y="${height - 8}">${d.name}</text>
    </g>`;
  }).join('');

  return `<svg class="${compact ? 'compact' : ''}" viewBox="0 0 ${width} ${height}" width="100%" height="${height}" role="group">
    <line class="base" x1="${pad.l}" x2="${width - pad.r}" y1="${base}" y2="${base}"/>${bars}${projLabels.join('')}</svg>`;
}
