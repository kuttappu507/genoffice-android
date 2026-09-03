import { useMemo } from 'react';
import type { Chart } from '../lib/sheet-model';
import { hasHeaderRow, parseNumberish, rangeToRect, refOf } from '../lib/sheet-model';

export const CHART_COLORS = ['#217346', '#2B579A', '#D24726', '#F2C811', '#7030A0', '#00B0F0', '#E97132', '#5B9BD5', '#70AD47', '#A5A5A5'];

export interface ChartData {
  labels: string[];
  series: { name: string; values: number[] }[];
}

/** Pull labels + numeric series out of the grid for a chart definition. */
export function chartData(chart: Chart, cells: Record<string, string>, display: Record<string, string>): ChartData {
  const rect = rangeToRect(chart.range);
  if (!rect) return { labels: [], series: [] };
  const header = chart.headerRow ?? hasHeaderRow(cells, rect);
  const firstDataRow = header ? rect.r1 + 1 : rect.r1;
  const firstDataCol = chart.labelsInFirstCol ? rect.c1 + 1 : rect.c1;
  const labels: string[] = [];
  for (let r = firstDataRow; r <= rect.r2; r++) labels.push(chart.labelsInFirstCol ? display[refOf(rect.c1, r)] ?? cells[refOf(rect.c1, r)] ?? '' : String(r - firstDataRow + 1));
  const series: ChartData['series'] = [];
  for (let c = firstDataCol; c <= rect.c2; c++) {
    const name = header ? display[refOf(c, rect.r1)] ?? cells[refOf(c, rect.r1)] ?? `Series ${c - firstDataCol + 1}` : `Series ${c - firstDataCol + 1}`;
    const values: number[] = [];
    for (let r = firstDataRow; r <= rect.r2; r++) values.push(parseNumberish(display[refOf(c, r)] ?? cells[refOf(c, r)] ?? '') ?? 0);
    series.push({ name, values });
  }
  return { labels, series };
}

function niceMax(v: number): number {
  if (v <= 0) return 1;
  const p = Math.pow(10, Math.floor(Math.log10(v)));
  const n = v / p;
  const m = n <= 1 ? 1 : n <= 2 ? 2 : n <= 2.5 ? 2.5 : n <= 5 ? 5 : 10;
  return m * p;
}

function fmtTick(v: number): string {
  if (Math.abs(v) >= 1e6) return `${parseFloat((v / 1e6).toPrecision(3))}M`;
  if (Math.abs(v) >= 1e3) return `${parseFloat((v / 1e3).toPrecision(3))}k`;
  return String(parseFloat(v.toPrecision(3)));
}

/** Pure-SVG chart renderer (column / bar / line / area / pie). Also used for PNG export. */
export function ChartView({ chart, data, width, height, dark }: { chart: Chart; data: ChartData; width: number; height: number; dark?: boolean }) {
  const fg = dark ? '#e8e8ea' : '#333';
  const grid = dark ? '#3a3a40' : '#e3e3e3';
  const bg = dark ? '#1e1e22' : '#ffffff';
  const titleH = chart.title ? 26 : 8;
  const legendH = data.series.length > 1 || chart.type === 'pie' ? 22 : 0;
  const m = { top: titleH + 8, right: 12, bottom: 30 + legendH, left: 44 };
  const pw = Math.max(10, width - m.left - m.right);
  const ph = Math.max(10, height - m.top - m.bottom);

  const { maxV, minV } = useMemo(() => {
    let mx = 0;
    let mn = 0;
    for (const s of data.series) for (const v of s.values) { if (v > mx) mx = v; if (v < mn) mn = v; }
    return { maxV: niceMax(mx), minV: mn < 0 ? -niceMax(-mn) : 0 };
  }, [data]);

  const n = data.labels.length;
  const legend = (
    legendH > 0 && (
      <g transform={`translate(${m.left}, ${height - legendH + 4})`}>
        {(chart.type === 'pie' ? data.labels.slice(0, 6) : data.series.map((s) => s.name)).map((name, i) => (
          <g key={i} transform={`translate(${i * Math.min(90, pw / Math.max(1, chart.type === 'pie' ? Math.min(6, data.labels.length) : data.series.length))}, 0)`}>
            <rect width="10" height="10" rx="2" fill={CHART_COLORS[i % CHART_COLORS.length]} />
            <text x="14" y="9" fontSize="10" fill={fg}>{String(name).slice(0, 12)}</text>
          </g>
        ))}
      </g>
    )
  );

  const title = chart.title && (
    <text x={width / 2} y={18} textAnchor="middle" fontSize="13" fontWeight="600" fill={fg}>{chart.title}</text>
  );

  if (chart.type === 'pie') {
    const s = data.series[0] ?? { values: [] as number[] };
    const total = s.values.reduce((a, b) => a + Math.max(0, b), 0) || 1;
    const cx = width / 2;
    const cy = m.top + ph / 2;
    const r = Math.min(pw, ph) / 2 - 4;
    let ang = -Math.PI / 2;
    const slices = s.values.map((v, i) => {
      const frac = Math.max(0, v) / total;
      const a0 = ang;
      const a1 = ang + frac * Math.PI * 2;
      ang = a1;
      const large = a1 - a0 > Math.PI ? 1 : 0;
      const x0 = cx + r * Math.cos(a0), y0 = cy + r * Math.sin(a0);
      const x1 = cx + r * Math.cos(a1), y1 = cy + r * Math.sin(a1);
      const mid = (a0 + a1) / 2;
      const lx = cx + r * 0.65 * Math.cos(mid), ly = cy + r * 0.65 * Math.sin(mid);
      return { d: frac >= 0.9999 ? `M ${cx - r} ${cy} A ${r} ${r} 0 1 1 ${cx + r} ${cy} A ${r} ${r} 0 1 1 ${cx - r} ${cy}` : `M ${cx} ${cy} L ${x0} ${y0} A ${r} ${r} 0 ${large} 1 ${x1} ${y1} Z`, i, frac, lx, ly };
    });
    return (
      <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ background: bg, borderRadius: 8 }}>
        {title}
        {slices.map((sl) => (
          <path key={sl.i} d={sl.d} fill={CHART_COLORS[sl.i % CHART_COLORS.length]} stroke={bg} strokeWidth="1.5" />
        ))}
        {slices.filter((sl) => sl.frac > 0.05).map((sl) => (
          <text key={`t${sl.i}`} x={sl.lx} y={sl.ly} textAnchor="middle" fontSize="10" fill="#fff" fontWeight="600">{Math.round(sl.frac * 100)}%</text>
        ))}
        {legend}
      </svg>
    );
  }

  const horizontal = chart.type === 'bar';
  const span = maxV - minV || 1;
  const yOf = (v: number) => m.top + ph - ((v - minV) / span) * ph;
  const xOf = (v: number) => m.left + ((v - minV) / span) * pw;
  const ticks = [0, 0.25, 0.5, 0.75, 1].map((t) => minV + t * span);
  const groupW = (horizontal ? ph : pw) / Math.max(1, n);
  const barW = Math.max(2, (groupW * 0.72) / Math.max(1, data.series.length));

  return (
    <svg width={width} height={height} viewBox={`0 0 ${width} ${height}`} style={{ background: bg, borderRadius: 8 }}>
      {title}
      {/* grid + axis labels */}
      {ticks.map((t, i) =>
        horizontal ? (
          <g key={i}>
            <line x1={xOf(t)} x2={xOf(t)} y1={m.top} y2={m.top + ph} stroke={grid} />
            <text x={xOf(t)} y={m.top + ph + 14} fontSize="9" textAnchor="middle" fill={fg}>{fmtTick(t)}</text>
          </g>
        ) : (
          <g key={i}>
            <line x1={m.left} x2={m.left + pw} y1={yOf(t)} y2={yOf(t)} stroke={grid} />
            <text x={m.left - 6} y={yOf(t) + 3} fontSize="9" textAnchor="end" fill={fg}>{fmtTick(t)}</text>
          </g>
        ),
      )}
      {/* category labels */}
      {data.labels.map((l, i) =>
        horizontal ? (
          <text key={i} x={m.left - 6} y={m.top + i * groupW + groupW / 2 + 3} fontSize="9" textAnchor="end" fill={fg}>{String(l).slice(0, 8)}</text>
        ) : (
          n <= 16 || i % Math.ceil(n / 16) === 0 ? (
            <text key={i} x={m.left + i * groupW + groupW / 2} y={m.top + ph + 14} fontSize="9" textAnchor="middle" fill={fg}>{String(l).slice(0, 8)}</text>
          ) : null
        ),
      )}
      {/* zero line */}
      {horizontal ? <line x1={xOf(0)} x2={xOf(0)} y1={m.top} y2={m.top + ph} stroke={fg} strokeOpacity="0.5" /> : <line x1={m.left} x2={m.left + pw} y1={yOf(0)} y2={yOf(0)} stroke={fg} strokeOpacity="0.5" />}
      {/* series */}
      {data.series.map((s, si) => {
        const color = CHART_COLORS[si % CHART_COLORS.length];
        if (chart.type === 'line' || chart.type === 'area') {
          const pts = s.values.map((v, i) => [m.left + i * groupW + groupW / 2, yOf(v)] as const);
          const path = pts.map((p, i) => `${i ? 'L' : 'M'} ${p[0]} ${p[1]}`).join(' ');
          return (
            <g key={si}>
              {chart.type === 'area' && pts.length > 1 && <path d={`${path} L ${pts[pts.length - 1][0]} ${yOf(0)} L ${pts[0][0]} ${yOf(0)} Z`} fill={color} fillOpacity="0.22" />}
              <path d={path} fill="none" stroke={color} strokeWidth="2" strokeLinejoin="round" />
              {pts.length <= 40 && pts.map((p, i) => <circle key={i} cx={p[0]} cy={p[1]} r="2.6" fill={color} />)}
            </g>
          );
        }
        return (
          <g key={si}>
            {s.values.map((v, i) =>
              horizontal ? (
                <rect
                  key={i}
                  x={Math.min(xOf(0), xOf(v))}
                  y={m.top + i * groupW + (groupW - barW * data.series.length) / 2 + si * barW}
                  width={Math.abs(xOf(v) - xOf(0))}
                  height={barW}
                  fill={color}
                  rx="1.5"
                />
              ) : (
                <rect
                  key={i}
                  x={m.left + i * groupW + (groupW - barW * data.series.length) / 2 + si * barW}
                  y={Math.min(yOf(0), yOf(v))}
                  width={barW}
                  height={Math.abs(yOf(v) - yOf(0))}
                  fill={color}
                  rx="1.5"
                />
              ),
            )}
          </g>
        );
      })}
      {legend}
    </svg>
  );
}
