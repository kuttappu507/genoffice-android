import type { CSSProperties, PointerEvent as RPointerEvent } from 'react';
import type { DeckSlide, ShapePara, SlideShape } from '../lib/deck-model';
import { isDark } from '../lib/deck-model';

const alignCss = (a: ShapePara['align']): CSSProperties['textAlign'] => (a === 'center' ? 'center' : a === 'right' ? 'right' : a === 'justify' ? 'justify' : 'left');

/** SVG path for a preset geometry in a 100×100 box (non-rect shapes). */
function geomPath(g: NonNullable<SlideShape['geom']>): string | null {
  switch (g) {
    case 'ellipse': return 'M50 0 A50 50 0 1 1 49.99 0 Z';
    case 'triangle': return 'M50 0 L100 100 L0 100 Z';
    case 'diamond': return 'M50 0 L100 50 L50 100 L0 50 Z';
    case 'rightArrow': return 'M0 30 L65 30 L65 5 L100 50 L65 95 L65 70 L0 70 Z';
    case 'star': return 'M50 0 L61 35 L98 35 L68 57 L79 91 L50 70 L21 91 L32 57 L2 35 L39 35 Z';
    case 'hexagon': return 'M25 0 L75 0 L100 50 L75 100 L25 100 L0 50 Z';
    case 'chevron': return 'M0 0 L75 0 L100 50 L75 100 L0 100 L25 50 Z';
    case 'line': return 'M0 50 L100 50';
    default: return null;
  }
}

export interface SlideViewProps {
  slide: DeckSlide;
  width: number;
  /** editing affordances: selection outline + handles + drag */
  edit?: boolean;
  selected?: string | null;
  onSelect?: (id: string | null) => void;
  onDoubleTap?: (id: string) => void;
  onDragStart?: (e: RPointerEvent, id: string, mode: 'move' | 'resize') => void;
  /** step of the entrance animation when presenting (shapes with anim beyond this step stay hidden) */
  animStep?: number;
  className?: string;
}

/**
 * Faithful slide renderer: positioned shapes on a cw × ch inch canvas scaled to
 * `width` px. Shared by the editor canvas, filmstrip, sorter, presenter and
 * image/PDF export so every view matches.
 */
export function SlideView({ slide, width, edit, selected, onSelect, onDoubleTap, onDragStart, animStep, className }: SlideViewProps) {
  const cw = slide.cw ?? 10;
  const ch = slide.ch ?? 5.63;
  if (width <= 0) return null;
  const dark = isDark(slide.bg);
  const ptScale = width / (cw * 72); // px per point
  const height = (width * ch) / cw;
  let animIdx = 0;
  return (
    <div
      className={`sv${edit ? ' editing' : ''}${className ? ` ${className}` : ''}`}
      style={{ width, height, background: slide.bg ?? '#FFFFFF', fontSize: Math.max(2, 12 * ptScale) }}
      onPointerDown={edit && onSelect ? (e) => { if (e.target === e.currentTarget) onSelect(null); } : undefined}
    >
      {(slide.shapes ?? []).map((sh, i) => {
        const id = sh.id ?? String(i);
        const isSel = edit && selected === id;
        const hasAnim = sh.anim && sh.anim !== 'none';
        const myStep = hasAnim ? ++animIdx : 0;
        const hiddenByAnim = animStep !== undefined && hasAnim && myStep > animStep;
        const style: CSSProperties = {
          left: `${(sh.x / cw) * 100}%`,
          top: `${(sh.y / ch) * 100}%`,
          width: `${(sh.w / cw) * 100}%`,
          height: `${(sh.h / ch) * 100}%`,
          transform: sh.rot ? `rotate(${sh.rot}deg)` : undefined,
          opacity: hiddenByAnim ? 0 : sh.opacity ?? 1,
          filter: sh.shadow ? 'drop-shadow(0 2px 4px rgba(0,0,0,.28))' : undefined,
          transition: animStep !== undefined && hasAnim ? 'opacity .45s ease, transform .45s ease' : undefined,
        };
        const path = sh.geom && sh.geom !== 'rect' && sh.geom !== 'roundRect' ? geomPath(sh.geom) : null;
        const boxStyle: CSSProperties = path
          ? {}
          : {
              background: sh.fill,
              border: sh.line ? `${Math.max(1, (sh.lineW ?? 1) * ptScale)}px solid ${sh.line}` : undefined,
              borderRadius: sh.geom === 'roundRect' ? `${Math.min(sh.w, sh.h) * 0.16 * ptScale * 72}px` : undefined,
            };
        const content =
          sh.kind === 'image' && sh.img ? (
            <img src={sh.img} alt="" draggable={false} />
          ) : (
            <div className={`sv-text v-${sh.valign ?? 'top'}`}>
              {sh.paras.map((p, j) => {
                const text = p.runs.map((r) => r.text).join('');
                const first = p.runs.find((r) => r.text.trim()) ?? p.runs[0];
                const size = Math.max(4, (first?.sz ?? 14) * ptScale);
                if (!text.trim()) return <p key={j} style={{ fontSize: size }}>{'\u00a0'}</p>;
                return (
                  <p
                    key={j}
                    style={{
                      fontSize: size,
                      textAlign: alignCss(p.align),
                      paddingLeft: p.bullet ? `${1 + (p.level ?? 0) * 0.9}em` : p.level ? `${p.level * 0.9}em` : undefined,
                      textIndent: p.bullet ? '-0.75em' : undefined,
                      fontFamily: first?.font,
                    }}
                  >
                    {p.bullet ? '\u2022 ' : ''}
                    {p.runs.map((r, k) => (
                      <span
                        key={k}
                        style={{
                          fontSize: r.sz && r.sz !== first?.sz ? Math.max(4, r.sz * ptScale) : undefined,
                          color: r.color ?? (dark ? '#F2F2F2' : '#333333'),
                          fontWeight: r.b ? 700 : 400,
                          fontStyle: r.i ? 'italic' : undefined,
                          textDecoration: r.u && r.s ? 'underline line-through' : r.u ? 'underline' : r.s ? 'line-through' : undefined,
                          fontFamily: r.font,
                          background: r.highlight,
                        }}
                      >
                        {r.text}
                      </span>
                    ))}
                  </p>
                );
              })}
            </div>
          );
        return (
          <div
            key={id}
            data-shape={id}
            className={`sv-sh${sh.kind === 'image' ? ' img' : ''}${edit && !sh.locked ? ' tappable' : ''}${isSel ? ' sel' : ''}${sh.locked ? ' locked' : ''}`}
            style={{ ...style, ...boxStyle }}
            onPointerDown={
              edit && !sh.locked && onDragStart
                ? (e) => { e.stopPropagation(); onSelect?.(id); onDragStart(e, id, 'move'); }
                : edit && onSelect ? (e) => { e.stopPropagation(); if (!sh.locked) onSelect(id); } : undefined
            }
            onDoubleClick={edit && onDoubleTap && !sh.locked ? () => onDoubleTap(id) : undefined}
          >
            {path && (
              <svg className="sv-geom" viewBox="0 0 100 100" preserveAspectRatio="none">
                <path d={path} fill={sh.geom === 'line' ? 'none' : sh.fill ?? 'transparent'} stroke={sh.line ?? (sh.geom === 'line' ? sh.fill ?? '#333' : 'none')} strokeWidth={sh.geom === 'line' ? Math.max(1.5, (sh.lineW ?? 2) * 1.5) : sh.line ? 1.5 : 0} vectorEffect="non-scaling-stroke" />
              </svg>
            )}
            {content}
            {isSel && onDragStart && (
              <>
                <span className="sv-handle br" onPointerDown={(e) => { e.stopPropagation(); onDragStart(e, id, 'resize'); }} />
                <span className="sv-outline" />
              </>
            )}
          </div>
        );
      })}
    </div>
  );
}
