import type { ReactNode } from 'react';
import { Icon } from './Icon';

/**
 * Shared Office-style ribbon building blocks.
 * A ribbon row holds labeled groups; each group stacks its buttons in a row
 * with a small caption underneath — the way Word/Excel mobile trays work.
 */

export interface RBtnDef {
  icon: string;
  label?: string;
  onRun: () => void;
  active?: boolean;
  disabled?: boolean;
  /** keeps text selection alive (contentEditable tools) */
  keepFocus?: boolean;
  /** small color bar under the icon (current font/fill color) */
  colorBar?: string;
}

export function RGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="rgroup">
      <div className="rgroup-row">{children}</div>
      <div className="rgroup-label">{label}</div>
    </div>
  );
}

export function RBtn({ icon, label, onRun, active, disabled, keepFocus, colorBar }: RBtnDef) {
  // preventDefault on pointerdown/mousedown stops the button from stealing
  // focus — on Android the contentEditable would otherwise blur and lose the
  // selection before the command runs.
  const guard = keepFocus
    ? (e: React.SyntheticEvent) => {
        e.preventDefault();
      }
    : undefined;
  return (
    <button
      className={`rbtn${active ? ' active' : ''}${label ? ' labeled' : ''}`}
      title={label}
      aria-label={label}
      aria-pressed={active}
      disabled={disabled}
      onPointerDown={guard}
      onMouseDown={guard}
      onClick={onRun}
    >
      <Icon name={icon} size={19} />
      {colorBar ? <span className="color-bar" style={{ background: colorBar }} /> : null}
      {label ? <span className="rbtn-label">{label}</span> : null}
    </button>
  );
}

export function RWide({ icon, label, onRun, active, disabled }: {
  icon: string; label: string; onRun: () => void; active?: boolean; disabled?: boolean;
}) {
  return (
    <button className={`rbtn wide${active ? ' active' : ''}`} disabled={disabled} onClick={onRun}>
      <Icon name={icon} size={18} />
      <span>{label}</span>
    </button>
  );
}

export function RSelect({ value, options, onChange, width, title }: {
  value: string;
  options: { v: string; t: string }[];
  onChange: (v: string) => void;
  width?: number;
  title?: string;
}) {
  return (
    <select
      className="rselect"
      style={width ? { width } : undefined}
      title={title}
      value={value}
      onChange={(e) => onChange(e.target.value)}
    >
      {options.map((o) => (
        <option key={o.v} value={o.v}>
          {o.t}
        </option>
      ))}
    </select>
  );
}

/** Office-standard palette used for font color, highlight, cell fill, backgrounds. */
export const PALETTE = [
  '#000000', '#404040', '#8B0000', '#C00000', '#FFC000', '#FFFF00',
  '#92D050', '#00B050', '#00B0F0', '#0070C0', '#1F4E79', '#7030A0',
  '#FFFFFF', '#F3F3F3', '#FBE2D5', '#E2F0D9', '#DEEBF7', '#FFE699',
];

export function Palette({ onPick, auto }: { onPick: (c: string) => void; auto?: () => void }) {
  const guard = (e: React.SyntheticEvent) => e.preventDefault();
  return (
    <div className="palette-row">
      {auto && (
        <button className="swatch auto" aria-label="Automatic / no color" onPointerDown={guard} onMouseDown={guard} onClick={auto}>
          <Icon name="close" size={12} />
        </button>
      )}
      {PALETTE.map((c) => (
        <button
          key={c}
          className="swatch"
          style={{ background: c }}
          aria-label={`Color ${c}`}
          onPointerDown={guard}
          onMouseDown={guard}
          onClick={() => onPick(c)}
        />
      ))}
    </div>
  );
}
