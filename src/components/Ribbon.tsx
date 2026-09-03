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
  /** shows a tiny caret: the button opens a picker */
  menu?: boolean;
}

export function RGroup({ label, children }: { label: string; children: ReactNode }) {
  return (
    <div className="rgroup">
      <div className="rgroup-row">{children}</div>
      <div className="rgroup-label">{label}</div>
    </div>
  );
}

const guardEvt = (e: React.SyntheticEvent) => e.preventDefault();

export function RBtn({ icon, label, onRun, active, disabled, keepFocus, colorBar, menu }: RBtnDef) {
  // preventDefault on pointerdown/mousedown stops the button from stealing
  // focus — on Android the contentEditable would otherwise blur and lose the
  // selection before the command runs.
  const guard = keepFocus ? guardEvt : undefined;
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
      {label ? (
        <span className="rbtn-label">
          {label}
          {menu ? <span className="rbtn-caret">▾</span> : null}
        </span>
      ) : null}
    </button>
  );
}

export function RWide({ icon, label, onRun, active, disabled, keepFocus }: {
  icon: string; label: string; onRun: () => void; active?: boolean; disabled?: boolean; keepFocus?: boolean;
}) {
  const guard = keepFocus ? guardEvt : undefined;
  return (
    <button className={`rbtn wide${active ? ' active' : ''}`} disabled={disabled} onPointerDown={guard} onMouseDown={guard} onClick={onRun}>
      <Icon name={icon} size={18} />
      <span>{label}</span>
    </button>
  );
}

export function RSelect({ value, options, onChange, width, title, keepFocus }: {
  value: string;
  options: { v: string; t: string }[];
  onChange: (v: string) => void;
  width?: number;
  title?: string;
  keepFocus?: boolean;
}) {
  return (
    <select
      className="rselect"
      style={width ? { width } : undefined}
      title={title}
      aria-label={title}
      value={value}
      onMouseDown={keepFocus ? (e) => e.stopPropagation() : undefined}
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

/** "- 11 +" stepper used for font size / decimals / zoom. */
export function RStepper({ value, onDec, onInc, title, keepFocus, width = 34 }: {
  value: string; onDec: () => void; onInc: () => void; title: string; keepFocus?: boolean; width?: number;
}) {
  const guard = keepFocus ? guardEvt : undefined;
  return (
    <div className="rstepper" title={title} aria-label={title}>
      <button className="rstep" aria-label={`${title}: smaller`} onPointerDown={guard} onMouseDown={guard} onClick={onDec}>
        −
      </button>
      <span className="rstep-val" style={{ minWidth: width }}>{value}</span>
      <button className="rstep" aria-label={`${title}: larger`} onPointerDown={guard} onMouseDown={guard} onClick={onInc}>
        +
      </button>
    </div>
  );
}

/** Segmented control (e.g. Portrait | Landscape). */
export function RSeg<T extends string>({ value, options, onChange, keepFocus }: {
  value: T; options: { v: T; t: string; icon?: string }[]; onChange: (v: T) => void; keepFocus?: boolean;
}) {
  const guard = keepFocus ? guardEvt : undefined;
  return (
    <div className="rseg" role="radiogroup">
      {options.map((o) => (
        <button
          key={o.v}
          role="radio"
          aria-checked={o.v === value}
          className={`rseg-opt${o.v === value ? ' on' : ''}`}
          onPointerDown={guard}
          onMouseDown={guard}
          onClick={() => onChange(o.v)}
        >
          {o.icon && <Icon name={o.icon} size={15} />}
          {o.t}
        </button>
      ))}
    </div>
  );
}

/** Office-standard palette used for font color, highlight, cell fill, backgrounds. */
export const PALETTE = [
  '#000000', '#404040', '#7F7F7F', '#BFBFBF', '#FFFFFF',
  '#C00000', '#FF0000', '#FFC000', '#FFFF00', '#92D050',
  '#00B050', '#00B0F0', '#0070C0', '#002060', '#7030A0',
  '#FBE2D5', '#FFF2CC', '#E2F0D9', '#DEEBF7', '#E9D7F1',
  '#F4B183', '#FFD966', '#A9D18E', '#9DC3E6', '#C9A0DC',
  '#843C0C', '#7F6000', '#385723', '#1F4E79', '#4B2A6B',
];

export function Palette({ onPick, auto, autoLabel = 'No color', current }: {
  onPick: (c: string) => void;
  auto?: () => void;
  autoLabel?: string;
  current?: string;
}) {
  return (
    <div className="palette-wrap">
      <div className="palette-grid">
        {auto && (
          <button className="swatch auto" aria-label={autoLabel} title={autoLabel} onPointerDown={guardEvt} onMouseDown={guardEvt} onClick={auto}>
            <Icon name="close" size={12} />
          </button>
        )}
        {PALETTE.map((c) => (
          <button
            key={c}
            className={`swatch${current && current.toLowerCase() === c.toLowerCase() ? ' on' : ''}`}
            style={{ background: c }}
            aria-label={`Color ${c}`}
            onPointerDown={guardEvt}
            onMouseDown={guardEvt}
            onClick={() => onPick(c)}
          />
        ))}
        <label className="swatch custom" title="Custom color" onPointerDown={(e) => e.stopPropagation()}>
          <Icon name="plus" size={13} />
          <input type="color" onChange={(e) => onPick(e.target.value)} />
        </label>
      </div>
    </div>
  );
}

/** Expandable gallery row that lives directly above the ribbon tabs. */
export function RibbonPanel({ title, onClose, children }: { title: string; onClose: () => void; children: ReactNode }) {
  return (
    <div className="rpanel">
      <div className="rpanel-head">
        <span>{title}</span>
        <button className="icon-btn" aria-label="Close panel" onPointerDown={guardEvt} onMouseDown={guardEvt} onClick={onClose}>
          <Icon name="chevronDown" size={18} />
        </button>
      </div>
      <div className="rpanel-body">{children}</div>
    </div>
  );
}

/** Tab strip of the ribbon; `accent` var is set by the parent editor. */
export function RibbonTabs<T extends string>({ tabs, value, onChange }: { tabs: { id: T; label: string }[]; value: T; onChange: (t: T) => void }) {
  return (
    <div className="ribbon-tabs" role="tablist">
      {tabs.map((t) => (
        <button
          key={t.id}
          role="tab"
          aria-selected={value === t.id}
          className={`ribbon-tab${value === t.id ? ' active' : ''}`}
          onPointerDown={guardEvt}
          onMouseDown={guardEvt}
          onClick={() => onChange(t.id)}
        >
          {t.label}
        </button>
      ))}
    </div>
  );
}

/** Shared editor app bar (brand-colored) used by Docs / Sheets / Slides. */
export function AppBar({
  kindIcon,
  title,
  onTitle,
  placeholder,
  onBack,
  children,
  saved,
}: {
  kindIcon: ReactNode;
  title: string;
  onTitle: (t: string) => void;
  placeholder: string;
  onBack?: () => void;
  children?: ReactNode;
  saved?: 'saved' | 'saving' | 'dirty';
}) {
  return (
    <header className="appbar">
      <button className="icon-btn light" aria-label="Back to Home" onClick={onBack}>
        <Icon name="arrowLeft" size={22} />
      </button>
      <span className="appbar-kind">{kindIcon}</span>
      <div className="appbar-titlewrap">
        <input className="appbar-title" value={title} onChange={(e) => onTitle(e.target.value)} placeholder={placeholder} aria-label="Document title" />
        {saved && <span className={`appbar-saved ${saved}`}>{saved === 'saved' ? 'Saved on device' : saved === 'saving' ? 'Saving…' : 'Unsaved changes'}</span>}
      </div>
      {children}
    </header>
  );
}
