import { useEffect, useRef, useState } from 'react';
import type { ReactNode } from 'react';
import { Icon } from './Icon';
import { onBack } from '../lib/native';

/**
 * Material-style bottom sheet used for menus, pickers and small forms —
 * replaces window.prompt/confirm so every dialog is touch friendly, themed
 * and dismissible with the Android back button.
 */
export function BottomSheet({
  open,
  onClose,
  title,
  children,
  tall,
}: {
  open: boolean;
  onClose: () => void;
  title?: string;
  children: ReactNode;
  tall?: boolean;
}) {
  const [shown, setShown] = useState(open);
  const [entering, setEntering] = useState(false);

  useEffect(() => {
    if (open) {
      setShown(true);
      requestAnimationFrame(() => setEntering(true));
    } else {
      setEntering(false);
      const t = setTimeout(() => setShown(false), 190);
      return () => clearTimeout(t);
    }
  }, [open]);

  useEffect(() => {
    if (!open) return;
    return onBack(() => {
      onClose();
      return true;
    });
  }, [open, onClose]);

  if (!shown) return null;
  return (
    <div className={`sheet-root${entering ? ' in' : ''}`} role="dialog" aria-modal="true" aria-label={title}>
      <div className="sheet-backdrop" onClick={onClose} />
      <div className={`sheet${tall ? ' tall' : ''}`}>
        <div className="sheet-grip" />
        {title && (
          <div className="sheet-head">
            <span>{title}</span>
            <button className="icon-btn" aria-label="Close" onClick={onClose}>
              <Icon name="close" size={18} />
            </button>
          </div>
        )}
        <div className="sheet-body">{children}</div>
      </div>
    </div>
  );
}

export interface SheetItem {
  icon?: string;
  label: string;
  hint?: string;
  danger?: boolean;
  disabled?: boolean;
  active?: boolean;
  onRun: () => void;
}

/** A list of tappable rows inside a BottomSheet (overflow menus). */
export function SheetMenu({ items, onClose }: { items: (SheetItem | 'divider')[]; onClose: () => void }) {
  return (
    <div className="sheet-menu">
      {items.map((it, i) =>
        it === 'divider' ? (
          <div key={i} className="sheet-divider" />
        ) : (
          <button
            key={i}
            className={`sheet-item${it.danger ? ' danger' : ''}${it.active ? ' active' : ''}`}
            disabled={it.disabled}
            onClick={() => {
              onClose();
              it.onRun();
            }}
          >
            {it.icon && <Icon name={it.icon} size={20} />}
            <span className="sheet-item-text">
              <span>{it.label}</span>
              {it.hint && <small>{it.hint}</small>}
            </span>
            {it.active && <Icon name="check" size={18} className="sheet-check" />}
          </button>
        ),
      )}
    </div>
  );
}

/** Text prompt sheet (replacement for window.prompt). */
export function PromptSheet({
  open,
  title,
  label,
  initial = '',
  placeholder,
  confirmLabel = 'OK',
  multiline,
  onSubmit,
  onClose,
  validate,
}: {
  open: boolean;
  title: string;
  label?: string;
  initial?: string;
  placeholder?: string;
  confirmLabel?: string;
  multiline?: boolean;
  onSubmit: (value: string) => void;
  onClose: () => void;
  validate?: (v: string) => string | null;
}) {
  const [v, setV] = useState(initial);
  const [err, setErr] = useState<string | null>(null);
  const ref = useRef<HTMLInputElement & HTMLTextAreaElement>(null);

  useEffect(() => {
    if (open) {
      setV(initial);
      setErr(null);
      setTimeout(() => {
        ref.current?.focus();
        ref.current?.select?.();
      }, 220);
    }
  }, [open, initial]);

  const submit = () => {
    const e = validate?.(v) ?? null;
    if (e) {
      setErr(e);
      return;
    }
    onSubmit(v);
    onClose();
  };

  return (
    <BottomSheet open={open} onClose={onClose} title={title}>
      <label className="field" style={{ margin: '4px 0 0' }}>
        {label && <span>{label}</span>}
        {multiline ? (
          <textarea ref={ref} className="input" rows={4} value={v} placeholder={placeholder} onChange={(e) => setV(e.target.value)} />
        ) : (
          <input
            ref={ref}
            className="input"
            value={v}
            placeholder={placeholder}
            onChange={(e) => setV(e.target.value)}
            onKeyDown={(e) => e.key === 'Enter' && submit()}
          />
        )}
      </label>
      {err && <p className="err">{err}</p>}
      <div className="btn-row end">
        <button className="btn" onClick={onClose}>
          Cancel
        </button>
        <button className="btn primary" onClick={submit}>
          {confirmLabel}
        </button>
      </div>
    </BottomSheet>
  );
}

/** Confirmation sheet (replacement for window.confirm). */
export function ConfirmSheet({
  open,
  title,
  message,
  confirmLabel = 'Delete',
  danger = true,
  onConfirm,
  onClose,
}: {
  open: boolean;
  title: string;
  message?: string;
  confirmLabel?: string;
  danger?: boolean;
  onConfirm: () => void;
  onClose: () => void;
}) {
  return (
    <BottomSheet open={open} onClose={onClose} title={title}>
      {message && <p className="hint" style={{ margin: '2px 0 6px' }}>{message}</p>}
      <div className="btn-row end">
        <button className="btn" onClick={onClose}>
          Cancel
        </button>
        <button
          className={`btn ${danger ? 'danger-solid' : 'primary'}`}
          onClick={() => {
            onConfirm();
            onClose();
          }}
        >
          {confirmLabel}
        </button>
      </div>
    </BottomSheet>
  );
}

/** Small transient status message. */
export function Toast({ msg }: { msg: string }) {
  if (!msg) return null;
  return (
    <div className="toast" role="status">
      {msg}
    </div>
  );
}

/** Hook: toast state + flash helper shared by all screens. */
export function useToast(): [string, (m: string, ms?: number) => void] {
  const [msg, setMsg] = useState('');
  const timer = useRef<ReturnType<typeof setTimeout> | null>(null);
  const flash = (m: string, ms = 2600) => {
    setMsg(m);
    if (timer.current) clearTimeout(timer.current);
    timer.current = setTimeout(() => setMsg(''), ms);
  };
  return [msg, flash];
}
