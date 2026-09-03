/**
 * Thin wrappers around Capacitor plugins that no-op on the web so the same
 * code runs in the browser preview and inside the Android shell.
 */
import { Capacitor } from '@capacitor/core';
import { getPrefs } from './storage';

export const isNative = (): boolean => Capacitor.isNativePlatform();

/** Light tap feedback (respects the haptics preference). */
export async function tap(style: 'light' | 'medium' | 'heavy' = 'light'): Promise<void> {
  if (!getPrefs().haptics) return;
  try {
    if (isNative()) {
      const { Haptics, ImpactStyle } = await import('@capacitor/haptics');
      const map = { light: ImpactStyle.Light, medium: ImpactStyle.Medium, heavy: ImpactStyle.Heavy };
      await Haptics.impact({ style: map[style] });
    } else {
      navigator.vibrate?.(style === 'light' ? 10 : style === 'medium' ? 20 : 35);
    }
  } catch {
    /* unsupported */
  }
}

export async function selectionTick(): Promise<void> {
  if (!getPrefs().haptics) return;
  try {
    if (isNative()) {
      const { Haptics } = await import('@capacitor/haptics');
      await Haptics.selectionChanged();
    } else navigator.vibrate?.(6);
  } catch {
    /* unsupported */
  }
}

/** Color the Android status bar to match the current app bar. */
export async function setStatusBar(color: string, lightText: boolean): Promise<void> {
  if (!isNative()) return;
  try {
    const { StatusBar, Style } = await import('@capacitor/status-bar');
    await StatusBar.setBackgroundColor({ color });
    await StatusBar.setStyle({ style: lightText ? Style.Dark : Style.Light });
  } catch {
    /* plugin missing */
  }
}

type BackHandler = () => boolean; // return true when handled

const backStack: BackHandler[] = [];
let wired = false;

/**
 * Register a hardware back-button handler. Handlers are consulted newest
 * first; the first one that returns true consumes the event. When nothing
 * consumes it on Android the app moves to the background (never force-quits
 * mid-edit). Returns an unsubscribe function.
 */
export function onBack(handler: BackHandler): () => void {
  backStack.push(handler);
  void wireBack();
  return () => {
    const i = backStack.lastIndexOf(handler);
    if (i >= 0) backStack.splice(i, 1);
  };
}

async function wireBack(): Promise<void> {
  if (wired) return;
  wired = true;
  const dispatch = (): boolean => {
    for (let i = backStack.length - 1; i >= 0; i--) {
      try {
        if (backStack[i]()) return true;
      } catch {
        /* ignore handler errors */
      }
    }
    return false;
  };
  if (isNative()) {
    try {
      const { App } = await import('@capacitor/app');
      await App.addListener('backButton', () => {
        if (!dispatch()) void App.minimizeApp();
      });
    } catch {
      /* plugin missing */
    }
  } else {
    // Browser: Escape key mirrors back for desktop testing.
    window.addEventListener('keydown', (e) => {
      if (e.key === 'Escape') dispatch();
    });
  }
}

/** Ask the OS to keep the screen awake while presenting (Wake Lock API). */
export async function keepAwake(on: boolean): Promise<void> {
  try {
    const nav = navigator as Navigator & { wakeLock?: { request: (t: 'screen') => Promise<{ release: () => Promise<void> }> } };
    if (!nav.wakeLock) return;
    if (on) {
      lock = await nav.wakeLock.request('screen');
    } else {
      await lock?.release();
      lock = null;
    }
  } catch {
    /* not permitted */
  }
}
let lock: { release: () => Promise<void> } | null = null;

/** Copy text to the clipboard with a fallback for older WebViews. */
export async function copyText(text: string): Promise<boolean> {
  try {
    await navigator.clipboard.writeText(text);
    return true;
  } catch {
    try {
      const ta = document.createElement('textarea');
      ta.value = text;
      ta.style.position = 'fixed';
      ta.style.opacity = '0';
      document.body.appendChild(ta);
      ta.select();
      const ok = document.execCommand('copy');
      ta.remove();
      return ok;
    } catch {
      return false;
    }
  }
}

/** Share plain text via the system share sheet (Android) or Web Share API. */
export async function shareText(title: string, text: string): Promise<string> {
  try {
    if (isNative()) {
      const { Share } = await import('@capacitor/share');
      await Share.share({ title, text, dialogTitle: title });
      return 'Share sheet opened';
    }
    if (navigator.share) {
      await navigator.share({ title, text });
      return 'Shared';
    }
  } catch {
    /* cancelled */
  }
  return (await copyText(text)) ? 'Copied to clipboard' : 'Sharing is not available here';
}
