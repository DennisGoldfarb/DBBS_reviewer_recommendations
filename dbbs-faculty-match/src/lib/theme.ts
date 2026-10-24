import type { ThemePreference } from "../types";

export const DEFAULT_THEME: ThemePreference = "dark";
// Apply the dark default on upgrade; subsequent explicit choices remain saved.
export const THEME_STORAGE_KEY = "washu-theme-v2";

export function getStoredThemePreference(): ThemePreference {
  try {
    const stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    return stored === "light" ? "light" : DEFAULT_THEME;
  } catch {
    return DEFAULT_THEME;
  }
}

export function applyTheme(theme: ThemePreference): void {
  document.documentElement.dataset.theme = theme;
  if (document.body) document.body.dataset.theme = theme;
  try {
    window.localStorage.setItem(THEME_STORAGE_KEY, theme);
  } catch {
    /* Storage can be unavailable in restricted webviews. */
  }
}
