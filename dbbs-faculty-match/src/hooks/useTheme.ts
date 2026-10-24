import { useEffect, useState } from "react";
import { isTauri } from "@tauri-apps/api/core";
import { getCurrentWindow } from "@tauri-apps/api/window";
import { applyTheme, getStoredThemePreference } from "../lib/theme";

export function useTheme() {
  const [theme, setTheme] = useState(getStoredThemePreference);
  useEffect(() => {
    applyTheme(theme);
    if (isTauri())
      void getCurrentWindow()
        .setTheme(theme)
        .catch(() => {});
  }, [theme]);
  const toggleTheme = () =>
    setTheme((current) => (current === "light" ? "dark" : "light"));
  const themeToggleLabel =
    theme === "light"
      ? "Switch to dark WashU theme"
      : "Switch to light WashU theme";
  return { theme, toggleTheme, themeToggleLabel };
}
