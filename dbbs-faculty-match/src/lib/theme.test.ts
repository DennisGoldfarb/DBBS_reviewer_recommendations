import { describe, expect, it } from "vitest";
import {
  DEFAULT_THEME,
  getStoredThemePreference,
  THEME_STORAGE_KEY,
} from "./theme";

describe("dark default", () => {
  it("starts dark for a new installation and migrates the previous light default", () => {
    expect(DEFAULT_THEME).toBe("dark");
    expect(getStoredThemePreference()).toBe("dark");
    window.localStorage.setItem("washu-theme", "light");
    expect(getStoredThemePreference()).toBe("dark");
  });

  it("preserves an explicit light-mode choice after the default changes", () => {
    window.localStorage.setItem(THEME_STORAGE_KEY, "light");
    expect(getStoredThemePreference()).toBe("light");
  });
});
