import React from "react";
import ReactDOM from "react-dom/client";
import App, { THEME_STORAGE_KEY } from "./App";

const applyInitialThemePreference = () => {
  if (typeof document === "undefined") {
    return;
  }

  let initialTheme: "light" | "dark" = "dark";

  if (typeof window !== "undefined") {
    const stored = window.localStorage.getItem(THEME_STORAGE_KEY);
    if (stored === "light" || stored === "dark") {
      initialTheme = stored;
    }
  }

  document.documentElement.dataset.theme = initialTheme;
  if (document.body) {
    document.body.dataset.theme = initialTheme;
  }
};

applyInitialThemePreference();

ReactDOM.createRoot(document.getElementById("root") as HTMLElement).render(
  <React.StrictMode>
    <App />
  </React.StrictMode>,
);
