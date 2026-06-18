import { Theme, definePreset } from "@primeuix/styled";
import Aura from "@primeuix/themes/aura";
import { ref, type Ref } from "vue";

import type { MindooDBAppHostTheme, MindooDBAppThemeMode } from "mindoodb-app-sdk";

const MINDOO_PRIMARY_SCALE = {
  50: "#eef2ff",
  100: "#dce5ff",
  200: "#bfceff",
  300: "#98adff",
  400: "#6f88f6",
  500: "#4d67d9",
  600: "#354fb9",
  700: "#253c92",
  800: "#1f3a8a",
  900: "#182a63",
  950: "#0b1020",
} as const;

const MINDOO_THEME_PRESET = definePreset(Aura, {
  semantic: {
    primary: MINDOO_PRIMARY_SCALE,
    focusRing: {
      width: "2px",
      style: "solid",
      color: "rgba(212, 160, 23, 0.38)",
      offset: "2px",
      shadow: "none",
    },
  },
});

export const DEFAULT_THEME_PRESET = "mindoo";
export const DEFAULT_THEME_MODE: MindooDBAppThemeMode = "dark";

/**
 * Reactive mirror of the host theme mode. Updated by `applyAppTheme` whenever the
 * Haven host pushes a `theme-changed` event, so components (e.g. the DocxEditor's
 * `colorMode`) can stay in sync with light/dark mode without re-reading the DOM.
 */
export const themeMode: Ref<MindooDBAppThemeMode> = ref(DEFAULT_THEME_MODE);

export function buildPrimeVueTheme() {
  return {
    preset: MINDOO_THEME_PRESET,
    options: {
      darkModeSelector: "[data-theme='dark']",
    },
  };
}

export function applyAppTheme(theme?: Partial<MindooDBAppHostTheme> | null) {
  const mode = theme?.mode === "light" ? "light" : DEFAULT_THEME_MODE;
  themeMode.value = mode;
  if (typeof document !== "undefined") {
    document.documentElement.dataset.theme = mode;
    document.documentElement.dataset.themePreset = DEFAULT_THEME_PRESET;
  }
  Theme.setTheme(buildPrimeVueTheme());
  return { mode, preset: DEFAULT_THEME_PRESET };
}
