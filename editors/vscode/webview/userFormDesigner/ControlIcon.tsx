import type { ComponentChildren } from "preact";
import type { ToolboxType } from "../../src/userForm/toolbox";

// Inline 16x16 SVG glyphs that look like the control they create. They use
// currentColor so they follow the theme and the toolbox selection colors.
// "fill" marks the shapes that are painted solid instead of outlined.
const glyphs: Record<ToolboxType | "Pointer" | "Page", ComponentChildren> = {
  Pointer: <path d="M4 2.5v10l2.6-2.4 1.8 4 1.7-.8-1.8-3.9H12z" stroke-linejoin="round" />,
  Label: (
    <>
      <path d="M3 13 7 3h1.2l4 10" />
      <path d="M4.6 9.3h5.9" />
    </>
  ),
  TextBox: (
    <>
      <rect x="1.5" y="4" width="13" height="8" rx="1.5" />
      <path d="M4.5 6v4M3.7 6h1.6M3.7 10h1.6" />
    </>
  ),
  ComboBox: (
    <>
      <rect x="1.5" y="4" width="13" height="8" rx="1.5" />
      <path d="M10 4v8" />
      <path d="m11.2 7.2 1.05 1.1 1.05-1.1" />
      <path d="M3.8 8h4" />
    </>
  ),
  ListBox: (
    <>
      <rect x="2" y="1.5" width="12" height="13" rx="1.5" />
      <rect class="fill" x="3.5" y="6.5" width="9" height="3" rx=".5" />
      <path d="M4.5 4.5h7M4.5 12h7" />
    </>
  ),
  CommandButton: (
    <>
      <rect x="1.5" y="3" width="13" height="8" rx="2" />
      <path d="M4.5 7h5" />
      <path class="fill" d="M10 8.5v6l1.5-1.4 1 2.2 1-.5-1-2.1h2z" stroke-width=".8" />
    </>
  ),
  CheckBox: (
    <>
      <rect x="2.5" y="2.5" width="11" height="11" rx="2" />
      <path d="m5.2 8.2 1.9 1.9 3.8-4.2" />
    </>
  ),
  OptionButton: (
    <>
      <circle cx="8" cy="8" r="5.5" />
      <circle class="fill" cx="8" cy="8" r="2.3" />
    </>
  ),
  ToggleButton: (
    <>
      <rect x="1.5" y="4.5" width="13" height="7" rx="3.5" />
      <circle class="fill" cx="11" cy="8" r="2" />
    </>
  ),
  SpinButton: (
    <>
      <rect x="4.5" y="1.5" width="7" height="13" rx="1.5" />
      <path d="M4.5 8h7" />
      <path d="m6.6 5.6 1.4-1.4 1.4 1.4M6.6 10.4 8 11.8l1.4-1.4" />
    </>
  ),
  ScrollBar: (
    <>
      <rect x="1.5" y="5" width="13" height="6" rx="1.5" />
      <path d="M3.9 6.9 2.9 8l1 1.1M12.1 6.9l1 1.1-1 1.1" />
      <rect class="fill" x="6" y="6.5" width="3.5" height="3" rx=".75" />
    </>
  ),
  Image: (
    <>
      <rect x="1.5" y="2.5" width="13" height="11" rx="1.5" />
      <circle cx="5.3" cy="6" r="1.2" />
      <path d="m1.8 12.2 4-3.6 2.7 2.4 2-1.7 3.7 3" />
    </>
  ),
  Frame: (
    <>
      <path d="M4 3.5H2.5a1 1 0 0 0-1 1v8a1 1 0 0 0 1 1h11a1 1 0 0 0 1-1v-8a1 1 0 0 0-1-1h-3" />
      <path d="M5.6 3.5h3.3" />
    </>
  ),
  MultiPage: (
    <>
      <path d="M1.5 5v8a1 1 0 0 0 1 1h11a1 1 0 0 0 1-1V6a1 1 0 0 0-1-1H6.5V3a1 1 0 0 0-1-1h-3a1 1 0 0 0-1 1z" />
      <path d="M6.5 3.5h3a1 1 0 0 1 1 1V5" />
      <path d="M4 8.5h8M4 11h5" />
    </>
  ),
  TabStrip: (
    <>
      <path d="M.8 11.5h.7V5.5a1 1 0 0 1 1-1h3a1 1 0 0 1 1 1v6h8.7" />
      <path d="M6.5 6.5h2.5a1 1 0 0 1 1 1v4M10 6.5h2.5a1 1 0 0 1 1 1v4" />
    </>
  ),
  ProgressBar: (
    <>
      <path d="M2 5.5h5M11 5.5h3" />
      <rect x="1.5" y="9" width="13" height="3" rx="1.5" />
      <rect class="fill" x="1.5" y="9" width="8" height="3" rx="1.5" />
    </>
  ),
  FilePicker: (
    <>
      <rect x="1" y="4.5" width="14" height="7" rx="1" />
      <path d="M10.5 4.5v7" />
      <path class="fill" d="M12 8h.01M13.6 8h.01" stroke-width="1.4" stroke-linecap="round" />
      <path d="M3 8h5" />
    </>
  ),
  Page: (
    <>
      <path d="M3.5 1.5h6l3 3v10h-9z" />
      <path d="M9.5 1.5v3h3" />
    </>
  ),
};

export type ControlIconName = keyof typeof glyphs;

export function isControlIconName(name: string): name is ControlIconName {
  return Object.prototype.hasOwnProperty.call(glyphs, name);
}

export function ControlIcon({ name }: { name: ControlIconName }) {
  return (
    <svg
      class="control-icon"
      viewBox="0 0 16 16"
      width="16"
      height="16"
      fill="none"
      stroke="currentColor"
      stroke-width="1.2"
      stroke-linecap="round"
      stroke-linejoin="round"
      aria-hidden="true"
    >
      {glyphs[name]}
    </svg>
  );
}
