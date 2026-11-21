/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { EditorDescriptor } from "./types";
import { colorStyle, systemColorTable } from "./appearance";
import { useEffect, useId, useRef, useState } from "preact/hooks";
import { Icon } from "../Icon";

// Presentation shortcuts, not a validation schema. Arbitrary RGB and OLE values
// remain editable below; system choices always come from the LSP descriptor.
// Columns run from light to dark, rows by hue, as in the VBE palette.
const palette = [
  "#FFFFFF",
  "#FFC0C0",
  "#FFE0C0",
  "#FFFFC0",
  "#C0FFC0",
  "#C0FFFF",
  "#C0C0FF",
  "#FFC0FF",
  "#E0E0E0",
  "#FF8080",
  "#FFC080",
  "#FFFF80",
  "#80FF80",
  "#80FFFF",
  "#8080FF",
  "#FF80FF",
  "#C0C0C0",
  "#FF0000",
  "#FF8000",
  "#FFFF00",
  "#00FF00",
  "#00FFFF",
  "#0000FF",
  "#FF00FF",
  "#808080",
  "#C00000",
  "#C06000",
  "#C0C000",
  "#00C000",
  "#00C0C0",
  "#0000C0",
  "#C000C0",
  "#404040",
  "#800000",
  "#804000",
  "#808000",
  "#008000",
  "#008080",
  "#000080",
  "#800080",
  "#000000",
  "#400000",
  "#402000",
  "#404000",
  "#004000",
  "#004040",
  "#000040",
  "#400040",
];
const COLUMNS = 8;

type Tab = "palette" | "system" | "ole";
interface Props {
  descriptor: EditorDescriptor;
  value: Record<string, unknown>;
  disabled: boolean;
  unset: string;
  paletteLabel: string;
  systemLabel: string;
  onChange: (value: Record<string, unknown>, valid?: boolean) => void;
}

/** Black or white, whichever reads on `hex`. */
function contrastOn(hex: string | undefined): string {
  if (!hex || !/^#[\da-f]{6}$/i.test(hex)) return "inherit";
  const n = Number.parseInt(hex.slice(1), 16);
  const [r, g, b] = [(n >> 16) & 255, (n >> 8) & 255, n & 255];
  return 0.299 * r + 0.587 * g + 0.114 * b > 150 ? "#000" : "#fff";
}

/** A color's name for people: the system color's Japanese name, or its value. */
function colorLabel(value: Record<string, unknown>, unset: string): string {
  if (value.kind === "system" && typeof value.value === "string")
    return systemColorTable[value.value]?.label ?? value.value;
  return typeof value.value === "string" && value.value !== "" ? value.value.toUpperCase() : unset;
}

function Swatch({
  value,
  unset,
  caption,
}: {
  value: Record<string, unknown>;
  unset: string;
  caption: string;
}) {
  const css = colorStyle(value);
  const technical =
    value.kind === "system" && typeof value.value === "string"
      ? `${value.value} · ${css ?? ""}`
      : value.kind === "ole"
        ? "OLE"
        : "RGB";
  return (
    <div class="color-compare-item">
      <span class="color-compare-caption">{caption}</span>
      <span
        class={`color-compare-chip ${css ? "" : "unknown"}`}
        style={{ backgroundColor: css, color: contrastOn(css) }}
      >
        Aa
      </span>
      <span class="color-compare-name">{colorLabel(value, unset)}</span>
      <span class="color-compare-detail">{value.kind === undefined ? "" : technical}</span>
    </div>
  );
}

export function ColorEditor({
  descriptor,
  value,
  disabled,
  unset,
  paletteLabel,
  systemLabel,
  onChange,
}: Props) {
  const id = useId();
  const [original] = useState(value);
  const kindDescriptor = descriptor.children?.find((child) => child.field === "kind");
  const kinds =
    kindDescriptor?.allowedValues ??
    kindDescriptor?.suggestions?.map((choice) => String(choice.value)) ??
    [];
  const [tab, setTab] = useState<Tab>(
    value.kind === "system" ? "system" : value.kind === "ole" ? "ole" : "palette",
  );
  const [focusIndex, setFocusIndex] = useState(() =>
    Math.max(0, palette.indexOf(String(value.value).toUpperCase())),
  );
  const [filter, setFilter] = useState("");
  const systemList = useRef<HTMLDivElement>(null);
  useEffect(() => {
    if (tab === "system")
      systemList.current
        ?.querySelector<HTMLElement>('[aria-selected="true"]')
        ?.scrollIntoView({ block: "nearest" });
  }, [tab]);
  const valueDescriptor = descriptor.children?.find((child) => child.field === "value");
  const systemDescriptor = descriptor.children?.find((child) => child.field === "system");
  const systemChoices =
    systemDescriptor?.suggestions ??
    valueDescriptor?.suggestions ??
    (systemDescriptor?.allowedValues ?? valueDescriptor?.allowedValues ?? []).map((choice) => ({
      value: choice,
      label: choice,
    }));
  const tabs = (
    [
      { key: "palette", label: paletteLabel, kind: "rgb" },
      { key: "system", label: systemLabel, kind: "system" },
      { key: "ole", label: "OLE", kind: "ole" },
    ] as const
  ).filter((item) => kinds.includes(item.kind));
  const rgbHex =
    value.kind === "rgb" && typeof value.value === "string" ? value.value.toUpperCase() : undefined;
  const needle = filter.trim().toLowerCase();
  const visibleSystem = systemChoices.filter((choice) => {
    if (!needle) return true;
    const key = String(choice.value);
    return [key, choice.label, systemColorTable[key]?.label ?? ""].some((text) =>
      String(text).toLowerCase().includes(needle),
    );
  });
  return (
    <div class="color-editor">
      <div class="segmented" role="tablist">
        {tabs.map((item) => (
          <button
            type="button"
            role="tab"
            id={`${id}-${item.key}-tab`}
            aria-selected={tab === item.key}
            aria-controls={`${id}-${item.key}-panel`}
            tabIndex={tab === item.key ? 0 : -1}
            disabled={disabled}
            onClick={() => setTab(item.key)}
            onKeyDown={(event) => {
              if (
                !["ArrowLeft", "ArrowRight", "Home", "End"].includes(event.key) ||
                event.isComposing
              )
                return;
              const index = tabs.findIndex((candidate) => candidate.key === item.key);
              const next =
                event.key === "Home"
                  ? 0
                  : event.key === "End"
                    ? tabs.length - 1
                    : (index + (event.key === "ArrowRight" ? 1 : tabs.length - 1)) % tabs.length;
              event.preventDefault();
              event.stopPropagation();
              setTab(tabs[next].key);
              event.currentTarget.parentElement
                ?.querySelectorAll<HTMLButtonElement>("button")
                [next]?.focus();
            }}
          >
            {item.label}
          </button>
        ))}
      </div>
      <div
        class="color-panel"
        role="tabpanel"
        id={`${id}-${tab}-panel`}
        aria-labelledby={`${id}-${tab}-tab`}
      >
        {tab === "palette" && (
          <>
            <div class="color-palette" role="group" aria-label={paletteLabel}>
              {palette.map((hex, index) => (
                <button
                  type="button"
                  class="color-palette-swatch"
                  style={{ backgroundColor: hex, color: contrastOn(hex) }}
                  title={hex}
                  aria-label={hex}
                  aria-pressed={rgbHex === hex}
                  tabIndex={focusIndex === index ? 0 : -1}
                  onFocus={() => setFocusIndex(index)}
                  disabled={disabled}
                  onClick={() => onChange({ kind: "rgb", value: hex })}
                  onKeyDown={(event) => {
                    const offset = {
                      ArrowLeft: -1,
                      ArrowRight: 1,
                      ArrowUp: -COLUMNS,
                      ArrowDown: COLUMNS,
                    }[event.key];
                    const next =
                      event.key === "Home"
                        ? 0
                        : event.key === "End"
                          ? palette.length - 1
                          : offset === undefined
                            ? undefined
                            : Math.max(0, Math.min(palette.length - 1, index + offset));
                    if (next === undefined || event.isComposing) return;
                    event.preventDefault();
                    event.stopPropagation();
                    event.currentTarget.parentElement
                      ?.querySelectorAll<HTMLButtonElement>("button")
                      [next]?.focus();
                  }}
                >
                  {rgbHex === hex && <Icon name="check" />}
                </button>
              ))}
            </div>
            <div class="color-custom">
              <label class="color-custom-picker" title="カスタム カラー">
                <input
                  type="color"
                  aria-label="カスタム カラー"
                  disabled={disabled}
                  value={rgbHex && /^#[\dA-F]{6}$/.test(rgbHex) ? rgbHex.toLowerCase() : "#000000"}
                  onInput={(event) =>
                    onChange({ kind: "rgb", value: event.currentTarget.value.toUpperCase() })
                  }
                />
                <Icon name="symbol-color" />
              </label>
              <label class="color-hex" for={`${id}-hex`}>
                <span>#</span>
                <input
                  id={`${id}-hex`}
                  type="text"
                  spellcheck={false}
                  maxLength={6}
                  placeholder="RRGGBB"
                  disabled={disabled}
                  value={rgbHex ? rgbHex.replace(/^#/, "") : ""}
                  onInput={(event) => {
                    // Color syntax and applicability remain server decisions.
                    const text = event.currentTarget.value.trim().replace(/^#/, "");
                    onChange({ kind: "rgb", value: `#${text.toUpperCase()}` });
                  }}
                />
              </label>
            </div>
          </>
        )}
        {tab === "system" && (
          <>
            <input
              class="color-system-filter"
              type="search"
              placeholder="システム カラーを検索"
              aria-label="システム カラーを検索"
              disabled={disabled}
              value={filter}
              onInput={(event) => setFilter(event.currentTarget.value)}
            />
            <div ref={systemList} class="color-system-list" role="listbox" aria-label={systemLabel}>
              {visibleSystem.map((choice) => {
                const key = String(choice.value);
                const entry = systemColorTable[key];
                const selected = value.kind === "system" && value.value === key;
                return (
                  <button
                    type="button"
                    role="option"
                    class="color-system-choice"
                    aria-selected={selected}
                    disabled={disabled}
                    title={entry ? `${key} (${entry.hex})` : key}
                    onClick={() => onChange({ kind: "system", value: key })}
                  >
                    <span
                      class="color-swatch large"
                      style={{ backgroundColor: entry?.hex }}
                      aria-hidden="true"
                    />
                    <span class="color-system-text">
                      <span class="color-system-name">{entry?.label ?? choice.label}</span>
                      <span class="color-system-id">
                        {key}
                        {entry && ` · ${entry.hex}`}
                      </span>
                    </span>
                    {selected && <Icon name="check" />}
                  </button>
                );
              })}
              {visibleSystem.length === 0 && (
                <p class="notice">該当するシステム カラーはありません。</p>
              )}
            </div>
          </>
        )}
        {tab === "ole" && (
          <label class="property-editor-field" for={`${id}-ole`}>
            <span>OLE 値</span>
            <input
              id={`${id}-ole`}
              type="text"
              spellcheck={false}
              placeholder="0x8000000F"
              disabled={disabled}
              value={value.kind === "ole" && typeof value.value === "string" ? value.value : ""}
              onInput={(event) =>
                onChange({ kind: "ole", value: event.currentTarget.value.trim() })
              }
            />
          </label>
        )}
      </div>
      <div class="color-compare" aria-live="polite">
        <Swatch value={original} unset={unset} caption="変更前" />
        <Icon name="arrow-right" />
        <Swatch value={value} unset={unset} caption="変更後" />
      </div>
    </div>
  );
}
