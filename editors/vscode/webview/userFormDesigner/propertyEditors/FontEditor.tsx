/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import type { EditorDescriptor } from "./types";
import { fontStyle } from "./appearance";
import { useEffect, useId, useRef, useState } from "preact/hooks";
import { validNumericInput } from "./numeric";

// Suggestions only: any name may be typed. Each is drawn in itself, so a font
// the machine lacks shows in the fallback face.
const families: { group: string; names: string[] }[] = [
  {
    group: "日本語",
    names: [
      "MS UI Gothic",
      "MS Gothic",
      "MS PGothic",
      "MS Mincho",
      "MS PMincho",
      "Meiryo",
      "Meiryo UI",
      "Yu Gothic",
      "Yu Gothic UI",
      "Yu Mincho",
      "BIZ UDGothic",
      "BIZ UDPGothic",
    ],
  },
  {
    group: "欧文",
    names: [
      "Segoe UI",
      "Tahoma",
      "Arial",
      "Calibri",
      "Verdana",
      "Times New Roman",
      "Georgia",
      "Consolas",
      "Courier New",
    ],
  },
];
const sizes = [8, 9, 10, 11, 12, 14, 16, 18, 20, 24, 28, 36, 48];
// Windows character sets (LOGFONT lfCharSet), the "文字セット" of the font dialog.
const charsets: [number, string][] = [
  [0, "欧文 (ANSI)"],
  [1, "既定 (DEFAULT)"],
  [2, "記号 (SYMBOL)"],
  [128, "日本語 (SHIFTJIS)"],
  [129, "韓国語 (HANGUL)"],
  [134, "簡体中国語 (GB2312)"],
  [136, "繁体中国語 (BIG5)"],
  [161, "ギリシャ語"],
  [162, "トルコ語"],
  [163, "ベトナム語"],
  [177, "ヘブライ語"],
  [178, "アラビア語"],
  [186, "バルト言語"],
  [204, "キリル (RUSSIAN)"],
  [222, "タイ語"],
  [238, "中央ヨーロッパ"],
  [255, "OEM"],
];
const styleToggles: { field: string; label: string; glyph: string; css: string }[] = [
  { field: "bold", label: "太字", glyph: "B", css: "font-weight:700" },
  { field: "italic", label: "斜体", glyph: "I", css: "font-style:italic" },
  { field: "underline", label: "下線", glyph: "U", css: "text-decoration:underline" },
  { field: "strikeThrough", label: "取り消し線", glyph: "S", css: "text-decoration:line-through" },
];
const known = new Set([
  "name",
  "size",
  "charset",
  "pitchAndFamily",
  ...styleToggles.map((t) => t.field),
]);

interface Props {
  descriptor: EditorDescriptor;
  value: Record<string, unknown>;
  disabled: boolean;
  unset: string;
  sample: string;
  onChange: (field: string, value: unknown, valid?: boolean) => void;
}
export function FontEditor({ descriptor, value, disabled, unset, sample, onChange }: Props) {
  const id = useId();
  const children = descriptor.children ?? [];
  const child = (field: string) => children.find((item) => item.field === field);
  const name = typeof value.name === "string" ? value.name : "";
  const [filter, setFilter] = useState("");
  const [sizeText, setSizeText] = useState(() =>
    typeof value.size === "number" || typeof value.size === "string" ? String(value.size) : "",
  );
  const list = useRef<HTMLDivElement>(null);
  useEffect(() => {
    list.current
      ?.querySelector<HTMLElement>('[aria-selected="true"]')
      ?.scrollIntoView({ block: "center" });
  }, []);
  const needle = filter.trim().toLowerCase();
  const listed = families.flatMap((family) => family.names);
  const groups = [
    ...(name && !listed.includes(name) ? [{ group: "現在", names: [name] }] : []),
    ...families,
  ]
    .map((family) => ({
      ...family,
      names: family.names.filter((item) => !needle || item.toLowerCase().includes(needle)),
    }))
    .filter((family) => family.names.length > 0);
  const numeric = (field: string, text: string) => {
    const descriptor = child(field);
    if (!descriptor) return;
    const valid = validNumericInput(text, descriptor);
    onChange(field, valid ? Number(text) : text, valid);
  };
  const sizeDescriptor = child("size");
  const charsetDescriptor = child("charset");
  const pitchDescriptor = child("pitchAndFamily");
  const extra = children.filter((item) => !known.has(item.field));
  const charset = typeof value.charset === "number" ? value.charset : undefined;
  return (
    <div class="font-editor">
      <div class="font-editor-main">
        {child("name") && (
          <div class="font-family">
            <input
              id={`${id}-name`}
              class="font-family-input"
              type="text"
              aria-label="フォント名"
              placeholder={value.name === null ? "null" : unset}
              disabled={disabled}
              spellcheck={false}
              value={name}
              onInput={(event) => {
                onChange("name", event.currentTarget.value);
                setFilter(event.currentTarget.value);
              }}
              onFocus={(event) => event.currentTarget.select()}
            />
            <div ref={list} class="font-family-list" role="listbox" aria-label="フォント名">
              {groups.map((family) => (
                <div role="group" aria-label={family.group}>
                  <div class="font-family-group">{family.group}</div>
                  {family.names.map((item) => (
                    <button
                      type="button"
                      role="option"
                      class="font-family-option"
                      aria-selected={item === name}
                      disabled={disabled}
                      style={{ fontFamily: `${JSON.stringify(item)}, sans-serif` }}
                      onClick={() => {
                        onChange("name", item);
                        setFilter("");
                      }}
                    >
                      {item}
                    </button>
                  ))}
                </div>
              ))}
              {groups.length === 0 && (
                <p class="notice font-family-empty">一覧にないフォント名をそのまま使います。</p>
              )}
            </div>
          </div>
        )}
        <div class="font-side">
          {sizeDescriptor && (
            <div class="font-section">
              <label class="font-section-title" for={`${id}-size`}>
                サイズ
              </label>
              <div class="font-size-input">
                <input
                  id={`${id}-size`}
                  type="text"
                  inputMode="decimal"
                  disabled={disabled}
                  placeholder={value.size === null ? "null" : unset}
                  value={sizeText}
                  onInput={(event) => {
                    setSizeText(event.currentTarget.value);
                    numeric("size", event.currentTarget.value);
                  }}
                />
                <span>pt</span>
              </div>
              <div class="font-size-chips">
                {sizes.map((size) => (
                  <button
                    type="button"
                    class="chip"
                    aria-pressed={value.size === size}
                    disabled={disabled}
                    onClick={() => {
                      setSizeText(String(size));
                      numeric("size", String(size));
                    }}
                  >
                    {size}
                  </button>
                ))}
              </div>
            </div>
          )}
          {styleToggles.some((toggle) => child(toggle.field)) && (
            <div class="font-section">
              <span class="font-section-title">スタイル</span>
              <div class="font-style-toggles" role="group" aria-label="スタイル">
                {styleToggles
                  .filter((toggle) => child(toggle.field))
                  .map((toggle) => (
                    <button
                      type="button"
                      class="font-style-toggle"
                      aria-pressed={value[toggle.field] === true}
                      aria-label={toggle.label}
                      title={toggle.label}
                      disabled={disabled}
                      onClick={() => onChange(toggle.field, value[toggle.field] !== true)}
                    >
                      <span style={toggle.css}>{toggle.glyph}</span>
                    </button>
                  ))}
              </div>
            </div>
          )}
        </div>
      </div>
      <div class="font-editor-sample">
        <div
          class="font-editor-sample-text"
          style={fontStyle(value)}
          contentEditable={!disabled}
          spellcheck={false}
          data-handles-enter
          aria-label="サンプル（編集できます）"
        >
          {sample}
        </div>
        <span class="font-editor-sample-meta">
          {[name || unset, typeof value.size === "number" ? `${value.size} pt` : undefined]
            .filter(Boolean)
            .join(" · ")}
        </span>
      </div>
      {(charsetDescriptor || pitchDescriptor || extra.length > 0) && (
        <details class="font-advanced">
          <summary>詳細設定</summary>
          <div class="property-editor-fields">
            {charsetDescriptor && (
              <label class="property-editor-field" for={`${id}-charset`}>
                <span>文字セット</span>
                <select
                  id={`${id}-charset`}
                  disabled={disabled}
                  value={charset === undefined ? "__unset__" : String(charset)}
                  onChange={(event) => onChange("charset", Number(event.currentTarget.value))}
                >
                  {charset === undefined && (
                    <option value="__unset__" disabled>
                      {value.charset === null ? "null" : unset}
                    </option>
                  )}
                  {charset !== undefined && !charsets.some(([n]) => n === charset) && (
                    <option value={String(charset)}>{charset}</option>
                  )}
                  {charsets.map(([n, label]) => (
                    <option value={String(n)}>
                      {label} — {n}
                    </option>
                  ))}
                </select>
              </label>
            )}
            {pitchDescriptor && (
              <label class="property-editor-field" for={`${id}-pitch`}>
                <span>ピッチとファミリ</span>
                <input
                  id={`${id}-pitch`}
                  type="text"
                  inputMode="numeric"
                  disabled={disabled}
                  placeholder={value.pitchAndFamily === null ? "null" : unset}
                  defaultValue={
                    typeof value.pitchAndFamily === "number" ? String(value.pitchAndFamily) : ""
                  }
                  onInput={(event) => numeric("pitchAndFamily", event.currentTarget.value)}
                />
              </label>
            )}
            {extra.map((item) => (
              <label class="property-editor-field" for={`${id}-${item.field}`}>
                <span>{item.field}</span>
                <input
                  id={`${id}-${item.field}`}
                  type="text"
                  disabled={disabled}
                  defaultValue={
                    typeof value[item.field] === "string" || typeof value[item.field] === "number"
                      ? String(value[item.field])
                      : ""
                  }
                  placeholder={unset}
                  onInput={(event) =>
                    item.valueType === "number" || item.valueType === "integer"
                      ? numeric(item.field, event.currentTarget.value)
                      : onChange(item.field, event.currentTarget.value)
                  }
                />
              </label>
            ))}
          </div>
          <p class="notice font-advanced-note">
            文字セットは VBE のフォント ダイアログの「文字セット」と同じ Font.Charset
            です。通常はフォントに合わせて自動で決まるため、変更する必要はありません。
          </p>
        </details>
      )}
    </div>
  );
}
