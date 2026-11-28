/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
export function Icon({ name }: { name: string }) {
  // The chevrons are the page's own (the sheet view's): one shape, size and color for every ▾.
  if (name.startsWith("chevron-")) return <span class={`chev chev-${name.slice(8)}`} aria-hidden="true" />;
  const codiconName = name === "undo" ? "discard" : name;
  return <span class={`codicon codicon-${codiconName}`} aria-hidden="true" />;
}
