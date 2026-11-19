/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
export function Icon({ name }: { name: string }) {
  const codiconName = name === "undo" ? "discard" : name;
  return <span class={`codicon codicon-${codiconName}`} aria-hidden="true" />;
}
