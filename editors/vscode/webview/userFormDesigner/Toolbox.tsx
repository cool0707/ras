/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import { ControlIcon } from "./ControlIcon";
import { controlMetadata } from "../../src/userForm/toolbox";
import type { ToolboxType } from "../../src/userForm/toolbox";
import type { DesignerStrings } from "../../src/userForm/protocol";

export function Toolbox({
  selected,
  disabled,
  parentEditable = false,
  pageEditable = false,
  tabEditable = false,
  strings,
  onSelect,
}: {
  selected?: ToolboxType;
  disabled: boolean;
  parentEditable?: boolean;
  pageEditable?: boolean;
  tabEditable?: boolean;
  strings: DesignerStrings;
  onSelect: (type?: ToolboxType) => void;
}) {
  return (
    <aside
      class="toolbox"
      aria-label={strings.toolbox}
      onKeyDown={(event) => {
        if (event.key === "Escape") {
          event.preventDefault();
          onSelect(undefined);
        }
      }}
    >
      <h2>{strings.toolbox}</h2>
      <button
        type="button"
        title={strings.pointer}
        aria-pressed={!selected}
        onClick={() => onSelect(undefined)}
      >
        <ControlIcon name="Pointer" />
        <span class="toolbox-label">{strings.pointer}</span>
      </button>
      {(["commonControls", "containers"] as const).map((group) => (
        <section class="toolbox-group" key={group}>
          <h3>{strings[group]}</h3>
          {controlMetadata
            .filter(
              (control) =>
                ["Frame", "MultiPage", "TabStrip"].includes(control.type) ===
                (group === "containers"),
            )
            .map((control) => (
              <button
                type="button"
                key={control.type}
                data-toolbox-type={control.type}
                title={control.type}
                aria-pressed={selected === control.type}
                disabled={
                  disabled ||
                  (control.type === "MultiPage" && !pageEditable) ||
                  (control.type === "TabStrip" && !tabEditable)
                }
                onClick={() => onSelect(control.type)}
              >
                <ControlIcon name={control.type} />
                <span class="toolbox-label">{control.type}</span>
              </button>
            ))}
        </section>
      ))}
      {disabled && <p class="notice">{strings.structuralReadOnly}</p>}
      {selected && (
        <p class="notice" role="status">
          {parentEditable
            ? pageEditable
              ? strings.pagePlacement
              : strings.parentPlacement
            : strings.placement}
        </p>
      )}
    </aside>
  );
}
