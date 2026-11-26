/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import { useEffect, useRef, useState } from "preact/hooks";
import { numberRoundTrips } from "./numbers";
import { Icon } from "./Icon";
import { designerStrings } from "../../src/userForm/protocol";
import type {
  DesignerError,
  DesignerStrings,
  PropertyDescriptor,
  PropertyTarget,
  PropertyValue,
  ScalarValue,
} from "../../src/userForm/protocol";
import { objectValue } from "./propertyEditors/types";
import { editorRegistry } from "./propertyEditors/registry";
import { colorStyle, systemColorTable } from "./propertyEditors/appearance";
import { validNumericInput } from "./propertyEditors/numeric";
import { listToText, textToList } from "../../src/userForm/listItems";
import { toggleState } from "../../src/userForm/model";

// Presentation only. The server owns types, applicability and authoring values.
const categories = {
  identity: ["name", "tag"],
  appearance: [
    "caption",
    "text",
    "value",
    "list",
    "font",
    "foreColor",
    "backColor",
    "textAlign",
    "borderStyle",
    "picture",
    "pictureAlignment",
    "pictureSizeMode",
    "backStyle",
    "specialEffect",
    "borderColor",
    "barColor",
    "trackColor",
    "valueColor",
    "showValue",
    "barHeight",
    "alignment",
    "wordWrap",
    "autoSize",
    "multiLine",
    "scrollBars",
    "passwordChar",
    "style",
    "listStyle",
    "tabOrientation",
    "multiRow",
    "tabFixedWidth",
    "tabFixedHeight",
    "orientation",
    "controlTipText",
  ],
  layout: ["left", "top", "width", "height", "zoom", "scrollBars", "keepScrollBarsVisible", "scrollWidth", "scrollHeight", "scrollLeft", "scrollTop"],
  behavior: [
    "enabled",
    "visible",
    "locked",
    "default",
    "cancel",
    "takeFocusOnClick",
    "groupName",
    "maxLength",
    "multiSelect",
    "columnCount",
    "columnWidths",
    "boundColumn",
    "textColumn",
    "columnHeads",
    "matchEntry",
    "matchRequired",
    "showDropButtonWhen",
    "dropButtonStyle",
    "listRows",
    "listWidth",
    "min",
    "max",
    "smallChange",
    "largeChange",
    "enterKeyBehavior",
    "tabKeyBehavior",
    "imeMode",
    "tripleState",
    "controlSource",
    "rowSource",
    "startUpPosition",
    "showModal",
  ],
  navigation: ["tabIndex", "tabStop", "selectedIndex", "accelerator", "cycle"],
  advanced: [],
} as const;
type Category = keyof typeof categories;
const categoryOf = (field: string): Category =>
  (Object.keys(categories) as (keyof typeof categories)[]).find((key) =>
    (categories[key] as readonly string[]).includes(field),
  ) ?? "advanced";
const labels: Record<string, keyof DesignerStrings> = {
  name: "propertyName",
  caption: "propertyCaption",
  text: "propertyText",
  value: "propertyValue",
  left: "propertyLeft",
  top: "propertyTop",
  width: "propertyWidth",
  height: "propertyHeight",
  tabIndex: "propertyTabIndex",
  selectedIndex: "propertySelectedIndex",
  list: "propertyList",
  enabled: "propertyEnabled",
  visible: "propertyVisible",
  tag: "propertyTag",
  controlTipText: "propertyControlTipText",
  accelerator: "propertyAccelerator",
};

export interface PropertyReply {
  requestId: number;
  controlId?: string;
  field?: string;
  error?: DesignerError;
}
interface Props {
  drafts: PropertyDraftStore;
  target: PropertyTarget;
  controlId?: string;
  // A collection entry keeps its draft identity when its wire index changes.
  draftId?: string;
  title: string;
  targetType?: string;
  version: number;
  editable: boolean;
  busy: boolean;
  strings: DesignerStrings;
  reply?: PropertyReply;
  onCommit: (field: string, value: PropertyValue) => number | undefined;
  onOpenEditor?: (descriptor: PropertyDescriptor) => void;
  editorEnabled?: boolean;
  onPosition?: () => void;
}
type ValueKind = "string" | "number" | "boolean";
interface Draft {
  text: string;
  kind: ValueKind;
  dirty: boolean;
  version: number;
  submitted?: number;
  attempted?: boolean;
  error?: string;
}
// Owned by App so switching selections cannot discard a failed or pending edit.
export interface PropertyDraftStore {
  values: Map<string, Draft>;
  replies: Map<number, PropertyReply>;
}
function PropertyRow({ descriptor, ...props }: Props & { descriptor: PropertyDescriptor }) {
  const { target, version, editable, busy, strings, reply, controlId, onCommit, drafts } = props;
  const draftKey = JSON.stringify([props.draftId ?? controlId ?? null, descriptor.field]);
  const source = target.values[descriptor.field];
  // A list of strings is edited as text, one item per line.
  const isList = descriptor.valueType === "stringArray" || descriptor.valueType === "listItems";
  // The Value of a CheckBox, OptionButton or ToggleButton: True, False or Null.
  const isToggle = descriptor.editor === "toggle";
  const fromSource = (): Draft => {
    const value = source?.present ? source.value : undefined;
    if (isList) return { text: listToText(value), kind: "string", dirty: false, version };
    if (isToggle)
      return {
        text: value == null ? "" : { on: "True", off: "False", null: "Null" }[toggleState(value)],
        kind: "string",
        dirty: false,
        version,
      };
    return {
      text: value == null ? "" : String(value),
      kind:
        typeof value === "number"
          ? "number"
          : typeof value === "boolean"
            ? "boolean"
            : descriptor.valueType === "number" || descriptor.valueType === "integer"
              ? "number"
              : descriptor.valueType === "boolean"
                ? "boolean"
                : "string",
      dirty: false,
      version,
    };
  };
  const [draft, setDraft] = useState<Draft>(() => drafts.values.get(draftKey) ?? fromSource());
  const current = useRef(draft);
  const update = (next: Draft) => {
    current.current = next;
    if (next.dirty) drafts.values.set(draftKey, next);
    else drafts.values.delete(draftKey);
    setDraft(next);
  };
  const stale = draft.dirty && draft.version !== version && draft.submitted === undefined;
  useEffect(() => {
    const d = current.current;
    // A preview can arrive before a failed edit reply. Only the matching reply
    // may discard a submitted draft; a document update alone is not success.
    if (!d.dirty) update(fromSource());
  }, [target, version]);
  useEffect(() => {
    const response =
      current.current.submitted === undefined
        ? undefined
        : drafts.replies.get(current.current.submitted);
    if (!response || response.controlId !== controlId || response.field !== descriptor.field)
      return;
    drafts.replies.delete(response.requestId);
    if (response.error) {
      // This reply belongs to this row's single submitted operation. The server
      // may report canonical paths such as controls[0].caption or form.caption.
      const diagnostics = response.error.diagnostics?.filter((d) => d.operationIndex === 0);
      const message = diagnostics?.length
        ? diagnostics
            .map((d) => `${d.code}: ${d.message}${d.suggestion ? ` ${d.suggestion}` : ""}`)
            .join("\n")
        : `${response.error.code}: ${response.error.message}`;
      update({ ...current.current, submitted: undefined, error: message });
    } else update(fromSource());
  }, [reply]);
  const commit = () => {
    const d = current.current;
    if (
      !editable ||
      busy ||
      !d.dirty ||
      d.attempted ||
      d.submitted !== undefined ||
      d.version !== version
    )
      return;
    if (isList) {
      const items = textToList(d.text);
      const authored = source?.present && Array.isArray(source.value) ? source.value : null;
      if (JSON.stringify(items) === JSON.stringify(authored)) {
        update(fromSource());
        return;
      }
      // No items removes the field.
      const requestId = onCommit(descriptor.field, items);
      if (requestId !== undefined)
        update({ ...d, submitted: requestId, attempted: true, error: undefined });
      return;
    }
    let value: ScalarValue = d.text;
    if (isToggle) {
      // On and off are written as Boolean values, Null as the string "Null".
      value = d.text === "True" ? true : d.text === "False" ? false : "Null";
      if (source?.present && toggleState(source.value) === toggleState(value) && typeof source.value === typeof value) {
        update(fromSource());
        return;
      }
      const requestId = onCommit(descriptor.field, value);
      if (requestId !== undefined)
        update({ ...d, submitted: requestId, attempted: true, error: undefined });
      return;
    }
    if (d.kind === "boolean") value = d.text === "true";
    else if (d.kind === "number") {
      value = Number(d.text);
      if (!d.text.trim() || !Number.isFinite(value)) {
        update({ ...d, error: strings.invalidNumber });
        return;
      }
      if (descriptor.valueType === "integer" && !Number.isSafeInteger(value)) {
        update({ ...d, error: strings.invalidInteger });
        return;
      }
      if (!numberRoundTrips(d.text, value)) {
        update({ ...d, error: strings.invalidNumberPrecision });
        return;
      }
      if (!validNumericInput(d.text, descriptor)) {
        update({ ...d, error: strings.invalidValue });
        return;
      }
    }
    if (
      descriptor.allowedValues &&
      !descriptor.allowedValues.some((allowed) => {
        if (d.kind === "number")
          return (
            allowed.trim() !== "" && Number.isFinite(Number(allowed)) && Number(allowed) === value
          );
        if (d.kind === "boolean")
          return (allowed === "true" || allowed === "false") && (allowed === "true") === value;
        return allowed === value;
      })
    ) {
      update({ ...d, error: strings.invalidValue });
      return;
    }
    if (source?.present && source.value === value) {
      update(fromSource());
      return;
    }
    const requestId = onCommit(descriptor.field, value);
    if (requestId !== undefined)
      update({ ...d, submitted: requestId, attempted: true, error: undefined });
  };
  const change = (text: string, kind = current.current.kind) =>
    update({ text, kind, version, dirty: true });
  const id = `property-${controlId ?? "form"}-${descriptor.field}`;
  const error = stale
    ? [strings.staleProperty, draft.error].filter(Boolean).join("\n")
    : draft.error;
  const disabled = !editable || busy;
  // VBE property names are identifiers and stay English in every UI language.
  const fieldLabel =
    designerStrings[labels[descriptor.field]] ?? objectLabels[descriptor.field] ?? vbaName(descriptor.field);
  // Source null remains intact until a concrete edit. It is an input hint,
  // never an authorable Designer value or a reason to hide the actual editor.
  const unconfigured = (!source?.present || source.value === null) && !draft.dirty;
  const placeholder = source?.present && source.value === null ? strings.nullValue : strings.unset;
  return (
    <div class="property-row" data-property-field={descriptor.field}>
      <label for={id}>{fieldLabel}</label>
      <div class="property-value">
        {isList ? (
          <textarea
            id={id}
            rows={Math.min(8, Math.max(3, draft.text.split("\n").length + 1))}
            wrap="off"
            spellcheck={false}
            value={draft.text}
            placeholder={unconfigured ? placeholder : undefined}
            title={strings.listHelp}
            disabled={disabled}
            aria-invalid={!!error}
            aria-describedby={error ? `${id}-error` : undefined}
            onInput={(event) => change(event.currentTarget.value)}
            onBlur={commit}
            onKeyDown={(event) => {
              if (event.isComposing) return;
              // Tab separates the columns of an item (Shift+Tab still leaves the field).
              if (
                event.key === "Tab" &&
                !event.shiftKey &&
                !event.ctrlKey &&
                !event.altKey &&
                !event.metaKey
              ) {
                event.preventDefault();
                const field = event.currentTarget;
                field.setRangeText("\t", field.selectionStart, field.selectionEnd, "end");
                change(field.value);
                return;
              }
              if (event.key === "Enter" && (event.ctrlKey || event.metaKey)) {
                event.preventDefault();
                commit();
              }
              if (event.key === "Escape") {
                event.preventDefault();
                update(fromSource());
              }
            }}
          />
        ) : descriptor.editor === "boolean" && draft.kind === "boolean" ? (
          <input
            id={id}
            type="checkbox"
            checked={draft.text === "true"}
            ref={(element) => {
              if (element) element.indeterminate = unconfigured;
            }}
            disabled={disabled}
            aria-invalid={!!error}
            onChange={(event) => {
              change(String(event.currentTarget.checked));
              commit();
            }}
          />
        ) : isToggle ? (
          <select
            id={id}
            value={unconfigured ? "__unset__" : draft.text}
            disabled={disabled}
            aria-invalid={!!error}
            aria-describedby={error ? `${id}-error` : undefined}
            onChange={(event) => {
              change(event.currentTarget.value);
              commit();
            }}
          >
            {unconfigured && (
              <option value="__unset__" disabled>
                {placeholder}
              </option>
            )}
            {["True", "False", "Null"].map((choice) => (
              <option value={choice}>{choice}</option>
            ))}
          </select>
        ) : descriptor.allowedValues || draft.kind === "boolean" ? (
          <select
            id={id}
            value={unconfigured ? "__unset__" : draft.text}
            disabled={disabled}
            aria-invalid={!!error}
            aria-describedby={error ? `${id}-error` : undefined}
            onChange={(event) => {
              change(event.currentTarget.value);
              commit();
            }}
            onKeyDown={(event) => {
              if (event.key === "Escape") {
                event.preventDefault();
                update(fromSource());
              }
            }}
          >
            {unconfigured && (
              <option value="__unset__" disabled>
                {placeholder}
              </option>
            )}
            {(descriptor.allowedValues ?? ["true", "false"])
              .map((value) => ({
                value,
                label:
                  descriptor.suggestions?.find((hint) => String(hint.value) === value)?.label ??
                  value,
              }))
              .map((choice) => (
                <option value={String(choice.value)}>{choice.label}</option>
              ))}
          </select>
        ) : (
          <input
            id={id}
            type="text"
            inputMode={draft.kind === "number" ? "decimal" : undefined}
            value={draft.text}
            placeholder={unconfigured ? placeholder : undefined}
            disabled={disabled}
            aria-invalid={!!error}
            aria-describedby={error ? `${id}-error` : undefined}
            onInput={(event) => change(event.currentTarget.value)}
            onBlur={commit}
            onKeyDown={(event) => {
              if (event.isComposing) return;
              if (event.key === "Enter") {
                event.preventDefault();
                commit();
              }
              if (event.key === "Escape") {
                event.preventDefault();
                update(fromSource());
              }
            }}
          />
        )}
        {error && (
          <p id={`${id}-error`} role="alert" class="property-error">
            {error}
          </p>
        )}
      </div>
    </div>
  );
}

const objectLabels: Record<string, string> = {
  font: "Font",
  foreColor: "ForeColor",
  backColor: "BackColor",
  picture: "Picture",
  mouseIcon: "MouseIcon",
  pictureAlignment: "PictureAlignment",
  pictureSizeMode: "PictureSizeMode",
  textAlign: "TextAlign",
  borderStyle: "BorderStyle",
};
/** The VBA name of a field (`wordWrap` → `WordWrap`), for fields without a label of their own. */
const vbaName = (field: string) => field.charAt(0).toUpperCase() + field.slice(1);
function ObjectPropertyRow(props: Props & { descriptor: PropertyDescriptor }) {
  const { descriptor, target, strings } = props;
  const state = target.values[descriptor.field];
  const value = objectValue(state?.value);
  const label = objectLabels[descriptor.field] ?? vbaName(descriptor.field);
  const summary = !state?.present
    ? strings.unset
    : state.value === null
      ? strings.nullValue
      : descriptor.editor === "font"
        ? [value.name, typeof value.size === "number" ? `${value.size} pt` : undefined]
            .filter((item) => item !== undefined)
            .join(", ") || "Font"
        : descriptor.editor === "color"
          ? value.kind === "system" && typeof value.value === "string"
            ? (systemColorTable[value.value]?.label ?? value.value)
            : String(value.value ?? "Color")
          : descriptor.editor === "picture"
            ? typeof value.path === "string"
              ? value.path
              : strings.pictureNone
            : "";
  const supported = !!descriptor.editor && Object.hasOwn(editorRegistry, descriptor.editor);
  return (
    <div class="property-row" data-property-field={descriptor.field}>
      <label>{label}</label>
      <div class="property-value object-property-value">
        <span class="property-object-summary" title={summary}>
          {descriptor.editor === "color" && (
            <span class="color-swatch" style={{ backgroundColor: colorStyle(state?.value) }} />
          )}
          {summary}
        </span>
        <button
          type="button"
          class="property-editor-launch"
          aria-label={`${label}…`}
          title={`${label}…`}
          disabled={!supported || !props.editable || props.busy || !props.editorEnabled}
          onClick={() => props.onOpenEditor?.(descriptor)}
        >
          <Icon name="ellipsis" />
        </button>
      </div>
    </div>
  );
}

export function PropertyGrid(props: Props) {
  const groups = Object.keys(categories) as Category[];
  return (
    <aside class="property-grid" aria-label={props.strings.properties}>
      <h2>
        {props.strings.properties}: {props.title}
      </h2>
      {props.targetType && <p class="property-grid-target-type">{props.targetType}</p>}
      {props.onPosition && (
        <button
          class="position-editor-launch"
          type="button"
          onClick={props.onPosition}
          disabled={!props.editable || props.busy || !props.editorEnabled}
        >
          {props.strings.editorPosition}…
        </button>
      )}
      {!props.editable && <p class="notice">{props.strings.propertyReadOnly}</p>}
      {groups.map((category) => {
        const order: readonly string[] = categories[category];
        const descriptors = props.target.descriptors
          .filter((d) => categoryOf(d.field) === category)
          .sort((a, b) => {
            const ai = order.indexOf(a.field),
              bi = order.indexOf(b.field);
            return (ai < 0 ? order.length : ai) - (bi < 0 ? order.length : bi);
          });
        return (
          descriptors.length > 0 && (
            <details class="property-category" open>
              <summary>{props.strings[category]}</summary>
              {descriptors.map((descriptor) =>
                descriptor.valueType === "object" ? (
                  <ObjectPropertyRow
                    key={JSON.stringify([
                      props.draftId ?? props.controlId ?? null,
                      descriptor.field,
                    ])}
                    {...props}
                    descriptor={descriptor}
                  />
                ) : (
                  <PropertyRow
                    key={JSON.stringify([
                      props.draftId ?? props.controlId ?? null,
                      descriptor.field,
                    ])}
                    {...props}
                    descriptor={descriptor}
                  />
                ),
              )}
            </details>
          )
        );
      })}
    </aside>
  );
}
