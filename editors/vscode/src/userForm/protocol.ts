/* Ported from xlflow (https://github.com/harumiWeb/xlflow), MIT License,
   Copyright (c) 2026 harumiWeb. See THIRD_PARTY_NOTICES.md. */
import { containerColorOnCreate, controlMetadata, transparentOnCreate } from "./toolbox";
import { propertyEditorStrings } from "./propertyEditorStrings";
import type { NewControl, ToolboxType } from "./toolbox";

export interface DesignerError {
  code: string;
  message: string;
  line?: number;
  column?: number;
  diagnostics?: {
    operationIndex: number;
    operation?: string;
    controlId?: string;
    field?: string;
    code: string;
    message: string;
    suggestion?: string;
  }[];
}

export type ScalarValue = string | number | boolean | null;
export interface FontValue {
  name?: string | null;
  size?: number | null;
  bold?: boolean | null;
  italic?: boolean | null;
  underline?: boolean | null;
  strikeThrough?: boolean | null;
  charset?: number | null;
}
export interface ColorValue {
  kind: "rgb" | "system" | "ole";
  value: string;
}
export interface PictureValue {
  path?: string;
  remove?: true;
}
/** A property that holds a picture file: the picture of a control or form, or its mouse icon. */
export type PictureField = "picture" | "mouseIcon";
export type PropertyValue = ScalarValue | FontValue | ColorValue | PictureValue | string[] | (string | string[])[];
export interface AppearanceProperties {
  font?: FontValue | null;
  foreColor?: ColorValue | null;
  backColor?: ColorValue | null;
  textAlign?: number | null;
  borderStyle?: number | null;
  picture?: PictureValue | null;
  pictureAlignment?: number | null;
  pictureSizeMode?: number | null;
  /** `data:` URI of the picture file, read by the extension host for the canvas. */
  picturePreview?: string;
}
export interface PropertyDescriptor {
  field: string;
  valueType: "string" | "number" | "integer" | "boolean" | "any" | "object" | "stringArray" | "listItems";
  required: boolean;
  nullable: boolean;
  allowedValues?: string[];
  editor?: string;
  children?: PropertyDescriptor[];
  minimum?: number;
  maximum?: number;
  suggestions?: { value: string | number; label: string }[];
  exclusiveMinimum?: boolean;
}
export interface PropertyState {
  present: boolean;
  legacy?: boolean;
  value: PropertyValue;
}
export interface PropertyTarget {
  descriptors: PropertyDescriptor[];
  values: Record<string, PropertyState>;
}
export interface PropertyGridData {
  form: PropertyTarget;
  controls: Record<string, PropertyTarget>;
  tabs?: Record<string, PropertyTarget[]>;
}

export interface SpecTab {
  name: string;
  caption?: string;
  visible?: boolean;
  enabled?: boolean;
  tag?: string;
  controlTipText?: string;
  accelerator?: string;
}

export interface SpecControl extends AppearanceProperties {
  id: string;
  parentId?: string;
  type: string;
  progId?: string;
  name: string;
  caption?: string;
  text?: string;
  value?: unknown;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
  zIndex?: number;
  selectedIndex?: number;
  enabled?: boolean;
  visible?: boolean;
  /** ComboBox / ListBox items: a string, or the strings of an item's columns. */
  list?: (string | string[])[];
  tabs?: SpecTab[];
  /** ComboBox: 0 DropDownCombo, 2 DropDownList. TabStrip / MultiPage: 0 tabs, 1 buttons, 2 none. */
  style?: number | null;
  /** ListBox: 0 single, 1 multi, 2 extended. */
  multiSelect?: number | null;
  /** ComboBox / ListBox: 0 plain, 1 option (a box beside each item). */
  listStyle?: number | null;
  /** ComboBox / ListBox: 0 first letter, 1 complete, 2 none. */
  matchEntry?: number | null;
  /** ComboBox: 0 never, 1 with the focus, 2 always (omitted: always, as Excel makes it). */
  showDropButtonWhen?: number | null;
  /** ComboBox: rows of the drop-down list. */
  listRows?: number | null;
  /** ComboBox: 0 plain, 1 arrow, 2 ellipsis, 3 reduce. */
  dropButtonStyle?: number | null;
  /** 0 flat, 1 raised, 2 sunken, 3 etched, 6 bump. */
  specialEffect?: number | null;
  /** 0 transparent, 1 opaque. */
  backStyle?: number | null;
  borderColor?: ColorValue | null;
  /** CheckBox / OptionButton: 0 caption on the left, 1 on the right. */
  alignment?: number | null;
  wordWrap?: boolean | null;
  multiLine?: boolean | null;
  passwordChar?: string | null;
  /** TextBox: the gray hint shown while it is empty and has no focus (spec §14.11.18). */
  placeholder?: string | null;
  /** SpinButton / ScrollBar: -1 by the shape, 0 vertical, 1 horizontal. */
  orientation?: number | null;
  /** TabStrip / MultiPage tabs: 0 top, 1 bottom, 2 left, 3 right. */
  tabOrientation?: number | null;
  default?: boolean | null;
  /** ComboBox / ListBox: the cells the items come from. */
  rowSource?: string | null;
  /** Label, buttons, toggles: where the picture is beside the caption (fmPicturePosition 0–12). */
  picturePosition?: number | null;
  /** Frame / Page: the background picture repeats. */
  pictureTiling?: boolean | null;
  mouseIcon?: PictureValue | null;
  mousePointer?: number | null;
  /** ComboBox / ListBox: columns shown (-1 all), their widths ("72 pt;;1 in"), the column of Text. */
  columnCount?: number | null;
  columnWidths?: string | null;
  textColumn?: number | null;
  columnHeads?: boolean | null;
  /** Edges of the container the control follows: "top,left" (omitted), "top,bottom,left,right", "none" ... */
  anchor?: string | null;
  observed?: Partial<Omit<SpecControl, "observed">>;
}

export interface FormDimensions extends AppearanceProperties {
  /** The form's edge: the color of a Single border, and 0 flat, 1 raised, 2 sunken, 3 etched, 6 bump. */
  borderColor?: ColorValue | null;
  specialEffect?: number | null;
  /** The background picture repeats. */
  pictureTiling?: boolean | null;
  caption?: string | null;
  width?: number | null;
  height?: number | null;
  /** The least and greatest size of the form in points (spec §14.11.17); 0 is no limit. */
  sizeLimits?: FormSizeLimits;
}

export interface FormSizeLimits {
  minWidth: number;
  minHeight: number;
  maxWidth: number;
  maxHeight: number;
}

export interface ObservedFormDimensions {
  caption?: string | null;
  width?: number | null;
  height?: number | null;
  clientWidth?: number;
  clientHeight?: number;
  insideWidth?: number;
  insideHeight?: number;
}

export interface PreviewDocument {
  form: FormDimensions & { name: string; observed?: ObservedFormDimensions };
  controls: SpecControl[];
}

export interface SourcePosition {
  line: number;
  character: number;
}
export interface SourceRange {
  start: SourcePosition;
  end: SourcePosition;
}

/** Mirrors forms.ValidationIssue; range is added only by the LSP preview transport. */
export interface DesignerProblem {
  code?: string;
  severity?: string;
  message?: string;
  field?: string;
  suggestion?: string;
  support?: string;
  range?: SourceRange;
}

export interface PreviewResult {
  version: number;
  eventEditable?: boolean;
  editable?: boolean;
  propertyEditable?: boolean;
  propertyEditorVersion?: number;
  structuralEditable?: boolean;
  parentEditable?: boolean;
  pageEditable?: boolean;
  tabEditable?: boolean;
  productivityEditable?: boolean;
  propertyGrid?: PropertyGridData;
  propertyError?: DesignerError;
  document?: PreviewDocument;
  warnings?: DesignerProblem[];
  error?: DesignerError;
}

export interface DesignerControl extends Omit<SpecControl, "observed"> {
  left: number;
  top: number;
  width: number;
  height: number;
  approximate: boolean;
}

export interface DesignerDocument extends AppearanceProperties {
  borderColor?: ColorValue | null;
  specialEffect?: number | null;
  pictureTiling?: boolean | null;
  name: string;
  caption: string;
  width: number;
  height: number;
  sizeLimits?: FormSizeLimits;
  approximate: boolean;
  eventEditable: boolean;
  /** The font of controls made with the Toolbox when the form has none (the system UI font). */
  newControlFont?: FontValue;
  controls: DesignerControl[];
  warnings: DesignerProblem[];
}

export const designerStrings = {
  ...propertyEditorStrings,
  header: "RAS ユーザーフォーム デザイナー",
  designerView: "デザイナー",
  structure: "構造",
  commonControls: "標準",
  containers: "コンテナー",
  order: "順序",
  moreActions: "その他の操作",
  selectionBounds: "位置とサイズ",
  closePane: "ペインを閉じる",
  resizeProperties: "ドラッグしてプロパティ ペインの幅を変更（ダブルクリックで元に戻す）",
  openText: "JSON を横に開く",
  openEvent: "イベントを開く…",
  eventPending: "イベントを開いています…",
  eventError: "イベントを開けませんでした",
  save: "保存",
  undo: "元に戻す",
  redo: "やり直し",
  unsaved: "未保存の変更があります",
  sourceConflict: "JSON とデザイナーの変更が競合しています。",
  useSource: "JSON の変更を使う",
  useDesigner: "デザイナーの変更を使う",
  lastValid: "最後に有効だった内容を表示しています。",
  loading: "FormSpec を読み込んでいます…",
  approximate: "近似のプレビューです。寸法が無い項目は既定値で表示し、Page の範囲は MultiPage から決まります。",
  approximateBounds: "近似の範囲",
  grid: "グリッドを表示",
  snap: "グリッドに合わせる",
  smartGuides: "スマートガイド",
  smartGuidesHint: "移動するコントロールを、ほかのコントロールの端・中央、同じ間隔、ほかのコントロールやフォームの端から 8・16・24 の位置に合わせ、ガイドを表示します。Alt キーを押しながらドラッグすると合わせません。",
  zoom: "ズーム",
  readOnly: "このフォームは編集できません。",
  cannotMove: "このコントロールは親より大きいため移動できません。先に親に収まるサイズにしてください。",
  properties: "プロパティ",
  identity: "識別",
  appearance: "外観",
  layout: "レイアウト",
  behavior: "動作",
  navigation: "ナビゲーション",
  advanced: "詳細",
  propertyName: "Name",
  propertyCaption: "Caption",
  propertyText: "Text",
  propertyValue: "Value",
  propertyLeft: "Left",
  propertyTop: "Top",
  propertyWidth: "Width",
  propertyHeight: "Height",
  propertyTabIndex: "TabIndex",
  propertySelectedIndex: "SelectedIndex",
  propertyList: "List",
  listHelp: "1 行に 1 項目を入力します。複数の列はタブ（Tab キー。Excel からの貼り付けも可）で区切ります。空にすると list を削除します。",
  pictureNone: "画像なし",
  propertyEnabled: "Enabled",
  propertyVisible: "Visible",
  propertyTag: "Tag",
  propertyControlTipText: "ControlTipText",
  propertyAccelerator: "Accelerator",
  unset: "未設定",
  nullValue: "null",
  valueType: "値の型",
  typeString: "文字列",
  typeNumber: "数値",
  typeBoolean: "真偽値",
  invalidNumber: "有限の数値を入力してください。",
  invalidNumberPrecision: "この数値は精度が失われます。JSON を直接編集してください。",
  invalidInteger: "整数を入力してください。",
  invalidValue: "正しい値を入力してください。",
  staleProperty: "ドキュメントが変更されました。更新後のフォームで値を入力し直してください。",
  propertyReadOnly: "このプロパティは編集できません。",
  toolbox: "ツールボックス",
  pointer: "ポインター",
  deleteControl: "コントロールを削除",
  structuralReadOnly: "コントロールの追加と削除はできません。",
  cannotInsert: "このコントロールの既定のサイズはフォームより大きいため配置できません。",
  placement: "フォームをクリックすると既定のサイズで、ドラッグするとその大きさでコントロールを配置します。Esc で取り消します。",
  hierarchy: "選択のパス",
  invalidDrop: "このコントロールはこのコンテナーに配置できません。",
  dropTarget: "ドロップ先",
  resizeWidth: "フォームの幅を変更",
  resizeHeight: "フォームの高さを変更",
  resizeBoth: "フォームの幅と高さを変更",
  parentReadOnly: "親の変更はできません。",
  parentPlacement: "フォームか Frame をクリックすると既定のサイズで、ドラッグするとその大きさでコントロールを配置します。Esc で取り消します。",
  cannotInsertContainer: "このコントロールの既定のサイズはコンテナーより大きいため配置できません。",
  pages: "ページ",
  tabs: "タブ",
  addPage: "ページを追加",
  addTab: "タブを追加",
  deleteItem: "削除",
  moveUp: "上へ",
  moveDown: "下へ",
  moveToPage: "ページへ移動",
  choosePage: "移動先のページを選択",
  collectionReadOnly: "ページやタブの編集はできません。",
  pagePlacement: "フォーム・Frame・表示中のページをクリックすると既定のサイズで、ドラッグするとその大きさでコントロールを配置します。Esc で取り消します。",
  productivity: "整列",
  align: "整列",
  alignLeft: "左揃え",
  alignCenter: "左右中央揃え",
  alignRight: "右揃え",
  alignTop: "上揃え",
  alignMiddle: "上下中央揃え",
  alignBottom: "下揃え",
  sameSize: "サイズを揃える",
  sameWidth: "幅を揃える",
  sameHeight: "高さを揃える",
  sameDimensions: "幅と高さを揃える",
  alignToGrid: "グリッドに合わせる",
  sizeToGrid: "サイズをグリッドに合わせる",
  spacing: "間隔",
  horizontalSpacing: "左右",
  verticalSpacing: "上下",
  spacingEqual: "均等にする",
  spacingIncrease: "広げる",
  spacingDecrease: "狭める",
  spacingRemove: "詰める",
  centerInForm: "中央に配置",
  centerHorizontally: "左右の中央に配置",
  centerVertically: "上下の中央に配置",
  anchorEditor: "アンカー",
  anchorTop: "上",
  anchorBottom: "下",
  anchorLeft: "左",
  anchorRight: "右",
  anchorHelp: "親（フォーム・Frame・MultiPage）の大きさが変わったとき、コントロールが距離を保つ辺。左右の両方なら幅が、上下の両方なら高さが伸び縮みします。",
  cut: "切り取り",
  copy: "コピー",
  paste: "貼り付け",
  duplicate: "複製",
  bringToFront: "最前面へ移動",
  bringForward: "前面へ移動",
  sendBackward: "背面へ移動",
  sendToBack: "最背面へ移動",
  tabOrder: "タブ順",
  tabOrderParent: "親",
  applyTabOrder: "タブ順を適用",
  noEditableTabOrder: "この親の下には TabIndex を編集できるコントロールがありません。",
  selectionCount: "選択中のコントロール",
  propertiesAction: "プロパティ",
  selectControls: "コントロールを選択",
  problems: "問題",
  closeProblems: "問題パネルを閉じる",
  noProblems: "問題はありません",
  problemSeverityError: "エラー",
  problemSeverityWarning: "警告",
  problemSeverityInformation: "情報",
  problemSeverityHint: "ヒント",
};
export type DesignerStrings = typeof designerStrings;

export type HostMessage =
  | { type: "localization"; strings: DesignerStrings }
  | {
      type: "document";
      version: number;
      document: DesignerDocument;
      editable?: boolean;
      propertyEditable?: boolean;
      propertyEditorVersion?: number;
      structuralEditable?: boolean;
      parentEditable?: boolean;
      pageEditable?: boolean;
      tabEditable?: boolean;
      productivityEditable?: boolean;
      propertyGrid?: PropertyGridData;
      propertyError?: DesignerError;
    }
  | { type: "invalidDocument"; version: number; error: DesignerError }
  | { type: "editResult"; requestId: number; error?: DesignerError; selectionIds?: string[] }
  | { type: "eventResult"; requestId: number; error?: string }
  | {
      type: "pictureSelection";
      requestId: number;
      selection?: {
        token: string;
        name: string;
        /** Empty when the webview cannot draw the format. */
        dataURI: string;
        /** Size in pixels, when the host could read it. */
        width?: number;
        height?: number;
        /** The compiler converts the picture to a 24-bit BMP when it writes the form. */
        converted?: boolean;
      };
      error?: DesignerError;
    }
  | { type: "clipboardState"; available: boolean }
  | { type: "historyState"; canUndo: boolean; canRedo: boolean }
  | { type: "draftState"; dirty: boolean; conflicted: boolean; saving: boolean; error?: string }
  | { type: "editingUnavailable" }
  | { type: "themeChanged" };

export type GeometryOperation =
  | { type: "moveControl"; controlId: string; left: number; top: number }
  | { type: "resizeControl"; controlId: string; width: number; height: number };

export type PropertyOperation =
  | { type: "setFormProperty"; field: string; value: PropertyValue }
  | { type: "setControlProperty"; controlId: string; field: string; value: PropertyValue };
export interface SetGeometryOperation {
  type: "setGeometry";
  controlId?: string;
  left?: number;
  top?: number;
  width?: number;
  height?: number;
}
export type StructuralOperation =
  | { type: "addControl"; control: NewControl }
  | { type: "removeControl"; controlId: string; cascade: boolean };
export type ParentOperation = {
  type: "setParent";
  controlId: string;
  parentId: string;
  left: number;
  top: number;
};
export type PageOperation =
  | { type: "addPage"; control: Pick<SpecControl, "id" | "parentId" | "name" | "type" | "caption"> }
  | { type: "removePage"; controlId: string; cascade: boolean }
  | { type: "reorderPage"; controlId: string; index: number };
export type TabOperation =
  | { type: "addTab"; controlId: string; tab: SpecTab }
  | { type: "removeTab"; controlId: string; index: number }
  | { type: "reorderTab"; controlId: string; index: number; toIndex: number }
  | { type: "setTabProperty"; controlId: string; index: number; field: string; value: ScalarValue };
export interface ClipboardSnapshot {
  format: "yaml" | "json";
  text: string;
  controlIds: string[];
}
export interface CopyResult {
  version: number;
  snapshot?: ClipboardSnapshot;
  document?: PreviewDocument;
  hasPicturePaths?: boolean;
  error?: DesignerError;
}
export type ProductivityOperation =
  | { type: "moveControls"; moves: { controlId: string; left: number; top: number }[] }
  | {
      type: "alignControls";
      controlIds: string[];
      anchorId: string;
      alignment: "left" | "center" | "right" | "top" | "middle" | "bottom";
    }
  | { type: "distributeControls"; controlIds: string[]; axis: "horizontal" | "vertical" }
  | {
      type: "sameSizeControls";
      controlIds: string[];
      anchorId: string;
      dimension: "width" | "height" | "size";
    }
  | { type: "alignToGrid"; controlIds: string[] }
  | { type: "sizeToGrid"; controlIds: string[] }
  | {
      type: "spaceControls";
      controlIds: string[];
      anchorId: string;
      axis: "horizontal" | "vertical";
      command: "increase" | "decrease" | "remove";
    }
  | { type: "centerControls"; controlIds: string[]; axis: "horizontal" | "vertical" }
  | { type: "removeControls"; controlIds: string[]; cascade: boolean }
  | {
      type: "reorderControls";
      controlIds: string[];
      command: "front" | "forward" | "backward" | "back";
    }
  | { type: "setTabOrder"; parentId: string; controlIds: string[] }
  | {
      type: "pasteControls";
      snapshot: ClipboardSnapshot;
      allocations: { sourceId: string; id: string; name: string }[];
      parentId: string;
      offset: { left: number; top: number };
    };
export type SemanticOperation =
  | GeometryOperation
  | SetGeometryOperation
  | PropertyOperation
  | StructuralOperation
  | ParentOperation
  | PageOperation
  | TabOperation
  | ProductivityOperation;

export function isPageOperation(value: unknown): value is PageOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (op.type === "addPage") {
    if (Object.keys(op).length !== 2 || typeof op.control !== "object" || op.control === null)
      return false;
    const c = op.control as Record<string, unknown>;
    return (
      Object.keys(c).every((k) => ["id", "parentId", "name", "type", "caption"].includes(k)) &&
      typeof c.id === "string" &&
      !!c.id &&
      typeof c.parentId === "string" &&
      !!c.parentId &&
      typeof c.name === "string" &&
      /^[A-Za-z][A-Za-z0-9_]*$/.test(c.name) &&
      c.type === "Page" &&
      (!Object.hasOwn(c, "caption") || typeof c.caption === "string")
    );
  }
  if (typeof op.controlId !== "string" || !op.controlId) return false;
  if (op.type === "removePage")
    return Object.keys(op).length === 3 && typeof op.cascade === "boolean";
  return (
    op.type === "reorderPage" &&
    Object.keys(op).length === 3 &&
    Number.isSafeInteger(op.index) &&
    Number(op.index) >= 0
  );
}

export function isTabOperation(value: unknown): value is TabOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (typeof op.controlId !== "string" || !op.controlId) return false;
  if (op.type === "addTab") {
    if (
      Object.keys(op).length !== 3 ||
      typeof op.tab !== "object" ||
      op.tab === null ||
      Array.isArray(op.tab)
    )
      return false;
    const tab = op.tab as Record<string, unknown>;
    return (
      typeof tab.name === "string" &&
      !!tab.name &&
      Object.entries(tab).every(([key, v]) =>
        ["enabled", "visible"].includes(key)
          ? typeof v === "boolean"
          : ["name", "caption", "tag", "controlTipText", "accelerator"].includes(key) &&
            typeof v === "string",
      )
    );
  }
  if (!Number.isSafeInteger(op.index) || Number(op.index) < 0) return false;
  if (op.type === "removeTab") return Object.keys(op).length === 3;
  if (op.type === "reorderTab")
    return (
      Object.keys(op).length === 4 && Number.isSafeInteger(op.toIndex) && Number(op.toIndex) >= 0
    );
  return (
    op.type === "setTabProperty" &&
    Object.keys(op).length === 5 &&
    typeof op.field === "string" &&
    !!op.field &&
    (op.value === null ||
      typeof op.value === "string" ||
      typeof op.value === "boolean" ||
      (typeof op.value === "number" && Number.isFinite(op.value)))
  );
}

const isStringArray = (value: unknown, minimumLength = 0): value is string[] =>
  Array.isArray(value) &&
  value.length >= minimumLength &&
  value.every((item) => typeof item === "string" && item.length > 0);

export function isProductivityOperation(value: unknown): value is ProductivityOperation {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return false;
  const op = value as Record<string, unknown>;
  const exactKeys = (keys: string[]) =>
    Object.keys(op).length === keys.length && Object.keys(op).every((key) => keys.includes(key));
  const finite = (n: unknown) => typeof n === "number" && Number.isFinite(n);
  if (op.type === "moveControls") {
    return (
      exactKeys(["type", "moves"]) &&
      Array.isArray(op.moves) &&
      op.moves.length >= 2 &&
      op.moves.every((move) => {
        if (typeof move !== "object" || move === null || Array.isArray(move)) return false;
        const item = move as Record<string, unknown>;
        return (
          Object.keys(item).length === 3 &&
          Object.keys(item).every((key) => ["controlId", "left", "top"].includes(key)) &&
          typeof item.controlId === "string" &&
          item.controlId.length > 0 &&
          finite(item.left) &&
          finite(item.top)
        );
      }) &&
      new Set((op.moves as { controlId: string }[]).map((move) => move.controlId)).size ===
        (op.moves as unknown[]).length
    );
  }
  if (op.type === "alignControls")
    return (
      exactKeys(["type", "controlIds", "anchorId", "alignment"]) &&
      isStringArray(op.controlIds, 2) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      typeof op.anchorId === "string" &&
      op.controlIds.includes(op.anchorId) &&
      ["left", "center", "right", "top", "middle", "bottom"].includes(String(op.alignment))
    );
  if (op.type === "distributeControls")
    return (
      exactKeys(["type", "controlIds", "axis"]) &&
      isStringArray(op.controlIds, 3) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      ["horizontal", "vertical"].includes(String(op.axis))
    );
  if (op.type === "sameSizeControls")
    return (
      exactKeys(["type", "controlIds", "anchorId", "dimension"]) &&
      isStringArray(op.controlIds, 2) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      typeof op.anchorId === "string" &&
      op.controlIds.includes(op.anchorId) &&
      ["width", "height", "size"].includes(String(op.dimension))
    );
  if (op.type === "alignToGrid" || op.type === "sizeToGrid")
    return (
      exactKeys(["type", "controlIds"]) &&
      isStringArray(op.controlIds, 1) &&
      new Set(op.controlIds).size === op.controlIds.length
    );
  if (op.type === "spaceControls")
    return (
      exactKeys(["type", "controlIds", "anchorId", "axis", "command"]) &&
      isStringArray(op.controlIds, 2) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      typeof op.anchorId === "string" &&
      op.controlIds.includes(op.anchorId) &&
      ["horizontal", "vertical"].includes(String(op.axis)) &&
      ["increase", "decrease", "remove"].includes(String(op.command))
    );
  if (op.type === "centerControls")
    return (
      exactKeys(["type", "controlIds", "axis"]) &&
      isStringArray(op.controlIds, 1) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      ["horizontal", "vertical"].includes(String(op.axis))
    );
  if (op.type === "removeControls")
    return (
      exactKeys(["type", "controlIds", "cascade"]) &&
      isStringArray(op.controlIds, 1) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      typeof op.cascade === "boolean"
    );
  if (op.type === "reorderControls")
    return (
      exactKeys(["type", "controlIds", "command"]) &&
      isStringArray(op.controlIds, 1) &&
      new Set(op.controlIds).size === op.controlIds.length &&
      ["front", "forward", "backward", "back"].includes(String(op.command))
    );
  if (op.type === "setTabOrder")
    return (
      exactKeys(["type", "parentId", "controlIds"]) &&
      typeof op.parentId === "string" &&
      isStringArray(op.controlIds) &&
      new Set(op.controlIds).size === op.controlIds.length
    );
  if (op.type === "pasteControls") {
    const snapshot = op.snapshot as Record<string, unknown> | null;
    const offset = op.offset as Record<string, unknown> | null;
    return (
      exactKeys(["type", "snapshot", "allocations", "parentId", "offset"]) &&
      typeof snapshot === "object" &&
      snapshot !== null &&
      Object.keys(snapshot).length === 3 &&
      Object.keys(snapshot).every((key) => ["format", "text", "controlIds"].includes(key)) &&
      ["yaml", "json"].includes(String(snapshot.format)) &&
      typeof snapshot.text === "string" &&
      isStringArray(snapshot.controlIds, 1) &&
      new Set(snapshot.controlIds).size === snapshot.controlIds.length &&
      Array.isArray(op.allocations) &&
      op.allocations.length >= 1 &&
      op.allocations.every((allocation) => {
        if (typeof allocation !== "object" || allocation === null || Array.isArray(allocation))
          return false;
        const item = allocation as Record<string, unknown>;
        return (
          Object.keys(item).length === 3 &&
          Object.keys(item).every((key) => ["sourceId", "id", "name"].includes(key)) &&
          typeof item.sourceId === "string" &&
          item.sourceId.length > 0 &&
          typeof item.id === "string" &&
          item.id.length > 0 &&
          typeof item.name === "string" &&
          /^[A-Za-z][A-Za-z0-9_]*$/.test(item.name)
        );
      }) &&
      typeof op.parentId === "string" &&
      typeof offset === "object" &&
      offset !== null &&
      Object.keys(offset).length === 2 &&
      Object.keys(offset).every((key) => ["left", "top"].includes(key)) &&
      finite(offset.left) &&
      finite(offset.top)
    );
  }
  return false;
}

function isClipboardMessage(value: Record<string, unknown>): boolean {
  return (
    value.type === "clipboard" &&
    Object.keys(value).length === 6 &&
    Object.keys(value).every((key) =>
      ["type", "requestId", "version", "action", "controlIds", "parentId"].includes(key),
    ) &&
    Number.isSafeInteger(value.requestId) &&
    Number(value.requestId) >= 0 &&
    Number.isSafeInteger(value.version) &&
    Number(value.version) >= 0 &&
    ["copy", "cut", "paste", "duplicate"].includes(String(value.action)) &&
    isStringArray(value.controlIds) &&
    new Set(value.controlIds).size === value.controlIds.length &&
    typeof value.parentId === "string"
  );
}

export function isStructuralOperation(value: unknown): value is StructuralOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (op.type === "removeControl")
    return (
      Object.keys(op).length === 3 &&
      Object.keys(op).every((key) => ["type", "controlId", "cascade"].includes(key)) &&
      typeof op.controlId === "string" &&
      op.controlId.length > 0 &&
      typeof op.cascade === "boolean"
    );
  if (
    op.type !== "addControl" ||
    Object.keys(op).length !== 2 ||
    typeof op.control !== "object" ||
    op.control === null ||
    Array.isArray(op.control)
  )
    return false;
  const control = op.control as Record<string, unknown>;
  return (
    Object.keys(control).every((key) =>
      [
        "id",
        "parentId",
        "name",
        "type",
        "left",
        "top",
        "width",
        "height",
        "caption",
        "selectedIndex",
        "tabs",
        "backStyle",
        "backColor",
        "font",
      ].includes(key),
    ) &&
    typeof control.id === "string" &&
    control.id.length > 0 &&
    typeof control.name === "string" &&
    /^[A-Za-z][A-Za-z0-9_]*$/.test(control.name) &&
    (!Object.hasOwn(control, "parentId") ||
      (typeof control.parentId === "string" && control.parentId.length > 0)) &&
    controlMetadata.some((m) => m.type === control.type) &&
    ["left", "top", "width", "height"].every(
      (key) => typeof control[key] === "number" && Number.isFinite(control[key]),
    ) &&
    Number(control.width) > 0 &&
    Number(control.height) > 0 &&
    (!Object.hasOwn(control, "caption") || typeof control.caption === "string") &&
    (!Object.hasOwn(control, "backStyle") ||
      (control.backStyle === 0 && transparentOnCreate.has(control.type as ToolboxType))) &&
    (!Object.hasOwn(control, "backColor") ||
      (containerColorOnCreate.has(control.type as ToolboxType) && isNewBackColor(control.backColor))) &&
    (!Object.hasOwn(control, "font") ||
      (typeof control.font === "object" &&
        control.font !== null &&
        !Array.isArray(control.font) &&
        Object.keys(control.font).every((key) =>
          ["name", "size", "bold", "italic", "underline", "strikeThrough", "charset", "pitchAndFamily"].includes(key),
        ))) &&
    (((!Object.hasOwn(control, "selectedIndex") || control.selectedIndex === -1) &&
      (!Object.hasOwn(control, "tabs") ||
        (Array.isArray(control.tabs) && control.tabs.length === 0))) ||
      (control.type === "TabStrip" &&
        control.selectedIndex === 0 &&
        Array.isArray(control.tabs) &&
        control.tabs.length === 2 &&
        control.tabs.every(
          (tab) =>
            typeof tab === "object" &&
            tab !== null &&
            Object.keys(tab).every((key) => ["name", "caption"].includes(key)) &&
            typeof tab.name === "string" &&
            /^[A-Za-z][A-Za-z0-9_]*$/.test(tab.name) &&
            typeof tab.caption === "string",
        ) &&
        new Set(control.tabs.map((tab) => tab.name.toLowerCase())).size === 2))
  );
}

function isNewBackColor(value: unknown): boolean {
  if (typeof value !== "object" || value === null || Array.isArray(value)) return false;
  const color = value as Record<string, unknown>;
  return (
    Object.keys(color).every((key) => ["kind", "value"].includes(key)) &&
    ["rgb", "system", "ole"].includes(color.kind as string) &&
    typeof color.value === "string"
  );
}

export function isInitialMultiPageOperations(operations: unknown[]): boolean {
  if (operations.length !== 3) return false;
  const [owner, ...pages] = operations;
  if (
    !isStructuralOperation(owner) ||
    owner.type !== "addControl" ||
    owner.control.type !== "MultiPage" ||
    owner.control.selectedIndex !== -1
  )
    return false;
  if (
    !pages.every(
      (page) =>
        isPageOperation(page) &&
        page.type === "addPage" &&
        page.control.parentId === owner.control.id,
    )
  )
    return false;
  const controls = [
    owner.control,
    ...pages.map((page) => (page as Extract<PageOperation, { type: "addPage" }>).control),
  ];
  return (
    new Set(controls.map((control) => control.id)).size === 3 &&
    new Set(controls.map((control) => control.name.toLowerCase())).size === 3
  );
}

export function isPropertyOperation(value: unknown): value is PropertyOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  const control = op.type === "setControlProperty";
  if (!control && op.type !== "setFormProperty") return false;
  const keys = control ? ["type", "controlId", "field", "value"] : ["type", "field", "value"];
  return (
    Object.keys(op).length === keys.length &&
    Object.keys(op).every((key) => keys.includes(key)) &&
    (!control || (typeof op.controlId === "string" && op.controlId.length > 0)) &&
    typeof op.field === "string" &&
    op.field.length > 0 &&
    (op.value === null ||
      typeof op.value === "string" ||
      typeof op.value === "boolean" ||
      (typeof op.value === "number" && Number.isFinite(op.value)) ||
      (["font", "foreColor", "backColor", "picture"].includes(op.field as string) &&
        typeof op.value === "object" &&
        op.value !== null &&
        !Array.isArray(op.value)) ||
      (control &&
        op.field === "list" &&
        Array.isArray(op.value) &&
        op.value.every(
          (item) =>
            typeof item === "string" ||
            (Array.isArray(item) && item.every((column) => typeof column === "string")),
        )))
  );
}

export interface SourceTextEdit {
  range: { start: { line: number; character: number }; end: { line: number; character: number } };
  newText: string;
}
export interface EditResult {
  version: number;
  edits?: SourceTextEdit[];
  error?: DesignerError;
  selectionIds?: string[];
}
export type WebviewMessage =
  | { type: "ready" }
  | { type: "openText" }
  | { type: "openProblem"; version: number; index: number }
  | { type: "save" }
  | {
      type: "openEvent";
      requestId: number;
      version: number;
      targetId?: string;
      choose: boolean;
    }
  | { type: "history"; action: "undo" | "redo"; requestId: number; version: number }
  /** `controlId` "" is the form; `field` the picture property (`picture`, `mouseIcon`). */
  | { type: "picturePick"; requestId: number; version: number; controlId: string; field: PictureField }
  | {
      type: "pictureCommit";
      requestId: number;
      version: number;
      controlId: string;
      field: PictureField;
      token: string;
    }
  | { type: "pictureCancel" }
  | { type: "resolveConflict"; choice: "source" | "designer" }
  | {
      type: "clipboard";
      requestId: number;
      version: number;
      action: "copy" | "cut" | "paste" | "duplicate";
      controlIds: string[];
      parentId: string;
    }
  | { type: "edit"; requestId: number; version: number; operations: SemanticOperation[] };

export function isGeometryOperation(value: unknown): value is GeometryOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  if (typeof op.controlId !== "string" || !op.controlId) return false;
  const fields =
    op.type === "moveControl"
      ? ["left", "top"]
      : op.type === "resizeControl"
        ? ["width", "height"]
        : [];
  return (
    fields.length === 2 &&
    Object.keys(op).every((key) => ["type", "controlId", ...fields].includes(key)) &&
    fields.every((key) => typeof op[key] === "number" && Number.isFinite(op[key]))
  );
}

export function isSetGeometryOperation(value: unknown): value is SetGeometryOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  const axes = ["left", "top", "width", "height"];
  return (
    op.type === "setGeometry" &&
    Object.keys(op).every((key) => ["type", "controlId", ...axes].includes(key)) &&
    (!Object.hasOwn(op, "controlId") ||
      (typeof op.controlId === "string" && op.controlId.length > 0)) &&
    (op.controlId !== undefined || (!Object.hasOwn(op, "left") && !Object.hasOwn(op, "top"))) &&
    axes.some((key) => Object.hasOwn(op, key)) &&
    axes.every(
      (key) =>
        !Object.hasOwn(op, key) ||
        (typeof op[key] === "number" &&
          Number.isFinite(op[key]) &&
          (!["width", "height"].includes(key) || Number(op[key]) > 0)),
    )
  );
}

export function isParentOperation(value: unknown): value is ParentOperation {
  if (typeof value !== "object" || value === null) return false;
  const op = value as Record<string, unknown>;
  return (
    Object.keys(op).length === 5 &&
    Object.keys(op).every((key) =>
      ["type", "controlId", "parentId", "left", "top"].includes(key),
    ) &&
    op.type === "setParent" &&
    typeof op.controlId === "string" &&
    op.controlId.length > 0 &&
    typeof op.parentId === "string" &&
    typeof op.left === "number" &&
    Number.isFinite(op.left) &&
    typeof op.top === "number" &&
    Number.isFinite(op.top)
  );
}

export function isWebviewMessage(value: unknown): value is WebviewMessage {
  if (typeof value !== "object" || value === null) return false;
  const message = value as Record<string, unknown>;
  if (message.type === "openProblem")
    return (
      Object.keys(message).length === 3 &&
      Number.isSafeInteger(message.version) &&
      Number(message.version) >= 0 &&
      Number.isSafeInteger(message.index) &&
      Number(message.index) >= 0
    );
  if (message.type === "history")
    return (
      Object.keys(message).length === 4 &&
      (message.action === "undo" || message.action === "redo") &&
      Number.isSafeInteger(message.requestId) &&
      Number(message.requestId) >= 0 &&
      Number.isSafeInteger(message.version) &&
      Number(message.version) >= 0
    );
  if (message.type === "ready" || message.type === "openText" || message.type === "save")
    return Object.keys(message).length === 1;
  if (message.type === "openEvent")
    return (
      Object.keys(message).every((key) =>
        ["type", "requestId", "version", "targetId", "choose"].includes(key),
      ) &&
      Object.keys(message).length === (Object.hasOwn(message, "targetId") ? 5 : 4) &&
      Number.isSafeInteger(message.requestId) &&
      Number(message.requestId) >= 0 &&
      Number.isSafeInteger(message.version) &&
      Number(message.version) >= 0 &&
      (message.targetId === undefined ||
        (typeof message.targetId === "string" && message.targetId.length > 0)) &&
      typeof message.choose === "boolean"
    );
  if (message.type === "pictureCancel") return Object.keys(message).length === 1;
  if (message.type === "picturePick" || message.type === "pictureCommit") {
    const keys = [
      "type",
      "requestId",
      "version",
      "controlId",
      "field",
      ...(message.type === "pictureCommit" ? ["token"] : []),
    ];
    return (
      Object.keys(message).length === keys.length &&
      Object.keys(message).every((key) => keys.includes(key)) &&
      Number.isSafeInteger(message.requestId) &&
      Number(message.requestId) >= 0 &&
      Number.isSafeInteger(message.version) &&
      Number(message.version) >= 0 &&
      typeof message.controlId === "string" &&
      (message.field === "picture" || message.field === "mouseIcon") &&
      (message.type !== "pictureCommit" ||
        (typeof message.token === "string" && message.token.length > 0))
    );
  }
  if (message.type === "resolveConflict")
    return (
      Object.keys(message).length === 2 && ["source", "designer"].includes(String(message.choice))
    );
  if (message.type === "clipboard") return isClipboardMessage(message);
  const operations = message.operations;
  return (
    message.type === "edit" &&
    Object.keys(message).every((key) =>
      ["type", "requestId", "version", "operations"].includes(key),
    ) &&
    Number.isSafeInteger(message.requestId) &&
    Number(message.requestId) >= 0 &&
    Number.isSafeInteger(message.version) &&
    Number(message.version) >= 0 &&
    Array.isArray(operations) &&
    operations.length >= 1 &&
    operations.length <= 3 &&
    (isInitialMultiPageOperations(operations) ||
      (operations.length === 1 &&
        (isPropertyOperation(operations[0]) ||
          isSetGeometryOperation(operations[0]) ||
          (isStructuralOperation(operations[0]) &&
            (operations[0].type !== "addControl" ||
              (operations[0].control.type !== "MultiPage" &&
                (operations[0].control.type !== "TabStrip" ||
                  (operations[0].control.selectedIndex === 0 &&
                    operations[0].control.tabs?.length === 2))))) ||
          isParentOperation(operations[0]) ||
          isPageOperation(operations[0]) ||
          isTabOperation(operations[0]) ||
          (isProductivityOperation(operations[0]) && operations[0].type !== "pasteControls"))) ||
      (operations.length === 2 &&
        operations.every(
          (op) =>
            isPropertyOperation(op) &&
            op.type === "setFormProperty" &&
            ["width", "height"].includes(op.field) &&
            typeof op.value === "number",
        ) &&
        new Set(operations.map((op) => op.field)).size === 2) ||
      (operations.every(isGeometryOperation) &&
        operations.every((op) => op.controlId === operations[0].controlId) &&
        new Set(operations.map((op) => op.type)).size === operations.length))
  );
}
