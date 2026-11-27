# RAS 早見表

RAS は TypeScript に似た構文で書き、VBA（`.bas` / `.cls`）に変換して Excel で動かす言語です。詳しい仕様は [RAS_Language_Specification.md](RAS_Language_Specification.md)。この早見表は `ras cheatsheet` でも表示できます。コード例はすべて `ras check` を通ります。

## TypeScript と違うところ（先に読む）

- `number` はありません。整数は `int`（VBA の `Long`）、実数は `double`。`1` は `int`、`1.5` は `double`。
- 整数どうしの `/` は整数の割り算（切り捨て）です。実数にするには片方を `double` にします（`a / 2.0`、`a as double / b`）。
- 条件は `boolean` だけです。`if (count)`・`if (name)`・`if (obj)` は書けません（`count != 0`、`name != ""`、`obj != null`）。
- `struct`・タプル・配列は値型です（代入でコピーされます）。`class` のインスタンスと Excel のオブジェクトは参照です。
- アロー関数の中から、外側の変数を書き換えられません（`forEach` の中で合計を足すなど）。`for...of` か `reduce` を使います。
- 負の添字は末尾から数えません（`xs[-1]` は範囲外のエラー）。末尾からは C# と同じく `xs[^1]`、一部の取り出しは範囲 `xs[1..^1]`・`s[..3]`（終わりは含まない）、数の列挙は `for (const i of 0..n)` です。範囲・`slice`・`step(n)`・`toReversed()`・2 次元配列の `transpose()`（行と列の入れ替え）は新しい配列を返し、`reverse()`・`sort()` は配列そのものを並べ替えます。
- `undefined`・`?.`・スプレッド（`...`）・`async` / `await`・クラスの継承（`extends` / `super`）・関数の中の `function` はありません。
- Excel / VBA のメンバーとグローバルは camelCase で書きます（`ws.cells(1, 1).value`、`msgBox(...)`、`xlUp`）。VBA の綴り（`Cells`）は警告になり、`ras check --fix` で直せます。
- 名前は大文字小文字だけの違いでは区別できません（VBA が区別しないため。`Shape` という型と `shape` という変数は同時に使えません）。
- セルの値（`range.value`）は `cellValue` です（エラー値 `#N/A` などのことがある `any`）。`isError()` で確かめる（`if (isError(v)) continue;`）か `Excel.errorToNull(v)` を通すと `any` になります。そのまま比較・計算・型変換すると警告 W0313（エラー値のとき実行時エラーになるため）。型の付いた変数に入れると実行時に変換されます（空のセル → `0` / `""`）。
- VBA の `[A1]`・`[SUM(A1:A10)]`（`Evaluate`）は `` fx`SUM(A1:A10)` `` と書きます（`[...]` は配列です）。
- シートはコード名（`Sheet1`）で参照します。`range("A1")` のようにシートを付けない書き方は、選択中のシートに依存するので警告になります。

## 型

| RAS | VBA | 備考 |
|---|---|---|
| `int` / `long` / `short` / `byte` | `Long` / `LongLong` / `Integer` / `Byte` | 整数リテラルは `int`（`0xFF`、列番号 `0cAA` = 27 も）。`long` は 64bit Office だけ |
| `double` / `float` | `Double` / `Single` | 小数リテラルは `double` |
| `currency` / `decimal` | `Currency` / `Variant`（`CDec`） | `1.5m` は `decimal` |
| `string` / `boolean` / `date` | `String` / `Boolean` / `Date` | |
| `any` / `object` | `Variant` / `Object` | `object` のメンバーは遅延バインディング |
| `cellValue` | `Variant` | セルの値。エラー値のことがある `any`（`isError` で絞り込む） |
| `T[]` / `T[][]` | 動的配列 / 2 次元配列 | 添字は 0 始まり |
| `Map<K, V>` | `Scripting.Dictionary` 相当 | キーは数値・文字列・日付 |
| `struct` / `{ a: A; b: B }` / `[A, B]` | `Type` | 値型 |
| `A \| B` | `Variant` / `Object` | `instanceof` / `typeof` で絞り込む |
| `Excel.Range` など | `Range` など | 同梱の型定義 |

## 変数・制御構文

```typescript
export function basics(): string {
    const limit = 10;                 // 再代入しないものは const
    let total = 0;                    // int
    let ratio = 0.0;                  // double
    let name: string = "";
    const names: string[] = ["a", "b"];

    for (let i = 0; i < limit; i++) {
        if (i % 2 == 0) {
            continue;
        }
        total += i;
    }
    for (const n of names) {
        name += n.toUpperCase();
    }
    for (const i of 0..names.length) {   // i = 0, 1, ..., length - 1
        name += names[^(i + 1)][..1];     // 末尾から 1 つずつ、その先頭の 1 文字
    }
    while (total > 100) {
        total = total / 2;            // 整数の割り算
    }
    ratio = total as double / limit;  // 実数の割り算
    switch (names.length) {
        case 0:
            return "none";
        case 1:
        case 2:
            return `${name}: ${total} (${ratio})`;   // テンプレート文字列
        default:
            return name;
    }
}
```

## 関数・拡張メソッド・ラムダ

```typescript
/** ドキュメントコメント（VBA のコメントとエディタのヒントになる）。 */
export function area(width: double, height: double = 1.0): double {
    return width * height;
}

// byref: 呼び出し元の変数を書き換える
function increment(byref n: int, by: int = 1) {
    n += by;
}

// 拡張メソッド: 第 1 引数を this にすると ws.lastRow() と呼べる
export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.cells(this.rows.count, col).end(xlUp).row;
}

export function lambdas(): int {
    const xs = [3, 1, 2];
    const doubled = xs.map(x => x * 2).filter(x => x > 2);
    const sum = doubled.reduce((acc, x) => acc + x, 0);
    let count = 0;
    increment(count, by: 2);          // 名前付き引数
    return sum + count + area(2.0) as int;
}
```

## 構造体・タプル・クラス・列挙型

```typescript
struct Point { x: int; y: int; label?: string; }

enum Color { red, green = 5, blue }

interface Shape {
    area(): double;
}

export class Rect implements Shape {
    private w: double;
    private h: double;
    constructor(w: double, h: double) {
        this.w = w;
        this.h = h;
    }
    area(): double {
        return this.w * this.h;
    }
}

function divmod(a: int, b: int): [int, int] {
    return [a / b, a % b];
}

export function values(): double {
    const p: Point = { x: 1, y: 2 };
    const q = p;                      // コピー（値型）
    const [quot, rem] = divmod(17, 5);
    const { x, y } = p;
    const box: Shape = new Rect(2.0, 3.0);
    const ages = new Map<string, int>();
    ages.set("a", quot + rem + x + y + q.x);
    const c = Color.green;
    return box.area() + ages.get("a") + (c as int);
}
```

## エラー処理

```typescript
export function errors(path: string): string {
    try {
        if (path == "") {
            throw new Error("パスが空です");
        }
        const book = workbooks.open(path);
        book.close(false);
        return "ok";
    } catch (e) {
        return `エラー ${e.code}: ${e.description}`;   // VBA の Err.Number / Description
    } finally {
        status("");
    }
}

export function ignoreErrors() {
    // On Error Resume Next: catch には resume; だけを書く
    try {
        Sheet1.showAllData();
    } catch {
        resume;
    }
}
```

## Excel の操作

```typescript
// 関数全体で画面更新・イベント・自動計算を止める（終わると元に戻る）
@fast
export function fill() {
    const ws = Sheet1;                              // シートはコード名で参照する
    const last = ws.cells(ws.rows.count, 1).end(xlUp).row;
    for (let r = 2; r <= last; r++) {
        // セルを 1 つずつ読み書きするループは、Range と配列の一括転送に変換される
        // セルのエラー値（#N/A など）は計算・比較すると実行時エラーになるので、null にしてから使う
        ws.cells(r, 3).value = (Excel.errorToNull(ws.cells(r, 2).value) ?? 0) * 2;
    }
    ws.cells(1, 0cAA).value = "AA 列";            // 0cAA は列番号 27（A = 1。0x のような列の文字のリテラル）
    ws.range("C1").value = "2 倍";
    ws.range("C1").font.bold = true;
    ws.range("C2:C" + last).numberFormat = "#,##0";
    // 上・下・左・右に行・列を足した範囲（extend(up, down, left, right)。名前付きでも、負の数で縮める）
    ws.range("A2:C2").extend(up: 1, down: last - 2).borders.lineStyle = xlContinuous;
}

export function readAll(): double {
    const ws = Sheet1;
    const data = ws.range("A1").currentRegion.value;   // 2 次元配列（cellValue。添字は 0 始まり）
    let total = 0.0;
    for (let r = 1; r < data.length; r++) {
        const v = data[r][1];
        if (isError(v)) continue;                       // エラー値（#N/A など）を飛ばす
        total += v as double;
    }
    const names = worksheets.map(s => s.name);         // コレクションは配列のように使える
    debug.print(names.join(","), total);
    return total;
}

export function ask(): string {
    if (msgBox("続けますか", vbYesNo + vbQuestion) != vbYes) {
        return "";
    }
    return inputBox("月を入力してください", "入力", "2026/4");
}

// Excel の数式を評価する（VBA の [SUM(A1:A10)] / Evaluate）。${...} で値を埋め込める
export function sums(code: string): double {
    const ws = Sheet1;
    const total = fx`SUM(A1:A10)`;
    const hit = ws.fx`SUMIF(A:A, "${code}", B:B)`;
    if (isError(total) || isError(hit)) return 0;      // 数式の結果もエラー値のことがある
    return total + hit;
}
```

シートと `ThisWorkbook` のコード（イベントなど）は、`objects/` のファイルに書きます。ファイルがそのシートのモジュール全体になります。

```typescript
// objects/Sheet1.ras（コード名 Sheet1 のシートのモジュール）
// イベントは @on("Worksheet.イベント")。このシートは this
@on("Worksheet.Change")
function changed(target: Excel.Range) {
    if (target.column == 1) {
        this.range("D1").value = target.address;
    }
}

// export した関数は、ほかのモジュールから Sheet1.clearInput() で呼べる
export function clearInput() {
    this.range("A2:C100").clearContents();
}
```

```typescript
// objects/ThisWorkbook.ras（ブックのモジュール）
@on("Workbook.Open")
function opened() {
    status("");
    // 実行時にイベントハンドラを付ける: on(イベント, ハンドラ)。addEventListener も同じ
    const l = Sheet1.on("SelectionChange", (target) => { status(target.address); });
    this.on("BeforeClose", () => msgBox("閉じますか？", vbYesNo) == vbYes);   // false を返すと取り消す
    l.off();                                                                   // 外す
}
```

`on` は、コントロール（`this.btnOK.on("Click", () => { ... })`）・シート・ブック・`application` に使えます。配列・コレクションにも使えます（`this.controls.on("Change", (c) => { ... })`、`[this.txtA, this.txtB].on(...)`。ハンドラは先頭の引数でそのメンバーを受け取ります）。フォームのコードビハインドで付けたものは、フォームが閉じると外れます。図形には、マクロにする関数の名前を渡します（`shape.on("Click", reload)`。`OnAction` に設定されます）。決まったコントロールのイベントは `@on("btnOK.Click")` で書きます。

## ユーザーフォーム（`forms/<名前>.form.json` と `.ras`）

```typescript
// forms/frmInput.ras（コードビハインド。this がフォーム）
function constructor(name: string, limit: int = 20) {   // new frmInput(...) の引数
    this.txtName.text = name;
}

@on("btnOK.Click")
function ok() {
    this.close(this.txtName.text);    // showModal() がこの値を返す
}

@on("btnCancel.Click")
function cancel() {
    this.close();                     // 値なし（× ボタンも）: 型の既定値（""）
}
```

```typescript
// src/ の関数から
export function askName() {
    const name = new frmInput("山田").showModal();   // 閉じるとアンロードされる
    if (name == "") return;                          // キャンセル
    const f: frmInput = new frmInput(name, 10);      // フォーム名は型としても書ける
    f.show();
}
```

```typescript
// forms/frmEntry.ras（入力のバリデーション。zod 風の規則を付けると、入力の確定時に検証し、
// エラーを <コントロール名>Error のラベルに表示する。無ければコントロールの下に透明なラベルを作る）
import * as z from "validation";

const form = new z.FormValidator();

function constructor() {
    form.rule(this.txtName, z.string().required("名前を入力してください").max(20));
    // z.int().min(0).max(150) / z.string().email() / .regex(…) / .refine(v => …, "…") など。
    // 前の入力で規則が変わるときは、Change のハンドラで rule / remove を呼び直す
}

@on("btnOK.Click")
function submit() {
    if (!form.validateAll()) return;  // すべて検証し、最初のエラーへフォーカス
    this.close(this.txtName.text);
}
```

```typescript
// forms/frmWork.ras（プログレスバー。.form.json に
// { "type": "ProgressBar", "name": "pbWork", "caption": "処理中", "value": 0 } を置く。
// キャプション・割合・バーを表示し、バーの既定の色は Windows のアクセント カラー）
@on("btnOK.Click")
function run() {
    const total = 200;
    this.pbWork.max = total;                  // min / max（既定 0 / 100）
    for (let i = 1; i <= total; i++) {
        this.pbWork.value = i;                // バーと右上の割合（this.pbWork.text は "56%"）が動く
    }
    this.pbWork.caption = "完了";
    this.pbWork.barColor = rgb(16, 124, 16);  // trackColor / foreColor / valueColor / showValue / visible …
}
```

```typescript
// forms/frmPick.ras（ファイル・フォルダーの選択欄。.form.json に
// { "type": "FilePicker", "name": "fpBook", "title": "ブックを選択",
//   "filters": [{ "description": "Excel ブック", "extensions": "*.xlsx; *.xlsm" }] } を置く。
// 欄の右の「...」で Application.FileDialog を出す。欄には貼り付けもできる）
@on("fpBook.Change")                          // 欄のイベント（選んだ・貼り付けた）
function changed() {
    this.btnOK.enabled = this.fpBook.path != "";
}

@on("btnOK.Click")
function ok() {
    this.fpBook.dialogType = 4;               // FolderPicker（1 Open / 2 SaveAs / 3 FilePicker）
    const chosen = this.fpBook.browse();      // 選んだパスの配列。キャンセルは空の配列
    if (chosen.length == 0) return;
    this.close(this.fpBook.paths.join(";"));  // paths: 欄のパス（"…" は外す）
}
```

表示スケール（DPI）が 100% 以外でも、フォームの内側の大きさはデザインどおりになります（読み込み時に合わせます）。

フォームのデザインを作らずに、コードだけでダイアログを組み立てることもできます（`import … from "dialog"`。仕様 §14.11.15）。

```typescript
import { Dialog, choose, Progress } from "dialog";

export function confirmSave(): string {
    // 任意のボタン: 押したボタンの番号（0 から）。× と Esc（cancel のボタン）は -1
    const r = choose("変更を保存しますか？", ["変更を確定", "別名で確定", "キャンセル"], title: "保存", cancel: 2);
    if (r < 0) return "";
    return r == 0 ? "save" : "saveAs";
}

export function askExport() {
    const d = new Dialog("エクスポート", "出力の設定を選んでください");   // 上から 1 列に並ぶ。ボタンは OK・キャンセル
    const kind = d.select("形式", ["CSV", "TSV"]);                         // text / password / combo / list / radio / folder / label も
    const file = d.file("出力先", dialogType: 2).required("出力先を選んでください");   // rule(z.string()…) も付けられる
    const head = d.checkbox("見出しを出力する", true);
    if (d.show() != 0) return;                     // OK 以外（キャンセル・×）。規則に合わないと OK で閉じない
    debug.print(kind.text, kind.index, file.text, head.checked);
}

export function longWork() {
    const p = new Progress("集計中", max: 500);    // モードレスですぐ表示。「中止」か × で cancelled
    for (let i = 1; i <= 500; i++) {
        if (p.cancelled) break;
        p.value = i;
        p.message = `${i} / 500`;
    }
    p.close();
}
```

## ファイルの読み書き（`Open` / `Print #` / `Close`）

```typescript
export function copyLines(src: string, dest: string) {
    const inp = freeFile();
    fileOpen(inp, src, forInput);                     // Open src For Input As #inp
    const out = freeFile();
    fileOpen(out, dest, forAppend);                   // For Output / Append / Binary / Random
    while (!eof(inp)) {
        const line = lineInput(inp);                  // Line Input #inp, line
        printLine(out, line, "\t", line.length);     // Print #out, line; vbTab; …（print は改行しない）
    }
    writeLine(out, "合計", 3, true);                  // Write #out, "合計", 3, True
    fileClose(inp, out);                              // Close #inp, #out（引数なしで全部）
    if (dir(dest + ".bak") != "") kill(dest + ".bak");
    fileCopy(dest, dest + ".bak");
}
```

ほかに `input(n, a, b)`（`Input #`）、`inputString(n, 文字数)`、`fileGet` / `filePut`（`Get` / `Put`）、`seek`、`lock` / `unlock`、`fileWidth`、`rename(古い, 新しい)`（`Name … As …`）、`reset()`。`fileOpen` のモード（`forInput` など）・`accessRead`・`lockShared` は定数で書きます。

## テスト（`test/*.test.ras`）

```typescript
// test/report.test.ras
describe("fill", () => {
    beforeEach(() => {
        Sheet1.range("A1:B3").value = [["名前", "数量"], ["a", "1"], ["b", "2"]];
    });

    test("C 列に 2 倍を書く", () => {
        Sheet1.cells(2, 3).value = Sheet1.cells(2, 2).value * 2;
        expect(Sheet1.cells(2, 3).value).toBe(2);
        expect(Sheet1.range("A2:A3").value).toEqual([["a"], ["b"]]);
    });

    test("確認には「いいえ」と答える", () => {
        const box = mockMsgBox(() => vbNo);
        expect(msgBox("続けますか", vbYesNo)).toBe(vbNo);
        expect(box.calls.length).toBe(1);
    });

    test.each([[1, 2], [2, 4]])("%i の 2 倍は %i", (n: int, want: int) => {
        expect(n * 2).toBe(want);
    });
});
```

- ブックはテストごとに元に戻ります。別のブックで試すには `useWorkbook("test/fixtures/x.xlsx")`。
- `src/` の関数は `Module.f()` のようにモジュール名で呼ぶか、`import { f } from "src/module";` で使います（`export` されたものだけ。パスはプロジェクトのフォルダーから）。

## コマンド（Excel が無くても動くもの）

| コマンド | 内容 |
|---|---|
| `ras check` | 型チェックとリント（`--fix` で機械的に直せる警告を直す、`--message-format json` で JSON） |
| `ras check --target stand-in` | Excel なしの実行で使えない API も警告する |
| `ras explain E0301` | エラー・警告のコードの説明と直し方 |
| `ras run Module.func` | Excel なしで実行し、出力・エラー・変更されたセルを表示（`--json`、`--view` でグリッド表示の HTML） |
| `ras test` | `test/**/*.test.ras` を実行（`-t 名前`、`--reporter json`） |
| `ras book dump Book.xlsm` | シート・数式・名前の一覧（`--sheet Sheet1 --range A1:F20` でセル） |
| `ras book diff a.xlsx b.xlsx` | 2 つのブックの差分 |
| `ras book view Book.xlsx` | シートをグリッドで表示する HTML を作る |
| `ras book edit Book.xlsm --sheet Sheet1 --set B2=100` | セルを編集してブックに書き込む（`--clear A1:C3`、`--insert 3:4`、`--delete B:B`、`--op '<JSON>'`） |
| `ras sheet list` / `add 集計` / `copy Sheet1` / `rename Sheet1 --code-name Sales` / `remove Sheet3` | シートの一覧・追加・複製・名前の変更・削除（ブックと `objects/` のファイルを一緒に変える） |
| `ras build` / `ras pack` | VBA を出力する / ブックに書き込む |
