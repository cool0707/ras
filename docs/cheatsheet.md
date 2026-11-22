# RAS 早見表

RAS は TypeScript に似た構文で書き、VBA（`.bas` / `.cls`）に変換して Excel で動かす言語です。詳しい仕様は [RAS_Language_Specification.md](RAS_Language_Specification.md)。この早見表は `ras cheatsheet` でも表示できます。コード例はすべて `ras check` を通ります。

## TypeScript と違うところ（先に読む）

- `number` はありません。整数は `int`（VBA の `Long`）、実数は `double`。`1` は `int`、`1.5` は `double`。
- 整数どうしの `/` は整数の割り算（切り捨て）です。実数にするには片方を `double` にします（`a / 2.0`、`a as double / b`）。
- 条件は `boolean` だけです。`if (count)`・`if (name)`・`if (obj)` は書けません（`count != 0`、`name != ""`、`obj != null`）。
- `struct`・タプル・配列は値型です（代入でコピーされます）。`class` のインスタンスと Excel のオブジェクトは参照です。
- アロー関数の中から、外側の変数を書き換えられません（`forEach` の中で合計を足すなど）。`for...of` か `reduce` を使います。
- `undefined`・`?.`・スプレッド（`...`）・`async` / `await`・クラスの継承（`extends` / `super`）・関数の中の `function` はありません。
- Excel / VBA のメンバーとグローバルは camelCase で書きます（`ws.cells(1, 1).value`、`msgBox(...)`、`xlUp`）。VBA の綴り（`Cells`）は警告になり、`ras check --fix` で直せます。
- 名前は大文字小文字だけの違いでは区別できません（VBA が区別しないため。`Shape` という型と `shape` という変数は同時に使えません）。
- セルの値（`range.value`）は `any` です。型の付いた変数に入れると実行時に変換されます（空のセル → `0` / `""`）。
- シートはコード名（`Sheet1`）で参照します。`range("A1")` のようにシートを付けない書き方は、選択中のシートに依存するので警告になります。

## 型

| RAS | VBA | 備考 |
|---|---|---|
| `int` / `long` / `short` / `byte` | `Long` / `LongLong` / `Integer` / `Byte` | 整数リテラルは `int`。`long` は 64bit Office だけ |
| `double` / `float` | `Double` / `Single` | 小数リテラルは `double` |
| `currency` / `decimal` | `Currency` / `Variant`（`CDec`） | `1.5m` は `decimal` |
| `string` / `boolean` / `date` | `String` / `Boolean` / `Date` | |
| `any` / `object` | `Variant` / `Object` | `object` のメンバーは遅延バインディング |
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
        ws.cells(r, 3).value = ws.cells(r, 2).value * 2;
    }
    ws.range("C1").value = "2 倍";
    ws.range("C1").font.bold = true;
    ws.range("C2:C" + last).numberFormat = "#,##0";
}

export function readAll(): double {
    const ws = Sheet1;
    const data = ws.range("A1").currentRegion.value;   // 2 次元配列（any[][]。添字は 0 始まり）
    let total = 0.0;
    for (let r = 1; r < data.length; r++) {
        total += data[r][1] as double;
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
- `src/` の関数は `import { f } from "../src/module";` で使います（`export` されたものだけ）。

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
| `ras book edit Book.xlsm --sheet Sheet1 --op '{"op":"addControl","control":"button","range":"F2:G3","set":{"text":"実行","macro":"Main.run"}}'` | ボタンなどをシートに置く（`addShape` / `addActiveX` / `placeObject` / `setObject` / `deleteObject`） |
| `ras sheet list` / `add 集計` / `copy Sheet1` / `rename Sheet1 --code-name Sales` / `remove Sheet3` | シートの一覧・追加・複製・名前の変更・削除（ブックと `objects/` のファイルを一緒に変える） |
| `ras build` / `ras pack` | VBA を出力する / ブックに書き込む |
