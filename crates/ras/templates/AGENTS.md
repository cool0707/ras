# AI エージェント向けの手引き（RAS プロジェクト）

このプロジェクトは RAS（TypeScript に似た構文で書き、VBA に変換して Excel で動かす言語）で書かれています。Excel が無い環境でも、`ras` コマンドだけで「書く → 動かす → 確かめる」を回せます。

## 作業の手順

1. **ブックを知る**: `ras book dump`（シート・コード名・使用範囲・数式・名前・テーブル）、`ras book dump --sheet Sheet1 --range A1:F20`（セルの値と数式）。シートの構成やセルの内容を推測で書かない。
2. **書く**: `src/*.ras`（ユーザーフォームは `forms/`）。構文は `ras cheatsheet` で確かめる。
3. **検査する**: `ras check`。エラーのコードは `ras explain E0301` で説明と直し方を見られる。`ras check --fix` は機械的に直せる警告を直す。`--message-format json` で JSON。
4. **動かす**: `ras run Module.func`。`debug.print` の出力、実行時エラー（番号・説明・`.ras` の位置）、変更されたセルの差分が出る。ダイアログの応答は `--msgbox vbYes --inputbox "..."`。
5. **確かめる**: `ras test`（`test/*.test.ras`）。実行結果のブックは `ras book diff 元のブック build/debug/workbooks/ブック名` で比べる。人に見せるときは `ras run ... --view` / `ras book view`（グリッド表示の HTML）。

`ras run` と `ras test` は Excel の代わりに `ras.toml` の `workbook` を読み込んだ「ブックの代役」で動きます。元のブックは書き換えません。代役に無い機能を使っていないかは `ras check --target stand-in` で分かります。

## 書き方の決まり

- Excel / VBA のメンバーとグローバルは camelCase（`ws.cells(r, 1).value`、`msgBox`、`xlUp`、`vbYes`）。VBA の綴り（`Cells`、`MsgBox`）は警告。
- シートはコード名（`Sheet1`）で参照する。`range("A1")` / `cells(...)` / `activeSheet` をシートなしで使わない（選択中のシートに依存する）。
- 数値は `int`（整数）と `double`（実数）。`number` は無い。整数どうしの `/` は切り捨て。
- 条件は `boolean` だけ（`if (n != 0)`、`if (s != "")`、`if (obj != null)`）。
- `const` を基本にし、再代入するものだけ `let`。
- 名前は大文字小文字だけの違いでは区別できない（`Rect` という型と `rect` という変数は同時に使えない）。
- アロー関数の中から外側の変数を書き換えない（`for...of` か `reduce` を使う）。テストファイルの中と、`on` に渡すイベントハンドラの中（モジュール変数だけ）では書き換えられる。
- セルを 1 つずつ読み書きするループはそのまま書いてよい（配列との一括転送に変換される）。画面更新などを止めるには関数に `@fast` を付ける。
- シートを増やす・消す・名前を変えるときは、`objects/` のファイルを直接作ったり消したりせず、`ras sheet add` / `copy` / `remove` / `rename` を使う（ブックのシートとファイルを一緒に変える）。
- シートと `ThisWorkbook` のコードは `objects/<シートのコード名>.ras` / `objects/ThisWorkbook.ras` に書く（ファイルがそのモジュール全体になる。`this` がそのシート・ブック）。イベントは、そのファイルの中で `@on("Worksheet.Change")` / `@on("Workbook.Open")` を付けた関数で受け取る。`src/` のファイルには `@on` を書けない。`Worksheet_Change` という名前の関数は書かない。
- イベントハンドラを実行時に付けるとき（同じ処理を多くのコントロールに付ける、条件で付け外しする）は `対象.on("イベント", ハンドラ)`（`addEventListener` も同じ）。対象はコントロール・シート・ブック・`application`、ハンドラはアロー関数か関数の名前。戻り値の `off()` で外す。`this.controls` の要素は `instanceof MSForms.TextBox` などで絞り込んでから呼ぶ。図形には `shape.on("Click", マクロにする関数の名前)`。決まったコントロールのイベントは `@on("btnOK.Click")` で書く。
- エラー処理は `try` / `catch (e)`（`e.code`・`e.description`）。VBA の `On Error Resume Next` は `try { ... } catch { resume; }`。

## 使えないもの

- TypeScript の `number`・`undefined`・`?.`・スプレッド（`...`）・`async` / `await`・クラスの継承（`extends` / `super`）・関数の中の `function`・オーバーロード。
- VBA の構文（`Dim`、`Set`、`End If`、`GoTo`、`ReDim`）。
- Excel なしの実行では: 手書きの VBA の呼び出し、ファイルの選択ダイアログ、`Application.OnTime`、`Scripting.FileSystemObject`・`Scripting.Dictionary`・`VBScript.RegExp` 以外の COM オブジェクト、ユーザーフォームの表示。テストでは `mockRun` / `mockGetOpenFilename` / `mockCreateObject`、`ras run` では `--replace-run` / `--open-file` / `--replace-object` で置き換えられる。

## テストの書き方

```typescript
// test/report.test.ras
import { rebuildReport } from "../src/report";

describe("rebuildReport", () => {
    beforeEach(() => useWorkbook("test/fixtures/sales.xlsx"));

    test("C 列に 2 倍を書く", () => {
        rebuildReport();
        expect(Sheet1.cells(2, 3).value).toBe(4);
    });
});
```

- ブックはテストごとに元に戻る。`describe` / `test` の名前は文字列リテラルで書く。
- `ras test -t "名前"` で絞り込み、`ras test --reporter json` で結果を JSON で得られる。失敗したテストのブックは `build/test/workbooks/` に保存される。

## やってはいけないこと

- `build/` の中のファイルを直接編集しない（`ras build` / `ras run` / `ras test` が作り直す）。
- ブック（`.xlsm`）の VBA を直接書き換えない。`ras pack` がブックに書き込む（Excel でブックを閉じてから実行する）。
- 実行時エラーを、確かめずに推測で直さない。`ras run` で再現し、`ras book dump` でセルの中身を見る。
