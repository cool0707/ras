# RAS

TypeScript ライクな構文から VBA（`.bas` / `.cls`）へトランスパイルする Better VBA 言語。

- 言語仕様: [docs/RAS_Language_Specification.md](docs/RAS_Language_Specification.md)
- 今後実装する機能: [docs/ROADMAP.md](docs/ROADMAP.md)
- 対象: Windows 版 Excel（Office 2007 以降、32bit / 64bit）

```typescript
// src/report.ras
@fast
export function rebuildReport() {
    const ws = Sheet1;
    const last = ws.lastRow();
    for (let r = 2; r <= last; r++) {
        ws.cells(r, 3).value = r * 2;      // ループごと Range ⇔ 配列の一括転送に変換される
    }
}

export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.cells(this.rows.count, col).end(xlUp).row;
}

// コレクションは配列と同じように扱える（For Each のループに展開される）
export function visibleSheetNames(): string[] {
    return worksheets.filter(ws => ws.visible == xlSheetVisible).map(ws => ws.name);
}
```

Excel / VBA のメンバーとグローバルは、JavaScript の命名規則（camelCase）で書きます（`ws.cells(r, 1).value`、`msgBox(...)`）。生成される VBA には本来の綴り（`ws.Cells(r, 1).Value`）が出力されます。VBA の綴りのままでも動きますが警告になり、`ras check --fix` で一括して書き換えられます（仕様 §15.3.1）。

## ダウンロード

ビルド済みのバイナリと VS Code 拡張は GitHub Actions の成果物として取得できます。

1. リポジトリの **Actions** タブで、ワークフロー **CI** の成功した実行を開く
2. ページ下部の **Artifacts** からダウンロードする

| 成果物 | 内容 |
|---|---|
| `ras-windows-x64` | `ras.exe`、`ras_com.node`、`ras_book.node`（Windows x64） |
| `ras-windows-arm64` | `ras.exe`、`ras_com.node`、`ras_book.node`（Windows arm64） |
| `ras-linux-x64` | `ras` と `ras_book.node`（Linux x64。ビルド・チェック・オフライン書き込み・Excel を使わないデバッグ用） |
| `ras-macos-arm64` | `ras` と `ras_book.node`（macOS arm64。同上） |
| `ras-vscode` | VS Code 拡張（`.vsix`） |

`ras` は依存ランタイムのない単一バイナリです。PATH の通った場所に置くか、VS Code の設定 `ras.path` で場所を指定します。
`ras run` と `ras test` は、プログラムを JavaScript に変換して実行するため、Node.js（PATH の `node`。無ければ VS Code の実行ファイルを使います）と、`ras` と同じフォルダーに置いた `ras_book.node` が必要です。
VS Code 拡張は「拡張機能」ビューの `…` → 「VSIX からのインストール」で入れられます。

## 使い方

```sh
ras init myproject --workbook Book.xlsm   # ras.toml・src/main.ras・空のマクロ有効ブックを作成
cd myproject
ras check                                 # 型チェックとリント
ras check --fix                           # 機械的に直せる警告（VBA の綴りの名前など）をソースに書き戻す
ras build                                 # build/ に .bas / .cls / .rasmap を出力
ras pack                                  # Excel を起動せず、Book.xlsm の VBA プロジェクトへ直接書き込む
ras watch                                 # 変更を監視して再ビルド（ブックが閉じていれば pack も行う）
ras unpack                                # ブック内の手書き VBA を vba/ に取り出す
ras sync                                  # Excel で開いているブックへ反映する（Windows のみ）
ras form new frmInput                     # ユーザーフォームの雛形（forms/frmInput.form.json と .ras）を作成
ras form pull                             # ブックにあるユーザーフォームを forms/ に取り出す（コードは RAS に変換）
ras fmt                                   # 整形
ras build --target js                     # デバッグ実行用の JavaScript を build/debug/ に出力
ras run Report.rebuildReport              # 関数を Excel なしで実行し、出力・エラー・変更されたセルを表示
ras test                                  # test/**/*.test.ras のテストを Excel なしで実行
ras book dump Book.xlsm                   # ブックの中身（シート・数式・名前・テーブル）をテキストで出力
ras book diff before.xlsm after.xlsm      # 2 つのブックのセルと構造の差分
ras book view Book.xlsm                   # シートを Excel のようなグリッドで表示する HTML を作る
ras book edit Book.xlsm --set B2=100      # セルを Excel と同じように編集してファイルに書き込む（入力・クリア・コピー・行と列とセルの挿入と削除）
ras explain E0301                         # エラー・警告のコードの説明と直し方
ras cheatsheet                            # 構文と Excel の操作の早見表
```

サンプルプロジェクトは [examples/sales-report](examples/sales-report) にあります（`ras --project examples/sales-report build`）。

### 既存のブックの取り込み

`ras sync`（`ras watch`、VS Code の保存時の同期・「今すぐ同期」）と `ras pack` は、書き込む前にブックの VBA を読み、RAS がまだ管理していない手書きのコードを RAS に変換してプロジェクトに取り込みます（仕様 §14.7.1）。

| ブックにあるもの | 取り込み先 |
|---|---|
| 標準モジュール `Module1` | `src/Module1.ras` |
| シート `Sheet1` / `ThisWorkbook` のコード | `objects/Sheet1.ras` / `objects/ThisWorkbook.ras`（`Worksheet_Change` などは `@on("Worksheet.Change")` のハンドラに、`Me` と修飾なしの `Range(...)` は `this.range(...)` に変換。`Public` のプロシージャは `export function` になり、`Sheet1.Foo` のまま呼べます） |
| VBE で作ったユーザーフォーム `frmInput` | `forms/frmInput.form.json` と `forms/frmInput.ras`（`ras form pull` と同じ） |

- 元の VBA は `vba/` に保存します。クラスモジュールは取り込みません。
- 取り込んだものは、同じ同期でブック側を RAS の生成コードに置き換えます（上書きの指定は不要です）。シート・`ThisWorkbook` のモジュールも、モジュール全体を `objects/` のファイルから生成したコードに置き換えます。
- VBA から RAS に変換できなかった行（`GoTo`・`ReDim` など）は `// TODO(VBA):` のコメントとして残ります。この行が残っているファイルについては、ブックのコードを置き換えずにそのまま残します。直して `// TODO(VBA):` の行を消すと、次の同期で置き換えます。
- 変換したコードにコンパイル エラーがあると、ほかのビルド エラーと同じく同期しません（エディタに表示されるエラーを直してください）。
- RAS が生成したモジュールに、VBE でプロシージャや変数を足した場合も、次の同期（`ras pack`）で RAS に変換して、そのモジュールのファイルの末尾に取り込みます（仕様 §14.7.3）。RAS が生成したプロシージャを書き換えた・消した場合は、従来どおりドリフトとして警告します。
- 取り込みを止めるには、`ras.toml` に `[sync]` `pull-unmanaged = false` を書きます。

### シートと `ThisWorkbook` のコード（`objects/`）

RAS のファイルと VBA のモジュールは一対一に対応します（仕様 §12.1.1）。`src/` のファイルは標準モジュールに、プロジェクト直下の `objects/` のファイルは、シートと `ThisWorkbook` のモジュール（VBE の「Microsoft Excel Objects」）になります。ファイルがそのモジュール全体になります（仕様 §12.2）。

```typescript
// objects/Sheet1.ras（コード名 Sheet1 のシートのモジュール）
@on("Worksheet.Change")
function onChange(target: Excel.Range) {
    this.range("D1").value = target.address;      // this はこのシート（VBA の Me）
}

/** ほかのモジュールから Sheet1.clearInput() で呼べる。 */
export function clearInput() {
    this.range("A2:C100").clearContents();
}
```

- ファイル名はシートのコード名です（シートの見出しの名前ではなく、VBE のプロジェクトに表示される名前）。ブックにそのコード名のシートが無いとエラーになります。ブックのモジュールは `objects/ThisWorkbook.ras` です。
- イベントは、シートのファイルでは `@on("Worksheet.<イベント>")`、`ThisWorkbook.ras` では `@on("Workbook.<イベント>")` で受け取ります。`src/` のファイルには書けません。
- 実行時に付けるときは、どのファイルからでも `Sheet1.on("Change", (target) => { ... })` / `application.on("SheetActivate", ...)` / `this.CommandButton1.on("Click", ...)` と書けます（`addEventListener` も同じ。戻り値の `off()` で外します。仕様 §12.2.3）。図形やフォーム コントロールのボタンには `shape.on("Click", reload)` でマクロを登録します（`OnAction`）。
- VBA がドキュメントモジュールに置けないもの（`export` した変数・定数、構造体・クラス、`@udf` など）はエラーになります。共有するものは `src/` のモジュールに置いて `import` します。
- シートの追加・複製・削除・名前の変更は、`ras sheet`（VS Code では「RAS: シートを追加」など）で、ブックと `objects/` のファイルに一緒に行えます（仕様 §14.16）。

  ```sh
  ras sheet list                                   # シートの一覧（見出しの名前・コード名・ファイル）
  ras sheet add 集計 --after Sheet1                # シートを追加し、objects/Sheet3.ras を作る
  ras sheet copy Sheet1 4月                        # シートを複製し、objects/Sheet1.ras も複製する
  ras sheet rename Sheet1 --name 売上 --code-name Sales   # ファイル名とソースの参照（Sheet1.…）も変える
  ras sheet remove Sheet3                          # コードや参照があると理由を表示して止まる（--force で削除）
  ```

  ブックが閉じていればブックのファイルを書き換え、Excel で開いていれば Excel を操作します（Windows のみ）。
- 同期（`ras sync`、`ras watch`、VS Code）では、ブックの各シートと `ThisWorkbook` について、ファイル `objects/Sheet1.ras` / `objects/ThisWorkbook.ras` が無ければ作成します（仕様 §14.7.2。`[sync]` `sheet-files = false` で止められます）。作成されるファイルは書き方の例のコメントだけで、コードを書くまではブックのモジュールを変えません。一般に、関数・変数などのコードが無い `.ras` ファイルからは VBA のモジュールを生成しません。

`ras pack` はブックが Excel で開かれていると書き込みません（閉じてから実行してください）。
ブックに RAS が生成したものではない同名のモジュールがある場合や、手書きのイベントプロシージャ（`Workbook_Open` など）と生成コードの名前が重なる場合も、書き込まずにエラーにします（前者は `--overwrite` で置き換えられます）。
置き換え前のブックは `build/backup/` に保存されます。

### ユーザーフォーム

フォームのデザインは `forms/<名前>.form.json`（FormSpec）、イベント処理などのコードは RAS で `forms/<名前>.ras` に書きます（仕様 §14.11）。VS Code 拡張のデザイナーで編集でき、`ras pack` / `ras sync` がブックへ書き込みます。

```typescript
// forms/frmInput.ras — `this` がフォーム、コントロールは this.<名前>
@on("btnOK.Click")
function ok() {
    if (this.txtName.text == "") {
        msgBox("名前を入力してください");
        return;
    }
    unload(this);
}

@on("txtName.Exit")
function leave(cancel: MSForms.ReturnBoolean) {
    cancel.value = this.txtName.text.length > 20;
}
```

- イベントは `@on("<コントロール名>.<イベント>")`（フォーム自身は `UserForm.Initialize` など）を付けた関数で受け取ります。コンパイラが `Private Sub btnOK_Click()` などのイベントプロシージャを生成します。同じ処理を多くのコントロールに付けるときは、`Initialize` のハンドラで `this.controls.on("Change", (c) => { ... })`（コレクション・配列のすべてのメンバーに付く）/ `this.btnOK.addEventListener("Click", ok)` のように実行時に付けます（仕様 §12.2.3。`WithEvents` の変数を持つクラスを生成します。フォームが閉じると外れます）。`export function` はフォームの公開メンバーになり、他のモジュールから `frmInput.show()` / `frmInput.enteredName()` のように使えます。

- `Image` の画像は `"picture": { "path": "forms/assets/logo.png" }` のように、プロジェクト内の画像ファイル（BMP / JPEG / PNG / GIF / WebP / TIFF / ICO）を指定します。画像はブックに埋め込まれます。MS-Forms がそのまま保存できる BMP（24 / 32 ビット無圧縮）と JPEG は内容のまま、それ以外は書き込み時に 24 ビットの BMP に変換して埋め込みます（透過部分はコントロールの `backColor` で塗りつぶします。元のファイルは変更しません）。
- `ComboBox` / `ListBox` の項目は `"list": ["A", "B"]`、最初の選択は `"selectedIndex": 0` で指定します。項目はデザイナの情報としては保存されないため、フォームの読み込み時に項目を設定するコード（`RAS_FormInit`）を生成してコードモジュールに加えます。
- VBE で作った既存のフォームは `ras form pull` で `forms/` に取り出せます（デザイン・コード・画像）。コードは VBA から RAS に変換して `forms/<名前>.ras` に書きます。変換できなかった行（`GoTo`・`ReDim` など）は `// TODO(VBA):` のコメントとして残り、警告として表示されます。FormSpec で表せないプロパティも警告として表示され、`forms/` から書き戻すと既定値になります。取り出した後、`ras pack --overwrite` でブックのフォームを `forms/` の内容に置き換えます。
- 以前の版の VBA のコードビハインド `forms/<名前>.bas` も、`.ras` が無ければそのまま使えます。`ras form pull --force` で取り込み直すと `.ras` に変換されます。
- ブックを Excel で開いたままでも、`ras sync`（および `ras watch`、VS Code の保存時の同期）がフォームを反映します。デザインが変わったフォームは取り込み直し、コードだけが変わったフォームはコードだけを書き換えます。
- `forms/<名前>.ras` の先頭の CodeLens「▶ Excelで実行 | デバッグ実行」（デザイナーではタイトルバーのボタン）から、フォームを直接動かせます。「Excelで実行」はブックへ同期してから Excel 本体でフォームを表示し（Windows のみ）、「デバッグ実行」はフォームを VS Code のパネルに表示して、コードビハインドをブレークポイント付きでデバッグします（次の「デバッグ」を参照）。

## Excel なしで動かす・確かめる（CLI）

Mac・Linux や Excel の無い Windows、CI、AI エージェントの作業では、`ras` だけで「書く → 動かす → 確かめる」を回せます。どれも Excel を起動せず、`ras.toml` の `workbook` を読み込んだブックの代役（下の「デバッグ」と同じもの）に対して実行し、元のブックは書き換えません。

### 実行（`ras run`）

```sh
ras run --list                                        # 実行できる関数の一覧
ras run Report.rebuildReport                          # 実行する（モジュール名は、1 つに決まれば省ける）
ras run ask --msgbox vbYes --inputbox "2026/4"        # ダイアログの応答を、表示される順に指定する
ras run Report.rebuildReport --json                   # 結果を JSON で出力する
```

`debug.print` の出力、表示したダイアログと応答、実行時エラー（VBA のエラー番号・説明・`.ras` の位置）に続けて、変更されたセルを表示します。変更されたブックは `build/debug/workbooks/` に保存します。

```
rows 3
変更されたセル（Book.xlsx → build/debug/workbooks/Book.xlsx）:
Sheet1:
  D2	（空） → 6
  D3	（空） → 10
  D4	（空） → 4
差分 3 件（構造 0、セル 3、名前 0）
```

`application.run` で呼ぶ手書き VBA と、代役に無い COM オブジェクトは、プロジェクトの関数・クラスで置き換えられます（`--replace-run Legacy=Mocks.legacy`、`--replace-object ADODB.Connection=Mocks.FakeConnection`）。`--view` を付けると、実行後のブックをグリッドで表示する HTML（下の `ras book view`）も作ります。

応答を指定しなかった `msgBox` は `vbOK`、`inputBox` は既定値（無ければキャンセル）で応答します。終了コードは、正常に終われば 0、実行時エラーで終われば 1、実行できなければ 2 です（仕様 §14.13）。

### テスト（`ras test`）

テストは `test/` の `*.test.ras` に、vitest と同じ形で書きます（仕様 §14.12）。テストのファイルとヘルパー（`test/` の `.ras`）はブックには入りません。

```typescript
// test/report.test.ras
import { rebuildReport, lastRow } from "../src/report";

describe("rebuildReport", () => {
    beforeEach(() => useWorkbook("test/fixtures/sales.xlsx"));

    test("D 列に数量の 2 倍を書く", () => {
        rebuildReport();
        expect(Sheet1.cells(2, 4).value).toBe(6);
        expect(Sheet1.range("D2:D4").value).toEqual([[6], [10], [4]]);
    });

    test("確認で「いいえ」を選ぶと何もしない", () => {
        const box = mockMsgBox(() => vbNo);
        expect(ask()).toBe("");
        expect(box.calls.length).toBe(1);
    });
});
```

```sh
ras test                          # すべてのテストを実行
ras test test/report.test.ras     # ファイル（またはフォルダー）を指定
ras test -t "rebuildReport"       # 名前（describe > test）に含まれる文字列で絞る
ras test --list                   # 実行せずに一覧を出す
ras test --reporter json          # 結果を JSON Lines で出力（junit で JUnit XML）
ras test --watch                  # 変更を監視して、影響するテストファイルを実行し直す
ras test --ci                     # test.only と、置き換えていないダイアログをエラーにする
```

- ブックの状態はテストごとに、そのテストを囲む `beforeAll` を実行し終えた時点に戻ります。失敗したテストのブックは `build/test/workbooks/<テストファイル名>/` に保存します。
- 使える API: `describe` / `test`（`it`）とその `.skip` / `.only` / `.each`、`test.todo`、`beforeAll` / `afterAll` / `beforeEach` / `afterEach`、`expect(...)` のマッチャ（`toBe`・`toEqual`・`toBeCloseTo`・大小の比較・`toBeTruthy`・`toBeFalsy`・`toBeEmpty`・`toBeNothing`・`toContain`・`toHaveLength`・`toMatch`・`toThrow` と `.not`）、`useWorkbook`、`tempDir`、`printed`、`mockMsgBox` / `mockInputBox` / `mockGetOpenFilename` / `mockGetSaveAsFilename` / `mockRun` / `mockCreateObject`。
- 終了コードは、すべて成功なら 0、失敗があれば 1、コンパイル エラーなどで実行できなければ 2 です。`ras check` はテストのファイルも検査します。

### ブックの中身を見る・比べる（`ras book`）

```sh
ras book dump Book.xlsm                                   # シートの一覧、使用範囲、数式、名前、テーブル、入力規則、条件付き書式
ras book dump Book.xlsm --sheet Sheet1 --range A1:F20     # セルの値・数式（--formats で表示形式とセルの表示も）
ras book dump Book.xlsm --sheet Sheet1 --format csv       # text（既定）/ json / csv / markdown
ras book diff before.xlsm build/debug/workbooks/Book.xlsm # セル（値・数式・表示形式）と構造の差分
```

```
Book.xlsx
シート:
  1. Sheet1（コード名 Sheet1）  使用範囲 A1:C5、値のあるセル 13 個（うち数式 4 個）
     数式: C2:C4 =B2*100（同じ形の数式 3 個）
     数式: C5 =SUM(C2:C4)
名前:
  合計 = Sheet1!$C$5
```

- 下や右へコピーされた同じ形の数式は 1 行にまとめます。大きなシートは既定で 200 行 × 50 列まで出力し、打ち切った範囲を示します（`--max-rows` / `--max-cols` / `--all`）。
- `diff` は、シートの追加・削除・名前の変更と、行・列の挿入・削除（ひとまとまりのもの。推定）も出力します。差分が無ければ終了コード 0、あれば 1 です（仕様 §14.14）。

### シートを目で見る（`ras book view`）

```sh
ras book view Book.xlsx --open                                 # グリッド表示の HTML（build/view/Book.html）を作ってブラウザーで開く
ras book view build/debug/workbooks/Book.xlsx --diff Book.xlsx # 変更されたセルを強調する
ras run Report.rebuildReport --view                            # 実行して、結果のブックを変更つきで表示する
```

値・数式・主な書式（フォント・塗りつぶし・罫線・配置・表示形式）・列幅と行の高さ・結合セル・フィルターを、Excel に近い見た目のグリッドで表示します。データのない範囲にも枠線を引き、スクロールに合わせてグリッドが続きます。図形・テキストボックス・ワードアート・SmartArt・画像（EMF も）・グラフ・埋め込みオブジェクトと、フォームコントロール・ActiveX コントロール（ボタン・チェックボックスなど）もセルの上に描きます（クリックすると名前と登録されたマクロを表示）。グラフの値は参照するセルの今の値から読むので、実行で変わった値もグラフに出ます。数式のセルを選ぶと参照先を色付きの枠で示し、「数式の表示」（Ctrl+`）でセルに数式を表示します。`--diff` を付けると、変更されたセルに印が付き、変更の一覧からそのセルへ移動できます。条件付き書式・テーブルのスタイルは描きません。

VS Code では「RAS: ブックを表示（グリッド）」（エクスプローラーのブックの右クリックにもあります）で同じ表示をパネルに出せます。シートのファイル（`objects/Sheet1.ras`）では、エディターのタイトルバーのボタンやエクスプローラーの右クリックの「RAS: シートのグリッドを表示」で、プロジェクトのブックをそのシートで開けます（`ras book view --sheet Sheet1` と同じ）。Excel を使わないデバッグ実行では、ブレークポイントやステップ実行で止まるたびに、そのときのシートをパネルに表示し、前回から変わったセルを強調します（設定 `ras.debug.sheetView`）。

### ブックを編集する（`ras book edit`、VS Code のグリッド）

VS Code のグリッド（「RAS: ブックを表示（グリッド）」「RAS: シートのグリッドを表示」）では、ブックのファイルを Excel と同じ操作で編集できます。

| 操作 | キー・マウス |
|---|---|
| セルの入力 | 文字を打つ（上書き）、F2・ダブルクリック・数式バー（今の内容を編集）。Enter / Tab で確定、Ctrl+Enter で選択範囲の全セルに入力、Alt+Enter でセル内の改行、Esc で取り消し |
| クリア | Delete / Backspace（数式と値。書式は残ります） |
| コピー・切り取り・貼り付け | Ctrl+C / Ctrl+X / Ctrl+V（数式の参照は Excel と同じにずれ、書式も貼り付きます。ほかのアプリケーションのタブ区切りテキストも貼り付けられます） |
| 下・右へコピー | Ctrl+D / Ctrl+R |
| 行・列・セルの挿入と削除 | 行番号・列見出し・セルの右クリックのメニュー、Ctrl++ / Ctrl+-（セルを選んでいるときは、動かす向きを選ぶダイアログが出ます） |
| 書式 | ツールバー（太字・斜体・下線・取り消し線・フォントの大きさ・文字の色・塗りつぶし・罫線・配置・折り返し・表示形式・書式のクリア）、Ctrl+B / Ctrl+I / Ctrl+U |
| 結合 | ツールバーの「結合」（セルを結合して中央揃え・横方向に結合・結合の解除） |
| 列の幅・行の高さ | 列見出し・行番号の境目をドラッグ（ダブルクリックで内容に合わせる）。見出しの右クリックで数値の指定・非表示・再表示 |
| 並べ替え・フィルター | ツールバーの A→Z / Z→A / フィルター（Ctrl+Shift+L）、見出しの ▾ から値を選んで絞り込み |
| 連続データ（フィルハンドル） | 選択範囲の右下の ■ をドラッグ（数値・日付・「第1回」のような番号は続きを、数式は参照をずらして入力）。ダブルクリックで隣の列の下端まで |
| 元に戻す・やり直し | Ctrl+Z / Ctrl+Y |

行・列・セルを挿入・削除すると、Excel と同じく、全シートの数式・名前の定義・結合セル・条件付き書式・入力規則・フィルター・メモ・図形やグラフの位置とグラフの参照が付いていき、編集したセルに依存する数式は再計算されます。結合セルの一部だけを動かす挿入など、Excel が拒む操作は理由を表示して何も変えません。

同じ編集をコマンドラインから行えます（AI エージェントがブックのデータを直すときにも使えます）。

```sh
ras book edit Book.xlsm --sheet Sheet1 --set B2=100 --set "C2==B2*2"   # セルに入力する（数式・日付も Excel の入力と同じに読みます）
ras book edit Book.xlsm --sheet Sheet1 --clear A1:C3                    # 数式と値をクリアする
ras book edit Book.xlsm --sheet Sheet1 --insert 3:4                     # 行を挿入する（B:C は列、B2:C3 はセル。--shift right で右へ動かす）
ras book edit Book.xlsm --sheet Sheet1 --delete B:B                     # 列を削除する
ras book edit Book.xlsm --op '{"op":"copy","sheet":"Sheet1","from":"A1:B2","to":"D1"}'   # JSON で指定する編集（copy / move / fill / values など）
ras book edit Book.xlsm --op '{"op":"format","sheet":"Sheet1","range":"A1:D1","set":{"bold":true,"fill":"#DDEEFF","border":"all"}}'   # 書式
ras book edit Book.xlsm --op '{"op":"sort","sheet":"Sheet1","range":"A1","by":"B","desc":true}'    # 並べ替え（filter / filterBy / size / hide / merge / autofill も同様）
ras book edit Book.xlsm --set A1=x -o build/edited.xlsm                 # 別のファイルに書く（元のブックは変えない）
```

- ファイルは、変わったセルだけを元のブックの XML に書き戻します。VBA プロジェクト（署名も）、図形・画像・コントロール、スタイル、変更のないシートは、元のバイト列のまま残ります。VBA の関数を呼ぶ数式など、Excel なしでは計算できない数式も、数式のまま保ちます。
- 書き込む前に、書き換えたファイルを読み戻してセルが編集の結果と一致することを確かめます。プロジェクトのブックは、前のファイルを `build/backup/` に残します。
- Excel で開いているブックは書き換えません（エラーになります。Excel で閉じてから行ってください）。
- テーブルの列をまたぐ列の挿入・削除と、ピボットテーブルのあるシートでの挿入・削除は、ファイルの書き換えでは行えません（Excel で行ってください）。
- 新しい書式は、元のスタイル（`xl/styles.xml`）の末尾に足します。既存の書式の番号は変わりません。
- 形式を選択して貼り付け、条件付き書式・入力規則の編集、複数キーの並べ替え、値の一覧以外の条件での絞り込みは、まだ行えません。`ras book view` が作る HTML（ブラウザーで開くもの）は読み取り専用です。

### エラーを調べる・AI に渡す

```sh
ras check --message-format json           # 診断を 1 行 1 件の JSON で出力（build / pack / run / test でも同じ）
ras explain E0320                         # コードの説明、典型的な原因、直し方
ras check --target stand-in               # Excel なしの実行で使えない API（代役に無いメンバーなど）も警告する
ras cheatsheet                            # 構文の早見表
```

- `ras check --target stand-in` は、ブックの代役に無いメンバー（W0510）、Excel が必要な操作（W0511）、作れない COM オブジェクト（W0512）を、実行する前に警告します。`ras.toml` の `[lint]` に `stand-in = true` と書くと、常に警告します。
- `ras init` は、AI エージェント向けの手引き `AGENTS.md`（作業の手順、書き方の決まり、使えないもの）もプロジェクトに作ります。早見表は [docs/cheatsheet.md](docs/cheatsheet.md)、仕様書の節の一覧は [llms.txt](llms.txt) にあります。

## デバッグ

VS Code 拡張を入れた状態で `.ras` ファイルにブレークポイントを置き、F5（または「RAS: デバッグを開始」）で実行します。
ステップ実行、変数・ウォッチ式の確認、例外での停止が使えます。

仕組みは VBA のデバッグではありません。プログラムを JavaScript に変換して VS Code 同梱のランタイムで動かし、Excel の操作だけを実際の Excel に対して行います（仕様 §14.10）。

- Excel の操作は、Windows では拡張に同梱の `ras_com.node` を通じて実際の Excel に対して行います。
- Mac・Linux や Excel の無い Windows では、Excel の代わりに `ras.toml` の `workbook` のブックを読み込んで、そのセルを読み書きします（拡張に同梱の `ras_book.node`。仕様 §14.10.8）。数式は xljs と同じ GridCraft の計算エンジンで計算し、セルを書き換えると Excel の自動計算と同じく再計算します。プログラムが変更したブックは、終了時に `build/debug/workbooks/` に保存します（元のブックは書き換えません）。Windows でも設定 `ras.debug.excel` を `standIn` にするとこの方法で実行できます。セルの挿入・削除、Excel の中での切り取り・コピー・貼り付け、並べ替え・フィルター、書式（保存します。描画はしません）、手動計算に対応し、プログラムの操作で起きるブック・シートのイベントは `@on` のハンドラを呼び出します。`Scripting.FileSystemObject`・`Scripting.Dictionary`・`VBScript.RegExp` も使えます（ほかの COM オブジェクトは作れません）。手書き VBA・ダイアログなど Excel そのものが必要な操作はエラーになります。
- 対象のブック（`ras.toml` の `workbook`）が Excel で開かれていればそこに接続し、開かれていなければ開きます。
- プログラムが Excel の応答を待っている間に Excel がダイアログ（手書き VBA の実行時エラー・コンパイル エラー・`MsgBox`、Excel 自身の確認）を出すと、その内容を VS Code に通知し、ダイアログを前面に出します。`launch.json` の `"dialogs": "dismiss"` を指定すると、キャンセルに相当するボタンで自動的に閉じます（仕様 §14.10.6）。
- 実行できるのは引数なしの関数とトップレベルの文、およびユーザーフォームです。ブック・シートの `@on` のイベントハンドラや UDF のように Excel から呼ばれる関数は起動できません。Excel を使わない実行では、プログラムの操作（セルの書き込み、シートの切り替え、保存など）で起きるイベントのハンドラが呼ばれ、ステップ実行できます。
- ユーザーフォームは Excel ではなく VS Code のパネルに表示され、クリックや入力でフォームのイベント（`Click`・`Change`・`Exit` など）が起き、そのハンドラをステップ実行できます。フォームの見た目はブラウザーでの近似で、キー入力・マウスのイベントは起きません。対応範囲は仕様 §14.10.7 にまとめています。
- VBA ビルドの最適化（一括転送・インライン展開など）は通らないため、実行タイミングや速度は実機と異なります。`long` / `currency` / `decimal` の精度、未対応の VBA 標準関数など、実機との違いは仕様 §14.10.5 にまとめています。

関数の上の CodeLens「▶ Excelで実行 | デバッグ実行」から、その関数を直接起動することもできます。「Excelで実行」はブックへ同期してから Excel 本体でマクロとして実行し（Windows のみ）、「デバッグ実行」は上記のデバッグ実行をその関数で開始します。フォームのコードビハインド（`forms/<名前>.ras`）では、ファイルの先頭に同じ CodeLens が出て、フォームを表示します。

`launch.json` で関数を固定することもできます。

```json
{ "type": "ras", "request": "launch", "name": "RAS: rebuildReport", "entry": "Report.rebuildReport" }
```

| コマンド | 状態 |
|---|---|
| `init` / `check` / `build` / `build --target js` / `watch` / `pack` / `unpack` / `fmt` / `lsp --stdio` | 実装済み |
| `form new` / `form check` / `form pull` | 実装済み |
| `sync` | 実装済み（Windows のみ。Excel で開いているブックへ、モジュールとユーザーフォームを反映する。`--pull` は未実装） |
| `setup` / `typegen` / `pack --via-excel` | 未実装（実行するとその旨を表示して終了） |
| `run` / `book dump` / `book diff` / `book view` / `book edit` | 実装済み（仕様 §14.13、§14.14） |
| `explain` / `cheatsheet` | 実装済み（仕様 §14.15） |
| `test` | 実装済み（仕様 §14.12。`--excel` とエディタの連携は未実装） |

## 実装状況

仕様書（Draft v0.7）に対する現状です。

| 仕様 | 状態 | 補足 |
|---|---|---|
| §2 型・数値リテラル・型変換 | 実装済み | `long` を `#If Win64` の外で使った場合のエラーは未実装（宣言の条件付き出力のみ） |
| §2.4 ジェネリクス | 実装済み | 関数・拡張メソッド・クラス。クラス内のジェネリックメソッドは未対応（エラー）。既定の型引数（`T = 型`）はクラスとインターフェイスのみ |
| §2.6 ユニオン型・型の絞り込み | 実装済み | `instanceof` / `typeof` による絞り込みはローカル変数と引数のみ（フィールド・モジュール変数は対象外）。ユニオン型の推論（条件演算子の両辺の型が異なる場合など）は行わない。同梱の型定義では `activeSheet`・シートのコレクション・`Worksheet.next`・各種 `parent`・`getOpenFilename` などをユニオン型にした（§2.6.5） |
| §3 関数・`byref`・拡張メソッド・名前付き引数 | 実装済み | 名前付き引数（`f(a: 123, c: "あいう")`）は宣言の順序に並べ替えて出力する。ラムダ・遅延バインディングの呼び出しには使えない |
| §4 ラムダ（ループ融合・クロージャクラス） | 実装済み | `map` / `filter` / `forEach` / `reduce` / `some` / `every` / `find` / `findIndex`。ホストのコレクション（`worksheets` など、§4.4）でも同じメソッドと `toArray()` を使える。配列・文字列メソッドは仕様 §3.2.4 の「標準ライブラリの拡張メソッド」ではなくコンパイラ組み込みとして実装 |
| §5 `try` / `catch` / `finally` / `using` / `throw` | 実装済み | catch 本体はハンドラ内ではなく `Resume` 後のラベルで実行する（同一プロシージャ内の外側 `try` へ再送出を届けるため） |
| §6 リンター | 一部 | ルール 1〜6・8〜10・14・15 は実装。7（予約語のリネーム）は自動リネームのみで警告なし、11 は変換のみで通知なし |
| §7 CLI・VS Code 拡張 | 一部 | 上の表のとおり |
| §8 インライン展開・ループ最適化 | 一部 | `inline`、ループ内自動展開、評価順序の保存、`For...Next` 化、`$` 付き関数は実装。ループ不変式の巻き上げ（上限式以外）・`With` 化・文字列バッファは未実装 |
| §9 Range ⇔ 配列の一括転送 | 実装済み | `@bulk` の「既存内容を読んでから書き戻す」方式は未実装（条件 5 を満たさないループは変換しない） |
| §10 アプリケーション状態・`DoEvents` 自動挿入 | 実装済み | |
| §11 `declare`（32/64bit） | 一部 | `declare struct` のアラインメント計算と `@fallback` は未実装 |
| §12 ボイラープレート・イベント・`@init`・マクロ入口 | 一部 | シート・`ThisWorkbook` のファイル（`objects/`、§12.1.1、§12.2）は実装済み。シートの上の ActiveX コントロール（§12.2.1）と `WithEvents`（`@on("<変数名>.<イベント>")`、§12.2.2）も実装済み（Excel のないデバッグ実行・テストでも、変数がプロジェクトのブック・シート・`Application` を指していれば、ハンドラを呼ぶ）。実行時に付けるイベントハンドラ（`on` / `addEventListener`、§12.2.3）も実装済み（生成する VBA は、実際の Excel では未確認。フォームのデバッグ実行・Excel のないデバッグ実行・テストでは動く）。リボン（`@ribbon`、§12.5）は未実装 |
| §13 VBA の罠の回避 | 一部 | `Collection` ラッパ、`Range.Find` ラッパ、文字列で指定される名前の存在確認、一時変数の再利用は未実装 |
| §14.1〜14.8 ライブ同期 | 一部 | Excel で開いているブックへ、保存のたびに反映する（`ras sync`、`ras watch`、VS Code の言語サーバー。Windows のみ）。ras が COM 経由で VBE のコードモジュールを直接書き換える方式で、共有メモリと RASBridge.xlam（§14.2、§14.4、§14.5）は使わない。適用規則（§14.6）のうち差分適用・ドキュメントモジュール全体の書き換え・属性付きクラスの `Import`・失敗時の巻き戻し・実行中／中断モード中の保留と、ドリフト検出（§14.7）は実装。ブックの手書きの標準モジュール・シート / `ThisWorkbook` のコード・フォームを RAS に変換して取り込み、RAS 管理に移す処理（§14.7.1）は実装。RAS が生成したモジュールに VBE で足したプロシージャ・宣言を RAS のファイルに取り込む処理（§14.7.3。書き込んだコードを `build/sync/` に保存して比べる）、Excel で複製したシートにコピー元のファイルを複製する処理（§14.7.2）も実装。適用後のコンパイル確認（`sync.compileCheck`）、Excel → ローカル方向（`--pull`、非管理モジュールの双方向同期）、クラスモジュールの取り込み、既存クラスの属性変更、`ras setup` は未実装 |
| §14.9 オフライン書き込み・読み出し | 実装済み | 再署名は未実装。`vba/` の非管理モジュールをブックへ書き戻す処理は未実装（取り出しのみ。ユーザーフォームは `forms/` で扱う） |
| §14.10 デバッグ実行（JavaScript ランタイム） | 実装済み | ユーザーフォーム（§14.10.7）は VS Code のパネルに表示する。キー入力・マウス・ドラッグのイベント、`ControlSource` は未対応。Excel への接続（`ras_com.node`）は Excel 2021（64bit）の実機で確認済み。ダイアログの監視（§14.10.6）は Excel の無い Windows 上の CI でメッセージ ボックスを相手に確認しており、Excel の実機では未確認。CI では Excel の無い Windows 上で `Scripting.Dictionary` 等を使って COM 呼び出しを確認している。Excel の無い環境でのブックの代役（§14.10.8、`ras_book.node`）は、RAS のプログラムを Node で動かしてセル・数式・保存結果を確かめる自動テスト（Linux / Windows）で確認しているが、VS Code 上での確認はまだ。Excel との突き合わせは `tools/standin-check` で行う（下の「検証状況」）。セルの挿入・削除、切り取り・コピー・貼り付け、並べ替え・フィルター、書式の保存、手動計算、ブック・シートのイベントも同じ自動テストで確認しており、イベントの順序、コピーが解除される操作、保存したブックを Excel で開いたときの書式は、Excel の実機と突き合わせて合わせた。`super` 呼び出しを伴う継承は対象外 |
| §14.11 ユーザーフォーム（`forms/`） | 実装済み | FormSpec からの生成（`Image` の画像、`ComboBox` / `ListBox` の項目を含む）、ブックからの取り出し（`ras form pull`）、ライブ同期、Excel での実行（§14.11.8）。カスタム ActiveX コントロールの生成、VBE での編集の検出（ドリフト）は未実装 |
| §14.12 テスト（`ras test`） | 一部 | テストの列挙（`--list`）、実行、`expect`、フック、ブックの巻き戻し、`useWorkbook` / `tempDir` / `printed`、置き換え（`mock*`）、タイムアウト、`--reporter pretty / json / junit`、`-t`、`--ci`、`--bail`、`--watch`、失敗したテストのブックの保存、`[test]` の設定は実装。実際の Excel での実行（`--excel`、`[test] excel = "com"`）、エディタの連携（§14.12.10。CodeLens・テスト エクスプローラー、言語サーバーでのテストファイルの型チェックと補完）は未実装。テストファイルは 1 つずつ実行する（並行しない）。`-t` で選ばれなかったテストは結果に出さない。`toBeEmpty` と `toBeNothing` は、ブックの代役では `Empty` と `Nothing` をどちらも成功にする（ランタイムが区別しないため）。`expect` の型は `toBe` の期待値だけを検査する |
| §14.13 CLI からの実行（`ras run`） | 実装済み | ブックの代役でだけ実行する（実際の Excel に対する実行は VS Code のデバッグ実行を使う）。ユーザーフォームの表示はエラーになる |
| §14.16 シートの追加・複製・削除・名前の変更（`ras sheet`） | 一部 | ブックのファイルを直接書き換える操作（`.xlsx` / `.xlsm`。`.xlsb` は不可）と、VS Code のコマンドは実装済み。Excel で開いているブックに対する操作は、Windows 専用のコードを書いたが、実際の Excel での確認はまだ行っていない |
| §14.14 ブックの中身を見る・比べる・編集する（`ras book dump` / `diff` / `view` / `edit`） | 実装済み | 行・列の挿入・削除の推定は、ひとまとまりの挿入・削除だけ（複数箇所は、セルの差分として出る）。`dump` / `diff` は書式（フォント・色・罫線）を扱わない。`view` は条件付き書式・テーブルのスタイル・画像・図形・グラフ・斜線・文字の回転を描かず、列幅は既定のフォントを日本語のフォント（11pt）とみなして計算する |
| §14.15 機械可読の診断・エラーの説明・代役で使えない API の警告 | 一部 | `--message-format json`、`ras explain`（すべてのコード）、`ras check --target stand-in`（W0510〜W0512）は実装。診断の「関連する位置」は、コンパイラが持っていないため出力しない。ユーザーフォームの検査の結果は JSON にならない。代役で使えるかどうかを言語サーバーのクイック情報に表示することは未実装 |
| §15 プロジェクト構成・`ras.toml`・同梱の型定義 | 一部 | Excel / VBA / Office の型定義を同梱し、メンバーとグローバルを camelCase で参照する（§15.3.1）。`Scripting`（`FileSystemObject`・`Dictionary`）と `VBScript.RegExp` の型定義も同梱し、VBA には遅延バインディングで出力する（§15.3）。`ras typegen` と `vba:` インポートは未実装 |
| §16.1 JSDoc | 実装済み | 生成 VBA へのコメント・`VB_Description` 出力、`@param` 名の検査 |
| §16.2 IntelliSense（LSP） | 一部 | 診断・補完（名前付き引数の引数名、未インポートの候補と `import` の自動追加を含む）・ホバー・パラメーター情報・定義へ移動・型定義へ移動・参照の検索・強調表示・名前の変更・シンボル一覧（ファイル・ワークスペース）・インレイヒント・セマンティックトークン・折りたたみ・整形・CodeLens（最適化の通知、関数ごとの「Excelで実行」「デバッグ実行」）・「生成 VBA を表示」と、クイックフィックス（VBA の綴りの名前を直す、未インポートの名前の `import` を追加する）は実装。`completionItem/resolve`（JSDoc の遅延読み込み）、範囲の整形、その他のコードアクション（`byref` 引数へのリテラル渡しの一時変数化・縮小変換への `as` 追加）、`@deprecated` の使用箇所の警告は未実装 |
| §17 採用ライブラリ | 一部 | `salsa`（増分計算。現状は変更のたびに全体を再解析）・`line-index`・`annotate-snippets`・`rayon`・`insta`・`lsp-types`・`windows` は未使用 |

仕様の例と出力が異なる点:

- 小数リテラル `1.5m` は `CDec("1.5")` ではなく `(CDec("15") / CDec("10"))` を出力します（小数点がカンマのロケールでも値が変わらないようにするため）。
- スクリプト形式（トップレベルの文）の `Main` にも、§12.4 のマクロ入口のラップが付きます。

## 検証状況

- 自動テスト: `cargo test --workspace`（構文、型チェック、仕様書のコード例と生成 VBA の照合、CLI の一連の流れ、LSP のセッション、`vbaProject.bin` の読み書き）。
- `ras book view` のページは、自動テストで描画のもとになるデータ（セルの表示・数式・書式の CSS・結合セル・変更の一覧・図形とコントロールの HTML と位置）を確かめ、Chromium（Playwright）で開いて見た目（スクロールで続くグリッド、図形・画像・コントロール・ワードアート）とセルの選択・変更の一覧からの移動を手で確認しています。図形を含むテスト用のブックは、Excel が保存する形をまねて作ったもので、Excel が保存したフォームコントロール・ActiveX コントロール・SmartArt・埋め込みオブジェクトのブックでは未確認です。グラフは xljs のテスト用ブック（12 種類）、EMF は Office が作った EMF（python-pptx の埋め込みオブジェクトのアイコン）で見た目を確かめています。VS Code のパネル（「ブックを表示」、デバッグの一時停止中の表示）と `launch.json` の `answers` は、型チェックと拡張のビルドだけで、実機の VS Code では未確認です。一時停止中にランタイムがブックを書き出す部分（`$rasBookSnapshot`）は自動テストで確かめています。
- `ras book edit` は、XlsxWriter で作ったブック（数式・VBA の関数を呼ぶ数式・結合セル・条件付き書式・入力規則・フィルター・メモ・ボタン・ハイパーリンク・名前の定義・グラフ・テーブル・VBA プロジェクト）を編集する自動テスト（`crates/ras-workbook/src/patch/tests.rs`）で、書き換えたパートと変えていないパートを確かめ、書き換えたブックを openpyxl で読めることを手で確認しています。**Excel が保存したブックの編集と、書き換えたブックを Excel で開くことは未確認です。** グリッドでの編集は、Chromium（Playwright）でページを操作し、`ras book edit` につないで確認しています（入力・再計算・行の挿入・元に戻す・コピーと貼り付け・エラーの表示、ツールバーの書式・表示形式・フィルハンドル・列幅のドラッグ・並べ替え・フィルター・結合）。VS Code のパネルでの編集（キーの受け取り・クリップボード・IME・元に戻す）は、型チェックと拡張のビルドだけで、実機の VS Code では未確認です。
- `ras run`・`ras test`・`ras book dump` / `diff` は、プロジェクトを作って `ras` を実行する自動テスト（`crates/ras/tests/offline.rs`。Linux の Node 22）で確認しています。Windows・macOS と、VS Code の実行ファイルをランタイムに使う経路は未確認です。`ras book dump` の出力は xljs のテスト用ブック（入力規則・条件付き書式・テーブル）でも確かめています。Excel が保存した大きなブックでの速度は測っていません。
- `ras pack` が出力したブックは、別実装の `ovba` クレートと oletools（`olevba`）で読み直して確認しています（CI でも実行）。
- 実機の Excel（Windows 11、Excel 2021 64bit）で次を確認しています（CI には Excel が無いため手動）:
  - `ras pack` で書き込んだブックを開いてマクロを実行できること。`crates/ras-compiler/tests/js/semantics.ras` の VBA ビルドの出力が `semantics.out` と一致すること
  - `RAS_Runtime.bas` の Win32 API・`SAFEARRAY` を扱う部分（`RAS_Rebase` のその場での下限書き換え、`RAS_Yield` の間引き、`sleep` / `waitUntil`）、Range ⇔ 配列の一括転送、`@on` のイベント、アプリケーション状態の復元
  - Excel が保存したブック（コンパイル済みの p-code 入り）への再 `pack`、開いているブックへの書き込み拒否、`unpack`
  - デバッグ実行（`ras_com.node`）が開いているブックに接続し、VBA ビルドと同じ結果・エラー番号になること
- ユーザーフォーム:
  - 生成するデザイナストレージ（画像の埋め込みを含む）は、読み戻して書き直すと同じになること、全プロパティについて書いた値が読み戻せることを自動テストで確認しています。実物の Excel との突き合わせは `tools/excel-check` のスクリプトで行います（Windows と Excel が必要）。PNG などから変換した画像は、変換後の BMP を直接指定した場合と同じ結果になることを確認しています。
  - ブックからの取り出しは、Excel が保存したフォーム 10 個について、読み取り結果が期待どおりであること、取り出したものを書き戻すと同じデザインになることを自動テストで確認しています。
  - ライブ同期でのフォームの取り込み（`VBComponents.Import` に渡す `.frm` / `.frx`）と、`ComboBox` / `ListBox` の項目を設定する生成コードは、実機の Excel では未検証です。`.frx` の構成は Excel が書き出したファイルと自動テストで照合しています。
- フォームの実行:
  - デバッグ実行（§14.10.7）は、フォームのランタイムを自動テストで、VS Code のパネルを Chromium（Playwright）で、ランタイムとつないで確認しています（入力・クリック・選択・タブ・フォーカスの取り消し・閉じるボタン）。実機の VS Code では未確認です。
  - 「Excelで実行」（`RAS_ShowForm`、§14.11.8）は実機の Excel では未検証です。
- 既存のブックの取り込み（§14.7.1）: `ras pack` での取り込みと置き換えは自動テストで確認しています。ライブ同期での取り込み（COM でのコードの読み取り、`VBComponent.Export` で書き出したフォームの読み取り）は実機の Excel では未検証です。`.frm` / `.frx` の読み取りは、Excel が保存したフォームを書き出して読み直すと同じデザインになることを自動テストで確認しています。
- デバッグ実行のダイアログ監視は、Excel / VBE の実際のダイアログ（実行時エラー、コンパイル エラー、Office 独自のダイアログ）では未検証です。判定規則は [xlflow](https://github.com/harumiWeb/xlflow) が実機で確立したもの（ボタンをコントロール ID とアクセスキーで識別する、実行時エラーは「終了」で閉じる、など）に合わせています。
- ブックの代役（§14.10.8）は、`tools/standin-check` で Excel 2021（64bit、日本語）と突き合わせています（手動。自動テストと同じシナリオを代役・COM 経由のデバッグ実行・Excel の中の VBA で実行して、出力と実行後のブックを比べる）。合わせたもの: コピーが解除される操作、貼り付け後の選択範囲、数式が参照先から引き継ぐ表示形式、日付の表示形式（`NumberFormat` は `m/d/yyyy`、表示と `NumberFormatLocal` は `yyyy/m/d`）、保存したブックの列幅、`Union` のまとめ方、シートを削除するときの `Deactivate`、追加したシートの名前、`VBScript.RegExp` のエラー番号と `TypeName`。残っている違い:
  - セルの数値の `typeName` は整数なら `Long` になる（VBA は `Double`）。空のセルと `Dictionary` に無いキーは、COM 経由では `Nothing` になる（VBA は `Empty`）。ランタイムが `Empty` と `Nothing` を区別しないため。
  - 列幅に収まらない数値・日付の `.text` は `#####` にならない。`Interior.ThemeColor` は `null` を返す。
  - VBE で開いたことのないブックに追加したシートの `codeName` は、Excel では空文字列（代役は `Sheet2` など）。
  - COM 経由のデバッグ実行では、閉じていない `TextStream` が変数の寿命で解放されないため、そのファイルのあるフォルダーの削除がエラー 70 になることがある（VBA では参照がなくなった時点で閉じる）。
- 初期化の入口（§12.3）は、Excel 2021 の実機で次を確認しています（手動）: ブックをマクロの中で開いてすぐ閉じても、Excel がブックを開き直さないこと（`Application.OnTime` で遅延していた版では開き直し、セキュリティの通知で止まった）。保護ビューで開いて編集を有効にしたとき（`ProtectedViewWindow.Edit`）、`@on("Workbook.Open")` のハンドラが `Workbook_Activate` で呼ばれ、`Range.Select` などが成功すること（`Workbook_Open` から直接呼ぶ手書きのブックではエラー 1004 になる）。リボンから「編集を有効にする」を押す操作と、「コンテンツの有効化」の後の動作は未確認です。
- 32bit Office では未検証です（`RAS_Rebase` の 32bit 側のオフセットなど）。
- VS Code 拡張は実機の VS Code 上で動作を確認しています。

## 開発

```sh
cargo test --workspace
cargo clippy --workspace --all-targets -- -D warnings
cargo run -p ras -- --project path/to/project build

cd editors/vscode && npm ci && npm test && npm run package
```

| ディレクトリ | 内容 |
|---|---|
| `crates/ras-syntax` | 字句解析（`logos`）、構文解析（手書き再帰下降、`rowan` のロスレス CST）、AST、フォーマッタ |
| `crates/ras-compiler` | 宣言収集、型チェック、VBA への変換、ランタイム、出力 |
| `crates/ras-ovba` | MS-OVBA 圧縮、`vbaProject.bin` の読み書き、OPC パッケージ |
| `crates/ras-lsp` | 言語サーバ |
| `crates/ras` | CLI |
| `crates/ras-workbook` | ブックの代役のブック（umya-spreadsheet と GridCraft の計算エンジン）、`ras book dump` / `diff` / `view`、`ras book edit`（`patch`: 編集をブックのファイルに書き戻す） |
| `crates/ras-book` | ブックの代役を Node に渡すアドオン（`ras_book.node`） |
| `editors/vscode` | VS Code 拡張 |
| `examples/sales-report` | サンプルプロジェクト |
