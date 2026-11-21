# ブックの代役を Excel と突き合わせる

同じ RAS のプログラムを、ブックの代役（`ras_book.node`、仕様 §14.10.8）と本物の Excel（`ras_com.node` でのデバッグ実行）で動かし、`debug.print` の出力と、実行後のブック（`ras book diff`）を比べます。ロードマップの 11 にあたります。

## 何を比べるか

- `cells/src/main.ras`・`tools.ras`: 自動テスト（`crates/ras-book/tests/excel.rs`、`crates/ras-compiler/tests/js_backend.rs` の `SCRIPTING`）と同じシナリオ。
- `cells/src/probe.ras` と `Probe.bas`: 同じ問いを 1 行ずつ出力するもの。`Probe.bas` は Excel の中の VBA として実行するので、代役・COM 経由のデバッグ実行・VBA の 3 つを比べられます（コピーが解除される操作、貼り付け後の選択範囲、`Union`、入力した値の型と表示形式、`TypeName`、正規表現のエラー番号）。
- `events/`: イベントの順序。Excel はイベントを VBA に対して起こすので、こちらは `ras pack` したブックを Excel で実行して比べます（スクリプトにはまだ入れていません。関数が返す文字列を `Application.Run` で受け取ります）。

## 必要なものと実行

Windows とデスクトップ版の Excel、Node.js、`cargo build --release` 済みの `target\release`（`ras.exe`・`ras_com.node`・`ras_book.dll`）、「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」。

```bat
pwsh -File tools\standin-check\Check-StandIn.ps1
pwsh -File tools\standin-check\Check-StandIn.ps1 -Only clipboard,probe -KeepExcel
```

結果は `out\results.json` と、`out\standin\<シナリオ>\Book.xlsx`・`out\excel\<シナリオ>.xlsx`。違いがあると終了コード 1 です。

## 見方の注意

- `failing` の 1 行目（手書き VBA の実行）と `regex` の最後（`ADODB.Connection`）は、代役では使えないものなので必ず違います。
- `manual` のブックの `F1` は、スクリプトが Excel の計算方法を自動に戻してから保存するために違います。
- Excel 側の `TypeName` が `IFile` などになるのは、COM 経由のデバッグ実行の側の違いです（VBA は `File`）。
