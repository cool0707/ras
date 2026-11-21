# 実物の Excel で確かめるスクリプト

`tools/` には、RAS が書いたものを本物の Excel で確かめるスクリプトを置きます。CI には Excel が無いので、手元の Windows で実行します。

| フォルダー | 確かめるもの |
| --- | --- |
| [`excel-check/`](excel-check/README.md) | ユーザーフォーム（`ras pack` / `ras form pull`）の全プロパティと見た目 |
| [`standin-check/`](standin-check/README.md) | ブックの代役（Excel なしの実行）と Excel の動作の違い |

## 書き方の決まり

- **PowerShell 7（`pwsh`）で書く。** Windows PowerShell 5.1 との互換は考えなくてよい（三項演算子・`??`・`ForEach-Object -Parallel` なども使える）。実行は `pwsh -File tools\<フォルダー>\<スクリプト>.ps1`。
  - 例外: `excel-check/Check-Forms.ps1` は 5.1 向けに書いたもの。`powershell.exe`（5.1）で実行する（7 での動作は確かめていない）。
- Excel は COM（`New-Object -ComObject Excel.Application`）で操作する。表示したまま動かし、終わったら閉じる（`-KeepExcel` のような引数で残せるようにする）。
- 必要なものと実行方法、結果の見方は、フォルダーごとの `README.md` に書く。共通して必要なのは次のとおり。
  - Windows とデスクトップ版の Excel
  - `ras`（PATH に置くか、`-Ras` でパスを渡す）
  - VBA プロジェクトを触る場合は、Excel のオプション［トラスト センター］［マクロの設定］の「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」
- 結果は各フォルダーの `out/` に書く（git では無視される）。人が見る `report.html` と、機械で読む `results.json` を基本にし、違いがあれば終了コード 1 にする。
- Excel に読み込ませる VBA のモジュール（`.bas`）は ASCII だけで書き、改行は CRLF にする（`VBComponents.Import` はシステムのコード ページで読み、LF だけの改行では壊れる。`.gitattributes` で CRLF にしている）。
- CI（Windows）では、`tools/` の `.ps1` を構文の解析だけ行う（`pwsh` のパーサー）。
