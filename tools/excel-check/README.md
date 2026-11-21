# Excel でユーザーフォームを確かめる

RAS が書き込むユーザーフォーム（`ras pack`）と読み込むユーザーフォーム（`ras form pull`）が、本物の Excel と食い違わないかを自動で確かめるスクリプトです。手でフォームを開いて見比べる代わりに使います。

## 何を比べるか

`forms/*.form.json` は、FormSpec の全プロパティを既定と違う値にしたテスト用のフォームです（14 個）。漏れがないことは `cargo test -p ras-form`（`the_excel_check_forms_cover_every_property`）が確かめます。

スクリプトはこれを使って 2 冊のブックを作ります。

- **RAS のブック**: `ras pack` で書き込んだもの。
- **参照ブック**: 同じフォームを Excel の VBA エディターのデザイナー（`VBComponent.Designer`）で一から作ったもの。人が手で作ったのと同じ状態です。

そのうえで、プロパティごとに次を比べます。

| 列 | 中身 |
| --- | --- |
| 期待値 | `forms/*.form.json` に書いた値 |
| RAS→Excel | RAS のブックを Excel で開き、デザイナーから読んだ値 |
| Excel（参照） | 参照ブックのフォームから読んだ値 |
| Excel→RAS | 参照ブックを `ras form pull` で取り込んだ FormSpec の値 |

- **NG**: RAS→Excel が期待値と違う、または Excel→RAS が Excel（参照）と違う。RAS の読み書きの違いです。
- **注意**: Excel 自身が値を変える（参照も同じ値になる）、参照のフォームをうまく作れなかった、AutoSize のコントロールの大きさ（Excel が決める）など、RAS の違いとは言えないもの。

さらに、両方のブックのフォームを実際に表示して画面を撮り、ピクセル単位で比べます（見た目）。ComboBox / ListBox の項目（RAS が `UserForm_Initialize` で入れるもの、`RowSource` でセルから読むもの）もこのときに比べます。

## 必要なもの

- Windows と Excel（デスクトップ版）
- Windows PowerShell 5.1（`powershell.exe`。Windows に最初から入っているもの）
- `ras`（PATH に置くか、`-Ras` でパスを指定）
- Excel の［ファイル］［オプション］［トラスト センター］［トラスト センターの設定］［マクロの設定］で「**VBA プロジェクト オブジェクト モデルへのアクセスを信頼する**」をオン

## 実行

リポジトリーのルートで:

```bat
powershell -ExecutionPolicy Bypass -File tools\excel-check\Check-Forms.ps1
```

ras.exe を指定する、一部のフォームだけ確かめる:

```bat
powershell -ExecutionPolicy Bypass -File tools\excel-check\Check-Forms.ps1 -Ras target\release\ras.exe -Forms CheckLabels,CheckButtons
```

| 引数 | 意味 |
| --- | --- |
| `-Ras` | ras の実行ファイル（既定 `ras`） |
| `-Out` | 結果のフォルダー（既定 `tools\excel-check\out`、毎回作り直す） |
| `-Forms` | 確かめるフォームの名前（既定はすべて） |
| `-Tolerance` | 見た目で違っていてもよいピクセル数（既定 40。TextBox のカーソルの点滅など） |
| `-SkipVisual` | 見た目を比べない |
| `-KeepExcel` | 終わった後も Excel を開いたままにする |

実行中は Excel が表示され、フォームが 1 つずつ開いては閉じます。その間はマウスやキーボードに触らないでください。

## 結果

- `out\report.html`: フォームごとの画像（RAS / Excel / 違いを赤で示したもの）と、全プロパティの表。「NG と注意だけ表示」で絞り込めます。
- `out\results.json`: 同じ内容の JSON。
- `out\work\`: 作ったブック（`ras\RasForms.xlsm`、`ExcelForms.xlsm`）と、取り込んだ FormSpec（`pull\forms\`）。

NG があると終了コード 1 です。NG の行はコンソールにも出ます。

## ファイル

- `Check-Forms.ps1`: 本体
- `RasCheck.bas`: 両方のブックに読み込む VBA の補助関数（デザイナーでの読み書き、フォームの表示）。`VBComponents.Import` がシステムのコード ページで読むので ASCII だけで書いています。
- `forms/`: テスト用のフォームと、その画像（`forms/assets/`）
