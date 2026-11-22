# RAS (Better VBA)

[RAS](https://github.com/cool0707/ras) は、TypeScript ライクな構文から VBA（`.bas` / `.cls`）へトランスパイルする言語です。この拡張機能は、Visual Studio Code に RAS（`.ras` / `.d.ras`）の言語サポートを追加します。

拡張機能自体は意図的に薄く作られています。行うのは、シンタックスハイライトの提供と、`ras` バイナリを言語サーバー（`ras lsp --stdio`）として起動して接続することだけです。補完・クイック情報・パラメーター情報・診断・定義へ移動・名前の変更・インレイヒント・セマンティックハイライト・CodeLens・フォーマットなどは、すべて言語サーバーが提供します。

## 必要なもの

- Visual Studio Code 1.90 以降
- `ras` 実行ファイル（Windows では `ras.exe`）。`PATH` の通った場所に置くか、設定 `ras.path` でパスを指定してください。

`ras` が見つからない場合はエラーメッセージを表示します。メッセージの［設定を開く］から `ras.path` を設定できます。`ras` が無くてもシンタックスハイライトは利用できます。

## 機能

- `.ras` ファイルのシンタックスハイライト（TypeScript の TextMate 文法をもとに、`byref` / `byval` / `inline` / `noinline` / `resume`、RAS の基本型、数値リテラルのサフィックス `123L` / `1.5f` / `2d` / `1.5m`、デコレータ、`declare module` / `declare struct` に対応）
- コメント・括弧の補完、`/** */` ドキュメントコメントの継続入力、インデント、折りたたみ
- 言語サーバーによる IntelliSense 一式（`ras.toml` を含むワークスペース、または `.ras` ファイルを開くと自動的に起動します）
- ステータスバーへの同期状態の表示
- ブレークポイントを使ったデバッグ実行（ステップ実行、変数・ウォッチ式、例外での停止）。ユーザーフォームは VS Code のパネルに表示して動かします
- ユーザーフォーム デザイナー（`forms/*.form.json` をビジュアルに編集）

## コマンド

コマンドパレットから実行できます。いずれも言語サーバーに処理を依頼するだけのコマンドです。

`ras.toml` を含むフォルダーを開いていると、エクスプローラーのタイトルバー（プロジェクト フォルダー名の行）に［新しい標準モジュール］と［新しいユーザーフォーム］のアイコンボタンを表示します。どちらもフォルダーの右クリックメニューからも実行できます。

| コマンド | 内容 |
|---|---|
| `RAS: 生成 VBA を表示` | 編集中の RAS ファイルから生成される VBA を、読み取り専用のエディターで横に表示し、カーソル位置に対応する行をハイライトします。エディターのタイトルバーと右クリックメニューからも実行できます。 |
| `RAS: 今すぐ同期` | Excel で開いているブックへ、保存済みのソースから生成した VBA を今すぐ反映します。ブックが Excel で開かれていない場合（Windows 以外を含む）は、`ras pack` と同じくブックファイルへ直接書き込みます（オフライン書き込み。実行結果は「出力」パネルの RAS に表示します）。Excel（VBE）側で編集されたモジュールがある場合は、上書きするかどうかを確認します。 |
| `RAS: Excel で開く` | 対象のブックを Excel で開き（すでに開いていれば前面に出し）、同期します。 |
| `RAS: マクロを実行` | マクロの一覧から選び、同期してから Excel 上で実行します。 |
| `RAS: フォームを Excel で実行` | ユーザーフォームを選び、同期してから Excel 上で表示します（Windows のみ）。デザイナーのタイトルバーからも実行できます。 |
| `RAS: フォームをデバッグ実行` | デザイナーで開いているユーザーフォームを VS Code のパネルに表示してデバッグ実行します。デザイナーのタイトルバーから実行できます。 |
| `RAS: ブックを表示（グリッド）` | ブック（`.xlsx` / `.xlsm`）を Excel のようなグリッドでパネルに表示します。エクスプローラーのブックの右クリックからも実行できます。表示したブックは、グリッドで編集できます（下の「グリッドでの編集」。`build/` の下の実行結果のブックと、デバッグ中の表示は読み取り専用です）。 |
| `RAS: シートのグリッドを表示` | シートのファイル（`objects/Sheet1.ras` のように、`objects/` にシートのコード名で置いたファイル）で、プロジェクトのブック（`ras.toml` の `project.workbook`）をそのシートを開いた状態でグリッド表示します。エディターのタイトルバー（表のアイコン）・エディターの右クリック・エクスプローラーのファイルの右クリックから実行できます。表示するのは保存されたブックファイルの内容で、グリッドで編集できます。 |
| `RAS: シートを追加` / `シートを複製` / `シートを削除` / `シートの名前を変更` | ブックのシートと、そのファイル（`objects/<コード名>.ras`）を一緒に変えます。エクスプローラーの `objects/` とその中のファイルの右クリック、シートのファイルのタブの右クリック、シートのグリッドの見出しの右クリックから実行できます。ブックが閉じていればブックのファイルを書き換え、Excel で開いていれば Excel を操作します（Windows のみ）。削除は、ファイルにコードがある・ほかから参照されている場合に理由を示して確認します。コード名を変えると、ファイル名とソースの参照（`Sheet1.…`）も変えます。エクスプローラーで `objects/` のファイルの名前を変えた・削除したときは、シートにも同じ操作をするかを尋ねます。 |
| `RAS: 言語サーバーを再起動` | 言語サーバーを再起動します。 |
| `RAS: デバッグを開始` | 関数を選んでデバッグ実行します（F5 と同じ）。エディターのタイトルバーからも実行できます。 |
| `RAS: 新しい標準モジュール` | `src/<名前>.ras` を作成して開きます。VBA のモジュール名はファイル名から決まります（`str_util` → `StrUtil`）。エクスプローラーで `src/` の下のフォルダーを右クリックして実行すると、そのフォルダーに作成します。 |
| `RAS: 新しいユーザーフォーム` | `forms/<名前>.form.json` と RAS のコードビハインド `forms/<名前>.ras` を作成し、デザイナーで開きます。 |
| `RAS: デザイナーで開く` | テキストエディターで開いている FormSpec をデザイナーで開きます。 |
| `RAS: JSON を開く` | デザイナーで開いている FormSpec を JSON のテキストエディターで横に開きます。 |

## グリッドでの編集

「RAS: ブックを表示（グリッド）」「RAS: シートのグリッドを表示」で表示したブックは、Excel と同じ操作で編集できます。編集は `ras book edit` がブックのファイルに書き込みます（仕様 §14.14）。変わったセルだけを書き戻すので、VBA プロジェクト・図形・コントロール・書式はそのまま残ります。

| 操作 | キー・マウス |
|---|---|
| 選択 | クリック・Shift+クリック・ドラッグ、行番号・列見出し（行・列全体）、矢印キー・Tab・Enter・Home・Ctrl+Home・Ctrl+矢印・PageUp / PageDown・Ctrl+A |
| セルの入力 | 文字を打つ（上書き）、F2・ダブルクリック・数式バー（今の内容を編集）。Enter / Tab で確定、Ctrl+Enter で選択範囲の全セルに入力、Alt+Enter でセル内の改行、Esc で取り消し |
| クリア | Delete / Backspace（数式と値） |
| コピー・切り取り・貼り付け | Ctrl+C / Ctrl+X / Ctrl+V（ほかのアプリケーションのタブ区切りテキストも貼り付けられます） |
| 下・右へコピー | Ctrl+D / Ctrl+R |
| 行・列・セルの挿入と削除 | 右クリックのメニュー、Ctrl++ / Ctrl+-（セルを選んでいるときは、動かす向きを選びます） |
| 書式 | ツールバー（太字・斜体・下線・取り消し線・フォントの大きさ・文字の色・塗りつぶし・罫線・配置・折り返し・表示形式・書式のクリア）、Ctrl+B / Ctrl+I / Ctrl+U / Ctrl+5 |
| 結合 | ツールバーの「結合」（セルを結合して中央揃え。▾ から横方向に結合・セルの結合・結合の解除） |
| 列の幅・行の高さ | 列見出しの右端・行番号の下端をドラッグ（ダブルクリックで内容に合わせます）。見出しの右クリックの「列の幅…」「行の高さ…」「非表示」「再表示」 |
| 並べ替え・フィルター | ツールバーの A→Z / Z→A / フィルター（Ctrl+Shift+L）。フィルターの見出しの ▾ から、並べ替えと、表示する値の選択 |
| 連続データ | 選択範囲の右下の ■（フィルハンドル）をドラッグ。ダブルクリックで隣の列のデータの下端まで |
| 元に戻す・やり直し | Ctrl+Z / Ctrl+Y（ブックごとに 50 回分。Excel や同期がファイルを変えると、履歴は消えます） |

- 編集のたびにブックのファイルを書き換えます（保存の操作はありません）。プロジェクトのブックは、前のファイルが `build/backup/` に残ります。
- Excel で開いているブックは編集できません（理由をグリッドの下に表示します）。Excel で閉じてから行ってください。
- Excel が拒む操作（結合セルの一部だけを動かす挿入など）と、ファイルの書き換えでは行えない操作（テーブルの列をまたぐ列の挿入・削除、ピボットテーブルのあるシートでの挿入・削除）は、理由を表示して何も変えません。
- 形式を選択して貼り付け、条件付き書式・入力規則の編集、複数キーの並べ替え、数値フィルターなど値の一覧以外の条件での絞り込み、フォント名の選択は、まだ行えません。
- グリッドにフォーカスがある間、Ctrl+R・Ctrl+D・Ctrl+Y・Ctrl++・Ctrl+-・Ctrl+B・Ctrl+I・Ctrl+U・Ctrl+5・Ctrl+Shift+L は VS Code のコマンドではなくグリッドの操作になります。

## デバッグ

`.ras` ファイルにブレークポイントを置いて F5 を押すと、プロジェクトをビルドし、実行する関数を選んでデバッグを開始します。
関数の上に表示される CodeLens からも、その関数を直接起動できます。

```
▶ Excelで実行 | ⚙ デバッグ実行
export function rebuildReport() {
```

- **Excelで実行**（Windows のみ）: ブックへ同期してから、Excel 本体でマクロとして実行します（`RAS: マクロを実行` と同じ。ブックを Excel で開いている必要があります）。マクロとして公開される関数（引数なし・戻り値なしで `export` された関数、`@macro` の関数、トップレベルの文）に表示されます。
- **デバッグ実行**: その関数を指定してデバッグを開始します（F5 と同じ。関数を選ぶ手順が省けます）。引数なしの関数とトップレベルの文に表示されます。

ユーザーフォームのコード（`forms/<名前>.ras`）では、ファイルの先頭に同じ CodeLens が表示され、フォームを表示して動かします。デザイナーではタイトルバーのボタン（▶ と ⚙）が同じ働きをします。

- **Excelで実行**（Windows のみ）: ブックへ同期してから、Excel 本体でフォームを表示します。モーダルのフォームを閉じるまで、次の同期は待たされます。
- **デバッグ実行**: フォームを VS Code のパネルに表示し、フォームのコードをデバッグ実行します。パネルでのクリック・入力・選択がフォームのイベント（`Click`・`Change`・`Enter` / `Exit`・`QueryClose` など）になり、ハンドラにブレークポイントを置いて止められます。パネルの ✕（またはパネルのタブを閉じる）はフォームの閉じるボタンです。ハンドラで起きたエラーはパネルとデバッグ コンソールに表示され、実行は続きます（止めたい場合はブレークポイントの「捕捉された例外」を有効にします）。他のモジュールの関数から `frmInput.show()` で表示した場合も、同じパネルに表示されます。

CodeLens は設定 `ras.codeLens.run` で非表示にできます。

`launch.json` を使う場合は、構成の種類に `ras` を指定します。

```json
{
  "type": "ras",
  "request": "launch",
  "name": "RAS: rebuildReport",
  "entry": "Report.rebuildReport"
}
```

| 属性 | 内容 |
|---|---|
| `entry` | 実行する関数（`モジュール名.関数名` または関数名のみ）。省略すると起動時に選びます。 |
| `project` | `ras.toml` のあるフォルダー。省略時はワークスペースフォルダー。 |
| `stopOnEntry` | プログラムの先頭で停止します。 |
| `dialogs` | プログラムが Excel の応答を待っている間に Excel がダイアログを出したときの扱い。`notify`（既定。内容を通知してダイアログを前面に出す）/ `dismiss`（通知したうえでキャンセルに相当するボタンで閉じる。VBA の実行時エラーは「終了」）/ `off`。 |
| `excel` | プログラムが操作する Excel。`auto`（既定。Windows では Excel、Mac・Linux ではブックの代役）/ `excel`（Excel だけ）/ `standIn`（Windows でもブックの代役）。省略時は設定 `ras.debug.excel`。 |
| `runtimeExecutable` / `comAddon` / `bookAddon` / `mockHost` | 実行に使う Node.js、Excel 接続用アドオン、ブックの代役のアドオン、Excel の代役モジュールを差し替えます（通常は不要）。 |

デバッグは VBA 上ではなく、プログラムを JavaScript に変換して VS Code 同梱のランタイムで実行します。Excel の操作は、同梱のアドオン（`ras_com.node`、Windows 専用）を通じて実際の Excel に対して行われます。

Mac・Linux や Excel の無い Windows では、Excel の代わりに `ras.toml` の `workbook` のブックを読み込み、そのセルを読み書きして実行します（同梱のアドオン `ras_book.node`）。

- 数式は xljs と同じ GridCraft の計算エンジンで計算し、セルを書き換えると Excel の自動計算と同じく計算し直します。
- プログラムが変更したブックは、終了時に `build/debug/workbooks/` に保存します。元のブックは書き換えません。
- セルの挿入・削除、切り取り・コピー・貼り付け（Excel の中だけ。ほかのアプリケーションとは受け渡ししません）、並べ替え・フィルターに対応しています。
- フォント・色・罫線・列幅などの書式はブックに保存します（画面が無いので描画はしません）。
- プログラムの操作で起きるブック・シートのイベント（`Change`・`SelectionChange`・`Activate`・`BeforeSave` など）は、`@on` のハンドラを呼び出します。
- 計算方法を手動（`Application.Calculation`）にすると、`Calculate` を呼ぶまで数式を計算し直しません。
- `Scripting.FileSystemObject`・`Scripting.Dictionary`・`VBScript.RegExp` は使えます（テキスト ファイルは Windows と同じく Shift_JIS か UTF-16 で読み書きします）。`ADODB` や `WScript.Shell` など、ほかの COM オブジェクトは作れません。
- 手書き VBA（`Application.Run`）、ダイアログなど Excel そのものが必要な操作はエラーになります。変数に入れた `Range` は、Excel と違ってセルの挿入・削除に追従しません。
- 対応範囲と Excel との違いは言語仕様 §14.10.8 にまとめています。

- 実行できるのは、引数なしの関数とトップレベルの文です。イベントハンドラ（`@on`）や UDF は起動できません。
- 対象のブックが Excel で開かれていればそのブックに接続し、開かれていなければ開きます。
- 実行タイミングや速度、一部の数値型の精度、未対応の VBA 標準関数など、Excel 上で動かした場合との違いがあります（言語仕様 §14.10.5）。
- プログラムの標準出力と未処理のエラーは「出力」パネルの **RAS Debug** に表示します。`console.log` の出力はデバッグコンソールにも表示されます。

## 設定

| 設定 | 既定値 | 内容 |
|---|---|---|
| `ras.path` | `""` | `ras` 実行ファイルのパス。空の場合は `PATH` から `ras`（Windows では `ras.exe`）を検索します。相対パスは最初のワークスペースフォルダーを基準に解決し、先頭の `~` はホームディレクトリに展開します。変更すると言語サーバーを再起動します。 |
| `ras.server.extraArgs` | `[]` | 言語サーバーの起動時に `ras lsp --stdio` の後ろへ追加する引数。変更すると言語サーバーを再起動します。 |
| `ras.trace.server` | `"off"` | VS Code と言語サーバーの間の通信を「出力」パネルの「RAS Language Server」に記録します（`off` / `messages` / `verbose`）。 |
| `ras.codeLens.run` | `true` | 単独で起動できる関数の上に「Excelで実行」（Windows のみ）と「デバッグ実行」の CodeLens を表示します。 |
| `ras.debug.runtimeExecutable` | `""` | デバッグ実行に使う Node.js のパス。空の場合は VS Code 同梱のランタイムを使います。 |
| `ras.debug.comAddon` | `""` | Excel 接続用アドオン（`ras_com.node`）のパス。空の場合は拡張機能に同梱のもの、無ければ `ras` と同じフォルダーのものを使います。 |
| `ras.debug.excel` | `auto` | デバッグ実行でプログラムが操作する Excel。`auto`（Windows では Excel、それ以外ではブックの代役）/ `excel` / `standIn`（Windows でもブックの代役）。 |
| `ras.debug.bookAddon` | `""` | ブックの代役のアドオン（`ras_book.node`）のパス。空の場合は拡張機能に同梱のもの、無ければ `ras` と同じフォルダーのものを使います。 |
| `ras.debug.dialogs` | `notify` | Excel がダイアログを出して応答を待っているときの扱い（`notify` / `dismiss` / `off`）。`launch.json` の `dialogs` で構成ごとに上書きできます。 |

RAS ファイルでは、言語サーバーによるセマンティックハイライト（`editor.semanticHighlighting.enabled`）を既定で有効にしています。

## 言語サーバーの機能

| 機能 | 操作 |
|---|---|
| 補完 | `.` の後のメンバー、名前付き引数（`(` `,` の後）。まだインポートしていない他のモジュールの `export` も候補に出し、確定すると `import` を追加します |
| クイック情報・パラメーター情報 | ホバー。関数の `(` や `,` の後ではシグネチャと入力中の引数（`@param` の説明付き）を表示します |
| 定義へ移動・型定義へ移動 | F12 / 右クリック →［型定義へ移動］ |
| 参照の検索・強調表示 | Shift+F12（［すべての参照を検索］は Shift+Alt+F12）。カーソル位置の名前を使っている箇所をファイル内で強調表示します |
| 名前の変更 | F2。プロジェクト内のすべての参照（`import`、名前付き引数、型注釈、`@param`、オブジェクト リテラルのキーを含む）を書き換えます。インターフェイスのメンバーは、それを実装するクラスの同名のメンバーも一緒に変わります。VBA の予約語、大文字小文字だけが異なる名前（VBA では区別されません）、同じ範囲で既に使われている名前、型定義（`.d.ras`）で宣言された名前やフォームのコントロール名への変更は理由を表示して行いません |
| シンボル | アウトライン、［シンボルへ移動］（Ctrl+Shift+O）、［ワークスペース内のシンボルへ移動］（Ctrl+T） |
| インレイヒント | 推論された変数・ラムダ引数の型、ジェネリック関数の推論された型引数、参照渡しの引数の前の `byref`（`editor.inlayHints.enabled` で切り替え） |
| セマンティックハイライト | 引数・ローカル変数・プロパティ・メソッド・Excel / VBA の名前などを区別して色分けします。参照渡しの引数は下線、`@deprecated` の宣言は取り消し線で表示します（`editor.semanticTokenColorCustomizations` の `*.byref:ras` / `*.extension:ras` などで変更できます） |
| クイック フィックス | VBA の綴りの名前を camelCase に直す、インポートしていない名前の `import` を追加する。`"editor.codeActionsOnSave": { "source.fixAll": "explicit" }` を設定すると、保存時に VBA の綴りの名前をすべて直します |
| 折りたたみ | ブロック、複数行のコメント、`import` の並び、`// #region` ～ `// #endregion` |
| 整形・CodeLens | 「ドキュメントのフォーマット」（VBA の綴りの名前（`W0304`）も camelCase に直してから整形します。`ras check --fix` と同じ修正です）、関数の上の「Excelで実行」「デバッグ実行」と最適化の適用状況 |

## Excel とのライブ同期（Windows）

`ras.toml` の `project.workbook` に指定したブックを Excel で開いていると、`.ras` やユーザーフォームのファイル（`forms/*.form.json`、`forms/*.ras`、`forms/assets/` の画像）を保存するたびに、生成した VBA とユーザーフォームをそのブックの VBA プロジェクトへ反映します（`[sync] enabled = false` で無効にできます）。ブックを後から開いた場合も、数秒以内に接続して同期します。

- **必要な設定:** Excel の［ファイル］→［オプション］→［トラスト センター］→［トラスト センターの設定］→［マクロの設定］で「VBA プロジェクト オブジェクト モデルへのアクセスを信頼する」を有効にしてください。
- 反映するのは保存済みの内容です。ビルド エラーがある間は同期しません。
- コードを書き換えると、VBA のグローバル変数はリセットされます。
- ユーザーフォームもブックへ書き込みます。デザインが変わったフォームはブックへ取り込み直し、コード（`forms/<名前>.ras`）だけが変わったフォームはコードだけを書き換えます。Excel 側でフォームが読み込まれたまま（`Hide` で隠しただけ、モードレスで表示中）のときや、VBE でデザイナーを開いているときは、書き込む前にそのフォームを閉じます（`QueryClose` などのイベントは実行されます）。
- マクロの実行中・中断モード中、セルの編集中、ダイアログの表示中は同期を保留し、Excel が操作を受け付けるようになった時点で反映します。
- Excel（VBE）側で編集された RAS 管理モジュール（ドリフト）は上書きしません。「今すぐ同期」を実行すると、上書きするかどうかを選べます。
- ブックにある手書きの標準モジュール・シート / `ThisWorkbook` のコード・VBE で作ったユーザーフォームは、同期のときに RAS に変換してプロジェクト（`src/`・`forms/`）に取り込み、ブック側を RAS の生成コードに置き換えます。変換できなかった行（`// TODO(VBA):`）が残っているものは、直すまでブック側をそのまま残します（`[sync] pull-unmanaged = false` で取り込みを止められます）。
- 同期のときに、ブックの各シートと `ThisWorkbook` について、そのモジュールのファイル（`objects/Sheet1.ras`、`objects/ThisWorkbook.ras`）が無ければ作成します。コードを書くまではブックのモジュールを変えません（`[sync] sheet-files = false` で止められます）。
- RAS が生成したものではない同名のモジュールがブックにある場合は、何も書き込まずに警告します。
- ブックを閉じている間は自動では反映されません。ステータスバーをクリック（「今すぐ同期」）すると、ブックファイルへ直接書き込みます（`ras pack` と同じ）。

## ステータスバー

左側に同期状態を表示します。クリックすると「今すぐ同期」を実行します（未接続のときはブックファイルへのオフライン書き込み、言語サーバーが停止しているときは再起動）。Windows では、未接続の間はその右に［Excel で開く］ボタンを表示します。

| 表示 | 意味 |
|---|---|
| `RAS: Book.xlsm` | Excel で開いているブックに接続しています。 |
| `RAS: 未接続` | 対象のブックが Excel で開かれていません（Windows 以外では常にこの表示です）。理由はツールチップに表示します。 |
| 上向き矢印と数字 | まだブックに反映できていない保存の回数です（ビルド エラー、Excel が操作中、など。理由はツールチップに表示します）。 |
| 警告アイコンと数字（背景が警告色） | ドリフト警告です。Excel（VBE）側で編集された RAS 管理モジュールがあります。 |
| `RAS: 停止中` | 言語サーバーが起動していません。 |

## ユーザーフォーム デザイナー

ユーザーフォームのデザインは、プロジェクト直下の `forms/<名前>.form.json`（FormSpec v1。[xlflow](https://github.com/harumiWeb/xlflow) の FormSpec を元にした JSON）に、イベント処理などのコードは RAS で `forms/<名前>.ras` に書きます（`this` がフォームで、コントロールは `this.<名前>`。イベントは `@on("<コントロール名>.<イベント>")` を付けた関数で受け取ります。`this.` の補完、型チェック、ホバーなどは `src/` のファイルと同じように使えます）。`forms/` 直下の `*.form.json` は既定でデザイナーで開きます（テキストで編集したいときは［JSON を開く］か「エディターを再度開く…」）。デザイナーは [xlflow](https://github.com/harumiWeb/xlflow) の UserForm Designer を移植したものです。

```json
{
  "schemaVersion": 1,
  "kind": "ras.userform",
  "basis": "designer",
  "form": { "name": "OrderForm", "caption": "注文", "width": 240, "height": 180 },
  "controls": [
    { "id": "ok", "name": "OkButton", "type": "CommandButton", "caption": "OK", "left": 160, "top": 144, "width": 72, "height": 24 }
  ]
}
```

- **配置と編集**: ツールボックスで種類を選んでフォーム（または Frame・表示中のページ）をクリックすると配置します。ドラッグで移動、8 つのハンドルでサイズ変更、フォームの右端・下端でフォームのサイズ変更ができます。矢印キーで 1pt、Shift+矢印キーで 10pt 移動します。グリッドの表示・スナップ、ズーム（50〜200%）はフッターで切り替えます。Frame やページへドラッグすると親が変わります。
- **選択と整列**: Ctrl / Shift+クリックや範囲選択で同じ親のコントロールを複数選択し、ツールバーから左右・上下の整列、等間隔の配置、サイズを揃える、重なり順の変更ができます。切り取り・コピー・貼り付け・複製（Ctrl+X / C / V / D）と削除（Delete）にも対応します。子を持つコンテナーを削除するときは確認します。
- **プロパティ**: 右のペインの［プロパティ］で名前・Caption・位置とサイズ・Enabled・Tag・ControlTipText などを編集します。Font と ForeColor / BackColor は専用のダイアログで編集できます。［構造］では重なったコントロールや非表示のコントロールも名前で選べ、［タブ順］では TabIndex をまとめて設定できます。MultiPage のページと TabStrip のタブは、選択したときに表示される一覧で追加・削除・並べ替えができます。
- **画像**: Image・Label・CommandButton・CheckBox・OptionButton・ToggleButton・Frame・Page とフォームの Picture の［…］で開くダイアログの「画像を選択…」からファイルを選ぶと、`picture.path` にプロジェクト ルート（`ras.toml` のあるフォルダー。見つからないときは `forms` の親フォルダー）からの相対パスを書き込みます。プロジェクトの中にあるファイルはその場所のまま参照し、外にあるファイルは `forms/assets/` にコピーします（同じ名前で内容の違うファイルがあるときは `logo-2.bmp` のように番号を付け、同じ内容のファイルがあればそれを参照します）。選べるのは BMP・JPEG・PNG・GIF・WebP・TIFF・ICO のファイル（16 MiB 以下）です。24 / 32 ビットの無圧縮 BMP と JPEG（1,600 万ピクセル以下）はそのままフォームに埋め込みます。それ以外の画像は、元のファイルと `picture.path` はそのままで、フォームをブックへ書き込むときに 24 ビットの BMP に変換します（変換後の BMP が 16 MiB 以下、およそ 550 万ピクセルまで。GIF は最初のフレーム）。透明な部分は、その Image コントロールの BackColor（指定が無ければ buttonFace の #F0F0F0）に重ねた色になります。このため、変換された画像を `ras form pull` で取り込むと BMP になります。形式はファイルの内容で判別し、対応していない形式や大きすぎる画像は選んだ時点でエラーにします（画像データの破損などはビルド時に報告されます）。キャンバスには PictureSizeMode と PictureAlignment に合わせて画像を表示します（TIFF は表示できないため、ファイル名だけを表示します）。Label などの画像はキャプションの横（PicturePosition）に、Frame・Page・フォームの画像は背景（PictureAlignment・PictureSizeMode・PictureTiling）として表示します。MouseIcon（MousePointer が Custom のときのポインター）には、アイコン（.ico）かカーソル（.cur）のファイルを同じ方法で選びます。デバッグ実行のフォームでは、これらの画像とマウス アイコンも表示します。「画像を削除」は `picture` を取り除きます（コピーしたファイルは、削除や元に戻す操作をしても `forms/assets/` に残ります）。`picture.path` のファイルが見つからないときや画像として使えないときは、［問題］パネルに警告を表示します。
- **リストの項目**: ComboBox / ListBox の List には、項目を 1 行に 1 つずつ入力します（`list` に文字列の配列として書き込みます。空にすると `list` を削除します）。複数の列を持つ項目は、列をタブで区切ります（Tab キーで入力でき、Excel からコピーしたセル範囲をそのまま貼り付けられます。`list` には列ごとの文字列の配列として書き込みます）。ColumnCount・ColumnWidths・BoundColumn・TextColumn・ColumnHeads も［プロパティ］で設定でき、キャンバスとデバッグ実行のフォームは列と見出しの行を表示します。SelectedIndex（`selectedIndex`）には最初に選択する項目の番号を 0 から数えて指定し、選択なしは -1 です。指定できるのは `list` にある項目の番号だけで、`list` が無いときは -1 だけです。List の項目を減らして選択中の番号が範囲を外れると、SelectedIndex は -1 に戻ります。項目と選択は、ビルドで生成されるコードがフォームの初期化時（`UserForm_Initialize`）に設定します。キャンバスの ListBox には項目を、ComboBox には選択中の項目を表示します。ComboBox の Style（DropDownCombo / DropDownList）・ShowDropButtonWhen・ListRows、ListBox の MultiSelect、両方の ListStyle・MatchEntry なども［プロパティ］で選べます。RowSource（`rowSource`）にセル範囲（`Sheet1!A2:A10`）を指定すると、項目はフォームを開いたときにそのセルから読み込まれます（List と同時には指定できません。キャンバスにはセル範囲を表示します）。値を持つコントロールの ControlSource（`controlSource`）には、値を連動させるセルを指定します。デバッグ実行でも、これらのセルを Excel から読み、ControlSource には確定した値を書き込みます。
- **コントロールのプロパティ**: Locked・MultiLine・PasswordChar・WordWrap・AutoSize・BackStyle・SpecialEffect・BorderColor・Alignment、CheckBox / OptionButton / ToggleButton の TripleState、CommandButton の Default・Cancel・TakeFocusOnClick、SpinButton / ScrollBar の Min・Max・SmallChange・LargeChange・Orientation、TabStrip / MultiPage の Style・TabOrientation・MultiRow、TabStop・MousePointer・IMEMode など、値が 1 つの VBA のプロパティは［プロパティ］で設定できます（選択肢は VBA の定数名で表示します）。指定しないものは Excel でコントロールを作ったときと同じ値になります。キャンバスとデバッグ実行のフォームは、SpecialEffect・BackStyle・Alignment・PasswordChar・ドロップ ボタン・タブの Style と TabOrientation（左右のタブの文字は縦書きに回転）などを反映して表示します。
- **CheckBox / OptionButton / ToggleButton の Value**: True・False・Null から選びます（Null は `"Null"` と書き込みます）。Null は Windows と同じく、網掛けの箱に灰色の印で表示します。デバッグ実行では、TripleState が True の CheckBox / ToggleButton はクリックのたびに False → Null → True → False の順に変わります（OptionButton には TripleState は効きません）。
- **フォームのプロパティ**: フォームを選んだときの［プロパティ］で、StartUpPosition・ShowModal・Tag・Enabled・BorderColor・SpecialEffect・Cycle・MousePointer・Zoom・ScrollBars・KeepScrollBarsVisible・ScrollWidth / ScrollHeight / ScrollLeft / ScrollTop・HelpContextID・WhatsThisButton / WhatsThisHelp・RightToLeft・DrawBuffer を設定できます（`form` に書き込みます）。キャンバスはフォームの枠（BorderStyle・BorderColor・SpecialEffect）を、デバッグ実行のフォームはそれに加えて Zoom・スクロール・マウス ポインターを反映します。
- **TextBox の高さ**: 1 行の TextBox / ComboBox がフォントの 1 行分より低いと、Excel では文字が欠けて端に線が残ります。デザイナーでフォントやサイズを変えると、足りない高さは自動で広げます。JSON を直接書いて低くなっているものは［問題］パネルに警告（`UFV022`）を表示します。
- **枠と立体の効果**: MSForms は BorderStyle が Single の枠と SpecialEffect（立体）を同時には描けず、VBE でも一方を設定するともう一方が戻ります。両方を指定したものは［問題］パネルに警告（`UFV023`）を表示し、ブックには枠を優先して（SpecialEffect を平面にして）書き込みます。
- **イベント**: コントロールやフォームの背景をダブルクリックすると、既定のイベント（CommandButton なら `Click`、TextBox なら `Change`、フォームなら `UserForm_Click`）のハンドラー（`@on("OkButton.Click")` を付けた関数）を `forms/<名前>.ras` に開きます。無ければ空のハンドラーを、イベントの引数（`cancel: MSForms.ReturnBoolean` など）付きで追加します。右クリックの［イベントを開く…］で他のイベントも選べます。
- **問題**: 下部の［問題］パネルに FormSpec の検証結果（重複した ID、存在しない親、種類に合わないプロパティなど）を表示し、選ぶとその項目を JSON 上で表示します。新しい検証エラーを生む編集は反映しません。
- **元に戻す・保存**: デザイナーでの操作はすべて JSON ドキュメントの編集になるため、VS Code の元に戻す / やり直し・保存・自動保存がそのまま使えます。JSON をテキストエディターで編集すると、デザイナーにもすぐに反映されます。

VBE で作ったフォームは、CLI の `ras form pull [名前…] [--book <ファイル>] [--force]` で `forms/` に取り込めます（`forms/<名前>.form.json` と `forms/<名前>.ras` を書き出し、画像は `forms/assets/` に書き出します。コードは VBA から RAS に変換し、変換できなかった行は `// TODO(VBA):` のコメントとして残します）。YAML 形式の FormSpec は扱いません。

## 制限モード（信頼されていないワークスペース）

信頼されていないワークスペースでは、シンタックスハイライトのみが有効です。言語サーバーは、ワークスペースを信頼した時点で起動します。

## 開発

```sh
npm ci
npm run compile   # 型チェック（tsc --noEmit）と esbuild によるバンドル
npm test          # TextMate 文法とユーザーフォーム デザイナーのテスト
npm run package   # ras-vscode-<version>.vsix を作成
```

Node.js 22 以降が必要です。`syntaxes/ras.tmLanguage.json` は `scripts/build-grammar.mjs` が生成するファイルです。デザイナーの画面（`webview/userFormDesigner/`）は esbuild で `dist/webview/` にバンドルされます。直接編集せず、スクリプトを変更して `npm run build:grammar` を実行してください。

## ライセンス

MIT License。この拡張機能が含む第三者の著作物（TypeScript の TextMate 文法、xlflow のユーザーフォーム デザイナーなど）については `THIRD_PARTY_NOTICES.md` を参照してください。
