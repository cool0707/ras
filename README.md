# RAS

TypeScript ライクな構文から VBA（`.bas` / `.cls`）へトランスパイルする Better VBA 言語。

- 言語仕様: [docs/RAS_Language_Specification.md](docs/RAS_Language_Specification.md)
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
| `ras-windows-x64` | `ras.exe` と `ras_com.node`（Windows x64） |
| `ras-windows-arm64` | `ras.exe` と `ras_com.node`（Windows arm64） |
| `ras-linux-x64` | `ras`（Linux x64。ビルド・チェック・オフライン書き込み用） |
| `ras-macos-arm64` | `ras`（macOS arm64。同上） |
| `ras-vscode` | VS Code 拡張（`.vsix`） |

`ras` は依存ランタイムのない単一バイナリです。PATH の通った場所に置くか、VS Code の設定 `ras.path` で場所を指定します。
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
ras fmt                                   # 整形
ras build --target js                     # デバッグ実行用の JavaScript を build/debug/ に出力
```

サンプルプロジェクトは [examples/sales-report](examples/sales-report) にあります（`ras --project examples/sales-report build`）。

`ras pack` はブックが Excel で開かれていると書き込みません（閉じてから実行してください）。
ブックに RAS が生成したものではない同名のモジュールがある場合や、手書きのイベントプロシージャ（`Workbook_Open` など）と生成コードの名前が重なる場合も、書き込まずにエラーにします（前者は `--overwrite` で置き換えられます）。
置き換え前のブックは `build/backup/` に保存されます。

## デバッグ

VS Code 拡張を入れた状態で `.ras` ファイルにブレークポイントを置き、F5（または「RAS: デバッグを開始」）で実行します。
ステップ実行、変数・ウォッチ式の確認、例外での停止が使えます。

仕組みは VBA のデバッグではありません。プログラムを JavaScript に変換して VS Code 同梱のランタイムで動かし、Excel の操作だけを実際の Excel に対して行います（仕様 §14.10）。

- Excel の操作には Windows と Excel が必要です。接続には拡張に同梱の `ras_com.node` を使います。Excel を使わない処理だけなら、どの OS でもデバッグできます。
- 対象のブック（`ras.toml` の `workbook`）が Excel で開かれていればそこに接続し、開かれていなければ開きます。
- プログラムが Excel の応答を待っている間に Excel がダイアログ（手書き VBA の実行時エラー・コンパイル エラー・`MsgBox`、Excel 自身の確認）を出すと、その内容を VS Code に通知し、ダイアログを前面に出します。`launch.json` の `"dialogs": "dismiss"` を指定すると、キャンセルに相当するボタンで自動的に閉じます（仕様 §14.10.6）。
- 実行できるのは引数なしの関数とトップレベルの文です。`@on` のイベントハンドラや UDF のように Excel から呼ばれる関数は起動できません。
- VBA ビルドの最適化（一括転送・インライン展開など）は通らないため、実行タイミングや速度は実機と異なります。`long` / `currency` / `decimal` の精度、未対応の VBA 標準関数など、実機との違いは仕様 §14.10.5 にまとめています。

`launch.json` で関数を固定することもできます。

```json
{ "type": "ras", "request": "launch", "name": "RAS: rebuildReport", "entry": "Report.rebuildReport" }
```

| コマンド | 状態 |
|---|---|
| `init` / `check` / `build` / `build --target js` / `watch` / `pack` / `unpack` / `fmt` / `lsp --stdio` | 実装済み |
| `sync` / `setup` / `typegen` / `pack --via-excel` | 未実装（実行するとその旨を表示して終了） |

## 実装状況

仕様書（Draft v0.7）に対する現状です。

| 仕様 | 状態 | 補足 |
|---|---|---|
| §2 型・数値リテラル・型変換 | 実装済み | `long` を `#If Win64` の外で使った場合のエラーは未実装（宣言の条件付き出力のみ） |
| §2.4 ジェネリクス | 実装済み | 関数・拡張メソッド・クラス。クラス内のジェネリックメソッドは未対応（エラー） |
| §3 関数・`byref`・拡張メソッド | 実装済み | |
| §4 ラムダ（ループ融合・クロージャクラス） | 実装済み | `map` / `filter` / `forEach` / `reduce` / `some` / `every` / `find` / `findIndex`。ホストのコレクション（`worksheets` など、§4.4）でも同じメソッドと `toArray()` を使える。`filter` の後段のコールバックが受け取る通し番号は、絞り込み後ではなく元の並びでの番号になる（JavaScript・デバッグ実行と異なる）。配列・文字列メソッドは仕様 §3.2.4 の「標準ライブラリの拡張メソッド」ではなくコンパイラ組み込みとして実装 |
| §5 `try` / `catch` / `finally` / `using` / `throw` | 実装済み | catch 本体はハンドラ内ではなく `Resume` 後のラベルで実行する（同一プロシージャ内の外側 `try` へ再送出を届けるため） |
| §6 リンター | 一部 | ルール 1〜6・8〜10・14 は実装。7（予約語のリネーム）は自動リネームのみで警告なし、11 は変換のみで通知なし |
| §7 CLI・VS Code 拡張 | 一部 | 上の表のとおり |
| §8 インライン展開・ループ最適化 | 一部 | `inline`、ループ内自動展開、評価順序の保存、`For...Next` 化、`$` 付き関数は実装。ループ不変式の巻き上げ（上限式以外）・`With` 化・文字列バッファは未実装 |
| §9 Range ⇔ 配列の一括転送 | 実装済み | `@bulk` の「既存内容を読んでから書き戻す」方式は未実装（条件 5 を満たさないループは変換しない） |
| §10 アプリケーション状態・`DoEvents` 自動挿入 | 実装済み | |
| §11 `declare`（32/64bit） | 一部 | `declare struct` のアラインメント計算と `@fallback` は未実装 |
| §12 ボイラープレート・イベント・`@init`・マクロ入口 | 一部 | リボン（`@ribbon`、§12.5）は未実装 |
| §13 VBA の罠の回避 | 一部 | `Collection` ラッパ、`Range.Find` ラッパ、文字列で指定される名前の存在確認、一時変数の再利用は未実装 |
| §14.1〜14.8 ライブ同期（RASBridge） | **未実装** | ドリフト検出用ハッシュ、ソースマップ、`Erl` 行番号は実装済み |
| §14.9 オフライン書き込み・読み出し | 実装済み | 新規ユーザーフォームの作成と再署名は未実装。`vba/` の非管理モジュールをブックへ書き戻す処理は未実装（取り出しのみ） |
| §14.10 デバッグ実行（JavaScript ランタイム） | 実装済み | Excel への接続（`ras_com.node`）は Excel 2021（64bit）の実機で確認済み。ダイアログの監視（§14.10.6）は Excel の無い Windows 上の CI でメッセージ ボックスを相手に確認しており、Excel の実機では未確認。CI では Excel の無い Windows 上で `Scripting.Dictionary` 等を使って COM 呼び出しを確認している。`super` 呼び出しを伴う継承は対象外 |
| §15 プロジェクト構成・`ras.toml`・同梱の型定義 | 一部 | Excel / VBA / Office の型定義を同梱し、メンバーとグローバルを camelCase で参照する（§15.3.1）。`ras typegen` と `vba:` インポートは未実装 |
| §16.1 JSDoc | 実装済み | 生成 VBA へのコメント・`VB_Description` 出力、`@param` 名の検査 |
| §16.2 IntelliSense（LSP） | 一部 | 診断・補完・ホバー・定義へ移動・シンボル一覧・整形・CodeLens・「生成 VBA を表示」と、VBA の綴りの名前を直すクイックフィックスは実装。パラメーター情報・参照検索・名前変更・インレイヒント・セマンティックトークン・その他のコードアクション・未インポート候補は未実装 |
| §17 採用ライブラリ | 一部 | `salsa`（増分計算。現状は変更のたびに全体を再解析）・`line-index`・`annotate-snippets`・`rayon`・`insta`・`lsp-types`・`windows` は未使用 |

仕様の例と出力が異なる点:

- 小数リテラル `1.5m` は `CDec("1.5")` ではなく `(CDec("15") / CDec("10"))` を出力します（小数点がカンマのロケールでも値が変わらないようにするため）。
- スクリプト形式（トップレベルの文）の `Main` にも、§12.4 のマクロ入口のラップが付きます。

## 検証状況

- 自動テスト: `cargo test --workspace`（構文、型チェック、仕様書のコード例と生成 VBA の照合、CLI の一連の流れ、LSP のセッション、`vbaProject.bin` の読み書き）。
- `ras pack` が出力したブックは、別実装の `ovba` クレートと oletools（`olevba`）で読み直して確認しています（CI でも実行）。
- 実機の Excel（Windows 11、Excel 2021 64bit）で次を確認しています（CI には Excel が無いため手動）:
  - `ras pack` で書き込んだブックを開いてマクロを実行できること。`crates/ras-compiler/tests/js/semantics.ras` の VBA ビルドの出力が `semantics.out` と一致すること
  - `RAS_Runtime.bas` の Win32 API・`SAFEARRAY` を扱う部分（`RAS_Rebase` のその場での下限書き換え、`RAS_Yield` の間引き、`sleep` / `waitUntil`）、Range ⇔ 配列の一括転送、`@on` のイベント、アプリケーション状態の復元
  - Excel が保存したブック（コンパイル済みの p-code 入り）への再 `pack`、開いているブックへの書き込み拒否、`unpack`
  - デバッグ実行（`ras_com.node`）が開いているブックに接続し、VBA ビルドと同じ結果・エラー番号になること
- デバッグ実行のダイアログ監視は、Excel / VBE の実際のダイアログ（実行時エラー、コンパイル エラー、Office 独自のダイアログ）では未検証です。判定規則は [xlflow](https://github.com/harumiWeb/xlflow) が実機で確立したもの（ボタンをコントロール ID とアクセスキーで識別する、実行時エラーは「終了」で閉じる、など）に合わせています。
- 32bit Office では未検証です（`RAS_Rebase` の 32bit 側のオフセットなど）。
- VS Code 拡張はパッケージ化と文法テストまで確認済みで、実機の VS Code 上では動かしていません。

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
| `editors/vscode` | VS Code 拡張 |
| `examples/sales-report` | サンプルプロジェクト |
