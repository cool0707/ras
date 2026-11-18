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
        ws.Cells(r, 3).Value = r * 2;      // ループごと Range ⇔ 配列の一括転送に変換される
    }
}

export function lastRow(this: Excel.Worksheet, col: int = 1): int {
    return this.Cells(this.Rows.Count, col).End(xlUp).Row;
}
```

## ダウンロード

ビルド済みのバイナリと VS Code 拡張は GitHub Actions の成果物として取得できます。

1. リポジトリの **Actions** タブで、ワークフロー **CI** の成功した実行を開く
2. ページ下部の **Artifacts** からダウンロードする

| 成果物 | 内容 |
|---|---|
| `ras-windows-x64` | `ras.exe`（Windows x64） |
| `ras-windows-arm64` | `ras.exe`（Windows arm64） |
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
ras build                                 # build/ に .bas / .cls / .rasmap を出力
ras pack                                  # Excel を起動せず、Book.xlsm の VBA プロジェクトへ直接書き込む
ras watch                                 # 変更を監視して再ビルド（ブックが閉じていれば pack も行う）
ras unpack                                # ブック内の手書き VBA を vba/ に取り出す
ras fmt                                   # 整形
```

`ras pack` はブックが Excel で開かれていると書き込みません（閉じてから実行してください）。
置き換え前のブックは `build/backup/` に保存されます。

| コマンド | 状態 |
|---|---|
| `init` / `check` / `build` / `watch` / `pack` / `unpack` / `fmt` / `lsp --stdio` | 実装済み |
| `sync` / `setup` / `typegen` / `pack --via-excel` | 未実装（実行するとその旨を表示して終了） |

## 実装状況

仕様書（Draft v0.7）に対する現状です。

| 仕様 | 状態 | 補足 |
|---|---|---|
| §2 型・数値リテラル・型変換 | 実装済み | `long` を `#If Win64` の外で使った場合のエラーは未実装（宣言の条件付き出力のみ） |
| §2.4 ジェネリクス | 実装済み | 関数・拡張メソッド・クラス。クラス内のジェネリックメソッドは未対応（エラー） |
| §3 関数・`byref`・拡張メソッド | 実装済み | |
| §4 ラムダ（ループ融合・クロージャクラス） | 実装済み | `map` / `filter` / `forEach` / `reduce` / `some` / `every` / `find` / `findIndex`。配列・文字列メソッドは仕様 §3.2.4 の「標準ライブラリの拡張メソッド」ではなくコンパイラ組み込みとして実装 |
| §5 `try` / `catch` / `finally` / `using` / `throw` | 実装済み | catch 本体はハンドラ内ではなく `Resume` 後のラベルで実行する（同一プロシージャ内の外側 `try` へ再送出を届けるため） |
| §6 リンター | 一部 | ルール 1〜6・8〜10 は実装。7（予約語のリネーム）は自動リネームのみで警告なし、11 は変換のみで通知なし |
| §7 CLI・VS Code 拡張 | 一部 | 上の表のとおり |
| §8 インライン展開・ループ最適化 | 一部 | `inline`、ループ内自動展開、評価順序の保存、`For...Next` 化、`$` 付き関数は実装。ループ不変式の巻き上げ（上限式以外）・`With` 化・文字列バッファは未実装 |
| §9 Range ⇔ 配列の一括転送 | 実装済み | `@bulk` の「既存内容を読んでから書き戻す」方式は未実装（条件 5 を満たさないループは変換しない） |
| §10 アプリケーション状態・`DoEvents` 自動挿入 | 実装済み | |
| §11 `declare`（32/64bit） | 一部 | `declare struct` のアラインメント計算と `@fallback` は未実装 |
| §12 ボイラープレート・イベント・`@init`・マクロ入口 | 一部 | リボン（`@ribbon`、§12.5）は未実装 |
| §13 VBA の罠の回避 | 一部 | `Collection` ラッパ、`Range.Find` ラッパ、文字列で指定される名前の存在確認、一時変数の再利用は未実装 |
| §14.1〜14.8 ライブ同期（RASBridge） | **未実装** | ドリフト検出用ハッシュ、ソースマップ、`Erl` 行番号は実装済み |
| §14.9 オフライン書き込み・読み出し | 実装済み | 新規ユーザーフォームの作成と再署名は未実装。`vba/` の非管理モジュールをブックへ書き戻す処理は未実装（取り出しのみ） |
| §15 プロジェクト構成・`ras.toml`・同梱の型定義 | 一部 | Excel / VBA / Office の型定義を同梱。`ras typegen` と `vba:` インポートは未実装 |
| §16.1 JSDoc | 実装済み | 生成 VBA へのコメント・`VB_Description` 出力、`@param` 名の検査 |
| §16.2 IntelliSense（LSP） | 一部 | 診断・補完・ホバー・定義へ移動・シンボル一覧・整形・CodeLens・「生成 VBA を表示」は実装。パラメーター情報・参照検索・名前変更・インレイヒント・セマンティックトークン・コードアクション・未インポート候補は未実装 |
| §17 採用ライブラリ | 一部 | `salsa`（増分計算。現状は変更のたびに全体を再解析）・`line-index`・`annotate-snippets`・`rayon`・`insta`・`lsp-types`・`windows` は未使用 |

仕様の例と出力が異なる点:

- 小数リテラル `1.5m` は `CDec("1.5")` ではなく `(CDec("15") / CDec("10"))` を出力します（小数点がカンマのロケールでも値が変わらないようにするため）。
- スクリプト形式（トップレベルの文）の `Main` にも、§12.4 のマクロ入口のラップが付きます。

## 検証状況

- 自動テスト: `cargo test --workspace`（構文、型チェック、仕様書のコード例と生成 VBA の照合、CLI の一連の流れ、LSP のセッション、`vbaProject.bin` の読み書き）。
- `ras pack` が出力したブックは、別実装の `ovba` クレートと oletools（`olevba`）で読み直して確認しています（CI でも実行）。
- **実機の Excel では未検証です。** この開発環境にも CI にも Excel が無いため、生成した VBA を実際に実行したこと、`ras pack` で書き込んだブックを Excel で開いたことはありません。特に `RAS_Runtime.bas` のうち Win32 API や `SAFEARRAY` を直接扱う部分（`RAS_Rebase`、`RAS_Yield`）は、実機での確認が必要です。
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
