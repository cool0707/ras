# RAS Playground

RAS をブラウザで書いて VBA に変換し、Excel なしで実行・テストするページです。GitHub Pages
（https://cool0707.github.io/ras/）で公開しています。

## しくみ

| 部分 | 中身 |
|---|---|
| `crates/ras-wasm` | コンパイラー（`compile`・`format`・`explain`）、`ras pack`（`packBook`）、ブックの代役（`Book`）、`ras book diff` / `view` / `edit`（`editBook`）を WebAssembly にしたもの。ブックタブの画面（`view.js`）は、実行前のブックへの編集を親のページに `postMessage` し、`main.ts` が `editBook` で書き換えて描き直します |
| `src/main.ts` | 画面: エディタ（CodeMirror）、VBA・問題・実行結果・ブックの表示。ブックタブの図形・コントロールのイベント（`objects/<コード名>.ras` のハンドラー、マクロの登録）も扱います |
| `src/designer.ts` | ユーザーフォームのデザイナー。VS Code 拡張のデザイナー（`editors/vscode/webview/userFormDesigner`）を `designer.html` のフレームに表示し、拡張のプロバイダー（`src/userForm/provider.ts`）が VS Code のドキュメントで行うことを、ページのファイルで行います |
| `src/forms.ts` | 実行中のユーザーフォームの表示。VS Code 拡張のフォームのパネル（`webview/formRunner`）を `runner.html` のフレームでページの上に表示し、ユーザーの操作を Worker に返します（Worker は `Atomics.wait` で待ちます） |
| `src/worker.ts` | 実行用の Web Worker。止まらないプログラムでもページは固まらず、「停止」で終了します |
| `src/runtime/` | ブラウザの中の小さな Node.js（`fs` はメモリ上、`require`、`process`）。`ras build --target js` が出力したデバッグ用の JavaScript（`ras-main.js` / `ras-test-main.js` とランタイム）を、CLI と同じものをそのまま動かします。`ras_book` アドオンの代わりは wasm の `Book` です |
| `static/sw.js` | GitHub Pages ではヘッダーを付けられないため、Service Worker で COOP / COEP を付けて `SharedArrayBuffer` を使えるようにします。実行中の `msgBox` / `inputBox` は、Worker が `Atomics.wait` で待つ間にページがダイアログを表示して答えます（使えないブラウザでは `ras run` と同じ既定の応答） |
| `samples/` | サンプル（`examples/sales-report` も含めます） |
| `static/designer.html`・`runner.html`・`frame.css` | VS Code 拡張の Webview をフレームで動かすページ。`acquireVsCodeApi` を親のページへの `postMessage` にし、VS Code のテーマの色（`--vscode-*`）をプレイグラウンドの色にします |

`ras_book` アドオン（`crates/ras-book`）とブラウザの `Book` は、どちらも `ras_workbook::js` で値を変換します。アドオンにメソッドを足したら、`ras_workbook::js::call` にも足してください（`cargo test -p ras-workbook` が確かめます）。

## 開発

Rust（`rust-toolchain.toml` のもの）に加えて、次が必要です。

```sh
rustup target add wasm32-unknown-unknown
cargo install wasm-bindgen-cli --version 0.2.130 --locked   # Cargo.lock の wasm-bindgen と同じ版
```

```sh
cd web
npm ci
npm run wasm          # crates/ras-wasm をビルドして web/pkg/ に出力（--dev でデバッグビルド）
npm test              # サンプルのコンパイル・実行・テスト・.xlsm への書き込みを Node で確かめる
npm run check-types
npm run serve         # web/dist/ を作って http://localhost:8000/ で表示
npm run build         # web/dist/ を作る（GitHub Pages に置くもの）
```

`.github/workflows/pages.yml` が、すべての push と pull request でビルドとテストを行い、`main` では GitHub Pages に公開します。
初めて公開するときは、リポジトリの Settings → Pages の Source を「GitHub Actions」にしてください。
