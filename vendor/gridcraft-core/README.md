# vendor/gridcraft-core

`gridcraft-core` of xljs (`vendor/gridcraft/core` of https://github.com/cool0707/xljs at commit
`60e7e7f`, the one `crates/ras-workbook` uses), put in place of it by the `[patch]` of the
workspace's `Cargo.toml`.

The only change: the dependency on `js-sys` for wasm32 (`Cargo.toml`), which `src/date.rs` uses
there (`js_sys::Date::now()`) and xljs's copy does not declare. Without it the stand-in for Excel
does not build for the browser (`crates/ras-wasm`, the playground). Once xljs declares it, this
directory and the `[patch]` can go.

License: MIT OR Apache-2.0 (`LICENSE-MIT`, `LICENSE-APACHE`, `NOTICE`).
