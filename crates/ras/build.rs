// Embeds the icon (assets/ras.ico, the same picture as the VS Code extension's
// editors/vscode/images/icon.png) and the version information in ras.exe.
fn main() {
    println!("cargo:rerun-if-changed=assets/ras.ico");
    if std::env::var("CARGO_CFG_TARGET_OS").as_deref() != Ok("windows") {
        return;
    }
    let mut res = winresource::WindowsResource::new();
    res.set_icon("assets/ras.ico");
    res.set(
        "FileDescription",
        "RAS: a TypeScript-like language that transpiles to VBA",
    );
    res.set("ProductName", "RAS");
    if let Err(e) = res.compile() {
        panic!("could not embed the icon in ras.exe: {e}");
    }
}
