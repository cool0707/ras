//! Parses the given files and prints syntax errors (`cargo run -p ras-syntax --example parsecheck -- a.ras`).
fn main() {
    let mut failed = false;
    for path in std::env::args().skip(1) {
        let text = match std::fs::read_to_string(&path) {
            Ok(t) => t,
            Err(e) => {
                eprintln!("{path}: {e}");
                failed = true;
                continue;
            }
        };
        let parse = ras_syntax::parse(&text);
        for e in &parse.errors {
            let offset = u32::from(e.range.start()) as usize;
            let line = text[..offset].matches('\n').count() + 1;
            let col = offset - text[..offset].rfind('\n').map_or(0, |i| i + 1) + 1;
            println!("{path}:{line}:{col}: {}", e.message);
            failed = true;
        }
        println!("{path}: {} top-level items, {} errors", parse.ast.items.len(), parse.errors.len());
    }
    if failed {
        std::process::exit(1);
    }
}
