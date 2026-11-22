//! `ras sheet`: adding, copying, removing and renaming the sheets of the project's workbook
//! together with their files in `objects/` (spec §14.16).

use std::path::Path;
use std::process::ExitCode;

use ras_ovba::{Place, SheetEdit};
use ras_sync::sheets::{self, Failure, Outcome};

#[derive(clap::Args, Clone, Default)]
pub struct PlaceArgs {
    /// このシートの前に置く（コード名か見出しの名前）
    #[arg(long, value_name = "SHEET", conflicts_with = "after")]
    before: Option<String>,
    /// このシートの後ろに置く（コード名か見出しの名前）
    #[arg(long, value_name = "SHEET")]
    after: Option<String>,
}

impl PlaceArgs {
    fn place(&self) -> Option<Place> {
        match (&self.before, &self.after) {
            (Some(s), _) => Some(Place::Before(s.clone())),
            (_, Some(s)) => Some(Place::After(s.clone())),
            _ => None,
        }
    }
}

#[derive(clap::Subcommand)]
pub enum SheetCommand {
    /// シートの一覧（見出しの名前・コード名・ファイル・表示状態）を表示する
    List {
        /// 結果を JSON で出力する
        #[arg(long)]
        json: bool,
    },
    /// 空のワークシートを追加し、objects/<コード名>.ras を作る
    Add {
        /// 見出しの名前
        name: String,
        /// コード名（省略時は、空いている最初の Sheet<番号>）
        #[arg(long, value_name = "NAME")]
        code_name: Option<String>,
        #[command(flatten)]
        place: PlaceArgs,
        /// 結果を JSON で出力する
        #[arg(long)]
        json: bool,
    },
    /// ワークシートを複製し、そのファイルも複製する
    Copy {
        /// 複製するシート（コード名か見出しの名前）
        sheet: String,
        /// 複製先の見出しの名前（省略時は「<元の名前> (2)」）
        name: Option<String>,
        /// 複製先のコード名（省略時は、空いている最初の Sheet<番号>）
        #[arg(long, value_name = "NAME")]
        code_name: Option<String>,
        #[command(flatten)]
        place: PlaceArgs,
        /// 結果を JSON で出力する
        #[arg(long)]
        json: bool,
    },
    /// ワークシートと、そのファイルを削除する
    Remove {
        /// 削除するシート（コード名か見出しの名前）
        sheet: String,
        /// ファイルにコードがある・ほかから参照されている場合も削除する
        #[arg(long)]
        force: bool,
        /// 結果を JSON で出力する
        #[arg(long)]
        json: bool,
    },
    /// ワークシートの見出しの名前・コード名を変える（コード名を変えると、ファイル名とソースの参照も変える）
    Rename {
        /// 名前を変えるシート（コード名か見出しの名前）
        sheet: String,
        /// 新しい見出しの名前
        #[arg(long, value_name = "NAME")]
        name: Option<String>,
        /// 新しいコード名
        #[arg(long, value_name = "NAME")]
        code_name: Option<String>,
        /// 結果を JSON で出力する
        #[arg(long)]
        json: bool,
    },
}

/// The outcome as one JSON object.
pub fn outcome_json(outcome: &Outcome) -> serde_json::Value {
    serde_json::json!({
        "ok": true,
        "action": outcome.action,
        "name": outcome.name,
        "codeName": outcome.code_name,
        "before": outcome.before.as_ref().map(|(name, code_name)| serde_json::json!({ "name": name, "codeName": code_name })),
        "live": outcome.live,
        "created": outcome.created,
        "removed": outcome.removed,
        "renamed": outcome.renamed.iter().map(|(from, to)| serde_json::json!({ "from": from, "to": to })).collect::<Vec<_>>(),
        "references": outcome.references,
        "formulas": outcome.formulas,
        "warnings": outcome.warnings,
    })
}

fn describe(outcome: &Outcome) -> String {
    let sheet = format!(
        "シート「{}」（コード名 {}）",
        outcome.name, outcome.code_name
    );
    let mut lines = vec![match outcome.action {
        "add" => format!("{sheet}を追加しました"),
        "copy" => format!(
            "{sheet}を、シート「{}」の複製として作りました",
            outcome
                .before
                .as_ref()
                .map(|b| b.0.as_str())
                .unwrap_or_default()
        ),
        "remove" => format!("{sheet}を削除しました"),
        _ => match &outcome.before {
            Some((name, code_name)) => {
                format!("シート「{name}」（コード名 {code_name}）を、{sheet}に変えました")
            }
            None => format!("{sheet}の名前を変えました"),
        },
    }];
    for file in &outcome.created {
        lines.push(format!("  作成: {file}"));
    }
    for file in &outcome.removed {
        lines.push(format!("  削除: {file}"));
    }
    for (from, to) in &outcome.renamed {
        lines.push(format!("  名前の変更: {from} → {to}"));
    }
    if outcome.references > 0 {
        lines.push(format!(
            "  ソースの参照を {} 箇所書き換えました",
            outcome.references
        ));
    }
    if outcome.formulas > 0 {
        lines.push(format!(
            "  ブックの数式・名前の定義を {} 箇所書き換えました",
            outcome.formulas
        ));
    }
    lines.join("\n")
}

pub fn run(root: &Path, command: &SheetCommand) -> Result<ExitCode, String> {
    let (edit, force, json) = match command {
        SheetCommand::List { json } => {
            let (rows, live) = sheets::list(root)?;
            if *json {
                let list: Vec<serde_json::Value> = rows.iter().map(|r| serde_json::json!({ "name": r.name, "codeName": r.code_name, "worksheet": r.worksheet, "visible": r.visible, "file": r.file })).collect();
                println!("{}", serde_json::json!({ "live": live, "sheets": list }));
            } else {
                if live {
                    println!("（Excel で開いているブックのシートです）");
                }
                for (i, r) in rows.iter().enumerate() {
                    let mut notes = Vec::new();
                    if !r.worksheet {
                        notes.push("ワークシート以外".to_string());
                    }
                    if !r.visible {
                        notes.push("非表示".to_string());
                    }
                    let notes = if notes.is_empty() {
                        String::new()
                    } else {
                        format!("  [{}]", notes.join("・"))
                    };
                    println!(
                        "{:>3}. {}（コード名 {}）  {}{notes}",
                        i + 1,
                        r.name,
                        r.code_name,
                        r.file.as_deref().unwrap_or("ファイルなし")
                    );
                }
            }
            return Ok(ExitCode::SUCCESS);
        }
        SheetCommand::Add {
            name,
            code_name,
            place,
            json,
        } => (
            SheetEdit::Add {
                name: name.clone(),
                code_name: code_name.clone(),
                place: place.place().unwrap_or(Place::End),
            },
            false,
            *json,
        ),
        SheetCommand::Copy {
            sheet,
            name,
            code_name,
            place,
            json,
        } => (
            SheetEdit::Copy {
                sheet: sheet.clone(),
                name: name.clone(),
                code_name: code_name.clone(),
                place: place.place(),
            },
            false,
            *json,
        ),
        SheetCommand::Remove { sheet, force, json } => (
            SheetEdit::Remove {
                sheet: sheet.clone(),
            },
            *force,
            *json,
        ),
        SheetCommand::Rename {
            sheet,
            name,
            code_name,
            json,
        } => (
            SheetEdit::Rename {
                sheet: sheet.clone(),
                name: name.clone(),
                code_name: code_name.clone(),
            },
            false,
            *json,
        ),
    };
    match sheets::apply(root, &edit, force) {
        Ok(outcome) => {
            if json {
                println!("{}", outcome_json(&outcome));
            } else {
                println!("{}", describe(&outcome));
                for warning in &outcome.warnings {
                    eprintln!("warning: {warning}");
                }
            }
            Ok(ExitCode::SUCCESS)
        }
        Err(Failure::Refused(reasons)) => {
            if json {
                println!("{}", serde_json::json!({ "ok": false, "refused": reasons }));
            } else {
                eprintln!("error: シートを削除していません。次の理由があります（それでも削除するには --force を付けてください）:");
                for reason in &reasons {
                    eprintln!("  - {reason}");
                }
            }
            Ok(ExitCode::from(1))
        }
        Err(Failure::Error(message)) => {
            if json {
                println!("{}", serde_json::json!({ "ok": false, "error": message }));
                Ok(ExitCode::from(1))
            } else {
                Err(message)
            }
        }
    }
}
