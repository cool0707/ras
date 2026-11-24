//! Japanese eras (和暦), as Japanese Excel reads and shows them: dates typed with an era
//! (`R1/5/1`, `令和元年5月1日`), and the formats that show one (`ge.m.d`, `ggge"年"`).

/// The eras: the letter (`g`), the first character (`gg`), the name (`ggg`) and the year before
/// the first.
const ERAS: [(char, char, &str, i32); 5] = [
    ('M', '明', "明治", 1867),
    ('T', '大', "大正", 1911),
    ('S', '昭', "昭和", 1925),
    ('H', '平', "平成", 1988),
    ('R', '令', "令和", 2018),
];

/// The format Japanese Excel gives a cell a date is typed into as `R1/5/1`.
pub(crate) const ERA_SHORT: &str = "[$-411]ge.m.d";
/// The format Japanese Excel gives a cell a date is typed into as `令和1年5月1日`.
pub(crate) const ERA_LONG: &str = "[$-411]ggge\"年\"m\"月\"d\"日\"";
/// The formula bar's date in a cell shown with an era: `R1/5/1`.
pub(crate) const ERA_INPUT: &str = "[$-411]ge/m/d";

/// Text typed with an era as the same date in the Gregorian calendar (`R1/5/1 9:00` →
/// `2019/5/1 9:00`), with the format the cell gets when it has none (`None` with a time).
pub(crate) fn gregorian(text: &str) -> Option<(String, Option<&'static str>)> {
    let t = text.trim();
    let (base, rest) = ERAS.iter().find_map(|(letter, first, name, base)| {
        let rest = t
            .strip_prefix(name)
            .or_else(|| t.strip_prefix(*first))
            .or_else(|| t.strip_prefix(*letter))
            .or_else(|| t.strip_prefix(letter.to_ascii_lowercase()))?;
        Some((*base, rest))
    })?;
    let (year, rest) = match rest.strip_prefix('元') {
        Some(r) => (1, r),
        None => number(rest)?,
    };
    let kanji = rest.starts_with('年');
    let rest = rest.strip_prefix(['/', '.', '-', '年'])?;
    let (month, rest) = number(rest)?;
    let rest = rest.strip_prefix(['/', '.', '-', '月'])?;
    let (day, rest) = number(rest)?;
    let rest = rest.strip_prefix('日').unwrap_or(rest);
    if year < 1 || !(1..=12).contains(&month) || !(1..=31).contains(&day) {
        return None;
    }
    if !rest.is_empty() && !rest.starts_with(char::is_whitespace) {
        return None;
    }
    let format = rest
        .is_empty()
        .then_some(if kanji { ERA_LONG } else { ERA_SHORT });
    Some((format!("{}/{month}/{day}{rest}", base + year), format))
}

fn number(s: &str) -> Option<(i32, &str)> {
    let end = s.find(|c: char| !c.is_ascii_digit()).unwrap_or(s.len());
    if end == 0 || end > 4 {
        return None;
    }
    Some((s[..end].parse().ok()?, &s[end..]))
}

/// Whether a format shows an era (`g`, or `e` the year in it), outside quotes and brackets.
pub(crate) fn shows_era(code: &str) -> bool {
    if code.eq_ignore_ascii_case("General") {
        return false;
    }
    let cs: Vec<char> = code.chars().collect();
    let mut i = 0;
    while i < cs.len() {
        match cs[i] {
            '"' => {
                i += 1;
                while i < cs.len() && cs[i] != '"' {
                    i += 1;
                }
            }
            '\\' | '_' | '*' => i += 1,
            '[' => {
                while i < cs.len() && cs[i] != ']' {
                    i += 1;
                }
            }
            'g' | 'G' => {
                let rest: String = cs[i..].iter().take(7).collect();
                if rest.eq_ignore_ascii_case("general") {
                    i += 7;
                    continue;
                }
                if rest.starts_with("G/標準") {
                    i += 4;
                    continue;
                }
                return true;
            }
            'e' | 'E' if !matches!(cs.get(i + 1), Some('+' | '-')) => return true,
            _ => {}
        }
        i += 1;
    }
    false
}

#[cfg(test)]
mod tests {
    use super::*;

    #[test]
    fn eras_typed() {
        let g = |s: &str| gregorian(s);
        assert_eq!(g("R1/5/1"), Some(("2019/5/1".into(), Some(ERA_SHORT))));
        assert_eq!(g("r2.3.4"), Some(("2020/3/4".into(), Some(ERA_SHORT))));
        assert_eq!(
            g("令和元年5月1日"),
            Some(("2019/5/1".into(), Some(ERA_LONG)))
        );
        assert_eq!(
            g("平成31年4月30日"),
            Some(("2019/4/30".into(), Some(ERA_LONG)))
        );
        assert_eq!(g("S64/1/7"), Some(("1989/1/7".into(), Some(ERA_SHORT))));
        assert_eq!(g("H1/1/8 9:30"), Some(("1989/1/8 9:30".into(), None)));
        assert_eq!(g("R1/13/1"), None);
        assert_eq!(g("Rate"), None);
        assert_eq!(g("2019/5/1"), None);
    }

    #[test]
    fn formats_with_eras() {
        assert!(shows_era(ERA_SHORT));
        assert!(shows_era(ERA_LONG));
        assert!(shows_era("[$-ja-JP-x-gannen]ggge\"年\""));
        assert!(!shows_era("General"));
        assert!(!shows_era("G/標準"));
        assert!(!shows_era("0.00E+00"));
        assert!(!shows_era("yyyy\"年\"m\"月\"d\"日\""));
        assert!(!shows_era("[Red]#,##0"));
    }
}
