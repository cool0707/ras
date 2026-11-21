use super::*;

fn at(a1: &str) -> CellRef {
    CellRef::parse(a1).unwrap()
}

fn rng(a1: &str) -> RangeRef {
    RangeRef::parse(a1).unwrap()
}

fn num(n: f64) -> Val {
    Val::Num(n)
}

fn text(s: &str) -> Val {
    Val::Text(s.into())
}

fn put(b: &mut Book, a1: &str, v: Val) {
    b.set_values(0, rng(a1), &[vec![v]]).unwrap();
}

fn read(b: &mut Book, a1: &str, what: Read) -> Val {
    b.get(0, rng(a1), what).unwrap().remove(0).remove(0)
}

#[test]
fn formulas_follow_what_is_written() {
    let mut b = Book::new_book();
    put(&mut b, "A1", num(2.0));
    put(&mut b, "A2", num(3.0));
    b.set_formulas(0, rng("A3"), &[vec![text("=SUM(A1:A2)*10")]], false).unwrap();
    assert_eq!(read(&mut b, "A3", Read::Value), num(50.0));
    put(&mut b, "A1", num(7.0));
    assert_eq!(read(&mut b, "A3", Read::Value), num(100.0));
    assert_eq!(read(&mut b, "A3", Read::Formula), text("=SUM(A1:A2)*10"));
    // A constant's formula is its value.
    assert_eq!(read(&mut b, "A1", Read::Formula), text("7"));
    assert_eq!(read(&mut b, "B1", Read::Formula), text(""));
}

#[test]
fn text_is_read_as_typed_text() {
    let mut b = Book::new_book();
    put(&mut b, "A1", text("123"));
    assert_eq!(read(&mut b, "A1", Read::Value), num(123.0));
    put(&mut b, "A2", text("=A1+1"));
    assert_eq!(read(&mut b, "A2", Read::Value), num(124.0));
    put(&mut b, "A3", text("'123"));
    assert_eq!(read(&mut b, "A3", Read::Value), text("123"));
    put(&mut b, "A4", text("2026/3/4"));
    assert_eq!(read(&mut b, "A4", Read::NumberFormat), text(DATE_FORMAT));
    let Val::Date(d) = read(&mut b, "A4", Read::Value) else { panic!("not a date") };
    assert_eq!(read(&mut b, "A4", Read::Value2), num(d));
    assert_eq!(read(&mut b, "A4", Read::Text), text("2026/3/4"));
    put(&mut b, "A5", text("abc"));
    assert_eq!(read(&mut b, "A5", Read::Value), text("abc"));
    // A cell formatted as text keeps what is written.
    b.set_number_format(0, rng("A6"), "@").unwrap();
    put(&mut b, "A6", text("007"));
    assert_eq!(read(&mut b, "A6", Read::Value), text("007"));
}

#[test]
fn dates_get_a_date_format() {
    let mut b = Book::new_book();
    put(&mut b, "A1", Val::Date(46085.0));
    assert_eq!(read(&mut b, "A1", Read::Value), Val::Date(46085.0));
    assert_eq!(read(&mut b, "A1", Read::Value2), num(46085.0));
    assert_eq!(read(&mut b, "A1", Read::Text), text("2026/3/4"));
}

#[test]
fn arrays_are_laid_over_ranges() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:B2"), &[vec![num(1.0), num(2.0)], vec![num(3.0), num(4.0)]]).unwrap();
    assert_eq!(b.get(0, rng("A1:B2"), Read::Value).unwrap(), vec![vec![num(1.0), num(2.0)], vec![num(3.0), num(4.0)]]);
    // One value fills every cell; one row every row; beyond the array #N/A.
    b.set_values(0, rng("D1:E2"), &[vec![num(9.0)]]).unwrap();
    assert_eq!(b.get(0, rng("D1:E2"), Read::Value).unwrap(), vec![vec![num(9.0), num(9.0)], vec![num(9.0), num(9.0)]]);
    b.set_values(0, rng("G1:I2"), &[vec![num(1.0), num(2.0)]]).unwrap();
    assert_eq!(b.get(0, rng("G1:I1"), Read::Value).unwrap(), vec![vec![num(1.0), num(2.0), Val::Err(2042)]]);
    assert_eq!(b.get(0, rng("G2:H2"), Read::Value).unwrap(), vec![vec![num(1.0), num(2.0)]]);
}

#[test]
fn one_formula_over_a_range_moves_its_references() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A3"), &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]]).unwrap();
    b.set_formulas(0, rng("B1:B3"), &[vec![text("=A1*2")]], false).unwrap();
    assert_eq!(b.get(0, rng("B1:B3"), Read::Value).unwrap(), vec![vec![num(2.0)], vec![num(4.0)], vec![num(6.0)]]);
    assert_eq!(read(&mut b, "B3", Read::Formula), text("=A3*2"));
}

#[test]
fn formula2_spills_and_formula_intersects() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A3"), &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]]).unwrap();
    b.set_formulas(0, rng("C1"), &[vec![text("=A1:A3*10")]], true).unwrap();
    assert_eq!(b.get(0, rng("C1:C3"), Read::Value).unwrap(), vec![vec![num(10.0)], vec![num(20.0)], vec![num(30.0)]]);
    assert_eq!(read(&mut b, "C1", Read::Formula2), text("=A1:A3*10"));
    // Range.Formula: implicit intersection, so the formula in row 2 takes A2.
    b.set_formulas(0, rng("D2"), &[vec![text("=A1:A3*10")]], false).unwrap();
    assert_eq!(read(&mut b, "D2", Read::Value), num(20.0));
    assert_eq!(read(&mut b, "D3", Read::Value), Val::Empty);
    assert_eq!(read(&mut b, "D2", Read::Formula), text("=A1:A3*10"));
    assert_eq!(read(&mut b, "D2", Read::Formula2), text("=@A1:A3*10"));
}

#[test]
fn errors_are_cverr_numbers() {
    let mut b = Book::new_book();
    b.set_formulas(0, rng("A1"), &[vec![text("=1/0")]], false).unwrap();
    b.set_formulas(0, rng("A2"), &[vec![text("=NA()")]], false).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), Val::Err(2007));
    assert_eq!(read(&mut b, "A2", Read::Value), Val::Err(2042));
    assert_eq!(read(&mut b, "A1", Read::Text), text("#DIV/0!"));
}

#[test]
fn end_goes_where_ctrl_arrow_goes() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A5"), &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)], vec![Val::Empty], vec![num(5.0)]]).unwrap();
    assert_eq!(b.end(0, CellRef::new(MAX_ROWS - 1, 0), Dir::Up), at("A5"));
    assert_eq!(b.end(0, at("A1"), Dir::Down), at("A3"));
    assert_eq!(b.end(0, at("A3"), Dir::Down), at("A5"));
    assert_eq!(b.end(0, at("A5"), Dir::Down), CellRef::new(MAX_ROWS - 1, 0));
    assert_eq!(b.end(0, at("A5"), Dir::Up), at("A3"));
    assert_eq!(b.end(0, at("B1"), Dir::Down), CellRef::new(MAX_ROWS - 1, 1));
    assert_eq!(b.end(0, CellRef::new(MAX_ROWS - 1, 1), Dir::Up), at("B1"));
    assert_eq!(b.end(0, at("A1"), Dir::Right), CellRef::new(0, MAX_COLS - 1));
    assert_eq!(b.end(0, at("E1"), Dir::Left), at("A1"));
}

#[test]
fn regions_and_used_range() {
    let mut b = Book::new_book();
    assert_eq!(b.used_range(0), None);
    b.set_values(0, rng("B2:C3"), &[vec![num(1.0), num(2.0)], vec![num(3.0), num(4.0)]]).unwrap();
    put(&mut b, "D4", num(5.0));
    put(&mut b, "F9", num(6.0));
    assert_eq!(b.current_region(0, at("B2")), rng("B2:D4"));
    assert_eq!(b.used_range(0), Some(rng("B2:F9")));
    b.clear(0, rng("F9"), false).unwrap();
    assert_eq!(b.used_range(0), Some(rng("B2:D4")));
}

#[test]
fn sheets_and_names() {
    let mut b = Book::new_book();
    assert_eq!(b.sheet_name(0), Some("Sheet1"));
    assert_eq!(b.code_name(0).as_deref(), Some("Sheet1"));
    let i = b.add_sheet(1, Some("Data")).unwrap();
    assert_eq!(i, 1);
    assert_eq!(b.code_name(1).as_deref(), Some("Sheet2"));
    b.set_values(1, rng("A1"), &[vec![num(4.0)]]).unwrap();
    b.set_formulas(0, rng("A1"), &[vec![text("=Data!A1*2")]], false).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), num(8.0));
    b.rename_sheet(1, "売上").unwrap();
    assert_eq!(read(&mut b, "A1", Read::Formula), text("=売上!A1*2"));
    b.set_values(1, rng("A1"), &[vec![num(5.0)]]).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), num(10.0));

    b.add_name("Rate", None, "=売上!$A$1").unwrap();
    b.set_formulas(0, rng("A2"), &[vec![text("=Rate+1")]], false).unwrap();
    assert_eq!(read(&mut b, "A2", Read::Value), num(6.0));
    assert_eq!(b.resolve(0, "Rate"), Some((1, rng("A1"))));
    assert_eq!(b.resolve(0, "B2:C5"), Some((0, rng("B2:C5"))));
    assert_eq!(b.resolve(0, "'売上'!$C$3"), Some((1, rng("C3"))));
    assert_eq!(b.resolve(1, "A:A"), Some((1, RangeRef::cols(0, 0))));
    assert_eq!(b.resolve(0, "Nothing"), None);

    b.delete_sheet(1).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), Val::Err(2023));
    assert!(b.delete_sheet(0).is_err());
}

#[test]
fn evaluate_works_out_formulas() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A3"), &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]]).unwrap();
    assert_eq!(b.evaluate(0, "SUM(A1:A3)"), Value::Number(6.0));
    assert_eq!(b.evaluate(0, "=MATCH(2,A1:A3,0)"), Value::Number(2.0));
    assert_eq!(b.evaluate(0, "LENB(\"あいう\")"), Value::Number(6.0));
    assert!(matches!(b.evaluate(0, "A1:A3*2"), Value::Array(_)));
}

#[test]
fn saved_workbooks_read_back() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A2"), &[vec![num(2.0)], vec![text("x")]]).unwrap();
    b.set_formulas(0, rng("B1"), &[vec![text("=A1*3")]], false).unwrap();
    b.set_formulas(0, rng("C1"), &[vec![text("=SEQUENCE(3)")]], true).unwrap();
    put(&mut b, "D1", Val::Date(46085.0));
    b.add_name("Two", None, "Sheet1!$A$1").unwrap();
    let bytes = b.to_bytes().unwrap();

    let mut b = Book::open_bytes(&bytes).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), num(2.0));
    assert_eq!(read(&mut b, "A2", Read::Value), text("x"));
    assert_eq!(read(&mut b, "B1", Read::Value), num(6.0));
    assert_eq!(read(&mut b, "B1", Read::Formula), text("=A1*3"));
    assert_eq!(b.get(0, rng("C1:C3"), Read::Value).unwrap(), vec![vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]]);
    assert_eq!(read(&mut b, "C1", Read::Formula2), text("=SEQUENCE(3)"));
    assert_eq!(read(&mut b, "D1", Read::Value), Val::Date(46085.0));
    assert_eq!(b.resolve(0, "Two"), Some((0, rng("A1"))));
    // Still calculated after reading.
    put(&mut b, "A1", num(5.0));
    assert_eq!(read(&mut b, "B1", Read::Value), num(15.0));

    // What was emptied is empty in the file.
    b.clear(0, rng("A2"), false).unwrap();
    let bytes = b.to_bytes().unwrap();
    let mut b = Book::open_bytes(&bytes).unwrap();
    assert_eq!(read(&mut b, "A2", Read::Value), Val::Empty);
}

#[test]
fn formulas_the_engine_does_not_know_keep_their_values() {
    // A workbook saved with a formula calling a VBA function: its saved value stays.
    let mut file = umya_spreadsheet::new_file();
    let ws = file.sheet_mut(0).unwrap();
    ws.cell_mut((1, 1)).set_formula("MyVbaFunction(2)");
    ws.cell_mut((1, 1)).set_formula_result_number(42.0);
    ws.cell_mut((2, 1)).set_formula("A1+1");
    ws.cell_mut((2, 1)).set_formula_result_number(43.0);
    let mut buf = Vec::new();
    umya_spreadsheet::writer::xlsx::write_writer(&file, &mut buf).unwrap();
    let mut b = Book::open_bytes(&buf).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), num(42.0));
    assert_eq!(read(&mut b, "B1", Read::Value), num(43.0));
    b.calculate();
    assert_eq!(read(&mut b, "B1", Read::Value), num(43.0));
}

#[test]
fn r1c1_formulas() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A3"), &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]]).unwrap();
    b.set_formulas_r1c1(0, rng("B1:B3"), &[vec![text("=RC[-1]*2")]]).unwrap();
    assert_eq!(b.get(0, rng("B1:B3"), Read::Value).unwrap(), vec![vec![num(2.0)], vec![num(4.0)], vec![num(6.0)]]);
    assert_eq!(read(&mut b, "B2", Read::Formula), text("=A2*2"));
    assert_eq!(read(&mut b, "B2", Read::FormulaR1C1), text("=RC[-1]*2"));
    b.set_formulas_r1c1(0, rng("C1"), &[vec![text("=SUM(R1C1:R3C1)&\"RC\"")]]).unwrap();
    assert_eq!(read(&mut b, "C1", Read::Formula), text("=SUM($A$1:$A$3)&\"RC\""));
    assert_eq!(read(&mut b, "C1", Read::Value), text("6RC"));
}

#[test]
fn copies_move_relative_references() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A2"), &[vec![num(1.0)], vec![num(2.0)]]).unwrap();
    b.set_formulas(0, rng("B1"), &[vec![text("=A1*10")]], false).unwrap();
    put(&mut b, "C1", Val::Date(46085.0));
    b.copy(0, rng("B1:C1"), 0, at("B2")).unwrap();
    assert_eq!(read(&mut b, "B2", Read::Formula), text("=A2*10"));
    assert_eq!(read(&mut b, "B2", Read::Value), num(20.0));
    assert_eq!(read(&mut b, "C2", Read::Value), Val::Date(46085.0));
}
