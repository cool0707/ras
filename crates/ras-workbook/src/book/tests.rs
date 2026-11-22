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
    b.set_formulas(0, rng("A3"), &[vec![text("=SUM(A1:A2)*10")]], false)
        .unwrap();
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
    let Val::Date(d) = read(&mut b, "A4", Read::Value) else {
        panic!("not a date")
    };
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
    b.set_values(
        0,
        rng("A1:B2"),
        &[vec![num(1.0), num(2.0)], vec![num(3.0), num(4.0)]],
    )
    .unwrap();
    assert_eq!(
        b.get(0, rng("A1:B2"), Read::Value).unwrap(),
        vec![vec![num(1.0), num(2.0)], vec![num(3.0), num(4.0)]]
    );
    // One value fills every cell; one row every row; beyond the array #N/A.
    b.set_values(0, rng("D1:E2"), &[vec![num(9.0)]]).unwrap();
    assert_eq!(
        b.get(0, rng("D1:E2"), Read::Value).unwrap(),
        vec![vec![num(9.0), num(9.0)], vec![num(9.0), num(9.0)]]
    );
    b.set_values(0, rng("G1:I2"), &[vec![num(1.0), num(2.0)]])
        .unwrap();
    assert_eq!(
        b.get(0, rng("G1:I1"), Read::Value).unwrap(),
        vec![vec![num(1.0), num(2.0), Val::Err(2042)]]
    );
    assert_eq!(
        b.get(0, rng("G2:H2"), Read::Value).unwrap(),
        vec![vec![num(1.0), num(2.0)]]
    );
}

#[test]
fn one_formula_over_a_range_moves_its_references() {
    let mut b = Book::new_book();
    b.set_values(
        0,
        rng("A1:A3"),
        &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]],
    )
    .unwrap();
    b.set_formulas(0, rng("B1:B3"), &[vec![text("=A1*2")]], false)
        .unwrap();
    assert_eq!(
        b.get(0, rng("B1:B3"), Read::Value).unwrap(),
        vec![vec![num(2.0)], vec![num(4.0)], vec![num(6.0)]]
    );
    assert_eq!(read(&mut b, "B3", Read::Formula), text("=A3*2"));
}

#[test]
fn formula2_spills_and_formula_intersects() {
    let mut b = Book::new_book();
    b.set_values(
        0,
        rng("A1:A3"),
        &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]],
    )
    .unwrap();
    b.set_formulas(0, rng("C1"), &[vec![text("=A1:A3*10")]], true)
        .unwrap();
    assert_eq!(
        b.get(0, rng("C1:C3"), Read::Value).unwrap(),
        vec![vec![num(10.0)], vec![num(20.0)], vec![num(30.0)]]
    );
    assert_eq!(read(&mut b, "C1", Read::Formula2), text("=A1:A3*10"));
    // Range.Formula: implicit intersection, so the formula in row 2 takes A2.
    b.set_formulas(0, rng("D2"), &[vec![text("=A1:A3*10")]], false)
        .unwrap();
    assert_eq!(read(&mut b, "D2", Read::Value), num(20.0));
    assert_eq!(read(&mut b, "D3", Read::Value), Val::Empty);
    assert_eq!(read(&mut b, "D2", Read::Formula), text("=A1:A3*10"));
    assert_eq!(read(&mut b, "D2", Read::Formula2), text("=@A1:A3*10"));
}

#[test]
fn errors_are_cverr_numbers() {
    let mut b = Book::new_book();
    b.set_formulas(0, rng("A1"), &[vec![text("=1/0")]], false)
        .unwrap();
    b.set_formulas(0, rng("A2"), &[vec![text("=NA()")]], false)
        .unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), Val::Err(2007));
    assert_eq!(read(&mut b, "A2", Read::Value), Val::Err(2042));
    assert_eq!(read(&mut b, "A1", Read::Text), text("#DIV/0!"));
}

#[test]
fn end_goes_where_ctrl_arrow_goes() {
    let mut b = Book::new_book();
    b.set_values(
        0,
        rng("A1:A5"),
        &[
            vec![num(1.0)],
            vec![num(2.0)],
            vec![num(3.0)],
            vec![Val::Empty],
            vec![num(5.0)],
        ],
    )
    .unwrap();
    assert_eq!(b.end(0, CellRef::new(MAX_ROWS - 1, 0), Dir::Up), at("A5"));
    assert_eq!(b.end(0, at("A1"), Dir::Down), at("A3"));
    assert_eq!(b.end(0, at("A3"), Dir::Down), at("A5"));
    assert_eq!(b.end(0, at("A5"), Dir::Down), CellRef::new(MAX_ROWS - 1, 0));
    assert_eq!(b.end(0, at("A5"), Dir::Up), at("A3"));
    assert_eq!(b.end(0, at("B1"), Dir::Down), CellRef::new(MAX_ROWS - 1, 1));
    assert_eq!(b.end(0, CellRef::new(MAX_ROWS - 1, 1), Dir::Up), at("B1"));
    assert_eq!(
        b.end(0, at("A1"), Dir::Right),
        CellRef::new(0, MAX_COLS - 1)
    );
    assert_eq!(b.end(0, at("E1"), Dir::Left), at("A1"));
}

#[test]
fn regions_and_used_range() {
    let mut b = Book::new_book();
    assert_eq!(b.used_range(0), None);
    b.set_values(
        0,
        rng("B2:C3"),
        &[vec![num(1.0), num(2.0)], vec![num(3.0), num(4.0)]],
    )
    .unwrap();
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
    b.set_formulas(0, rng("A1"), &[vec![text("=Data!A1*2")]], false)
        .unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), num(8.0));
    b.rename_sheet(1, "売上").unwrap();
    assert_eq!(read(&mut b, "A1", Read::Formula), text("=売上!A1*2"));
    b.set_values(1, rng("A1"), &[vec![num(5.0)]]).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), num(10.0));

    b.add_name("Rate", None, "=売上!$A$1").unwrap();
    b.set_formulas(0, rng("A2"), &[vec![text("=Rate+1")]], false)
        .unwrap();
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
    b.set_values(
        0,
        rng("A1:A3"),
        &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]],
    )
    .unwrap();
    assert_eq!(b.evaluate(0, "SUM(A1:A3)"), Value::Number(6.0));
    assert_eq!(b.evaluate(0, "=MATCH(2,A1:A3,0)"), Value::Number(2.0));
    assert_eq!(b.evaluate(0, "LENB(\"あいう\")"), Value::Number(6.0));
    assert!(matches!(b.evaluate(0, "A1:A3*2"), Value::Array(_)));
}

#[test]
fn saved_workbooks_read_back() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A2"), &[vec![num(2.0)], vec![text("x")]])
        .unwrap();
    b.set_formulas(0, rng("B1"), &[vec![text("=A1*3")]], false)
        .unwrap();
    b.set_formulas(0, rng("C1"), &[vec![text("=SEQUENCE(3)")]], true)
        .unwrap();
    put(&mut b, "D1", Val::Date(46085.0));
    b.add_name("Two", None, "Sheet1!$A$1").unwrap();
    let bytes = b.to_bytes().unwrap();

    let mut b = Book::open_bytes(&bytes).unwrap();
    assert_eq!(read(&mut b, "A1", Read::Value), num(2.0));
    assert_eq!(read(&mut b, "A2", Read::Value), text("x"));
    assert_eq!(read(&mut b, "B1", Read::Value), num(6.0));
    assert_eq!(read(&mut b, "B1", Read::Formula), text("=A1*3"));
    assert_eq!(
        b.get(0, rng("C1:C3"), Read::Value).unwrap(),
        vec![vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]]
    );
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
    b.set_values(
        0,
        rng("A1:A3"),
        &[vec![num(1.0)], vec![num(2.0)], vec![num(3.0)]],
    )
    .unwrap();
    b.set_formulas_r1c1(0, rng("B1:B3"), &[vec![text("=RC[-1]*2")]])
        .unwrap();
    assert_eq!(
        b.get(0, rng("B1:B3"), Read::Value).unwrap(),
        vec![vec![num(2.0)], vec![num(4.0)], vec![num(6.0)]]
    );
    assert_eq!(read(&mut b, "B2", Read::Formula), text("=A2*2"));
    assert_eq!(read(&mut b, "B2", Read::FormulaR1C1), text("=RC[-1]*2"));
    b.set_formulas_r1c1(0, rng("C1"), &[vec![text("=SUM(R1C1:R3C1)&\"RC\"")]])
        .unwrap();
    assert_eq!(
        read(&mut b, "C1", Read::Formula),
        text("=SUM($A$1:$A$3)&\"RC\"")
    );
    assert_eq!(read(&mut b, "C1", Read::Value), text("6RC"));
}

#[test]
fn copies_move_relative_references() {
    let mut b = Book::new_book();
    b.set_values(0, rng("A1:A2"), &[vec![num(1.0)], vec![num(2.0)]])
        .unwrap();
    b.set_formulas(0, rng("B1"), &[vec![text("=A1*10")]], false)
        .unwrap();
    put(&mut b, "C1", Val::Date(46085.0));
    b.copy(0, rng("B1:C1"), 0, at("B2")).unwrap();
    assert_eq!(read(&mut b, "B2", Read::Formula), text("=A2*10"));
    assert_eq!(read(&mut b, "B2", Read::Value), num(20.0));
    assert_eq!(read(&mut b, "C2", Read::Value), Val::Date(46085.0));
}

// ---------------------------------------------------------------------------- cells inserted and deleted

fn formula(b: &mut Book, a1: &str, f: &str) {
    b.set_formulas(0, rng(a1), &[vec![text(f)]], false).unwrap();
}

/// A1:C4 of numbers (row * 10 + column), and formulas that read them.
fn grid() -> Book {
    let mut b = Book::new_book();
    for row in 1..=4u32 {
        for col in 1..=3u32 {
            let a1 = CellRef::new(row - 1, col - 1).a1();
            put(&mut b, &a1, num(f64::from(row * 10 + col)));
        }
    }
    formula(&mut b, "E1", "=SUM(A1:A4)");
    formula(&mut b, "E2", "=B3*2");
    formula(&mut b, "E3", "=$C$4");
    b
}

#[test]
fn rows_inserted_and_deleted_move_cells_and_references() {
    let mut b = grid();
    b.set_format(0, rng("A2:C2"), "font.bold", &Val::Bool(true))
        .unwrap();
    b.insert_cells(0, RangeRef::rows(2, 3), false, false)
        .unwrap();
    assert_eq!(read(&mut b, "A3", Read::Value), Val::Empty);
    assert_eq!(read(&mut b, "A5", Read::Value), num(31.0));
    assert_eq!(read(&mut b, "E1", Read::Formula), text("=SUM(A1:A6)"));
    assert_eq!(
        read(&mut b, "E1", Read::Value),
        num(11.0 + 21.0 + 31.0 + 41.0)
    );
    // The formulas of the rows that moved moved too.
    assert_eq!(read(&mut b, "E2", Read::Formula), text("=B5*2"));
    assert_eq!(read(&mut b, "E5", Read::Formula), text("=$C$6"));
    // The new rows take the format of the row above.
    assert_eq!(
        b.format_of(0, rng("A3:C4"), "font.bold").unwrap(),
        Some(Val::Bool(true))
    );
    assert_eq!(
        b.format_of(0, rng("A5"), "font.bold").unwrap(),
        Some(Val::Bool(false))
    );

    b.delete_cells(0, RangeRef::rows(0, 1), false).unwrap();
    assert_eq!(read(&mut b, "A3", Read::Value), num(31.0));
    assert_eq!(read(&mut b, "E3", Read::Formula), text("=$C$4"));
    // The cells a formula read are gone.
    b.delete_cells(0, RangeRef::rows(3, 3), false).unwrap();
    assert_eq!(read(&mut b, "E3", Read::Formula), text("=#REF!"));
    assert_eq!(read(&mut b, "E3", Read::Value), Val::Err(2023));
}

#[test]
fn cells_shift_within_their_columns_or_rows() {
    let mut b = grid();
    // B2 inserted, column B moving down from there.
    b.insert_cells(0, rng("B2"), false, false).unwrap();
    assert_eq!(
        b.get(0, rng("A2:C3"), Read::Value).unwrap(),
        vec![
            vec![num(21.0), Val::Empty, num(23.0)],
            vec![num(31.0), num(22.0), num(33.0)]
        ]
    );
    assert_eq!(read(&mut b, "B5", Read::Value), num(42.0));
    assert_eq!(read(&mut b, "E2", Read::Formula), text("=B4*2"));
    assert_eq!(read(&mut b, "E1", Read::Formula), text("=SUM(A1:A4)"));
    // A1 deleted, row 1 moving left: the formula in E1 comes to D1.
    b.delete_cells(0, rng("A1"), true).unwrap();
    assert_eq!(
        b.get(0, rng("A1:C1"), Read::Value).unwrap(),
        vec![vec![num(12.0), num(13.0), Val::Empty]]
    );
    assert_eq!(read(&mut b, "D1", Read::Formula), text("=SUM(A1:A4)"));
    assert_eq!(read(&mut b, "A2", Read::Value), num(21.0));
    // Columns.
    b.insert_cells(0, RangeRef::cols(0, 0), true, false)
        .unwrap();
    assert_eq!(read(&mut b, "B2", Read::Value), num(21.0));
    assert_eq!(read(&mut b, "F2", Read::Formula), text("=C4*2"));
}

#[test]
fn what_excel_refuses_to_move() {
    let mut b = grid();
    put(&mut b, &CellRef::new(MAX_ROWS - 1, 0).a1(), num(1.0));
    assert!(
        b.insert_cells(0, RangeRef::rows(0, 0), false, false)
            .unwrap_err()
            .contains("押し出す")
    );
    assert!(b.insert_cells(0, rng("B1"), false, false).is_ok());
    b.merge_cells(0, rng("B3:C3"), false).unwrap();
    assert!(
        b.insert_cells(0, rng("B1"), false, false)
            .unwrap_err()
            .contains("結合")
    );
    // Whole rows take merged cells along.
    b.delete_cells(0, RangeRef::rows(0, 0), false).unwrap();
    assert_eq!(b.merge_area(0, at("C2")), rng("B2:C2"));
}

// ---------------------------------------------------------------------------- copying and moving

#[test]
fn pasting_values_formats_and_operations() {
    let mut b = grid();
    b.set_number_format(0, rng("E1"), "#,##0").unwrap();
    b.set_format(0, rng("E1"), "interior.color", &num(255.0))
        .unwrap();
    let values = Paste {
        what: PasteWhat::Values,
        ..Paste::all()
    };
    b.paste(0, rng("E1:E2"), 0, at("G1"), &values).unwrap();
    assert_eq!(read(&mut b, "G1", Read::Formula), text("104"));
    assert_eq!(read(&mut b, "G1", Read::NumberFormat), text("General"));
    b.paste(
        0,
        rng("E1"),
        0,
        at("G1"),
        &Paste {
            what: PasteWhat::Formats,
            ..Paste::all()
        },
    )
    .unwrap();
    assert_eq!(read(&mut b, "G1", Read::Text), text("104"));
    assert_eq!(read(&mut b, "G1", Read::NumberFormat), text("#,##0"));
    assert_eq!(
        b.format_of(0, rng("G1"), "interior.color").unwrap(),
        Some(num(255.0))
    );
    // All: the formula with its references moved, and the format.
    b.copy(0, rng("E1"), 0, at("F2")).unwrap();
    assert_eq!(read(&mut b, "F2", Read::Formula), text("=SUM(B2:B5)"));
    assert_eq!(
        b.format_of(0, rng("F2"), "interior.color").unwrap(),
        Some(num(255.0))
    );
    // Operations and transposing.
    b.paste(
        0,
        rng("A1:A2"),
        0,
        at("B1"),
        &Paste {
            what: PasteWhat::Values,
            op: PasteOp::Add,
            ..Paste::all()
        },
    )
    .unwrap();
    assert_eq!(
        b.get(0, rng("B1:B2"), Read::Value).unwrap(),
        vec![vec![num(23.0)], vec![num(43.0)]]
    );
    b.paste(
        0,
        rng("A1:A3"),
        0,
        at("A7"),
        &Paste {
            what: PasteWhat::Values,
            transpose: true,
            ..Paste::all()
        },
    )
    .unwrap();
    assert_eq!(
        b.get(0, rng("A7:C7"), Read::Value).unwrap(),
        vec![vec![num(11.0), num(21.0), num(31.0)]]
    );
    // Empty cells are skipped when asked.
    put(&mut b, "H1", num(1.0));
    put(&mut b, "H2", num(2.0));
    b.paste(
        0,
        rng("I1:I2"),
        0,
        at("H1"),
        &Paste {
            skip_blanks: true,
            ..Paste::all()
        },
    )
    .unwrap();
    assert_eq!(read(&mut b, "H2", Read::Value), num(2.0));
    b.copy(0, rng("I1:I2"), 0, at("H1")).unwrap();
    assert_eq!(read(&mut b, "H2", Read::Value), Val::Empty);
}

#[test]
fn cut_cells_take_their_references_along() {
    let mut b = grid();
    formula(&mut b, "A5", "=A4+B1");
    b.move_cells(0, rng("A4:A5"), 0, at("G1")).unwrap();
    assert_eq!(read(&mut b, "A4", Read::Value), Val::Empty);
    assert_eq!(read(&mut b, "G1", Read::Value), num(41.0));
    // The formula that moved keeps what it read; the reference to the cell that moved follows.
    assert_eq!(read(&mut b, "G2", Read::Formula), text("=G1+B1"));
    assert_eq!(read(&mut b, "E1", Read::Formula), text("=SUM(A1:A4)"));
    // To another sheet.
    b.add_sheet(1, Some("別")).unwrap();
    b.move_cells(0, rng("G1:G2"), 1, at("A1")).unwrap();
    assert_eq!(
        b.get(1, rng("A2"), Read::Formula).unwrap(),
        vec![vec![text("=A1+Sheet1!B1")]]
    );
    assert_eq!(
        b.get(1, rng("A2"), Read::Value).unwrap(),
        vec![vec![num(53.0)]]
    );
    formula(&mut b, "F1", "=C1");
    b.move_cells(0, rng("C1"), 1, at("C3")).unwrap();
    assert_eq!(read(&mut b, "F1", Read::Formula), text("=別!C3"));
}

// ---------------------------------------------------------------------------- sorting and filtering

fn people() -> Book {
    let mut b = Book::new_book();
    let rows = [
        ("名前", "部署", "年齢"),
        ("佐藤", "営業", "41"),
        ("鈴木", "開発", "29"),
        ("高橋", "営業", "35"),
        ("田中", "総務", ""),
        ("伊藤", "開発", "52"),
    ];
    for (i, (a, c, d)) in rows.iter().enumerate() {
        b.set_values(
            0,
            RangeRef::new(CellRef::new(i as u32, 0), CellRef::new(i as u32, 2)),
            &[vec![text(a), text(c), text(d)]],
        )
        .unwrap();
    }
    b
}

fn column(b: &mut Book, a1: &str) -> Vec<Val> {
    b.get(0, rng(a1), Read::Value)
        .unwrap()
        .into_iter()
        .map(|mut r| r.remove(0))
        .collect()
}

#[test]
fn sorting_orders_rows_as_excel_does() {
    let mut b = people();
    formula(&mut b, "D2", "=C2+1");
    b.set_format(0, rng("A3"), "font.italic", &Val::Bool(true))
        .unwrap();
    assert!(b.guess_header(0, rng("A1:C6")));
    b.sort(
        0,
        rng("A2:D6"),
        &[SortKey {
            index: 2,
            descending: true,
        }],
        false,
        false,
    )
    .unwrap();
    // Empty cells last, also when descending.
    assert_eq!(
        column(&mut b, "C2:C6"),
        vec![num(52.0), num(41.0), num(35.0), num(29.0), Val::Empty]
    );
    assert_eq!(
        column(&mut b, "A2:A6"),
        vec![
            text("伊藤"),
            text("佐藤"),
            text("高橋"),
            text("鈴木"),
            text("田中")
        ]
    );
    // The formula and the format came along.
    assert_eq!(read(&mut b, "D3", Read::Formula), text("=C3+1"));
    assert_eq!(read(&mut b, "D3", Read::Value), num(42.0));
    assert_eq!(
        b.format_of(0, rng("A5"), "font.italic").unwrap(),
        Some(Val::Bool(true))
    );
    assert_eq!(
        b.format_of(0, rng("A3"), "font.italic").unwrap(),
        Some(Val::Bool(false))
    );
    // Two keys.
    b.sort(
        0,
        rng("A2:D6"),
        &[
            SortKey {
                index: 1,
                descending: false,
            },
            SortKey {
                index: 2,
                descending: false,
            },
        ],
        false,
        false,
    )
    .unwrap();
    assert_eq!(
        column(&mut b, "A2:A6"),
        vec![
            text("高橋"),
            text("佐藤"),
            text("田中"),
            text("鈴木"),
            text("伊藤")
        ]
    );
}

#[test]
fn filters_hide_the_rows_that_do_not_pass() {
    let mut b = people();
    formula(&mut b, "E1", "=SUBTOTAL(109,C2:C6)");
    b.set_auto_filter(0, Some(rng("A1:C6"))).unwrap();
    b.set_filter(
        0,
        1,
        Some(FilterCriterion::Values {
            values: vec!["営業".into(), "開発".into()],
            blanks: false,
        }),
    )
    .unwrap();
    assert!(b.filter_mode(0));
    assert_eq!(b.hidden(0, false, 4, 4), Some(true));
    assert_eq!(
        b.visible_areas(0, rng("A1:C6")),
        vec![rng("A1:C4"), rng("A6:C6")]
    );
    assert_eq!(
        read(&mut b, "E1", Read::Value),
        num(41.0 + 29.0 + 35.0 + 52.0)
    );
    b.set_filter(
        0,
        2,
        Some(FilterCriterion::Custom {
            a: (">=".into(), "35".into()),
            b: Some(("<".into(), "50".into())),
            and: true,
        }),
    )
    .unwrap();
    assert_eq!(b.visible_areas(0, rng("A2:A6")), vec![rng("A2"), rng("A4")]);
    assert_eq!(read(&mut b, "E1", Read::Value), num(76.0));
    // Saved with the filter, and read back with it.
    let bytes = b.to_bytes().unwrap();
    let mut saved = Book::open_bytes(&bytes).unwrap();
    assert_eq!(saved.auto_filter(0), Some((rng("A1:C6"), vec![1, 2])));
    assert_eq!(saved.hidden(0, false, 2, 2), Some(true));
    saved.show_all(0).unwrap();
    assert_eq!(saved.hidden(0, false, 0, 5), Some(false));
    // One field cleared, then the filter removed.
    b.set_filter(0, 1, None).unwrap();
    assert_eq!(b.visible_areas(0, rng("A2:A6")), vec![rng("A2"), rng("A4")]);
    b.set_filter(
        0,
        2,
        Some(FilterCriterion::Custom {
            a: ("=".into(), "*藤".into()),
            b: None,
            and: true,
        }),
    )
    .unwrap();
    assert_eq!(b.hidden(0, false, 1, 5), Some(true));
    b.set_filter(0, 2, None).unwrap();
    b.set_filter(
        0,
        0,
        Some(FilterCriterion::Custom {
            a: ("=".into(), "*藤".into()),
            b: None,
            and: true,
        }),
    )
    .unwrap();
    assert_eq!(b.visible_areas(0, rng("A2:A6")), vec![rng("A2"), rng("A6")]);
    b.set_auto_filter(0, None).unwrap();
    assert_eq!(b.hidden(0, false, 0, 5), Some(false));
    assert!(!b.filter_mode(0));
}

#[test]
fn special_cells_duplicates_and_advanced_filters() {
    let mut b = people();
    formula(&mut b, "D2", "=C2*2");
    assert_eq!(
        b.special_cells(0, rng("A1:D6"), Special::Formulas(23))
            .unwrap(),
        vec![rng("D2")]
    );
    assert_eq!(
        b.special_cells(0, rng("A1:D6"), Special::Constants(1))
            .unwrap(),
        vec![rng("C2:C4"), rng("C6")]
    );
    assert_eq!(
        b.special_cells(0, rng("A1:D6"), Special::Blanks).unwrap(),
        vec![rng("D1"), rng("D3:D4"), rng("C5:D5"), rng("D6")]
    );
    assert_eq!(
        b.special_cells(0, rng("A1"), Special::LastCell).unwrap(),
        vec![rng("D6")]
    );

    // The list filtered by a range of criteria, copied elsewhere.
    b.set_values(
        0,
        rng("G1:H2"),
        &[
            vec![text("部署"), text("年齢")],
            vec![text("営業"), text(">40")],
        ],
    )
    .unwrap();
    put(&mut b, "G3", text("開"));
    b.advanced_filter(
        0,
        rng("A1:C6"),
        Some((0, rng("G1:H3"))),
        Some((0, rng("J1"))),
        false,
    )
    .unwrap();
    assert_eq!(
        b.get(0, rng("J1:L4"), Read::Value).unwrap(),
        vec![
            vec![text("名前"), text("部署"), text("年齢")],
            vec![text("佐藤"), text("営業"), num(41.0)],
            vec![text("鈴木"), text("開発"), num(29.0)],
            vec![text("伊藤"), text("開発"), num(52.0)],
        ]
    );
    // In place, and unique values of one column.
    b.advanced_filter(0, rng("B1:B6"), None, None, true)
        .unwrap();
    assert!(b.filter_mode(0));
    assert_eq!(
        b.visible_areas(0, rng("B1:B6")),
        vec![rng("B1:B3"), rng("B5")]
    );
    b.show_all(0).unwrap();
    assert!(!b.filter_mode(0));
    assert_eq!(b.hidden(0, false, 0, 5), Some(false));

    assert_eq!(b.remove_duplicates(0, rng("A2:C6"), &[1]).unwrap(), 2);
    assert_eq!(
        column(&mut b, "A2:A6"),
        vec![
            text("佐藤"),
            text("鈴木"),
            text("田中"),
            Val::Empty,
            Val::Empty
        ]
    );
}

// ---------------------------------------------------------------------------- formats, manual calculation

#[test]
fn formats_are_kept_and_saved() {
    let mut b = grid();
    b.set_format(0, rng("A1:B2"), "font.bold", &Val::Bool(true))
        .unwrap();
    b.set_format(0, rng("A1"), "font.color", &num(255.0))
        .unwrap();
    b.set_format(0, rng("A1"), "font.name", &text("メイリオ"))
        .unwrap();
    b.set_format(0, rng("A1"), "interior.colorindex", &num(6.0))
        .unwrap();
    b.set_format(0, rng("A1:B2"), "borders.9.linestyle", &num(1.0))
        .unwrap();
    b.set_format(0, rng("A1:B2"), "borders.9.weight", &num(-4138.0))
        .unwrap();
    b.set_format(0, rng("C1"), "horizontalalignment", &num(-4108.0))
        .unwrap();
    b.set_format(0, RangeRef::cols(2, 2), "columnwidth", &num(20.0))
        .unwrap();
    b.set_format(0, RangeRef::rows(1, 1), "rowheight", &num(30.0))
        .unwrap();
    b.set_format(0, RangeRef::cols(5, 5), "font.italic", &Val::Bool(true))
        .unwrap();
    put(&mut b, "F3", num(1.0));
    b.merge_cells(0, rng("A6:C6"), false).unwrap();
    let bytes = b.to_bytes().unwrap();
    let mut b = Book::open_bytes(&bytes).unwrap();
    assert_eq!(
        b.format_of(0, rng("A1:B2"), "font.bold").unwrap(),
        Some(Val::Bool(true))
    );
    assert_eq!(b.format_of(0, rng("A1:C3"), "font.bold").unwrap(), None);
    assert_eq!(
        b.format_of(0, rng("A1"), "font.color").unwrap(),
        Some(num(255.0))
    );
    assert_eq!(
        b.format_of(0, rng("A1"), "font.name").unwrap(),
        Some(text("メイリオ"))
    );
    assert_eq!(
        b.format_of(0, rng("A1"), "interior.color").unwrap(),
        Some(num(65535.0))
    );
    assert_eq!(
        b.format_of(0, rng("A1"), "interior.colorindex").unwrap(),
        Some(num(6.0))
    );
    assert_eq!(
        b.format_of(0, rng("A2:B2"), "borders.9.weight").unwrap(),
        Some(num(-4138.0))
    );
    assert_eq!(
        b.format_of(0, rng("A1:B1"), "borders.-4107.linestyle")
            .unwrap(),
        Some(num(-4142.0))
    );
    assert_eq!(
        b.format_of(0, rng("C1"), "horizontalalignment").unwrap(),
        Some(num(-4108.0))
    );
    assert_eq!(
        b.format_of(0, rng("C1"), "columnwidth").unwrap(),
        Some(num(20.0))
    );
    assert_eq!(
        b.format_of(0, rng("A2"), "rowheight").unwrap(),
        Some(num(30.0))
    );
    // A cell written into a formatted column has the column's format.
    assert_eq!(
        b.format_of(0, rng("F3"), "font.italic").unwrap(),
        Some(Val::Bool(true))
    );
    assert_eq!(
        b.format_of(0, rng("F9"), "font.italic").unwrap(),
        Some(Val::Bool(true))
    );
    assert_eq!(
        b.format_of(0, rng("A6:C6"), "mergecells").unwrap(),
        Some(Val::Bool(true))
    );
    b.clear_formats(0, rng("A1:C6")).unwrap();
    assert_eq!(
        b.format_of(0, rng("A1:C6"), "font.bold").unwrap(),
        Some(Val::Bool(false))
    );
    assert_eq!(
        b.format_of(0, rng("A6:C6"), "mergecells").unwrap(),
        Some(Val::Bool(false))
    );
}

#[test]
fn manual_calculation_waits_for_calculate() {
    let mut b = grid();
    assert_eq!(read(&mut b, "E2", Read::Value), num(64.0));
    b.set_manual(true);
    put(&mut b, "B3", num(1.0));
    assert_eq!(read(&mut b, "E2", Read::Value), num(64.0));
    // A formula is worked out when it is entered; what reads it waits.
    formula(&mut b, "F1", "=E2+B3");
    formula(&mut b, "F2", "=F1*10");
    assert_eq!(read(&mut b, "F1", Read::Value), num(65.0));
    b.recalculate();
    assert_eq!(read(&mut b, "E2", Read::Value), num(2.0));
    assert_eq!(read(&mut b, "F2", Read::Value), num(30.0));
    b.set_manual(false);
    put(&mut b, "B3", num(5.0));
    assert_eq!(b.calculated_sheets(), vec![0]);
    assert_eq!(read(&mut b, "F2", Read::Value), num(150.0));
    put(&mut b, "H9", num(5.0));
    assert_eq!(b.calculated_sheets(), Vec::<usize>::new());
}

#[test]
fn sheets_are_copied() {
    let mut b = grid();
    b.set_format(0, rng("A1"), "font.bold", &Val::Bool(true))
        .unwrap();
    let i = b.copy_sheet(0, 1).unwrap();
    assert_eq!((i, b.sheet_name(1)), (1, Some("Sheet1 (2)")));
    assert_eq!(
        b.get(1, rng("E1"), Read::Formula).unwrap(),
        vec![vec![text("=SUM(A1:A4)")]]
    );
    b.set_values(1, rng("A1"), &[vec![num(0.0)]]).unwrap();
    assert_eq!(
        b.get(1, rng("E1"), Read::Value).unwrap(),
        vec![vec![num(93.0)]]
    );
    assert_eq!(read(&mut b, "E1", Read::Value), num(104.0));
    assert_eq!(
        b.format_of(1, rng("A1"), "font.bold").unwrap(),
        Some(Val::Bool(true))
    );
    assert_ne!(b.code_name(0), b.code_name(1));
}
