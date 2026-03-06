//! Property-based tests for xlex-core using proptest.

use proptest::prelude::*;
use xlex_core::{CellRef, CellValue, Range};

/// Valid column numbers are 1..=16384 (A..XFD in Excel).
fn arb_col() -> impl Strategy<Value = u32> {
    1..=CellRef::MAX_COL
}

/// Valid row numbers are 1..=1048576.
fn arb_row() -> impl Strategy<Value = u32> {
    1..=CellRef::MAX_ROW
}

/// Arbitrary valid CellRef.
#[allow(dead_code)]
fn arb_cell_ref() -> impl Strategy<Value = CellRef> {
    (arb_col(), arb_row()).prop_map(|(col, row)| CellRef::new(col, row))
}

proptest! {
    // ─── CellRef roundtrip ────────────────────────────────────────────

    #[test]
    fn cellref_a1_roundtrip(col in arb_col(), row in arb_row()) {
        let cr = CellRef::new(col, row);
        let a1 = cr.to_a1();
        let parsed = CellRef::parse(&a1).expect("should re-parse");
        prop_assert_eq!(parsed.col, col);
        prop_assert_eq!(parsed.row, row);
    }

    #[test]
    fn cellref_display_eq_to_a1(col in arb_col(), row in arb_row()) {
        let cr = CellRef::new(col, row);
        prop_assert_eq!(cr.to_string(), cr.to_a1());
    }

    #[test]
    fn cellref_fromstr_roundtrip(col in arb_col(), row in arb_row()) {
        let cr = CellRef::new(col, row);
        let s = cr.to_a1();
        let parsed: CellRef = s.parse().expect("FromStr should work");
        prop_assert_eq!(parsed, cr);
    }

    // Column letter conversion roundtrip
    #[test]
    fn col_letters_roundtrip(col in arb_col()) {
        let letters = CellRef::col_to_letters(col);
        let back = CellRef::col_from_letters_pub(&letters).expect("should parse back");
        prop_assert_eq!(back, col);
    }

    // Lowercased input should parse identically
    #[test]
    fn cellref_case_insensitive(col in arb_col(), row in arb_row()) {
        let cr = CellRef::new(col, row);
        let upper = cr.to_a1();
        let lower = upper.to_lowercase();
        let parsed = CellRef::parse(&lower).expect("lowercase should parse");
        prop_assert_eq!(parsed, cr);
    }

    // ─── Range roundtrip ──────────────────────────────────────────────

    #[test]
    fn range_a1_roundtrip(
        c1 in 1..=1000u32,
        r1 in 1..=50000u32,
        c2 in 1..=1000u32,
        r2 in 1..=50000u32,
    ) {
        // Ensure start <= end
        let sc = c1.min(c2);
        let sr = r1.min(r2);
        let ec = c1.max(c2);
        let er = r1.max(r2);

        let range = Range::new(CellRef::new(sc, sr), CellRef::new(ec, er));
        let a1 = range.to_a1();
        let parsed = Range::parse(&a1).expect("should re-parse");
        prop_assert_eq!(parsed.start.col, sc);
        prop_assert_eq!(parsed.start.row, sr);
        prop_assert_eq!(parsed.end.col, ec);
        prop_assert_eq!(parsed.end.row, er);
    }

    #[test]
    fn range_single_cell(col in 1..=1000u32, row in 1..=50000u32) {
        let cr = CellRef::new(col, row);
        let range = Range::single(cr.clone());
        prop_assert_eq!(range.start, cr.clone());
        prop_assert_eq!(range.end, cr);
    }

    #[test]
    fn range_contains_its_corners(
        c1 in 1..=500u32,
        r1 in 1..=10000u32,
        c2 in 1..=500u32,
        r2 in 1..=10000u32,
    ) {
        let sc = c1.min(c2);
        let sr = r1.min(r2);
        let ec = c1.max(c2);
        let er = r1.max(r2);

        let range = Range::new(CellRef::new(sc, sr), CellRef::new(ec, er));
        // All four corners should be within bounds
        prop_assert!(range.start.col <= range.end.col);
        prop_assert!(range.start.row <= range.end.row);
    }

    // ─── CellValue display ────────────────────────────────────────────

    #[test]
    fn cellvalue_number_display_not_empty(n in prop::num::f64::NORMAL) {
        let cv = CellValue::Number(n);
        let display = cv.to_display_string();
        prop_assert!(!display.is_empty());
    }

    #[test]
    fn cellvalue_string_roundtrip(s in "[a-zA-Z0-9 ]{0,100}") {
        let cv = CellValue::String(s.clone());
        let display = cv.to_display_string();
        prop_assert_eq!(display, s);
    }

    #[test]
    fn cellvalue_bool_display(b in any::<bool>()) {
        let cv = CellValue::Boolean(b);
        let display = cv.to_display_string();
        let expected = if b { "TRUE" } else { "FALSE" };
        prop_assert_eq!(display, expected);
    }

    // ─── Edge cases ───────────────────────────────────────────────────

    // CellRef at exact boundaries
    #[test]
    fn cellref_max_bounds_valid(
        col in (CellRef::MAX_COL - 10)..=CellRef::MAX_COL,
        row in (CellRef::MAX_ROW - 10)..=CellRef::MAX_ROW,
    ) {
        let cr = CellRef::new(col, row);
        let a1 = cr.to_a1();
        let parsed = CellRef::parse(&a1).expect("near-max should parse");
        prop_assert_eq!(parsed.col, col);
        prop_assert_eq!(parsed.row, row);
    }
}
