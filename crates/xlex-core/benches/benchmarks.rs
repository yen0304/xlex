//! Benchmarks for XLEX core operations.
//!
//! Run with: `cargo bench`

use criterion::{black_box, criterion_group, criterion_main, BenchmarkId, Criterion};
use std::path::PathBuf;
use tempfile::TempDir;
use xlex_core::{CellRef, CellValue, LazyWorkbook, Range, Style, Workbook};

/// Create a temporary test workbook with specified number of rows.
fn create_test_workbook(rows: usize) -> (TempDir, PathBuf) {
    let temp_dir = TempDir::new().unwrap();
    let path = temp_dir.path().join("test.xlsx");

    let mut workbook = Workbook::new();
    if let Some(sheet) = workbook.get_sheet_mut("Sheet1") {
        for i in 0..rows {
            let cell_ref = CellRef::new(0, i as u32); // Column A
            sheet.set_cell(cell_ref, CellValue::String(format!("Value {}", i)));
        }
    }
    workbook.save_as(&path).unwrap();

    (temp_dir, path)
}

fn bench_workbook_open(c: &mut Criterion) {
    let (_temp_dir, path) = create_test_workbook(1000);

    c.bench_function("workbook_open_1k_rows", |b| {
        b.iter(|| {
            let wb = Workbook::open(black_box(&path)).unwrap();
            black_box(wb)
        })
    });
}

fn bench_sheet_list(c: &mut Criterion) {
    let (_temp_dir, path) = create_test_workbook(1000);

    c.bench_function("sheet_list", |b| {
        let wb = Workbook::open(&path).unwrap();
        b.iter(|| {
            let names = wb.sheet_names();
            black_box(names)
        })
    });
}

fn bench_cell_get(c: &mut Criterion) {
    let (_temp_dir, path) = create_test_workbook(10000);
    let wb = Workbook::open(&path).unwrap();

    c.bench_function("cell_get_single", |b| {
        b.iter(|| {
            let sheet = wb.get_sheet("Sheet1").unwrap();
            let cell = sheet.get_cell(&CellRef::new(0, 500));
            black_box(cell)
        })
    });
}

fn bench_cell_set(c: &mut Criterion) {
    let (_temp_dir, path) = create_test_workbook(100);

    c.bench_function("cell_set_single", |b| {
        b.iter_batched(
            || Workbook::open(&path).unwrap(),
            |mut wb| {
                wb.set_cell(
                    "Sheet1",
                    CellRef::new(0, 50),
                    CellValue::String("Updated".to_string()),
                )
                .unwrap();
                black_box(wb)
            },
            criterion::BatchSize::SmallInput,
        )
    });
}

fn bench_row_append(c: &mut Criterion) {
    let mut group = c.benchmark_group("row_append");

    for size in [100, 1000, 10000].iter() {
        group.bench_with_input(BenchmarkId::from_parameter(size), size, |b, &size| {
            let temp_dir = TempDir::new().unwrap();
            let path = temp_dir.path().join("test.xlsx");

            let workbook = Workbook::new();
            workbook.save_as(&path).unwrap();

            b.iter(|| {
                let mut wb = Workbook::open(&path).unwrap();
                if let Some(sheet) = wb.get_sheet_mut("Sheet1") {
                    for i in 0..size {
                        sheet.set_cell(
                            CellRef::new(0, i as u32),
                            CellValue::String(format!("Row {}", i)),
                        );
                    }
                }
                black_box(&wb);
            })
        });
    }
    group.finish();
}

fn bench_range_read(c: &mut Criterion) {
    let (_temp_dir, path) = create_test_workbook(10000);
    let wb = Workbook::open(&path).unwrap();

    c.bench_function("range_read_100_cells", |b| {
        b.iter(|| {
            let sheet = wb.get_sheet("Sheet1").unwrap();
            let _range = Range::parse("A1:A100").unwrap();
            // Read individual cells in range
            let mut cells = Vec::with_capacity(100);
            for row in 0..100u32 {
                let cell = sheet.get_cell(&CellRef::new(0, row));
                cells.push(cell);
            }
            black_box(cells)
        })
    });
}

fn bench_workbook_save(c: &mut Criterion) {
    let mut group = c.benchmark_group("workbook_save");

    for rows in [100, 1000, 5000].iter() {
        group.bench_with_input(BenchmarkId::from_parameter(rows), rows, |b, &rows| {
            let temp_dir = TempDir::new().unwrap();
            let path = temp_dir.path().join("test.xlsx");

            let mut workbook = Workbook::new();
            if let Some(sheet) = workbook.get_sheet_mut("Sheet1") {
                for i in 0..rows {
                    sheet.set_cell(
                        CellRef::new(0, i as u32),
                        CellValue::String(format!("Value {}", i)),
                    );
                }
            }
            workbook.save_as(&path).unwrap();

            b.iter(|| {
                let output_path = temp_dir.path().join(format!("output_{}.xlsx", rows));
                workbook.save_as(&output_path).unwrap();
            })
        });
    }
    group.finish();
}

fn bench_cell_reference_parse(c: &mut Criterion) {
    c.bench_function("cell_ref_parse_simple", |b| {
        b.iter(|| {
            let cell_ref = CellRef::parse(black_box("A1")).unwrap();
            black_box(cell_ref)
        })
    });

    c.bench_function("cell_ref_parse_complex", |b| {
        b.iter(|| {
            let cell_ref = CellRef::parse(black_box("XFD1048576")).unwrap();
            black_box(cell_ref)
        })
    });
}

fn bench_range_parse(c: &mut Criterion) {
    c.bench_function("range_parse", |b| {
        b.iter(|| {
            let range = Range::parse(black_box("A1:Z1000")).unwrap();
            black_box(range)
        })
    });
}

fn bench_shared_string_deduplication(c: &mut Criterion) {
    let mut group = c.benchmark_group("shared_string_dedup");

    group.bench_function("1000_duplicates", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            for _ in 0..1000 {
                wb.add_shared_string("Duplicate String");
            }
            black_box(wb.shared_strings().len())
        })
    });

    group.bench_function("1000_unique", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            for i in 0..1000 {
                wb.add_shared_string(format!("Unique String {}", i));
            }
            black_box(wb.shared_strings().len())
        })
    });

    group.bench_function("1000_mixed", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            for i in 0..1000 {
                wb.add_shared_string(format!("String {}", i % 500));
            }
            black_box(wb.shared_strings().len())
        })
    });

    group.finish();
}

fn bench_lazy_workbook_open(c: &mut Criterion) {
    let (_temp_dir, path) = create_test_workbook(10000);

    let mut group = c.benchmark_group("workbook_open_comparison");

    group.bench_function("regular_10k_rows", |b| {
        b.iter(|| {
            let wb = Workbook::open(black_box(&path)).unwrap();
            black_box(wb)
        })
    });

    group.bench_function("lazy_10k_rows", |b| {
        b.iter(|| {
            let wb = LazyWorkbook::open(black_box(&path)).unwrap();
            black_box(wb)
        })
    });

    group.finish();
}

fn bench_lazy_workbook_stream_rows(c: &mut Criterion) {
    let (_temp_dir, path) = create_test_workbook(10000);

    c.bench_function("lazy_stream_rows_10k", |b| {
        let wb = LazyWorkbook::open(&path).unwrap();
        b.iter(|| {
            let mut count = 0;
            for row in wb.stream_rows("Sheet1").unwrap() {
                count += row.cells.len();
            }
            black_box(count)
        })
    });
}

fn bench_style_registry(c: &mut Criterion) {
    let mut group = c.benchmark_group("style_registry");

    group.bench_function("add_100_styles", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            for i in 0..100 {
                let mut style = Style::default();
                style.font.bold = i % 2 == 0;
                style.font.size = Some(10.0 + (i % 10) as f64);
                wb.style_registry_mut().add(style);
            }
            black_box(wb.style_registry().len())
        })
    });

    group.bench_function("apply_style_to_range", |b| {
        let temp_dir = TempDir::new().unwrap();
        let path = temp_dir.path().join("test.xlsx");

        let mut wb = Workbook::new();
        if let Some(sheet) = wb.get_sheet_mut("Sheet1") {
            for row in 0..100u32 {
                for col in 0..10u32 {
                    sheet.set_cell(CellRef::new(col, row), CellValue::String(format!("Cell")));
                }
            }
        }
        let mut style = Style::default();
        style.font.bold = true;
        let style_id = wb.style_registry_mut().add(style);
        wb.save_as(&path).unwrap();

        b.iter(|| {
            let mut wb = Workbook::open(&path).unwrap();
            if let Some(sheet) = wb.get_sheet_mut("Sheet1") {
                for row in 0..100u32 {
                    for col in 0..10u32 {
                        sheet.set_cell_style(&CellRef::new(col, row), Some(style_id));
                    }
                }
            }
            black_box(&wb);
        })
    });

    group.finish();
}

fn bench_sheet_operations(c: &mut Criterion) {
    let mut group = c.benchmark_group("sheet_operations");

    group.bench_function("add_10_sheets", |b| {
        b.iter(|| {
            let mut wb = Workbook::new();
            for i in 0..10 {
                wb.add_sheet(&format!("Sheet{}", i + 2)).unwrap();
            }
            black_box(wb.sheet_count())
        })
    });

    group.bench_function("rename_sheet", |b| {
        b.iter_batched(
            || {
                let mut wb = Workbook::new();
                wb.add_sheet("OldName").unwrap();
                wb
            },
            |mut wb| {
                wb.rename_sheet("OldName", "NewName").unwrap();
                black_box(wb)
            },
            criterion::BatchSize::SmallInput,
        )
    });

    group.finish();
}

criterion_group!(
    benches,
    bench_workbook_open,
    bench_sheet_list,
    bench_cell_get,
    bench_cell_set,
    bench_row_append,
    bench_range_read,
    bench_workbook_save,
    bench_cell_reference_parse,
    bench_range_parse,
    bench_shared_string_deduplication,
    bench_lazy_workbook_open,
    bench_lazy_workbook_stream_rows,
    bench_style_registry,
    bench_sheet_operations,
);

criterion_main!(benches);
