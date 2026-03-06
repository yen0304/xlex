window.BENCHMARK_DATA = {
  "lastUpdate": 1772808312320,
  "repoUrl": "https://github.com/yen0304/xlex",
  "entries": {
    "Benchmark": [
      {
        "commit": {
          "author": {
            "email": "yen0304@users.noreply.github.com",
            "name": "yen0304",
            "username": "yen0304"
          },
          "committer": {
            "email": "yen0304@users.noreply.github.com",
            "name": "yen0304",
            "username": "yen0304"
          },
          "distinct": true,
          "id": "c12a8f7125b23055261cc1fc788624c8e587035b",
          "message": "ci: re-trigger benchmark after gh-pages creation",
          "timestamp": "2026-03-06T22:38:44+08:00",
          "tree_id": "d810ef4003201a7e9d66ada8162568fcb47f98cf",
          "url": "https://github.com/yen0304/xlex/commit/c12a8f7125b23055261cc1fc788624c8e587035b"
        },
        "date": 1772808311468,
        "tool": "cargo",
        "benches": [
          {
            "name": "workbook_open_1k_rows",
            "value": 1406300,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "sheet_list",
            "value": 11,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_get_single",
            "value": 42,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_set_single",
            "value": 227,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/100",
            "value": 71674,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/1000",
            "value": 193780,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/10000",
            "value": 1307300,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "range_read_100_cells",
            "value": 1309,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/100",
            "value": 499020,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/1000",
            "value": 2639900,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/5000",
            "value": 12916000,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_ref_parse_simple",
            "value": 25,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_ref_parse_complex",
            "value": 38,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "range_parse",
            "value": 74,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "lazy_stream_rows_10k",
            "value": 8873100,
            "range": "± 0",
            "unit": "ns/iter"
          }
        ]
      }
    ]
  }
}