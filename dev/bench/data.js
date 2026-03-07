window.BENCHMARK_DATA = {
  "lastUpdate": 1772853110228,
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
      },
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
          "id": "8779cc4f3551d8cea7bf0d20cdab820bccc9da54",
          "message": "ci: remove debug step from benchmark job",
          "timestamp": "2026-03-06T22:48:25+08:00",
          "tree_id": "0df0acc33ce382789550895f0a911a497b2b6d7a",
          "url": "https://github.com/yen0304/xlex/commit/8779cc4f3551d8cea7bf0d20cdab820bccc9da54"
        },
        "date": 1772808844118,
        "tool": "cargo",
        "benches": [
          {
            "name": "workbook_open_1k_rows",
            "value": 1391700,
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
            "value": 40,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_set_single",
            "value": 214,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/100",
            "value": 72618,
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
            "value": 1305900,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "range_read_100_cells",
            "value": 1293,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/100",
            "value": 485880,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/1000",
            "value": 2650100,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/5000",
            "value": 12815000,
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
            "value": 8855700,
            "range": "± 0",
            "unit": "ns/iter"
          }
        ]
      },
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
          "id": "1d24023277e91490eaec2f47b5f7c18d41d41fba",
          "message": "feat: add session management (open/commit/close) and in-process batch execution\n\n- xlex open <file>: creates .xlex/ working copy for safe editing\n- xlex batch: executes multiple write commands in single open/save cycle\n  - Supports pipe, inline (-c), and script file (-s) modes\n  - In-process execution (no subprocess spawning)\n  - Supported: cell set/clear/formula, row append/insert/delete, sheet add/remove/rename\n- xlex commit: saves working copy back to original file\n- xlex close: discards changes and removes session\n- xlex status: shows active session info\n- Renamed 'session' to 'repl' (interactive read-only REPL)\n- Added .xlex/ to .gitignore\n- Updated SKILL.md, commands.md, examples.md, npm README",
          "timestamp": "2026-03-07T11:05:57+08:00",
          "tree_id": "691e49517129f4fe851e4ed684bd1687a6237b5e",
          "url": "https://github.com/yen0304/xlex/commit/1d24023277e91490eaec2f47b5f7c18d41d41fba"
        },
        "date": 1772853109948,
        "tool": "cargo",
        "benches": [
          {
            "name": "workbook_open_1k_rows",
            "value": 1393100,
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
            "value": 40,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_set_single",
            "value": 232,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/100",
            "value": 72872,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/1000",
            "value": 196220,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/10000",
            "value": 1316400,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "range_read_100_cells",
            "value": 1304,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/100",
            "value": 497580,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/1000",
            "value": 2638800,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/5000",
            "value": 12849000,
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
            "value": 39,
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
            "value": 8960100,
            "range": "± 0",
            "unit": "ns/iter"
          }
        ]
      }
    ]
  }
}