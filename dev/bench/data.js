window.BENCHMARK_DATA = {
  "lastUpdate": 1772859185257,
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
          "id": "12f5fa88e40d6f3a2ccb617e60b0b2a45e045c1f",
          "message": "docs: update README with session management and batch workflow (EN/ZH-TW)\n\n- Add Session Management and Batch Writes to Features list\n- Replace old Session Mode section with:\n  - Session Management (open/commit/close/status)\n  - Batch Writes (inline -c, script -s, stdin pipe)\n  - Interactive REPL (read-only, renamed from session)\n- Add Session & Batch Commands reference section\n- Remove old batch from Utility Commands\n- Update both English and Traditional Chinese versions",
          "timestamp": "2026-03-07T11:12:04+08:00",
          "tree_id": "4d1f3371f071cc5d35a08b666d5acf0abd724016",
          "url": "https://github.com/yen0304/xlex/commit/12f5fa88e40d6f3a2ccb617e60b0b2a45e045c1f"
        },
        "date": 1772853487509,
        "tool": "cargo",
        "benches": [
          {
            "name": "workbook_open_1k_rows",
            "value": 1397200,
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
            "value": 223,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/100",
            "value": 72983,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/1000",
            "value": 193870,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/10000",
            "value": 1319600,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "range_read_100_cells",
            "value": 1317,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/100",
            "value": 507600,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/1000",
            "value": 2685900,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/5000",
            "value": 12860000,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_ref_parse_simple",
            "value": 26,
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
            "value": 73,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "lazy_stream_rows_10k",
            "value": 8800300,
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
          "id": "f7672aae1728b18a2141b26d05e7cb66102f1e9e",
          "message": "fix: escape angle brackets in doc comments for rustdoc -D warnings",
          "timestamp": "2026-03-07T12:47:25+08:00",
          "tree_id": "2d723ca08ad60566863f105d66d14f95bc9133df",
          "url": "https://github.com/yen0304/xlex/commit/f7672aae1728b18a2141b26d05e7cb66102f1e9e"
        },
        "date": 1772859185019,
        "tool": "cargo",
        "benches": [
          {
            "name": "workbook_open_1k_rows",
            "value": 1388600,
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
            "value": 39,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "cell_set_single",
            "value": 221,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/100",
            "value": 72898,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/1000",
            "value": 194030,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "row_append/10000",
            "value": 1315700,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "range_read_100_cells",
            "value": 1287,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/100",
            "value": 518610,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/1000",
            "value": 2644200,
            "range": "± 0",
            "unit": "ns/iter"
          },
          {
            "name": "workbook_save/5000",
            "value": 12798000,
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
            "value": 8770300,
            "range": "± 0",
            "unit": "ns/iter"
          }
        ]
      }
    ]
  }
}