# xlsxrb-adapters

Compatibility adapters for migrating from peer XLSX libraries (e.g. `rubyXL`) to [`xlsxrb`](https://github.com/niku/xlsxrb) with zero downtime and drop-in safety.

## Overview

[`rubyXL`](https://github.com/weshatheleopard/rubyXL) is a mature, full-featured library for comprehensive OpenXML spreadsheet creation and manipulation in Ruby, maintained for more than a decade. Its intuitive object model (`sheet.add_cell`, row/cell indexing, and in-memory mutability) is familiar to many Ruby developers.

[`xlsxrb`](https://github.com/niku/xlsxrb) is a pure Ruby, zero-dependency, streaming-capable, low-memory XLSX engine designed for high throughput.

`xlsxrb-adapters` bridges the best of both worlds by providing drop-in compatible adapter layers to:
1. Support gradual migration (Strangler Fig pattern) from peer XLSX libraries to `xlsxrb` without rewriting application logic.
2. Enable high-throughput, low-memory execution in batch processing and resource-constrained environments while retaining `rubyXL`'s familiar and battle-tested API.
3. Construct interoperability test harnesses against peer libraries and real-world fixtures.
4. Keep `xlsxrb` core strictly zero-dependency, mutant-tested, and type-safe while providing rich compatibility layers.

## Design Principles

1. **No Global Hijacking**:
   Does not reopen or hijack top-level constants like `::RubyXL`. Instead, exposes namespaces like `Xlsxrb::Adapters::RubyXL` so you can use both in the same application during migration or run side-by-side comparison tests.
2. **Mutable-to-Immutable Boundary**:
   Maintains a mutable in-memory wrapper structure compatible with legacy workflows, translating into `xlsxrb`'s immutable data models (`Data.define`, frozen) upon save/export.
3. **Native Bridge Conversion**:
   Supports seamless conversions between adapter structures and native `xlsxrb` objects via `adapter_wb.to_xlsxrb` and `Xlsxrb::Adapters::RubyXL.from_xlsxrb(xlsxrb_wb)`.

## Supported Adapters

### RubyXL (`Xlsxrb::Adapters::RubyXL`)

- **Parser**:
  - `Xlsxrb::Adapters::RubyXL::Parser.parse(filepath)`
  - `Xlsxrb::Adapters::RubyXL::Parser.parse_buffer(buffer_or_string)`
- **Workbook**:
  - `wb[idx_or_name]`
  - `wb.worksheets`
  - `wb.add_worksheet(name)`
  - `wb.write(filepath)`
  - `wb.stream` (returns binary `StringIO`)
  - `wb.to_xlsxrb`
- **Worksheet**:
  - `ws[row_idx]`
  - `ws.add_cell(row_idx, col_idx, data, formula, overwrite)`
  - `ws.each` (iterates rows)
  - `ws.sheet_data[row_idx][col_idx]`
- **Row**:
  - `row[col_idx]`
  - `row.cells`
  - `row.each` (iterates cells)
- **Cell**:
  - `cell.value`
  - `cell.raw_value`
  - `cell.formula`
  - `cell.change_contents(data, formula_expression)`
  - `cell.remove_formula`

## Compatibility (100% Test Pass)

Thanks to the comprehensive and well-maintained RSpec suite provided by the `rubyXL` project, `Xlsxrb::Adapters::RubyXL` can systematically verify drop-in behavioral compatibility without monkey-patching:

| Spec File | Total Examples | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- |
| `worksheet_spec.rb` | 233 | 233 | 0 | 100.0% |
| `cell_spec.rb` | 75 | 75 | 0 | 100.0% |
| `workbook_spec.rb` | 26 | 26 | 0 | 100.0% |
| `reference_spec.rb` | 10 | 10 | 0 | 100.0% |
| `parser_spec.rb` | 8 | 8 | 0 | 100.0% |
| `stylesheet_spec.rb` | 4 | 4 | 0 | 100.0% |
| `color_spec.rb` | 3 | 3 | 0 | 100.0% |
| `rgb_color_spec.rb` | 2 | 2 | 0 | 100.0% |
| `text_spec.rb` | 2 | 2 | 0 | 100.0% |
| **Total** | **363** | **363** | **0** | **100.0%** |

Run the official test suite yourself:
```bash
bundle exec rake compatibility
```

## Performance & Memory Footprint

`rubyXL` is purposefully architected as a full Document Object Model (DOM) library, excelling at template preservation and complete XML-tree introspection.

`xlsxrb-adapters` complements this by pairing `rubyXL`'s familiar mutable API with `xlsxrb`'s low-overhead streaming/SST pipeline. For data-intensive pipelines (100,000+ rows) where throughput and RAM usage dominate, `xlsxrb-adapters` achieves **~7–8x faster execution** and **~77–78% lower memory usage**:

### Benchmark: 1,000,000 cells (100,000 rows × 10 cols)

| Operation | Library | Time (Median) | Peak Memory (VmHWM) | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- |
| **Write** | `rubyXL` (3.4.38) | 35.01 s | 2,198.8 MB | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters`** | **4.70 s** | **489.0 MB** | **~7.5x faster** | **77.8% reduced** |
| | `xlsxrb (In-Memory)` | 2.56 s | 240.6 MB | ~13.7x faster | 89.1% reduced |
| | `xlsxrb (Streaming)` | 0.95 s | 74.9 MB | ~36.9x faster | 96.6% reduced |
| **Read** | `rubyXL` (3.4.38) | 35.72 s | 2,469.5 MB | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters`** | **4.50 s** | **556.1 MB** | **~7.9x faster** | **77.5% reduced** |
| | `xlsxrb (In-Memory)` | 2.31 s | 392.4 MB | ~15.5x faster | 84.1% reduced |
| | `xlsxrb (Streaming)` | 1.97 s | 87.6 MB | ~18.1x faster | 96.5% reduced |

Detailed scaling analysis across 10,000, 100,000, and 1,000,000 cells is documented in [docs/BENCHMARK.md](docs/BENCHMARK.md).

## Development & Dev Container

The Dev Container configuration mounts the sibling `../xlsxrb` repository at `/workspaces/xlsxrb` inside the container:

```json
"mounts": [
  "source=${localWorkspaceFolder}/../xlsxrb,target=/workspaces/xlsxrb,type=bind"
]
```

Run test suite, compatibility harness, and benchmarks with:

```bash
bundle install

# Run unit tests
bundle exec rake test

# Run official rubyXL RSpec compatibility suite (363 tests)
bundle exec rake compatibility

# Run benchmarks
bundle exec rake benchmark
```

## Acknowledgements

`xlsxrb-adapters` would not exist without the following projects and the people behind them. We extend our sincere respect and gratitude to:

- Vivek Bhagwat, Wesha ([weshatheleopard](https://github.com/weshatheleopard)), and all contributors to [`rubyXL`](https://github.com/weshatheleopard/rubyXL) for pioneering comprehensive OpenXML DOM manipulation in Ruby and maintaining it for more than a decade.

`xlsxrb-adapters` builds directly on the APIs these projects designed and the test suites they published.

This is an independent project and is not affiliated with or endorsed by the authors or maintainers of the libraries listed above.

## License

MIT
