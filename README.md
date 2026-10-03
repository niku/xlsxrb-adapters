# xlsxrb-adapters

Compatibility adapters for migrating from peer XLSX libraries ([`rubyXL`](https://github.com/weshatheleopard/rubyXL), [`caxlsx`](https://github.com/caxlsx/caxlsx), [`roo`](https://github.com/roo-rb/roo)) to [`xlsxrb`](https://github.com/niku/xlsxrb) with zero downtime and drop-in safety.

## Overview

In the Ruby ecosystem, three libraries have long served as primary standards for spreadsheet processing:
- [`rubyXL`](https://github.com/weshatheleopard/rubyXL): A mature, full-featured Document Object Model (DOM) library widely trusted for comprehensive OpenXML spreadsheet creation, cell manipulation, and template editing.
- [`caxlsx`](https://github.com/caxlsx/caxlsx) (formerly `axlsx`): The de facto builder library for generating styled spreadsheets, rich typography, DrawingML charts, and financial reports.
- [`roo`](https://github.com/roo-rb/roo): The most widely used read-only spreadsheet reader library for consuming data, inspection, formulas, and formatters across various spreadsheet formats.

[`xlsxrb`](https://github.com/niku/xlsxrb) is a pure Ruby, zero-dependency, streaming-capable, low-memory XLSX engine designed for high-throughput batch workloads.

`xlsxrb-adapters` bridges the best of both worlds by providing drop-in compatible adapter layers to:
1. Support gradual migration (Strangler Fig pattern) from peer XLSX libraries to `xlsxrb` without rewriting application logic.
2. Enable high-throughput, low-memory execution in batch processing and resource-constrained environments while retaining `rubyXL`, `caxlsx`, and `roo` familiar and battle-tested APIs.
3. Construct interoperability test harnesses against peer libraries and real-world fixtures.
4. Keep `xlsxrb` core strictly zero-dependency, mutant-tested, and type-safe while providing rich compatibility layers.

## Design Principles

1. **No Global Hijacking**:
   Does not reopen or hijack top-level constants like `::RubyXL`, `::Axlsx`, or `::Roo`. Instead, exposes namespaces like `Xlsxrb::Adapters::RubyXL`, `Xlsxrb::Adapters::Caxlsx`, and `Xlsxrb::Adapters::Roo` so you can run side-by-side during migration or run comparison tests. Optional drop-in aliases (e.g. `Roo = Xlsxrb::Adapters::Roo`) are provided for seamless code transitions.
2. **Mutable-to-Immutable Boundary**:
   Maintains a mutable in-memory wrapper structure compatible with legacy workflows, translating into `xlsxrb`'s immutable data models (`Data.define`, frozen) upon save/export.
3. **Native Bridge Conversion**:
   Supports seamless bidirectional conversions between adapter structures and native `xlsxrb` objects via `adapter_wb.to_xlsxrb` and `AdapterClass.from_xlsxrb(xlsxrb_wb)`.

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

### Caxlsx (`Xlsxrb::Adapters::Caxlsx` / `Axlsx`)

Drop-in replacement for the popular [`caxlsx`](https://github.com/caxlsx/caxlsx) / `axlsx` gem without polluting the global `::Axlsx` namespace. Includes an internal alias `Xlsxrb::Adapters::Caxlsx::Axlsx = Xlsxrb::Adapters::Caxlsx` and opt-in top-level aliasing.

#### Drop-In Migration Example

```ruby
require "xlsxrb/adapters/caxlsx"

# Option A: Explicit namespace (recommended to avoid global pollution)
package = Xlsxrb::Adapters::Caxlsx::Package.new
wb = package.workbook

# Option B: Drop-in alias (existing Axlsx code works unmodified)
Axlsx = Xlsxrb::Adapters::Caxlsx::Axlsx
package = Axlsx::Package.new
wb = package.workbook

# Standard caxlsx DSL
header = wb.styles.add_style(b: true, sz: 12, bg_color: "4F81BD", fg_color: "FFFFFF")
wb.add_worksheet(name: "Monthly Report") do |sheet|
  sheet.add_row(["Product", "Qty", "Price", "Total"], style: header)
  sheet.add_row(["Widget", 10, 25.5, "=B2*C2"], types: [:string, :integer, :float, :formula])
  sheet.auto_filter.range = "A1:D2"
end

package.serialize("report.xlsx")
```

#### API Capabilities
- **Package**:
  - `p = Xlsxrb::Adapters::Caxlsx::Package.new`
  - `p.serialize("output.xlsx")`
  - `p.to_stream` (returns binary `StringIO`)
  - `p.use_autowidth = true`
  - `p.to_xlsxrb` (bridges mutable builder to native `xlsxrb` elements)
- **Workbook**:
  - `wb = p.workbook`
  - `wb.add_worksheet(name: "Report") { |ws| ... }`
  - `wb.date1904 = true` (Excel 1904 calendar date calculation)
  - `wb.add_defined_name("A1:B10", name: "Sales")`
- **Styles Engine**:
  - `wb.styles.add_style(b: true, i: true, sz: 14, font_name: "Arial", bg_color: "4F81BD", fg_color: "FFFFFF", alignment: { horizontal: :center, vertical: :center, wrap_text: true }, border: { style: :thin, color: "000000" }, num_fmt: 4)`
- **Worksheet & Cells**:
  - `ws.add_row(["Item", "Qty", "Price"], style: header_style, height: 24)`
  - `ws.add_row(["Widget", 10, 100], types: [:string, :integer, :integer])`
  - `ws.merge_cells("A1:C1")`
  - `ws.auto_filter.range = "A1:C2"`
  - `ws.sheet_view.pane { |pane| pane.state = :frozen; pane.y_split = 1 }`
  - `ws.sheet_protection.password = "secret"`
  - `ws.page_setup.orientation = :landscape`
  - `ws.page_margins { |m| m.left = 0.5; m.right = 0.5 }`
- **Advanced Features**:
  - **Charts**: `ws.add_chart(Xlsxrb::Adapters::Caxlsx::Bar3DChart, title: "Sales")` with series, labels, data, colors, and 3D rotation
  - **Tables**: `ws.add_table("A1:C10", name: "SalesTable")`
  - **Comments**: `ws.add_comment(ref: "A1", text: "Verified", author: "Auditor")` via legacy VML drawing
  - **Hyperlinks**: `ws.add_hyperlink(location: "https://example.com", ref: "A1")`
  - **Rich Text**: `cell.rich_text.add_run("Bold", b: true); cell.rich_text.add_run(" Normal")`
  - **Data Validation & Conditional Formatting**: list dropdowns, custom formulas, and color scales
- **Streaming Pipeline (`StreamingPackage`)**:
  For massive spreadsheets (100k+ to 1M+ rows), use `StreamingPackage` to write directly to disk in constant $O(1)$ memory with full Caxlsx DSL compatibility:
  ```ruby
  Xlsxrb::Adapters::Caxlsx::StreamingPackage.open("large_report.xlsx") do |pkg|
    header_style = pkg.workbook.styles.add_style(b: true, sz: 12)
    pkg.workbook.add_worksheet(name: "Records") do |sheet|
      sheet.add_row(["ID", "Name", "Score"], style: header_style)
      1_000_000.times { |i| sheet.add_row([i, "User #{i}", 95.5]) }
    end
  end
  ```

### Roo (`Xlsxrb::Adapters::Roo` / `Roo`)

Drop-in replacement for reading and inspecting XLSX / XLSM spreadsheets using the [`roo`](https://github.com/roo-rb/roo) API, fully compatible with the official [Excel (xlsx and xlsm) support](https://github.com/roo-rb/roo#excel-xlsx-and-xlsm-support) specification.

#### Drop-In Migration Example

```ruby
require "xlsxrb/adapters/roo"

# Option A: Explicit namespace (recommended to avoid global pollution)
xlsx = Xlsxrb::Adapters::Roo::Excelx.new("data.xlsx")

# Option B: Drop-in alias (existing Roo code works unmodified)
Roo = Xlsxrb::Adapters::Roo
xlsx = Roo::Spreadsheet.open("data.xlsx")

# Standard Roo inspection & navigation
puts xlsx.info
puts xlsx.sheets

xlsx.default_sheet = "Sheet1"
puts xlsx.first_row
puts xlsx.last_row
puts xlsx.first_column_as_letter
puts xlsx.last_column_as_letter

# Cell access and metadata
val  = xlsx.cell(1, 1)            # Typed Ruby object (String, Numeric, Date, etc.)
fmt  = xlsx.formatted_value(1, 1) # Excel formatted string representation (e.g. "$1,234.50")
type = xlsx.celltype(1, 1)        # :string, :float, :date, :datetime, :time, :boolean, :formula, :link
fml  = xlsx.formula(1, 1)         # Formula expression (e.g. "SUM(A1:A10)")
cmt  = xlsx.comment(1, 1)         # Cell comment or nil

# Iteration & querying
xlsx.each(header_search: ["ID", "Name"]) do |row_hash|
  puts "#{row_hash['ID']}: #{row_hash['Name']}"
end

# Named cells / ranges
xlsx.labels.each do |name, coord|
  puts "Named range #{name} is at #{coord.inspect}"
end

# Formatters / Export
csv_string = xlsx.to_csv
yaml_string = xlsx.to_yaml
matrix = xlsx.to_matrix
```

#### API Capabilities
- **Document Loading & Opening**:
  - `Xlsxrb::Adapters::Roo::Excelx.new(filepath_or_io, options)`
  - `Xlsxrb::Adapters::Roo::Spreadsheet.open(filepath_or_io, options)` (supports block yielding)
  - Options: `only_visible`, `cell_max`, `packed`, `file_warning`
- **Sheets & Navigation**:
  - `xlsx.sheets` (sheet names array)
  - `xlsx.default_sheet = name_or_index` (supports 1-based, 0-based index or sheet name)
  - `xlsx.sheet_for(name_or_index)`
  - `xlsx.sheet_hidden?(name_or_index)`
  - `xlsx.sheet_visible?(name_or_index)`
  - `xlsx.first_row`, `xlsx.last_row`, `xlsx.first_column`, `xlsx.last_column`
  - `xlsx.first_column_as_letter`, `xlsx.last_column_as_letter`
  - `xlsx.info` (detailed summary string of document and sheets)
- **Cell Reading & Metadata**:
  - `xlsx.cell(row, col, sheet = nil)`
  - `xlsx.celltype(row, col, sheet = nil)`
  - `xlsx.cell_post(row, col, sheet = nil)` (alias for cell)
  - `xlsx.excelx_value(row, col, sheet = nil)`
  - `xlsx.excelx_type(row, col, sheet = nil)`
  - `xlsx.excelx_format(row, col, sheet = nil)`
  - `xlsx.formatted_value(row, col, sheet = nil)`
  - `xlsx.formula(row, col, sheet = nil)`
  - `xlsx.formula?(row, col, sheet = nil)`
  - `xlsx.comment(row, col, sheet = nil)`
  - `xlsx.comments(sheet = nil)`
  - `xlsx.empty?(row, col, sheet = nil)`
  - `xlsx.row(row_number, sheet = nil)`
  - `xlsx.column(col_number, sheet = nil)`
- **Advanced Navigation & Querying**:
  - `xlsx.each(options) { |row| ... }` (supports `:header_search`, `:clean`, condition filtering)
  - `xlsx.row_with(query, return_headers = false)`
  - `xlsx.sheet(index, name = false)`
- **Named Cells (Defined Names)**:
  - `xlsx.labels` (array of `[name, [sheet, row, col]]`)
  - `xlsx.label(name)` (returns `[sheet, row, col]`)
- **Formatters & Export**:
  - `xlsx.to_csv(filename = nil, separator = ",", sheet = nil)`
  - `xlsx.to_matrix(sheet = nil)`
  - `xlsx.to_xml(sheet = nil)`
  - `xlsx.to_yaml(options = {}, sheet = nil)`

## Compatibility (100% Test Pass)

`xlsxrb-adapters` verifies 100% behavioral compatibility and drop-in safety against both `rubyXL` and `caxlsx` through dedicated compatibility suites, official upstream test suites, and cross-validation fixtures:

### Caxlsx Compatibility (`Xlsxrb::Adapters::Caxlsx`)

The Caxlsx adapter provides comprehensive coverage of the Caxlsx builder API with 100% pass rate.

> [!NOTE]
> Official `caxlsx` is exclusively a spreadsheet generation library (write-only) and has no reading or parsing APIs. Correspondingly, `xlsxrb-adapters` strictly focuses on 100% drop-in generation compatibility without adding unneeded reading extensions. To read XLSX spreadsheets, use `xlsxrb` directly (`Xlsxrb.read` / `Xlsxrb.stream`) or the RubyXL adapter (`Xlsxrb::Adapters::RubyXL::Parser.parse`).

#### Feature Compatibility Matrix

| Feature Domain | Caxlsx API / Construct | Compatibility Status | Notes |
| :--- | :--- | :--- | :--- |
| **Object Model** | `Package`, `Workbook`, `Worksheet`, `Row`, `Cell` | **100% Supported** | Drop-in compatible DOM; supports block and imperative building |
| **Types & Values** | String, Integer, Float, Date, Time, DateTime, Boolean, Formula | **100% Supported** | Type inference and explicit `:types` array; formula caching |
| **Styles Engine** | `styles.add_style` (fonts, fills, borders, cellXfs, alignments, num_fmts) | **100% Supported** | Native translation to `xlsxrb` style registry and OpenXML `styles.xml` |
| **Rich Text** | `RichText`, `RichTextRun` (multi-run formatting per cell) | **100% Supported** | Full SST `<si><r>` and inlineStr `<is><r>` styling support |
| **Date 1904** | `workbook.date1904 = true` | **100% Supported** | Native 1,462-day serial arithmetic offset handling |
| **Charts** | `BarChart`, `Bar3DChart`, `LineChart`, `PieChart` | **100% Supported** | Series customization, 3D rotations, markers, and DrawingML |
| **Comments / Notes** | `worksheet.add_comment` (ref, text, author) | **100% Supported** | Windows Excel compliant legacy VML 2-cell anchor positioning |
| **Tables & Filters** | `worksheet.add_table`, `worksheet.auto_filter` | **100% Supported** | OpenXML `<tableParts>` and `<autoFilter>` integration |
| **Data Validation** | `worksheet.add_data_validation` (list, range, prompt) | **100% Supported** | Dropdown lists and cell constraints |
| **Conditional Formatting** | `worksheet.add_conditional_formatting` | **100% Supported** | Expression rules, color scales, cell highlights |
| **Panes & Margins** | Freeze panes (`sheet_view.pane`), page setup, page margins | **100% Supported** | Standard OpenXML sheetViews and print configurations |
| **Memory Streaming** | `StreamingPackage.open` | **100% Supported** | Constant $O(1)$ memory streaming generation using `Xlsxrb::StreamWriter` |

#### Test Suite Breakdown

| Test Suite | Scope | Total Tests | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `caxlsx_test.rb` | Package, Workbook, Worksheet, Row, Cell APIs, formula handling, and types | 10 | 10 | 0 | 100.0% |
| `caxlsx_styles_test.rb` | Styles, fonts, fills, borders, cellXfs, alignments, num_fmts, and colors | 11 | 11 | 0 | 100.0% |
| `caxlsx_features_test.rb` | Merge cells, auto-filter, page setup/margins, freeze panes, comments, charts, rich text, Date1904 | 16 | 16 | 0 | 100.0% |
| `caxlsx_compatibility_test.rb` | Real-world fixture generation & cross-validation with official caxlsx | 5 | 5 | 0 | 100.0% |
| `caxlsx_dropin_test.rb` | End-to-end drop-in replacement workflow and streaming validation | 2 | 2 | 0 | 100.0% |
| `caxlsx_streaming_test.rb` | Constant $O(1)$ memory streaming generation via `StreamingPackage` | 5 | 5 | 0 | 100.0% |
| **Total** | **Caxlsx Adapter Compatibility Verification (329 assertions)** | **49** | **49** | **0** | **100.0%** |

Run the Caxlsx compatibility test suite:
```bash
bundle exec rake compatibility:caxlsx
# or
bundle exec rake test:caxlsx
```

### RubyXL Compatibility (`Xlsxrb::Adapters::RubyXL`)

Verified directly against the official `rubyXL` (3.4.38) RSpec test suite without monkey-patching:

#### Feature Compatibility Matrix

| Feature Domain | RubyXL API / Construct | Compatibility Status | Notes |
| :--- | :--- | :--- | :--- |
| **Object Model** | `Workbook`, `Worksheet`, `Row`, `Cell` | **100% Supported** | Mutable DOM with direct 2D indexing (`wb[0]`, `ws[r][c]`) |
| **Parsing** | `RubyXL::Parser.parse`, `parse_buffer` | **100% Supported** | Fast XLSX parsing from files, buffers, and binary `StringIO` |
| **Cell Operations** | `cell.value`, `raw_value`, `change_contents`, `formula` | **100% Supported** | Cell reading, type inference, formula mutation, and removal |
| **Worksheet Manipulation** | `add_cell`, `merge_cells`, `sheet_data`, row iteration | **100% Supported** | Dynamic row/column expansion and sparse cell access |
| **Styles Engine** | Fonts, fills, borders, alignments, num_fmts | **100% Supported** | Reads and maps OpenXML formatting on cells and rows |
| **Data Validation** | `ws.data_validations`, `ws.add_data_validation` | **100% Supported** | First-class validation rules, dropdowns, and prompt configurations |
| **Conditional Formatting** | `ws.conditional_formatting`, `ws.add_conditional_formatting` | **100% Supported** | Cell highlight rules, color scales, and expression evaluation |
| **Defined Names** | `wb.defined_names` | **100% Supported** | Named ranges and workbook-level references |
| **Serialization** | `wb.write(filepath)`, `wb.stream` | **100% Supported** | High-throughput export via low-overhead `xlsxrb` writer |
| **Native Bridge** | `wb.to_xlsxrb`, `RubyXL.from_xlsxrb` | **100% Supported** | Seamless bidirectional translation to native `xlsxrb` models |

#### Test Suite Breakdown

| Test Suite | Scope | Total Tests | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `worksheet_spec.rb` | Worksheet manipulation, cell addition, and row operations | 233 | 233 | 0 | 100.0% |
| `cell_spec.rb` | Cell values, types, formatting, and formulas | 75 | 75 | 0 | 100.0% |
| `workbook_spec.rb` | Workbook structure, sheet management, and serialization | 26 | 26 | 0 | 100.0% |
| `reference_spec.rb` | A1/R1C1 reference conversions and bounding box calculations | 10 | 10 | 0 | 100.0% |
| `parser_spec.rb` | Parsing XLSX files, buffers, and binary streams | 8 | 8 | 0 | 100.0% |
| `stylesheet_spec.rb` | Stylesheet definitions and formatting rules | 4 | 4 | 0 | 100.0% |
| `color_spec.rb` | Color translation and validation | 3 | 3 | 0 | 100.0% |
| `rgb_color_spec.rb` | RGB color conversions and hex definitions | 2 | 2 | 0 | 100.0% |
| `text_spec.rb` | Text escaping and rich text run elements | 2 | 2 | 0 | 100.0% |
| **Total** | **Official rubyXL RSpec Suite Verification** | **363** | **363** | **0** | **100.0%** |

Run the official rubyXL compatibility suite:
```bash
bundle exec rake compatibility:ruby_xl
```

### Roo Compatibility (`Xlsxrb::Adapters::Roo`)

Verified against official Roo (3.0.0) fixtures and Excelx behavior across the full feature scope documented in [Excel (xlsx and xlsm) support](https://github.com/roo-rb/roo#excel-xlsx-and-xlsm-support):

#### Feature Compatibility Matrix

| Feature Domain | Roo API / Construct | Compatibility Status | Notes |
| :--- | :--- | :--- | :--- |
| **Document Loading** | `Roo::Excelx.new`, `Roo::Spreadsheet.open` | **100% Supported** | Supports file paths, Pathnames, StringIO, open IO streams, and URLs |
| **Sheets & Bounds** | `sheets`, `default_sheet`, `first_row`, `last_row`, `first_column`, `last_column` | **100% Supported** | Sheet selection by name, 1-based index, or 0-based index; hidden sheet detection |
| **Cell Inspection** | `cell`, `celltype`, `excelx_type`, `excelx_value`, `formatted_value` | **100% Supported** | Native typed conversion (Integer, Float, Date, DateTime, Time, Boolean, String) |
| **Number & Date Formats**| Excel number & date formats, custom formatting patterns | **100% Supported** | Preserves formatted string representations (currencies, percentages, decimals, dates) |
| **Formulas** | `formula(row, col)`, `formula?(row, col)` | **100% Supported** | Returns formula expression string (without `=`); handles self-closing `<f ... />` |
| **Comments** | `comment(row, col)`, `comments(sheet)` | **100% Supported** | Reads and resolves cell comments from `xl/comments*.xml` |
| **Hyperlinks** | `cell.link?`, `cell.url`, `cell.hyperlink` | **100% Supported** | Returns `Roo::Link` with URL and text target |
| **Defined Names** | `labels`, `label(name)` | **100% Supported** | Parses `<definedNames>` in `xl/workbook.xml` into `[sheet, row, col]` coordinates |
| **Querying & Iteration** | `each(options)`, `row_with`, `header_search` | **100% Supported** | Header-based column mapping, condition matching, and clean string options |
| **Formatters / Export** | `to_csv`, `to_matrix`, `to_xml`, `to_yaml` | **100% Supported** | Exports sheet contents into CSV, Matrix, XML, and YAML strings or files |

#### Test Suite Breakdown

| Test Suite | Scope | Total Tests | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `roo_test.rb` | Excelx core methods, Spreadsheet.open, cell access, formats, formulas, info, each | 20 | 20 | 0 | 100.0% |
| `roo_compatibility_test.rb` | Cross-validation against official `::Roo::Excelx` across 12 diverse real-world fixtures | 12 | 12 | 0 | 100.0% |
| `roo_dropin_test.rb` | End-to-end drop-in replacement workflow and pipeline testing | 2 | 2 | 0 | 100.0% |
| **Total** | **Roo Adapter Compatibility Verification (1,747 assertions)** | **34** | **34** | **0** | **100.0%** |

> [!TIP]
> In addition to the test suite above, all 54 official `.xlsx` fixtures from upstream Roo's test repository were verified side-by-side against official `::Roo::Excelx`, achieving **0 mismatches (100% parity)** across all sheet names, cell values, and cell types.

Run the Roo compatibility test suite:
```bash
bundle exec rake compatibility:roo
# or
bundle exec rake test:roo
```

Run all compatibility suites (RubyXL, Caxlsx, and Roo):
```bash
bundle exec rake compatibility
# or
bin/compatibility_runner all
```

## Performance & Memory Footprint

`xlsxrb-adapters` organizes performance benchmarks around your specific migration path:

### 1. Caxlsx Migration: 1,000,000 cells (100,000 rows × 10 cols)

Official `caxlsx` is a write-only library designed for building and exporting spreadsheets. `xlsxrb-adapters (Caxlsx)` provides 100% drop-in compatibility for writing existing caxlsx workflows, and offers `StreamingPackage` for high-throughput, low-memory streaming generation:

| Operation | Library / Mode | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. caxlsx Speed | vs. caxlsx Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Write** | `caxlsx` (4.5.0) | 5.732 s | 186.9 MB | 20.0 | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters (Caxlsx)`** | **5.710 s** | **385.1 MB** | **19.0** | **~1.0x (matches)** | +106.0% (builder DOM) |
| | **`xlsxrb-adapters (Streaming)`** | **0.998 s** | **78.2 MB** | **18.0** | **~5.7x faster** | **58.2% reduced** |
| | `xlsxrb (In-Memory)` | 2.723 s | 240.7 MB | 12.0 | ~2.1x faster | +28.8% |
| | `xlsxrb (Streaming)` | 0.971 s | 75.3 MB | 27.0 | ~5.9x faster | **59.7% reduced** |

### 2. RubyXL Migration: 1,000,000 cells (100,000 rows × 10 cols)

`rubyXL` is architected as an in-memory Document Object Model (DOM) library. For data-intensive pipelines (100,000+ rows) where throughput and RAM usage dominate, `xlsxrb-adapters (RubyXL)` achieves **~9–11x faster execution** and **~77–78% lower memory usage**:

| Operation | Library / Mode | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. rubyXL Speed | vs. rubyXL Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Write** | `rubyXL` (3.4.38) | 48.733 s | 2,198.5 MB | 94.0 | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters (RubyXL)`** | **5.178 s** | **488.9 MB** | **16.0** | **~9.4x faster** | **77.8% reduced** |
| | `xlsxrb (In-Memory)` | 2.723 s | 240.7 MB | 12.0 | ~17.9x faster | 89.1% reduced |
| | `xlsxrb (Streaming)` | 0.971 s | 75.3 MB | 27.0 | ~50.2x faster | 96.6% reduced |
| **Read** | `rubyXL` (3.4.38) | 57.216 s | 2,462.4 MB | 140.0 | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters (RubyXL)`** | **5.075 s** | **556.5 MB** | **35.0** | **~11.3x faster** | **77.4% reduced** |
| | `xlsxrb (In-Memory)` | 2.779 s | 394.9 MB | 28.0 | ~20.6x faster | 84.0% reduced |
| | `xlsxrb (Streaming)` | 1.811 s | 87.7 MB | 51.0 | ~31.6x faster | 96.4% reduced |

### 3. Roo Migration: 1,000,000 cells (100,000 rows × 10 cols)

Official `roo` is a widely adopted spreadsheet extraction and parsing library. `xlsxrb-adapters (Roo)` provides 100% drop-in API compatibility for existing Roo codebases (supporting `Roo::Excelx.new`, `Roo::Spreadsheet.open`, cell lookups, formats, and `each_row_streaming`) while executing **~1.1x–1.6x faster**. For memory-critical pipelines, native `xlsxrb` streaming reader achieves **~4.5x faster throughput** with **~38% lower memory usage**:

| Operation | Library / Mode | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. roo Speed | vs. roo Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Read** | `roo` (3.0.0) | 8.387 s | 129.3 MB | 74.0 | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters (Roo)`** | **7.569 s** | **475.9 MB** | **77.0** | **~1.1x faster** | +268.1% (in-memory DOM) |
| | `xlsxrb (In-Memory)` | 3.363 s | 448.2 MB | 19.0 | ~2.5x faster | +246.6% |
| | `xlsxrb (Streaming)` | 1.879 s | 80.0 MB | 72.0 | ~4.5x faster | **38.1% reduced** |

Detailed scaling analysis across 10,000, 100,000, and 1,000,000 cells is documented in [docs/BENCHMARK.md](docs/BENCHMARK.md).

## Development & Dev Container

The Dev Container configuration mounts the sibling `../xlsxrb` repository at `/workspaces/xlsxrb` inside the container:

```json
"mounts": [
  "source=${localWorkspaceFolder}/../xlsxrb,target=/workspaces/xlsxrb,type=bind"
]
```

Run test suites, compatibility harnesses, and benchmarks with:

```bash
bundle install

# Run all unit and adapter tests (139 tests, 2,497 assertions)
bundle exec rake test

# Run Caxlsx compatibility test suite (49 tests, 329 assertions)
bundle exec rake test:caxlsx

# Run RubyXL adapter test suite (56 tests, 421 assertions)
bundle exec rake test:ruby_xl

# Run Roo adapter test suite (34 tests, 1,747 assertions)
bundle exec rake test:roo

# Run official rubyXL RSpec compatibility suite (363 tests)
bundle exec rake compatibility:ruby_xl

# Run all compatibility suites (RubyXL, Caxlsx, and Roo)
bundle exec rake compatibility

# Run static type checking with Steep
bundle exec rake typecheck

# Run linter
bundle exec rubocop

# Run benchmarks (usage: rake benchmark [rows=10000] [cols=10] [runs=3] [category=all|caxlsx|rubyxl|roo])
bundle exec rake benchmark
```

## Acknowledgements

We extend our sincere respect and gratitude to:
- Vivek Bhagwat, Wesha, and all contributors to [`rubyXL`](https://github.com/weshatheleopard/rubyXL) for pioneering comprehensive OpenXML DOM manipulation in Ruby over a decade.
- Randy Morgan, Jurriaan Pruis, and the [`caxlsx`](https://github.com/caxlsx/caxlsx) community for establishing the gold standard builder API for styled spreadsheets, charts, and reporting in Ruby.
- Thomas Preymesser, Hugh McGowan, and the [`roo`](https://github.com/roo-rb/roo) community for creating and maintaining the premier spreadsheet extraction library in the Ruby ecosystem.

`xlsxrb-adapters` builds directly upon the ergonomic foundations and test specifications they pioneered.

## License

MIT
