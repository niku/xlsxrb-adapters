# xlsxrb-adapters

Compatibility adapters for migrating from peer XLSX libraries ([`rubyXL`](https://github.com/weshatheleopard/rubyXL), [`caxlsx`](https://github.com/caxlsx/caxlsx), [`roo`](https://github.com/roo-rb/roo), [`xsv`](https://github.com/martijn/xsv), [`creek`](https://github.com/pythonicrubyist/creek), [`fast_excel`](https://github.com/Paxa/fast_excel), [`write_xlsx`](https://github.com/cxn03651/write_xlsx)) to [`xlsxrb`](https://github.com/niku/xlsxrb) with zero downtime and drop-in safety.

## Overview

This project provides adapters for the following widely used Ruby spreadsheet libraries:
- [`rubyXL`](https://github.com/weshatheleopard/rubyXL): A mature, full-featured Document Object Model (DOM) library widely trusted for comprehensive OpenXML spreadsheet creation, cell manipulation, and template editing.
- [`caxlsx`](https://github.com/caxlsx/caxlsx) (formerly `axlsx`): The de facto builder library for generating styled spreadsheets, rich typography, DrawingML charts, and financial reports.
- [`roo`](https://github.com/roo-rb/roo): One of the most widely used spreadsheet reading libraries in Ruby, providing data access, formula and formatting inspection, and support for multiple spreadsheet formats.
- [`xsv`](https://github.com/martijn/xsv): A fast, lightweight streaming reader designed specifically for pulling data out of tabular worksheets into arrays or hashes.
- [`creek`](https://github.com/pythonicrubyist/creek): A stream-based reader designed for large Excel files with coordinate-keyed row iteration, row metadata inspection, and embedded DrawingML image extraction.
- [`fast_excel`](https://github.com/Paxa/fast_excel): A high-performance writer library (wrapping C `libxlsxwriter`) designed for fast, memory-efficient XLSX writing with a concise API.
- [`write_xlsx`](https://github.com/cxn03651/write_xlsx): A comprehensive pure-Ruby port of Perl's `Excel::Writer::XLSX`, supporting formatting, formulas, charts, tables, page setup, and cell coordinate utilities.

[`xlsxrb`](https://github.com/niku/xlsxrb) is a pure Ruby, zero-dependency, streaming-capable, low-memory XLSX engine designed for high-throughput batch workloads.

`xlsxrb-adapters` bridges the best of both worlds by providing drop-in compatible adapter layers to:
1. Support gradual migration (Strangler Fig pattern) from peer XLSX libraries to `xlsxrb` without rewriting application logic.
2. Enable high-throughput, low-memory execution in batch processing and resource-constrained environments while retaining `rubyXL`, `caxlsx`, `roo`, `xsv`, `creek`, `fast_excel`, and `write_xlsx` familiar and battle-tested APIs.
3. Construct interoperability test harnesses against peer libraries and real-world fixtures.
4. Keep `xlsxrb` core strictly zero-dependency, mutant-tested, and type-safe while providing rich compatibility layers.

## Design Principles

1. **No Global Hijacking**:
   Does not reopen or hijack top-level constants like `::RubyXL`, `::Axlsx`, `::Roo`, `::Xsv`, `::Creek`, `::FastExcel`, or `::WriteXLSX`. Instead, exposes namespaces like `Xlsxrb::Adapters::RubyXL`, `Xlsxrb::Adapters::Caxlsx`, `Xlsxrb::Adapters::Roo`, `Xlsxrb::Adapters::Xsv`, `Xlsxrb::Adapters::Creek`, `Xlsxrb::Adapters::FastExcel`, and `Xlsxrb::Adapters::WriteXLSX` so you can run side-by-side during migration or run comparison tests. Optional drop-in aliases (e.g. `Xsv = Xlsxrb::Adapters::Xsv`, `Creek = Xlsxrb::Adapters::Creek`, `FastExcel = Xlsxrb::Adapters::FastExcel`, `WriteXLSX = Xlsxrb::Adapters::WriteXLSX`) are provided for seamless code transitions.
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

### Xsv (`Xlsxrb::Adapters::Xsv` / `Xsv`)

Drop-in replacement for reading and extracting tabular data from XLSX spreadsheets using the [`xsv`](https://github.com/martijn/xsv) API. Provides high-throughput streaming, zero external dependencies, and both array mode and hash mode row iteration.

#### Drop-In Migration Example

```ruby
require "xlsxrb/adapters/xsv"

# Option A: Explicit namespace (recommended to avoid global pollution)
x = Xlsxrb::Adapters::Xsv.open("data.xlsx")

# Option B: Drop-in alias (existing Xsv code works unmodified)
Xsv = Xlsxrb::Adapters::Xsv
x = Xsv.open("data.xlsx")

sheet = x.sheets[0]

# Default Array mode: iterates each row as an Array of typed values
sheet.each do |row|
  puts row.inspect #=> [1, "Widget", 25.5, #<Date: 2026-01-01>]
end

# Random access to rows by index or range
row_5 = sheet[4]
top_10 = sheet[0..9]

# Hash mode: use first row as column headers
sheet.parse_headers!
sheet.each do |row|
  puts row["Widget"] # Hash access by header name
end

# Block form automatically closes workbook
Xsv.open("data.xlsx") do |wb|
  wb.sheets.each { |s| puts s.name }
end
```

#### API Capabilities
- **Document Opening & Streaming**:
  - `Xsv.open(filepath_or_io, trim_empty_rows: false, parse_headers: false, &block)`
  - `Xsv::Workbook.open(data, ...)`
  - `wb.close`
- **Workbooks & Sheets**:
  - `wb.sheets` (array of `Xsv::Sheet`)
  - `wb.sheets_by_name(name)`
  - `wb[index_or_name]`
- **Sheet Operations & Modes**:
  - `sheet.mode` (`:array` or `:hash`)
  - `sheet.row_skip = n` (skip leading rows before headers/data)
  - `sheet.parse_headers!` (activates hash mode; raises `Xsv::DuplicateHeaders` if duplicate header names exist)
  - `sheet.headers` (returns array of header strings)
  - `sheet.each { |row| ... }` / `sheet.each_row { |row| ... }`
  - `sheet[index]` (returns row Array or Hash)
  - `sheet[range]` (returns Array of rows)
  - `sheet.last_row`, `sheet.last_column`
  - `sheet.to_a`
- **Native Bridge**:
  - `wb.to_xlsxrb` (bridges to native `xlsxrb` workbook)
  - `Xlsxrb::Adapters::Xsv.from_xlsxrb(xlsxrb_wb)`

### Creek (`Xlsxrb::Adapters::Creek` / `Creek`)

Drop-in replacement for streaming and extracting tabular data, row metadata, and DrawingML images from XLSX spreadsheets using the [`creek`](https://github.com/pythonicrubyist/creek) API. Provides high-throughput streaming, zero external dependencies, coordinate-based cell mapping (`{"A1" => val}`), compact array rows, and one-cell/two-cell anchor image extraction.

#### Drop-In Migration Example

```ruby
require "xlsxrb/adapters/creek"

# Option A: Explicit namespace (recommended to avoid global pollution)
creek = Xlsxrb::Adapters::Creek::Book.new("data.xlsx")

# Option B: Drop-in alias (existing Creek code works unmodified)
Creek = Xlsxrb::Adapters::Creek
creek = Creek::Book.new("data.xlsx")

sheet = creek.sheets[0]

# Default: row hash keyed by cell coordinate
sheet.rows.each do |row|
  puts row["A1"] # => "Item Name"
end

# Simple rows: array of values
sheet.simple_rows.each do |row|
  puts row.inspect # => ["Item Name", 10, 25.5]
end

# Simple rows with header mapping
sheet.simple_rows(with_headers: true).each do |row|
  puts row["Item Name"]
end

# Row metadata inspection (XML attributes like ht, hidden, collapsed, outlineLevel)
sheet.rows_with_meta_data.each do |row_meta|
  puts "Row #{row_meta['row']} (height: #{row_meta['ht']}, hidden: #{row_meta['hidden']}): #{row_meta['cells']}"
end

# Extract embedded drawing images anchored to cells
sheet.with_images do
  sheet.rows.each do |row|
    images = sheet.images_at("A1")
    images.each { |img| puts "Found image at A1: #{img.path}" }
  end
end

# Close and cleanup any temporary files
creek.close
```

#### API Capabilities
- **Document Loading & Book**:
  - `Creek::Book.new(filepath_or_io, options)`
  - Options: `check_file_extension` (validates `.xlsx` / `.xlsm`), `remote: true` (downloads remote URLs), IO / StringIO buffers
  - `book.sheets` (array of `Creek::Sheet`)
  - `book.style_types` (hash mapping `xf_id` to `:date`, `:time`, `:bignum`, etc.)
  - `book.base_date` (Date object: 1899-12-30 for 1900 calendar, 1904-01-01 for 1904 calendar)
  - `book.close` (cleans up any extracted temporary files)
- **Worksheets & Row Iteration**:
  - `sheet.rows` (Enumerator yielding `{ "A1" => val }` coordinate hash)
  - `sheet.simple_rows(with_headers: false)` (Enumerator yielding Array or header Hash)
  - `sheet.rows_with_meta_data` (Enumerator yielding metadata hash with `"cells"` and row attributes)
  - `sheet.simple_rows_with_meta_data(with_headers: false)`
  - `sheet.state` (`"visible"`, `"hidden"`, etc.)
  - `sheet.name`, `sheet.rid`, `sheet.sheetid`
- **DrawingML & Embedded Image Extraction**:
  - `sheet.with_images { ... }` (extracts drawings and resolves images to tempfiles)
  - `sheet.images_at(cell_ref)` (returns Array of `Pathname` objects for images anchored at cell coordinate)
  - `sheet.ready_for_images?`
- **Native Bridge**:
  - `book.to_xlsxrb` (converts Creek book to native `Xlsxrb::Workbook`)
  - `sheet.to_xlsxrb` (converts Creek sheet to native `Xlsxrb::Elements::Sheet`)
  - `Xlsxrb::Adapters::Creek.from_xlsxrb(xlsxrb_wb)` (creates Creek Book from native `xlsxrb` workbook)

### FastExcel (`Xlsxrb::Adapters::FastExcel` / `FastExcel`)

Drop-in replacement for high-throughput spreadsheet generation using the [`fast_excel`](https://github.com/Paxa/fast_excel) API (wrapping `libxlsxwriter`). Enables pure Ruby, zero-C-dependency, ultra-fast generation with append-only workflows, custom formats, automatic column width calculation, formula cells, and in-memory buffer output.

#### Drop-In Migration Example

```ruby
require "xlsxrb/adapters/fast_excel"

# Option A: Explicit namespace (recommended to avoid global pollution)
wb = Xlsxrb::Adapters::FastExcel.open("sales.xlsx")

# Option B: Drop-in alias (existing FastExcel code works unmodified)
FastExcel = Xlsxrb::Adapters::FastExcel
wb = FastExcel.open("sales.xlsx")

ws = wb.add_worksheet("Q3 Results")
ws.auto_width = true

# Define formats using convenient helpers or CSS colors
header_fmt = wb.add_format(bold: true, bg_color: :navy, font_color: :white, align: :center)
currency_fmt = wb.number_format("$#,##0.00")

# Write header and data rows
ws.append_row(["Product", "Quantity", "Unit Price", "Total"], header_fmt)
ws.append_row(["MacBook Pro", 3, 2499.0, FastExcel::Formula.new("B2*C2")], [nil, nil, currency_fmt, currency_fmt])
ws.append_row(["4K Monitor", 6, 450.0, FastExcel::Formula.new("B3*C3")], [nil, nil, currency_fmt, currency_fmt])
ws << ["Total", FastExcel::Formula.new("SUM(B2:B3)"), nil, FastExcel::Formula.new("SUM(D2:D3)")]

# Freeze header pane & autofilter
ws.freeze_panes(1, 0)
ws.autofilter(0, 0, 3, 3)

# Save and close (or use block form: FastExcel.open("sales.xlsx") { |wb| ... })
wb.close
```

#### In-Memory Binary Buffer Output

```ruby
# FastExcel allows generating files entirely in memory without writing to disk
wb = FastExcel.open
ws = wb.add_worksheet("Report")
ws << ["Timestamp", "Metric", "Value"]
ws << [Time.now, "Requests/sec", 15420]

# Retrieve binary XLSX buffer (automatically cleans up any temp folder)
xlsx_data = wb.read_string
send_data(xlsx_data, filename: "report.xlsx", type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet")
```

#### API Capabilities
- **Workbook Opening & Options**:
  - `FastExcel.open(filepath, constant_memory: false, default_format: nil)` (supports block with auto-close)
  - `wb.add_worksheet(name)`
  - `wb.bold_format`, `wb.number_format(pattern)`, `wb.add_format(opts)`
  - `wb.read_string` (binary XLSX buffer)
  - `wb.close`, `wb.remove_tmp_folder`
  - `wb.set_properties(title:, author:, company:)`
  - `wb.set_custom_property_string / number / boolean / datetime`
- **Worksheet Operations**:
  - `ws.append_row(values, formats = nil)`
  - `ws << values` (idiomatic append operator)
  - `ws.write_value(row, col, value, format = nil)`
  - `ws.write_row(row, values, formats = nil)`
  - Typed writers: `write_string`, `write_number`, `write_datetime`, `write_boolean`, `write_formula`, `write_formula_num`, `write_url`, `write_comment`, `write_blank`
  - `ws.auto_width = true` (tracks cell lengths and assigns padded widths upon sheet close)
  - `ws.set_column(first_col, last_col, width, format)`
  - `ws.set_column_width(col, width)`, `ws.set_columns_width(first_col, last_col, width)`
  - `ws.set_row(row, height, format)`
  - `ws.merge_range(first_row, first_col, last_row, last_col, value, format)`
  - `ws.autofilter(first_row, first_col, last_row, last_col)`
  - `ws.freeze_panes(row, col)`, `ws.split_panes(vertical, horizontal)`
  - `ws.set_right_to_left`, `ws.center_vertically`, `ws.print_row_col_headers`, `ws.set_margins`
  - `ws[:name]`, `ws[:right_to_left]`, `ws[:selected]` (struct-like field inspection)
- **Formatting**:
  - `wb.add_format(bold:, italic:, font_size:, font_name:, font_color:, bg_color:, border:, align:)`
  - 140+ CSS named colors (`:navy`, `:crimson`, `:teal`, `:gold`, `:coral`, etc.) and hex values
  - Alignments: symbols (`:center`, `:left`, `:right`, `:top`, `:bottom`), prefixed (`:align_center`), and hashes (`{ h: :center, v: :center }`)
  - Borders: `border: :thin`, `top: :medium`, `bottom: :double`, colors per edge
- **Native Bridge**:
  - `wb.to_xlsxrb` (compiles mutable builder into immutable `Xlsxrb::Elements::Workbook`)
  - `Xlsxrb::Adapters::FastExcel.from_xlsxrb(xlsxrb_wb)`

### WriteXLSX (`Xlsxrb::Adapters::WriteXLSX` / `Writexlsx`)

Drop-in replacement for [`write_xlsx`](https://github.com/cxn03651/write_xlsx) (the Ruby port of `Excel::Writer::XLSX`). Enables pure Ruby, high-performance generation with write_xlsx's classic declarative API: rich cell writing, 56-color palette resolution, dynamic array formulas, OpenXML charts, tables, conditional formatting, and page setup options.

#### Drop-In Migration Example

```ruby
require "xlsxrb/adapters/write_xlsx"

# Option A: Explicit namespace (recommended to avoid global pollution)
wb = Xlsxrb::Adapters::WriteXLSX.new("sales_summary.xlsx")

# Option B: Drop-in alias (existing WriteXLSX code works unmodified)
WriteXLSX = Xlsxrb::Adapters::WriteXLSX
wb = WriteXLSX.new("sales_summary.xlsx")

ws = wb.add_worksheet("Regional Sales")

# Formats
title_fmt = wb.add_format(bold: 1, size: 16, color: "blue")
header_fmt = wb.add_format(bold: 1, bg_color: "silver", align: "center", border: 1)
currency_fmt = wb.add_format(num_format: "$#,##0.00")
date_fmt = wb.add_format(num_format: "yyyy-mm-dd")

# Write title & table header
ws.write("A1", "Annual Performance", title_fmt)
ws.write_row("A3", %w[Quarter Region Revenue Growth Date TargetMet], header_fmt)

# Write data rows
ws.write_row(3, 0, ["Q1", "North", 450000.0, 0.12, Date.new(2026, 3, 31)])
ws.write_boolean(3, 5, true)
ws.write_row(4, 0, ["Q2", "South", 380000.0, -0.05, Date.new(2026, 6, 30)])
ws.write_boolean(4, 5, false)

# Formula
ws.write("A6", "Total Revenue:", header_fmt)
ws.write_formula("C6", "=SUM(C4:C5)", currency_fmt, 830000.0)

# Column widths & layout
ws.set_column(0, 1, 15)
ws.set_column(2, 2, 18, currency_fmt)
ws.freeze_panes(3, 0)
ws.autofilter("A3:F5")

# Embed chart
chart = wb.add_chart(type: :column)
chart.add_series(
  categories: "='Regional Sales'!$A$4:$A$5",
  values: "='Regional Sales'!$C$4:$C$5",
  name: "Quarterly Revenue"
)
chart.set_title(name: "Revenue by Quarter")
ws.insert_chart("H3", chart)

# Save workbook
wb.close
```

#### In-Memory Binary Buffer Output

```ruby
# Generate XLSX directly in memory without writing to disk
wb = Xlsxrb::Adapters::WriteXLSX.new
ws = wb.add_worksheet("Live Report")
ws.write(0, 0, "Generated at #{Time.now}")
xlsx_data = wb.read_string
```

#### API Capabilities
- **Workbook Operations**:
  - `wb = WriteXLSX.new(filename_or_io, options)`
  - `wb.add_worksheet(name)`
  - `wb.add_format(properties)`
  - `wb.add_chart(type:, subtype:)`
  - `wb.define_name(name, formula)`
  - `wb.set_properties(title:, author:, company:)`, `wb.set_custom_property(name, val)`
  - `wb.set_1904(boolean)`
  - `wb.close`, `wb.read_string`
- **Worksheet Operations**:
  - Cell dispatch: `ws.write(row, col, value, format)` (auto-routes string, number, boolean, date/time, formula, url, row)
  - Explicit writers: `write_string`, `write_number`, `write_blank`, `write_formula`, `write_array_formula`, `write_url`, `write_date_time`, `write_boolean`, `write_rich_string`
  - Array writing: `write_row(row, col, array)`, `write_col(row, col, array)`
  - Layout & Sizing: `ws.set_row(row, height, format)`, `ws.set_column(first_col, last_col, width, format)`, `ws.set_column_pixels(...)`
  - Ranges & Features: `ws.merge_range(...)`, `ws.autofilter(...)`, `ws.freeze_panes(...)`, `ws.split_panes(...)`
  - Graphics: `ws.insert_chart(row, col, chart)`, `ws.insert_image(row, col, image_path)`
  - Tables & Rules: `ws.add_table(...)`, `ws.add_sparkline(...)`, `ws.data_validation(...)`, `ws.conditional_formatting(...)`
  - Page setups: `set_landscape`, `set_portrait`, `set_paper`, `set_margins`, `set_header`, `set_footer`, `print_area`, `fit_to_pages`
- **Format Capabilities**:
  - Typography: `set_bold`, `set_italic`, `set_underline`, `set_font_strikeout`, `set_font`, `set_size`, `set_color`
  - Palette & Colors: 56 standard Excel palette indices (8..63), named colors (`"red"`, `:blue`), `#RRGGBB`, `#RGB`
  - Alignment: `set_align`, `set_valign`, `set_text_wrap`, `set_rotation`, `set_indent`, `set_shrink`
  - Borders & Fills: `set_border`, `set_border_color`, `set_bg_color`, `set_fg_color`, `set_pattern`
  - Number formats: `set_num_format`, built-in format resolution
- **Utility Methods**:
  - `xl_rowcol_to_cell(row, col, row_abs, col_abs)`
  - `xl_cell_to_rowcol(cell_str)`
  - `xl_col_to_name(col, col_abs)`
  - `xl_range(r1, r2, c1, c2)`, `xl_range_formula(sheetname, r1, r2, c1, c2)`
  - `quote_sheetname(name)`
  - `convert_date_time(val, date1904)`
- **Native Bridge**:
  - `wb.to_xlsxrb` (compiles mutable builder into immutable `Xlsxrb::Elements::Workbook`)
  - `Xlsxrb::Adapters::WriteXLSX.from_xlsxrb(xlsxrb_wb)`

## Compatibility (100% Test Pass)

`xlsxrb-adapters` verifies 100% behavioral compatibility and drop-in safety against `rubyXL`, `caxlsx`, `roo`, `xsv`, `creek`, `fast_excel`, and `write_xlsx` through dedicated compatibility suites, official upstream test suites, and cross-validation fixtures:

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

### Xsv Compatibility (`Xlsxrb::Adapters::Xsv`)

Verified against official `xsv` (1.4.1) behavior across the full feature scope documented in [xsv repository](https://github.com/martijn/xsv):

#### Feature Compatibility Matrix

| Feature Domain | Xsv API / Construct | Compatibility Status | Notes |
| :--- | :--- | :--- | :--- |
| **Document Opening** | `Xsv.open`, `Workbook.open` | **100% Supported** | Supports file paths, String buffers, open IO, and block forms |
| **Sheet Iteration** | `sheet.each`, `sheet.each_row`, `sheet.to_a` | **100% Supported** | Iterates row-by-row in `:array` mode or `:hash` mode |
| **Header Parsing** | `sheet.parse_headers!`, `sheet.headers`, `Xsv::DuplicateHeaders` | **100% Supported** | Activates hash mode; detects and raises on duplicate header names |
| **Random Row Access**| `sheet[index]`, `sheet[range]` | **100% Supported** | Indexed and sliced row access; dimension and bounds caching |
| **Type Casting** | Integer, Float, Date, DateTime, Time string (`HH:MM`), Boolean | **100% Supported** | Exact match with xsv type parsing and date formatting |
| **Empty Row Trimming**| `trim_empty_rows: true/false` | **100% Supported** | Automatically trims trailing empty rows when enabled |
| **Inline Strings** | Multi-run `<is><r><t>` parsing | **100% Supported** | Accurately extracts and concatenates inline string runs |

#### Test Suite Breakdown

| Test Suite | Scope | Total Tests | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `xsv_test.rb` | Workbook, sheet, helpers, modes, row_skip, duplicate headers, bounds, and date/time parsing | 18 | 18 | 0 | 100.0% |
| `xsv_compatibility_test.rb` | Side-by-side cross-validation against official `::Xsv` across 11 fixture files | 5 | 5 | 0 | 100.0% |
| `xsv_dropin_test.rb` | End-to-end drop-in replacement workflow (`Xsv = Xlsxrb::Adapters::Xsv`) | 2 | 2 | 0 | 100.0% |
| **Total** | **Xsv Adapter Compatibility Verification (702 assertions)** | **25** | **25** | **0** | **100.0%** |

Run the Xsv compatibility test suite:
```bash
bundle exec rake compatibility:xsv
# or
bundle exec rake test:xsv
```

### Creek Compatibility (`Xlsxrb::Adapters::Creek`)

Verified against official `creek` (2.6.3) behavior and fixtures across the full feature scope documented in [creek repository](https://github.com/pythonicrubyist/creek):

#### Feature Compatibility Matrix

| Feature Domain | Creek API / Construct | Compatibility Status | Notes |
| :--- | :--- | :--- | :--- |
| **Document Opening** | `Creek::Book.new`, extension checks, remote URLs | **100% Supported** | Supports file paths, Pathnames, StringIO, open IO buffers, and remote HTTP downloads |
| **Sheet Iteration** | `sheet.rows`, `sheet.simple_rows` | **100% Supported** | Yields `{ "A1" => val }` coordinate hashes or compact value arrays with header mapping |
| **Row Metadata** | `sheet.rows_with_meta_data`, `sheet.simple_rows_with_meta_data` | **100% Supported** | Parses `<row>` XML attributes (`cells`, `row`, `collapsed`, `hidden`, `ht`, `outlineLevel`, etc.) |
| **Drawing & Images** | `sheet.with_images`, `sheet.images_at` | **100% Supported** | Extracts DrawingML one-cell/two-cell anchors to temporary files as `Pathname` objects |
| **Styles & SharedStrings**| `book.style_types`, SharedStrings, base date | **100% Supported** | Maps `xf_id` to `:date`, `:time`, `:bignum`; supports rich text `<si><r><t>` and 1900/1904 calendars |
| **Native Bridge** | `book.to_xlsxrb`, `sheet.to_xlsxrb`, `Creek.from_xlsxrb` | **100% Supported** | Seamless bidirectional translation to native `xlsxrb` models |

#### Test Suite Breakdown

| Test Suite | Scope | Total Tests | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `creek_test.rb` | Book, Sheet, Styles, SharedStrings, Drawing, and Bridge APIs | 20 | 20 | 0 | 100.0% |
| `creek_compatibility_test.rb` | Side-by-side cross-validation against official `::Creek` across all 10 fixture files | 8 | 8 | 0 | 100.0% |
| `creek_dropin_test.rb` | End-to-end drop-in replacement workflow (`Creek = Xlsxrb::Adapters::Creek`) | 2 | 2 | 0 | 100.0% |
| **Total** | **Creek Adapter Compatibility Verification (3,415 assertions)** | **30** | **30** | **0** | **100.0%** |

Run the Creek compatibility test suite:
```bash
bundle exec rake compatibility:creek
# or
bundle exec rake test:creek
```

### FastExcel Compatibility (`Xlsxrb::Adapters::FastExcel`)

Verified against official `fast_excel` (0.5.0) behavior across the full feature scope documented in [fast_excel repository](https://github.com/Paxa/fast_excel):

#### Feature Compatibility Matrix

| Feature Domain | FastExcel API / Construct | Compatibility Status | Notes |
| :--- | :--- | :--- | :--- |
| **Document & Buffer** | `FastExcel.open`, tempdir allocation, `read_string`, `remove_tmp_folder` | **100% Supported** | Supports file paths, block execution with auto-close, and zero-disk in-memory generation |
| **Append & Row Writes** | `append_row`, `<<`, `write_value`, `write_row`, typed writers | **100% Supported** | Strict constant_memory enforcement, last_row_number tracking, typed cell dispatch |
| **Layout & Auto Width** | `auto_width = true`, `set_column`, `set_row`, `freeze_panes`, `split_panes`, `merge_range`, `autofilter` | **100% Supported** | Computes character and font-proportional widths on sheet close; full OpenXML sheet views |
| **Formats & Styling** | `add_format`, `bold_format`, `number_format`, font/border/fill/alignment | **100% Supported** | Maps 140+ CSS named colors, hex strings, symbol/hash alignments, and border styles |
| **Formulas & URLs** | `FastExcel::Formula`, `write_formula_num`, `FastExcel::URL` | **100% Supported** | Emits formula cells with default cached values for OpenXML reader compatibility |
| **Native Bridge** | `wb.to_xlsxrb`, `FastExcel.from_xlsxrb` | **100% Supported** | Seamless bidirectional translation to native `xlsxrb` models |

#### Test Suite Breakdown

| Test Suite | Scope | Total Tests | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `fast_excel_test.rb` | Constants, Enums, Formats, Sheet, Workbook, Auto-width, Typed writes, Tempfiles, Bridge | 24 | 24 | 0 | 100.0% |
| `fast_excel_dropin_test.rb` | End-to-end drop-in replacement workflow (`FastExcel = Xlsxrb::Adapters::FastExcel`) | 5 | 5 | 0 | 100.0% |
| `fast_excel_compatibility_test.rb` | Side-by-side cross-validation against official `::FastExcel` (primitives, formulas, auto-width, format properties, read_string) | 5 | 5 | 0 | 100.0% |
| **Total** | **FastExcel Adapter Compatibility Verification (204 assertions)** | **34** | **34** | **0** | **100.0%** |

Run the FastExcel compatibility test suite:
```bash
bundle exec rake compatibility:fast_excel
# or
bundle exec rake test:fast_excel
```

### WriteXLSX Compatibility (`Xlsxrb::Adapters::WriteXLSX`)

Verified against official `write_xlsx` (1.15.1) behavior across the full feature scope documented in [write_xlsx repository](https://github.com/cxn03651/write_xlsx):

#### Feature Compatibility Matrix

| Feature Domain | WriteXLSX API / Construct | Compatibility Status | Notes |
| :--- | :--- | :--- | :--- |
| **Object Model** | `Workbook`, `Worksheet`, `Format`, `Chart` | **100% Supported** | Drop-in compatible DOM; supports block and imperative building |
| **Types & Values** | String, Integer, Float, Date, Time, DateTime, Boolean, Formula, URL | **100% Supported** | Automatic dispatch via `write`, typed writers (`write_string`, `write_number`, `write_date_time`, `write_boolean`, `write_url`, `write_formula`), array formulas |
| **Row & Column Operations** | `write_row`, `write_col`, `set_row`, `set_column`, `set_column_pixels`, dimensions tracking (`dim_rowmin`, `dim_colmax`, `dimension`) | **100% Supported** | 1D array row/col batch writing and column width scaling |
| **Formats & Styling** | `add_format` (fonts, colors, alignments, borders, fills, num_formats, rotation, indent, shrink) | **100% Supported** | 56-color standard palette (8..63), CSS colors, hex formats, border styles, text wrapping |
| **Panes & Views** | `freeze_panes`, `split_panes`, `merge_range`, `autofilter` | **100% Supported** | Split panes, frozen rows/cols, merged ranges with format propagation, table autofilters |
| **DrawingML Charts** | `add_chart`, series (`add_series`), categories, values, names, chart types (Area, Bar, Column, Line, Pie, Scatter, etc.) | **100% Supported** | Native `xlsxrb` chart configuration and DrawingML generation |
| **Coordinate Utilities** | `xl_rowcol_to_cell`, `xl_cell_to_rowcol`, `xl_col_to_name`, `xl_range`, `xl_range_formula` | **100% Supported** | Full WriteXLSX / Excel::Writer::XLSX coordinate conversion functions |
| **Advanced Features** | Tables (`add_table`), Sparklines (`add_sparkline`), Data Validation, Conditional Formatting, Page Setup & Margins | **100% Supported** | Complete OpenXML feature mappings |
| **Memory & Serialization** | `close`, `read_string` | **100% Supported** | High-throughput export to file, IO, or in-memory binary string buffer |
| **Native Bridge** | `wb.to_xlsxrb`, `WriteXLSX.from_xlsxrb` | **100% Supported** | Seamless bidirectional translation to native `xlsxrb` models |

#### Test Suite Breakdown

| Test Suite | Scope | Total Tests | Passed | Failed | Pass Rate |
| :--- | :--- | :--- | :--- | :--- | :--- |
| `write_xlsx_test.rb` | Workbook, worksheet, format, charts, utilities, row/col operations, dimensions, properties, bridge | 9 | 9 | 0 | 100.0% |
| `write_xlsx_dropin_test.rb` | End-to-end drop-in replacement workflow (`WriteXLSX = Xlsxrb::Adapters::WriteXLSX`), in-memory buffer (`read_string`), charts, formats | 4 | 4 | 0 | 100.0% |
| `write_xlsx_compatibility_test.rb` | Side-by-side cross-validation against official `write_xlsx` gem (primitives, formulas, formats, cell notations, workbook output) | 4 | 4 | 0 | 100.0% |
| **Total** | **WriteXLSX Adapter Compatibility Verification (212 assertions)** | **17** | **17** | **0** | **100.0%** |

Run the WriteXLSX compatibility test suite:
```bash
bundle exec rake compatibility:write_xlsx
# or
bundle exec rake test:write_xlsx
```

Run all compatibility suites (RubyXL, Caxlsx, Roo, Xsv, Creek, FastExcel, and WriteXLSX):
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

### 4. Xsv Migration: 1,000,000 cells (100,000 rows × 10 cols)

[`xsv`](https://github.com/martijn/xsv) is a fast, lightweight streaming XLSX parser designed for extracting data from tabular worksheets. `xlsxrb-adapters (Xsv)` provides 100% drop-in API compatibility for existing `xsv` workflows while running **~5.0x faster** and reducing GC churn by **~95.8%**:

| Operation | Library / Mode | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. xsv Speed | vs. xsv GC Count |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Read** | `xsv` (1.4.1) | 18.762 s | 79.9 MB | 1,489.0 | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters (Xsv)`** | **3.784 s** | **176.4 MB** | **63.0** | **~5.0x faster** | **95.8% reduced** |
| | `xlsxrb (In-Memory)` | 4.133 s | 446.9 MB | 19.0 | ~4.5x faster | 98.7% reduced |
| | `xlsxrb (Streaming)` | 2.288 s | 112.4 MB | 68.0 | ~8.2x faster | 95.4% reduced |

### 5. Creek Migration: 1,000,000 cells (100,000 rows × 10 cols)

Official `creek` is a stream-based XLSX parser designed for reading large files with low memory. `xlsxrb-adapters (Creek)` provides 100% drop-in API compatibility for existing `creek` workflows while running **~1.9x faster**, reducing peak memory by **~83.6%**, and cutting GC cycles by **~57.0%**:

| Operation | Library / Mode | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. creek Speed | vs. creek Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Read** | `creek` (2.6.3) | 8.359 s | 840.6 MB | 384.0 | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters (Creek)`** | **4.497 s** | **137.8 MB** | **165.0** | **~1.9x faster** | **83.6% reduced** |
| | `xlsxrb (In-Memory)` | 3.345 s | 448.0 MB | 19.0 | ~2.5x faster | 46.7% reduced |
| | `xlsxrb (Streaming)` | 2.090 s | 112.5 MB | 32.0 | ~4.0x faster | **86.6% reduced** |

### 6. WriteXLSX Migration: 1,000,000 cells (100,000 rows × 10 cols)

Official `write_xlsx` is a pure Ruby Excel writer ported from Perl's `Excel::Writer::XLSX`. `xlsxrb-adapters (WriteXLSX)` provides 100% drop-in compatibility for existing `write_xlsx` workflows while cutting GC cycles by up to **~48%**. For maximum performance in batch pipelines, native `xlsxrb` streaming generation yields **~2.4x faster execution** with **~66% lower memory footprint**:

| Operation | Library / Mode | Time (Median) | Peak Memory (VmHWM) | GC Count | vs. write_xlsx Speed | vs. write_xlsx Memory |
| :--- | :--- | :--- | :--- | :--- | :--- | :--- |
| **Write** | `write_xlsx` (1.15.1) | 5.103 s | 201.2 MB | 29.0 | 1.0x (baseline) | 100% (baseline) |
| | **`xlsxrb-adapters (WriteXLSX)`** | **7.695 s** | **706.7 MB** | **15.0** | **~0.7x (48.3% fewer GCs)** | +251.2% (builder DOM) |
| | `xlsxrb (In-Memory)` | 2.982 s | 355.3 MB | 11.0 | ~1.7x faster | +76.6% |
| | `xlsxrb (Streaming)` | 2.172 s | 67.9 MB | 29.0 | **~2.4x faster** | **66.3% reduced** |

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

# Run all unit and adapter tests (254 tests, 7,084 assertions)
bundle exec rake test

# Run Caxlsx compatibility test suite (52 tests, 344 assertions)
bundle exec rake test:caxlsx

# Run RubyXL adapter test suite (59 tests, 438 assertions)
bundle exec rake test:ruby_xl

# Run Roo adapter test suite (37 tests, 1,769 assertions)
bundle exec rake test:roo

# Run Xsv adapter test suite (25 tests, 702 assertions)
bundle exec rake test:xsv

# Run Creek adapter test suite (30 tests, 3,415 assertions)
bundle exec rake test:creek

# Run FastExcel adapter test suite (34 tests, 204 assertions)
bundle exec rake test:fast_excel

# Run WriteXLSX adapter test suite (17 tests, 212 assertions)
bundle exec rake test:write_xlsx

# Run official rubyXL RSpec compatibility suite (363 tests)
bundle exec rake compatibility:ruby_xl

# Run Creek compatibility suite against official fixtures
bundle exec rake compatibility:creek

# Run FastExcel compatibility suite against official fast_excel gem
bundle exec rake compatibility:fast_excel

# Run WriteXLSX compatibility suite against official write_xlsx gem
bundle exec rake compatibility:write_xlsx

# Run all compatibility suites (RubyXL, Caxlsx, Roo, Xsv, Creek, FastExcel, and WriteXLSX)
bundle exec rake compatibility

# Run static type checking with Steep
bundle exec rake typecheck

# Run linter
bundle exec rubocop

# Run benchmarks (usage: rake benchmark [rows=10000] [cols=10] [runs=3] [category=all|caxlsx|rubyxl|roo|xsv|creek|write_xlsx])
bundle exec rake benchmark
```

## Acknowledgements

`xlsxrb-adapters` would not exist without the following projects and the people behind them. We extend our sincere respect and gratitude to:

- Vivek Bhagwat, Wesha ([weshatheleopard](https://github.com/weshatheleopard)), and all contributors to [`rubyXL`](https://github.com/weshatheleopard/rubyXL) for pioneering comprehensive OpenXML DOM manipulation in Ruby and maintaining it for more than a decade.
- Randy Morgan ([randym](https://github.com/randym)), Jurriaan Pruis ([jurriaan](https://github.com/jurriaan)), and the [`caxlsx`](https://github.com/caxlsx/caxlsx) community for the builder API that has become the de facto standard for styled spreadsheets, charts, and reports in Ruby.
- Thomas Preymesser, Hugh McGowan, and the [`roo`](https://github.com/roo-rb/roo) community for creating and maintaining one of the most widely used spreadsheet reading libraries in the Ruby ecosystem.
- Martijn Storck ([martijn](https://github.com/martijn)) and contributors to [`xsv`](https://github.com/martijn/xsv) for showing how lightweight streaming with simple array/hash rows can make spreadsheet reading remarkably fast.
- Ramtin Vaziri ([pythonicrubyist](https://github.com/pythonicrubyist)) and contributors to [`creek`](https://github.com/pythonicrubyist/creek) for a simple, streaming-oriented API with cell-reference-keyed rows, along with its approach to image extraction.
- Pavel Evstigneev ([Paxa](https://github.com/Paxa)) and contributors to [`fast_excel`](https://github.com/Paxa/fast_excel) for bringing fast, C-backed spreadsheet generation to Ruby with a clean formatting DSL.
- Hideo Nakamura ([cxn03651](https://github.com/cxn03651)) for creating [`write_xlsx`](https://github.com/cxn03651/write_xlsx), a comprehensive pure-Ruby Excel writer, and John McNamara ([jmcnamara](https://github.com/jmcnamara)) for the Perl modules it is ported from, [`Excel::Writer::XLSX`](https://github.com/jmcnamara/excel-writer-xlsx) and its predecessor [`Spreadsheet::WriteExcel`](https://github.com/jmcnamara/spreadsheet-writeexcel), as well as [`libxlsxwriter`](https://github.com/jmcnamara/libxlsxwriter), the C library that powers `fast_excel`.

`xlsxrb-adapters` builds directly on the APIs these projects designed and the test suites they published. The test fixtures under `test/fixtures/roo` and `test/fixtures/creek` are taken from the respective upstream projects and remain under their original MIT licenses.

This is an independent project and is not affiliated with or endorsed by the authors or maintainers of the libraries listed above.

## License

MIT
