# frozen_string_literal: true

require "test_helper"
require "stringio"

class XlsxrbAdaptersWriteXLSXTest < Test::Unit::TestCase
  def test_constants_and_limits
    assert_equal 1_048_576, Xlsxrb::Adapters::WriteXLSX::ROW_MAX
    assert_equal 16_384, Xlsxrb::Adapters::WriteXLSX::COL_MAX
    assert_equal 32_767, Xlsxrb::Adapters::WriteXLSX::STR_MAX
    assert_equal 2079, Xlsxrb::Adapters::WriteXLSX::URL_MAX

    assert_equal StandardError, Xlsxrb::Adapters::WriteXLSX::WriteXLSXInsufficientArgumentError.superclass
    assert_equal StandardError, Xlsxrb::Adapters::WriteXLSX::WriteXLSXDimensionError.superclass
    assert_equal StandardError, Xlsxrb::Adapters::WriteXLSX::WriteXLSXOptionParameterError.superclass
  end

  def test_colors_and_palette
    colors = Xlsxrb::Adapters::WriteXLSX::Colors

    assert_equal "FF0000", colors.to_hex("red")
    assert_equal "0000FF", colors.to_hex(:blue)
    assert_equal "00FF00", colors.to_hex("#00FF00")
    assert_equal "FFFFFF", colors.to_hex("#FFF")
    assert_equal "000000", colors.to_hex(8) # Palette 8 is black
    assert_equal "FFFFFF", colors.to_hex(9) # Palette 9 is white
    assert_equal "FF0000", colors.to_hex(10) # Palette 10 is red
    assert_equal "000000", colors.to_hex(0x000000)
    assert_equal "AABBCC", colors.to_hex(0xAABBCC)
    assert_nil colors.to_hex(nil)
  end

  def test_utility_cell_reference
    cell_ref = Xlsxrb::Adapters::WriteXLSX::Utility::CellReference

    assert_equal "A1", cell_ref.xl_rowcol_to_cell(0, 0)
    assert_equal "$A$1", cell_ref.xl_rowcol_to_cell(0, 0, 1, 1)
    assert_equal "$A1", cell_ref.xl_rowcol_to_cell(0, 0, 0, 1)
    assert_equal "A$1", cell_ref.xl_rowcol_to_cell(0, 0, 1, 0)
    assert_equal "B3", cell_ref.xl_rowcol_to_cell(2, 1)
    assert_equal "Z100", cell_ref.xl_rowcol_to_cell(99, 25)
    assert_equal "AA1", cell_ref.xl_rowcol_to_cell(0, 26)

    assert_equal [0, 0, false, false], cell_ref.xl_cell_to_rowcol("A1")
    assert_equal [2, 1, false, false], cell_ref.xl_cell_to_rowcol("B3")
    assert_equal [0, 0, true, true], cell_ref.xl_cell_to_rowcol("$A$1")
    assert_equal [99, 25, false, false], cell_ref.xl_cell_to_rowcol("Z100")

    assert_equal "A", cell_ref.xl_col_to_name(0)
    assert_equal "Z", cell_ref.xl_col_to_name(25)
    assert_equal "AA", cell_ref.xl_col_to_name(26)
    assert_equal "XFD", cell_ref.xl_col_to_name(16_383)

    assert_equal "A1:B2", cell_ref.xl_range(0, 1, 0, 1)
    assert_equal "=Sheet1!$A$1:$B$2", cell_ref.xl_range_formula("Sheet1", 0, 1, 0, 1)
    assert_equal "='My Sheet'!$A$1:$B$2", cell_ref.xl_range_formula("My Sheet", 0, 1, 0, 1)
    assert_equal "'Sheet 1'", cell_ref.quote_sheetname("Sheet 1")
    assert_equal "Sheet1", cell_ref.quote_sheetname("Sheet1")

    assert_equal [0, 0, 1, 1], cell_ref.row_col_notation("A1:B2")
    assert_equal [0, 0], cell_ref.row_col_notation("A1")
    assert_nil cell_ref.row_col_notation(0)
  end

  def test_utility_common_and_datetime
    common = Xlsxrb::Adapters::WriteXLSX::Utility::Common

    assert_equal true, common.ptrue?(1)
    assert_equal true, common.ptrue?(true)
    assert_equal true, common.ptrue?("1")
    assert_equal true, common.ptrue?("true")
    assert_equal false, common.ptrue?(0)
    assert_equal false, common.ptrue?(false)
    assert_equal false, common.ptrue?(nil)
    assert_equal false, common.ptrue?("0")

    dt_mod = Xlsxrb::Adapters::WriteXLSX::Utility::DateTime
    t = Time.utc(2026, 1, 1, 12, 0, 0)
    serial = dt_mod.convert_date_time(t, false)
    assert_in_delta 46_023.5, serial, 0.001

    d = Date.new(2026, 1, 1)
    serial_d = dt_mod.convert_date_time(d, false)
    assert_in_delta 46_023.0, serial_d, 0.001

    iso_serial = dt_mod.convert_date_time("2026-01-01T12:00:00", false)
    assert_in_delta 46_023.5, iso_serial, 0.001
  end

  def test_format_properties_and_styles
    wb = Xlsxrb::Adapters::WriteXLSX.new
    fmt = wb.add_format
    fmt.set_bold
    fmt.set_italic
    fmt.set_underline(1)
    fmt.set_strikeout
    fmt.set_font("Courier New")
    fmt.set_size(14)
    fmt.set_color("red")
    fmt.set_align("center")
    fmt.set_valign("vcenter")
    fmt.set_bg_color("yellow")
    fmt.set_border(1)
    fmt.set_border_color("blue")
    fmt.set_num_format("#,##0.00")
    fmt.set_text_wrap
    fmt.set_rotation(45)

    assert_equal 1, fmt.bold
    assert_equal true, fmt.bold?
    assert_equal 1, fmt.italic
    assert_equal true, fmt.italic?
    assert_equal 1, fmt.underline
    assert_equal 1, fmt.strikeout
    assert_equal "Courier New", fmt.font
    assert_equal 10, fmt.font_color
    assert_equal 2, fmt.align
    assert_equal 2, fmt.valign
    assert_equal 13, fmt.bg_color
    assert_equal 1, fmt.border
    assert_equal 12, fmt.border_color
    assert_equal "#,##0.00", fmt.num_format
    assert_equal 1, fmt.text_wrap
    assert_equal 45, fmt.rotation

    # Test copy
    fmt2 = wb.add_format
    fmt2.copy(fmt)
    assert_equal 1, fmt2.bold
    assert_equal "Courier New", fmt2.font
    assert_equal 14, fmt2.size

    # Test set_format_properties
    fmt3 = wb.add_format(bold: 1, color: "green", size: 16)
    assert_equal 1, fmt3.bold
    assert_equal 17, fmt3.font_color
    assert_equal 16, fmt3.size

    # Style generation
    style = fmt.to_style_hash
    assert_instance_of Hash, style
    assert_equal true, style.dig(:font, :bold)
    assert_equal true, style.dig(:font, :italic)
    assert_equal "Courier New", style.dig(:font, :name)
    assert_equal 14, style.dig(:font, :sz)
    assert_equal "center", style.dig(:alignment, :horizontal)
    assert_equal "center", style.dig(:alignment, :vertical)
    assert_equal 45, style.dig(:alignment, :text_rotation)
    assert_equal true, style.dig(:alignment, :wrap_text)
  end

  def test_chart_creation_and_subclasses
    wb = Xlsxrb::Adapters::WriteXLSX.new
    chart = wb.add_chart(type: :column)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Column, chart

    chart.add_series(
      categories: "=Sheet1!$A$2:$A$5",
      values: "=Sheet1!$B$2:$B$5",
      name: "Test Series"
    )
    chart.set_title(name: "My Chart")
    chart.set_x_axis(name: "Categories")
    chart.set_y_axis(name: "Values")
    chart.set_legend(position: "bottom")
    chart.set_style(10)

    opts = chart.to_chart_options
    assert_equal :col, opts[:type]
    assert_equal "My Chart", opts[:title]
    assert_equal "Categories", opts[:cat_axis_title]
    assert_equal "Values", opts[:val_axis_title]
    assert_equal "bottom", opts[:legend]
    assert_equal 10, opts[:style]
    assert_equal 1, opts[:series].size
    assert_equal "=Sheet1!$A$2:$A$5", opts[:series].first[:cat_ref]
    assert_equal "=Sheet1!$B$2:$B$5", opts[:series].first[:val_ref]

    # Other subclasses
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Area, wb.add_chart(type: :area)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Bar, wb.add_chart(type: :bar)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Line, wb.add_chart(type: :line)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Pie, wb.add_chart(type: :pie)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Scatter, wb.add_chart(type: :scatter)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Radar, wb.add_chart(type: :radar)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Chart::Doughnut, wb.add_chart(type: :doughnut)
  end

  def test_worksheet_writing_and_features
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws = wb.add_worksheet("Data")

    # String, number, blank, boolean, formula
    ws.write(0, 0, "Hello")
    ws.write("B1", 42)
    ws.write(0, 2, nil)
    ws.write("D1", true)
    ws.write(0, 4, "=SUM(A1:C1)")

    # Date and Time
    t = Time.utc(2026, 6, 15, 10, 30, 0)
    ws.write_date_time("A2", t)

    # URL
    ws.write_url("B2", "https://github.com", nil, "GitHub")

    # Rich string
    bold_fmt = wb.add_format(bold: 1)
    ws.write_rich_string("C2", "Normal and ", bold_fmt, "Bold")

    # Arrays: write_row, write_col
    ws.write_row("A3", %w[Row1 Row2 Row3])
    ws.write_col("E1", [100, 200, 300])

    # Columns and rows settings
    ws.set_column(0, 0, 15)
    ws.set_column_pixels(1, 1, 140)
    ws.set_row(0, 25)

    # Merge range
    merge_fmt = wb.add_format(align: "center")
    ws.merge_range("A5:C5", "Merged Header", merge_fmt)

    # Autofilter
    ws.autofilter("A3:C3")
    ws.filter_column(0, "Row1")

    # Panes
    ws.freeze_panes(1, 0)

    # Page setup
    ws.set_landscape
    ws.set_paper(9) # A4
    ws.set_margins(0.5, 0.5, 0.75, 0.75)
    ws.set_header("&CReport")
    ws.set_footer("&RPage &P")
    ws.fit_to_pages(1, 1)

    # Table & Data Validation
    ws.add_table("A8:C12", name: "MyTable")
    ws.data_validation("A8:A12", validate: "list", source: %w[High Med Low])

    # Verify dimensions
    dim = ws.dimension
    assert_equal 0, dim[:row_min]
    assert_equal 4, dim[:row_max]
    assert_equal 0, dim[:col_min]
    assert_equal 4, dim[:col_max]

    # Convert to xlsxrb
    xlsxrb_ws = ws.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Worksheet, xlsxrb_ws
    assert_equal "Data", xlsxrb_ws.name
    assert_equal 5, xlsxrb_ws.rows.first.cells.size
    assert_equal "Hello", xlsxrb_ws.rows.first.cells[0].value
    assert_equal 42, xlsxrb_ws.rows.first.cells[1].value
    assert_equal true, xlsxrb_ws.rows.first.cells[3].value
  end

  def test_workbook_file_output_and_stringio
    with_tempfile do |filepath|
      wb = Xlsxrb::Adapters::WriteXLSX.new(filepath)
      ws = wb.add_worksheet("Test")
      ws.write(0, 0, "Test Value")
      wb.close

      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      doc = Xlsxrb.read(filepath)
      assert_equal ["Test"], doc.sheets.map(&:name)
      assert_equal "Test Value", doc.sheets.first.cell(0, 0)&.value
    end

    # Test read_string
    wb_mem = Xlsxrb::Adapters::WriteXLSX.new
    ws_mem = wb_mem.add_worksheet("Memory")
    ws_mem.write(0, 0, "Memory Value")
    binary_data = wb_mem.read_string
    assert_instance_of String, binary_data
    assert binary_data.bytesize.positive?

    # Test StringIO
    sio = StringIO.new
    wb_sio = Xlsxrb::Adapters::WriteXLSX.new(sio)
    ws_sio = wb_sio.add_worksheet("Stream")
    ws_sio.write(0, 0, "Stream Value")
    wb_sio.close
    assert sio.string.bytesize.positive?
  end

  def test_bridge_roundtrip
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws1 = wb.add_worksheet("SheetA")
    fmt = wb.add_format(bold: 1, color: "blue")
    ws1.write(0, 0, "Key", fmt)
    ws1.write(0, 1, 999)
    ws2 = wb.add_worksheet("SheetB")
    ws2.write(0, 0, "Another")

    xlsxrb_wb = wb.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
    assert_equal 2, xlsxrb_wb.sheets.size
    assert_equal %w[SheetA SheetB], xlsxrb_wb.sheets.map(&:name)

    # from_xlsxrb
    roundtrip_wb = Xlsxrb::Adapters::WriteXLSX.from_xlsxrb(xlsxrb_wb)
    assert_instance_of Xlsxrb::Adapters::WriteXLSX::Workbook, roundtrip_wb
    assert_equal 2, roundtrip_wb.sheets.size
    assert_equal "Key", roundtrip_wb.sheets[0].table[[0, 0]].value
    assert_equal 999, roundtrip_wb.sheets[0].table[[0, 1]].value
    assert_equal "Another", roundtrip_wb.sheets[1].table[[0, 0]].value
  end

  def test_array_formula_serialization_and_to_xlsxrb
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws = wb.add_worksheet("ArrayFormula")
    ws.write_array_formula(0, 0, 0, 2, "SUM(B1:B10*C1:C10)", nil, 100)

    xlsxrb_ws = ws.to_xlsxrb
    c0 = xlsxrb_ws.rows.first.cells[0]
    assert_instance_of Xlsxrb::Elements::Formula, c0.formula
    assert_equal :array, c0.formula.type
    assert_equal "A1:C1", c0.formula.ref
    assert_equal "SUM(B1:B10*C1:C10)", c0.formula.expression
    assert_equal 100, c0.value

    with_tempfile do |filepath|
      wb_file = Xlsxrb::Adapters::WriteXLSX.new(filepath)
      ws_file = wb_file.add_worksheet("Sheet1")
      ws_file.write_array_formula(0, 0, 0, 2, "SUM(B1:B10*C1:C10)", nil, 100)
      wb_file.close

      entries = Xlsxrb::Ooxml::ZipReader.open(filepath, &:read_all)
      sheet_xml = entries["xl/worksheets/sheet1.xml"]
      assert_match(%r{<f t="array" ref="A1:C1">SUM\(B1:B10\*C1:C10\)</f>}, sheet_xml)
    end
  end

  def test_column_style_index_keyword
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws = wb.add_worksheet("Cols")
    fmt = wb.add_format(bold: 1, color: "red")
    ws.set_column("A:A", 25, fmt)

    xlsxrb_ws = ws.to_xlsxrb
    col = xlsxrb_ws.columns.first
    assert_not_nil col.style_index
    assert_equal 25, col.width
  end

  def test_sparkline_groups_and_serialization
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws = wb.add_worksheet("Sparklines")
    ws.write_row(0, 0, [10, 20, 30, 40])
    ws.add_sparkline(location: "E1", range: "A1:D1", type: "win_loss", markers: 1, high_point: 1)

    xlsxrb_ws = ws.to_xlsxrb
    assert_equal 1, xlsxrb_ws.sparkline_groups.size
    sg = xlsxrb_ws.sparkline_groups.first
    assert_equal "stacked", sg[:type]
    assert_equal true, sg[:markers]
    assert_equal true, sg[:high]
    assert_equal [{ location_ref: "E1", data_ref: "'Sparklines'!A1:D1" }], sg[:sparklines]

    with_tempfile do |filepath|
      wb_file = Xlsxrb::Adapters::WriteXLSX.new(filepath)
      ws_file = wb_file.add_worksheet("Sparklines")
      ws_file.write_row(0, 0, [10, 20, 30, 40])
      ws_file.add_sparkline(location: "E1", range: "A1:D1", type: "win_loss", markers: 1, high_point: 1)
      wb_file.close

      entries = Xlsxrb::Ooxml::ZipReader.open(filepath, &:read_all)
      sheet_xml = entries["xl/worksheets/sheet1.xml"]
      assert_match(/x14:sparklineGroup[^>]*type="stacked"/, sheet_xml)
      assert_match(/high="1"/, sheet_xml)
      assert_match(/markers="1"/, sheet_xml)
      assert_match(%r{<xm:sqref>E1</xm:sqref>}, sheet_xml)
    end
  end

  def test_print_area_and_repeat_titles_auto_generation
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws = wb.add_worksheet("Printing")
    ws.print_area("A1:H50")
    ws.repeat_rows(0, 1)
    ws.repeat_columns(0, 1)

    xlsxrb_ws = ws.to_xlsxrb
    assert_equal "A1:H50", xlsxrb_ws.print_area
    assert_equal({ rows: "1:2", cols: "A:B" }, xlsxrb_ws.print_titles)

    with_tempfile do |filepath|
      wb_file = Xlsxrb::Adapters::WriteXLSX.new(filepath)
      ws_file = wb_file.add_worksheet("Printing")
      ws_file.write(0, 0, "Top Left")
      ws_file.print_area("A1:H50")
      ws_file.repeat_rows(0, 1)
      ws_file.repeat_columns(0, 1)
      wb_file.close

      entries = Xlsxrb::Ooxml::ZipReader.open(filepath, &:read_all)
      wb_xml = entries["xl/workbook.xml"]
      assert_match(/_xlnm\.Print_Area/, wb_xml)
      assert_match(/_xlnm\.Print_Titles/, wb_xml)
      assert_match(/(?:'|&apos;)Printing(?:'|&apos;)!\$A\$1:\$H\$50/, wb_xml)
      assert_match(/(?:'|&apos;)Printing(?:'|&apos;)!\$A:\$B,(?:'|&apos;)Printing(?:'|&apos;)!\$1:\$2/, wb_xml)
    end
  end

  def test_rich_text_with_to_font_hash
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws = wb.add_worksheet("RichText")
    fmt_bold = wb.add_format(bold: 1)
    fmt_red = wb.add_format(color: "red")

    assert_respond_to fmt_bold, :to_font_hash
    assert_equal({ name: "Calibri", sz: 11, bold: true }, fmt_bold.to_font_hash)

    ws.write_rich_string(0, 0, fmt_bold, "Bold ", fmt_red, "Red")

    xlsxrb_ws = ws.to_xlsxrb
    c0 = xlsxrb_ws.rows.first.cells[0]
    assert_instance_of Xlsxrb::Elements::RichText, c0.value
    assert_equal "Bold Red", c0.value.to_s
    assert_equal 2, c0.value.runs.size
    assert_equal true, c0.value.runs[0].font[:bold]
    assert_equal "FF0000", c0.value.runs[1].font[:color]
  end

  def test_hyperlink_screentip_and_cell_tooltip
    wb = Xlsxrb::Adapters::WriteXLSX.new
    ws = wb.add_worksheet("Links")
    ws.write_url(0, 0, "https://github.com", nil, "GitHub", "Visit repository")

    xlsxrb_ws = ws.to_xlsxrb
    hl = xlsxrb_ws.hyperlinks["A1"]
    assert_equal "https://github.com", hl[:url]
    assert_equal "GitHub", hl[:display]
    assert_equal "Visit repository", hl[:tooltip]

    with_tempfile do |filepath|
      wb_file = Xlsxrb::Adapters::WriteXLSX.new(filepath)
      ws_file = wb_file.add_worksheet("Links")
      ws_file.write_url(0, 0, "https://github.com", nil, "GitHub", "Visit repository")
      wb_file.close

      entries = Xlsxrb::Ooxml::ZipReader.open(filepath, &:read_all)
      sheet_xml = entries["xl/worksheets/sheet1.xml"]
      assert_match(/tooltip="Visit repository"/, sheet_xml)
      assert_match(/display="GitHub"/, sheet_xml)
    end
  end
end
