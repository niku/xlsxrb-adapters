# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersFastExcelTest < Test::Unit::TestCase
  def test_constants_and_enums
    assert_equal 8.43, Xlsxrb::Adapters::FastExcel::DEF_COL_WIDTH
    assert_equal 86_400.0, Xlsxrb::Adapters::FastExcel::XLSX_DATE_DAY
    assert_equal 25_569, Xlsxrb::Adapters::FastExcel::XLSX_DATE_EPOCH_DIFF

    # Error enum
    assert_equal 0, Xlsxrb::Adapters::FastExcel::ERROR_ENUM.find(:no_error)
    assert_equal :no_error, Xlsxrb::Adapters::FastExcel::ERROR_ENUM.find(0)
    assert_includes Xlsxrb::Adapters::FastExcel::ERROR_ENUM.symbols, :error_sheetname_length_exceeded

    # Color enum
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel::COLOR_ENUM.find(:color_red)
    assert_equal :color_red, Xlsxrb::Adapters::FastExcel::COLOR_ENUM.find(0xFF0000)

    # Extra colors
    assert_equal 0xF0F8FF, Xlsxrb::Adapters::FastExcel::EXTRA_COLORS[:alice_blue]
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel::EXTRA_COLORS[:red]

    # Border enum
    assert_equal 1, Xlsxrb::Adapters::FastExcel::BORDER_ENUM.find(:border_thin)
    assert_equal :border_thin, Xlsxrb::Adapters::FastExcel::BORDER_ENUM.find(1)

    # Align enum
    assert_equal 2, Xlsxrb::Adapters::FastExcel::ALIGN_ENUM.find(:align_center)
    assert_equal :align_center, Xlsxrb::Adapters::FastExcel::ALIGN_ENUM.find(2)
  end

  def test_date_num
    t = Time.utc(2026, 1, 1, 12, 0, 0)
    num = Xlsxrb::Adapters::FastExcel.date_num(t)
    assert_in_delta 46_023.5, num, 0.001

    # Offset test
    num_offset = Xlsxrb::Adapters::FastExcel.date_num(t, 3600)
    assert_in_delta 46_023.5 + (1.0 / 24.0), num_offset, 0.001
  end

  def test_lxw_datetime_and_time
    d = Date.new(2026, 10, 4)
    lxw_d = Xlsxrb::Adapters::FastExcel.lxw_datetime(d)
    assert_equal 2026, lxw_d.year
    assert_equal 10, lxw_d.month
    assert_equal 4, lxw_d.day
    assert_equal 0, lxw_d.hour
    assert_equal 2026, lxw_d[:year]

    t = Time.utc(2026, 10, 4, 15, 30, 45)
    lxw_t = Xlsxrb::Adapters::FastExcel.lxw_time(t)
    assert_equal 15, lxw_t.hour
    assert_equal 30, lxw_t.min
    assert_equal 45, lxw_t.sec
  end

  def test_color_to_hex
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel.color_to_hex(0xFF0000)
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel.color_to_hex("0xFF0000")
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel.color_to_hex("#FF0000")
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel.color_to_hex("FF0000")
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel.color_to_hex("red")
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel.color_to_hex(:red)
    assert_equal 0xF0F8FF, Xlsxrb::Adapters::FastExcel.color_to_hex(:alice_blue)

    assert_raise(ArgumentError) do
      Xlsxrb::Adapters::FastExcel.color_to_hex(:non_existent_color_xyz)
    end

    assert_raise(ArgumentError) do
      Xlsxrb::Adapters::FastExcel.color_to_hex({ invalid: 123 })
    end
  end

  def test_formula_and_url_objects
    fml = Xlsxrb::Adapters::FastExcel::Formula.new("SUM(A1:A10)")
    assert_equal "SUM(A1:A10)", fml.fml
    assert_equal "SUM(A1:A10)", fml.to_s
    assert_equal fml, Xlsxrb::Adapters::FastExcel::Formula.new("SUM(A1:A10)")

    url = Xlsxrb::Adapters::FastExcel::URL.new("https://example.com")
    assert_equal "https://example.com", url.url
    assert_equal "https://example.com", url.to_s
    assert_equal url, Xlsxrb::Adapters::FastExcel::URL.new("https://example.com")
  end

  def test_open_temporary_file
    wb = Xlsxrb::Adapters::FastExcel.open
    assert_instance_of Xlsxrb::Adapters::FastExcel::Workbook, wb
    assert wb.tmp_file
    assert wb.is_open
    assert_not_nil wb.filename
    assert_equal 11, wb.default_format.font_size

    ws = wb.add_worksheet
    ws.append_row(%w[A B C])
    content = wb.read_string
    assert content.start_with?("PK")
    assert_false File.exist?(File.dirname(wb.filename))
  end

  def test_open_with_block
    closed = false
    with_tempfile do |file_path|
      Xlsxrb::Adapters::FastExcel.open(file_path) do |wb|
        ws = wb.add_worksheet("TestSheet")
        ws.append_row([1, 2, 3])
        closed = true
      end
      assert closed
      assert File.exist?(file_path)
      assert File.size(file_path).positive?
    end
  end

  def test_open_existing_non_empty_file_raises_argument_error
    with_tempfile do |file_path|
      File.binwrite(file_path, "dummy content")
      assert_raise(ArgumentError) do
        Xlsxrb::Adapters::FastExcel.open(file_path)
      end
    end
  end

  def test_worksheet_name_validation
    wb = Xlsxrb::Adapters::FastExcel.open

    # Normal names
    ws1 = wb.add_worksheet("Normal Name")
    assert_equal "Normal Name", ws1.name
    assert_equal "Normal Name", ws1[:name]

    # Empty name is allowed
    ws_empty = wb.add_worksheet("")
    assert_equal "", ws_empty.name

    # Length > 31
    assert_raise(ArgumentError) do
      wb.add_worksheet("a" * 32)
    end

    # Invalid characters: [ ] : * ? / \
    assert_raise(ArgumentError) do
      wb.add_worksheet("Test/Sheet")
    end
    assert_raise(ArgumentError) do
      wb.add_worksheet("Test[Sheet]")
    end
    assert_raise(ArgumentError) do
      wb.add_worksheet("Test:Sheet")
    end

    # Apostrophe start or end
    assert_raise(ArgumentError) do
      wb.add_worksheet("'Sheet'")
    end

    # Duplicate name
    assert_raise(ArgumentError) do
      wb.add_worksheet("Normal Name")
    end
    wb.close
  end

  def test_get_worksheet_by_name
    wb = Xlsxrb::Adapters::FastExcel.open
    ws = wb.add_worksheet("Report")
    assert_same ws, wb.get_worksheet_by_name("Report")
    assert_nil wb.get_worksheet_by_name("NonExistent")
    wb.close
  end

  def test_workbook_formats_and_shortcuts
    wb = Xlsxrb::Adapters::FastExcel.open
    bold = wb.bold_format
    assert bold.bold
    assert_equal 1, bold[:bold]

    num_fmt = wb.number_format("#,##0.00")
    assert_equal "#,##0.00", num_fmt.num_format

    custom = wb.add_format(italic: true, font_color: :blue)
    assert custom.italic
    assert_equal 0x0000FF, custom.font_color
    wb.close
  end

  def test_worksheet_write_value_and_last_row_number
    wb = Xlsxrb::Adapters::FastExcel.open
    ws = wb.add_worksheet
    assert_equal(-1, ws.last_row_number)

    ws.write_value(0, 2, "value at (0, 2)")
    assert_equal(0, ws.last_row_number)

    ws.append_row(%w[row1_col0 row1_col1])
    assert_equal(1, ws.last_row_number)

    ws.write_row(3, %w[row3_col0 row3_col1])
    assert_equal(3, ws.last_row_number)

    ws << %w[row4_col0 row4_col1]
    assert_equal(4, ws.last_row_number)

    # Writing to a lower column in the same row does not change or decrease last_row_number
    ws.write_value(0, 0, "value at (0, 0)")
    assert_equal(4, ws.last_row_number)
    wb.close
  end

  def test_constant_memory_mode_restrictions
    wb = Xlsxrb::Adapters::FastExcel.open(constant_memory: true)
    assert wb.constant_memory?

    ws = wb.add_worksheet
    ws.append_row(%w[r0_c0 r0_c1])
    ws.append_row(%w[r1_c0 r1_c1])

    # Attempting to write to row 0 after row 1 in constant_memory mode raises ArgumentError
    err = assert_raise(ArgumentError) do
      ws.write_value(0, 2, "should fail")
    end
    assert_match(/Can not write to saved row in constant_memory mode/, err.message)
    wb.close
  end

  def test_auto_width_calculation
    wb = Xlsxrb::Adapters::FastExcel.open
    ws = wb.add_worksheet
    assert_false ws.auto_width?

    ws.auto_width = true
    assert ws.auto_width?

    ws.append_row(["hello", "world of fast excel"])
    widths = ws.calculated_column_widths
    assert widths[0].positive?
    assert widths[1] > widths[0]

    ws.close
    # After close, padding of 0.2 is added to column width
    wb.close
  end

  def test_format_alignment
    wb = Xlsxrb::Adapters::FastExcel.open
    fmt = wb.add_format
    assert_equal({ horizontal: :align_none, vertical: :align_none }, fmt.align)

    fmt.align = :align_center
    assert_equal({ horizontal: :align_center, vertical: :align_none }, fmt.align)

    fmt.align = "center"
    assert_equal({ horizontal: :align_center, vertical: :align_none }, fmt.align)

    fmt.align = { h: :center, v: :center }
    assert_equal({ horizontal: :align_center, vertical: :align_vertical_center }, fmt.align)

    assert_raise(ArgumentError) do
      fmt.align = :invalid_alignment_val
    end

    assert_raise(ArgumentError) do
      fmt.align = { invalid_key: :center }
    end
    wb.close
  end

  def test_format_borders
    wb = Xlsxrb::Adapters::FastExcel.open
    fmt = wb.add_format

    fmt.bottom = :thin
    assert_equal :border_thin, fmt.bottom
    assert_equal 1, fmt[:bottom]

    fmt.top = "medium"
    assert_equal :border_medium, fmt.top
    assert_equal 2, fmt[:top]

    fmt.border = :thick
    assert_equal :border_thick, fmt.bottom
    assert_equal :border_thick, fmt.top
    assert_equal :border_thick, fmt.left
    assert_equal :border_thick, fmt.right

    fmt.border_bottom_color = :alice_blue
    assert_equal 0xF0F8FF, fmt.bottom_color
    assert_equal 0xF0F8FF, fmt.border_bottom_color

    assert_raise(ArgumentError) do
      fmt.bottom = :unknown_border_style
    end
    wb.close
  end

  def test_format_font_size_validation
    wb = Xlsxrb::Adapters::FastExcel.open
    fmt = wb.add_format
    fmt.font_size = 14
    assert_equal 14, fmt.font_size

    assert_raise(ArgumentError) do
      fmt.font_size = -1
    end
    wb.close
  end

  def test_merge_range_and_autofilter
    with_tempfile do |file_path|
      Xlsxrb::Adapters::FastExcel.open(file_path) do |wb|
        ws = wb.add_worksheet("Sheet1")
        bold = wb.bold_format
        ws.merge_range(0, 0, 1, 3, "Merged Header", bold)
        ws.write_row(2, %w[ID Name Score Active])
        ws.append_row([1, "Alice", 95, true])
        ws.append_row([2, "Bob", 88, false])
        ws.enable_filters!(end_col: 3)
      end

      # Verify contents with Roo
      roo = Roo::Excelx.new(file_path)
      assert_equal "Merged Header", roo.cell(1, 1)
      assert_nil roo.cell(1, 2)
      assert_nil roo.cell(2, 4)
      assert_equal "Alice", roo.cell(4, 2)
      assert_equal 95, roo.cell(4, 3)
    end
  end

  def test_column_formatting_inheritance
    with_tempfile do |file_path|
      Xlsxrb::Adapters::FastExcel.open(file_path) do |wb|
        ws = wb.add_worksheet
        price_fmt = wb.number_format("#,##0.00")
        ws.set_column(1, 1, 15, price_fmt)
        ws.write_value(0, 0, "Item")
        ws.write_value(0, 1, 1234.5) # Cell format is nil, inherits price_fmt
      end

      roo = Roo::Excelx.new(file_path)
      assert_equal "Item", roo.cell(1, 1)
      assert_equal 1234.5, roo.cell(1, 2)
    end
  end

  def test_dates_and_times_writing
    with_tempfile do |file_path|
      date_val = Date.new(2026, 10, 4)
      time_val = Time.utc(2026, 10, 4, 12, 0, 0)
      datetime_val = DateTime.new(2026, 10, 4, 15, 30, 0)

      Xlsxrb::Adapters::FastExcel.open(file_path) do |wb|
        ws = wb.add_worksheet
        date_fmt = wb.number_format("yyyy-mm-dd")
        datetime_fmt = wb.number_format("yyyy-mm-dd hh:mm:ss")

        ws.write_value(0, 0, date_val, date_fmt)
        ws.write_value(0, 1, time_val, datetime_fmt)
        ws.write_value(0, 2, datetime_val, datetime_fmt)
      end

      roo = Roo::Excelx.new(file_path)
      assert_equal date_val, roo.cell(1, 1)
      assert_equal :date, roo.celltype(1, 1)
    end
  end

  def test_formula_writing
    with_tempfile do |file_path|
      Xlsxrb::Adapters::FastExcel.open(file_path) do |wb|
        ws = wb.add_worksheet
        ws.append_row([10, 20, 30])
        ws.append_row([Xlsxrb::Adapters::FastExcel::Formula.new("SUM(A1:C1)")])
      end

      roo = Roo::Excelx.new(file_path)
      assert_equal 10, roo.cell(1, 1)
      assert_equal "SUM(A1:C1)", roo.formula(2, 1)
    end
  end

  def test_hyperlink_writing
    with_tempfile do |file_path|
      Xlsxrb::Adapters::FastExcel.open(file_path) do |wb|
        ws = wb.add_worksheet
        ws.write_url(0, 0, "https://github.com")
        ws.write_url_opt(1, 0, "https://google.com", nil, "Google Search", "Click to visit Google")
      end

      roo = Roo::Excelx.new(file_path)
      assert_equal "https://github.com", roo.cell(1, 1)
      assert_equal "Google Search", roo.cell(2, 1)
    end
  end

  def test_worksheet_struct_fields
    wb = Xlsxrb::Adapters::FastExcel.open
    ws = wb.add_worksheet("OptionsSheet")

    ws.set_right_to_left
    assert_equal 1, ws[:right_to_left]

    ws.center_vertically
    assert_equal 1, ws[:vcenter]
    assert_equal 1, ws[:print_options_changed]

    ws.print_row_col_headers
    assert_equal 1, ws[:print_headers]

    ws.set_margins(1.0, 1.0, 1.5, 1.5)
    assert_equal 1.0, ws[:margin_left]
    assert_equal 1.0, ws[:margin_right]
    assert_equal 1.5, ws[:margin_top]
    assert_equal 1.5, ws[:margin_bottom]

    ws.set_v_pagebreaks([10, 20, 30, 0])
    assert_equal 3, ws[:vbreaks_count]

    ws.set_h_pagebreaks([15, 25, 0])
    assert_equal 2, ws[:hbreaks_count]

    ws.freeze_panes(1, 0)
    wb.close
  end

  def test_native_bridge_to_and_from_xlsxrb
    wb = Xlsxrb::Adapters::FastExcel.open
    ws = wb.add_worksheet("BridgeSheet")
    ws.append_row(%w[Product Qty Price])
    ws.append_row(["Widget", 10, 25.5])

    # Convert to immutable Xlsxrb::Elements::Workbook
    elements_wb = wb.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Workbook, elements_wb
    assert_equal 1, elements_wb.sheets.size
    assert_equal "BridgeSheet", elements_wb.sheets.first.name

    # Convert back via FastExcel.from_xlsxrb
    reloaded_adapter = Xlsxrb::Adapters::FastExcel.from_xlsxrb(elements_wb)
    assert_instance_of Xlsxrb::Adapters::FastExcel::Workbook, reloaded_adapter
    assert_equal 1, reloaded_adapter.sheets.size
    reloaded_sheet = reloaded_adapter.sheets.first
    assert_equal "BridgeSheet", reloaded_sheet.name
    assert_equal 1, reloaded_sheet.last_row_number
    wb.close
  end

  def test_format_to_font_hash
    wb = Xlsxrb::Adapters::FastExcel.open
    fmt = wb.bold_format
    fmt.set_color("red")
    fmt.set_font_size(14)
    fmt.set_italic

    assert_respond_to fmt, :to_font_hash
    f_hash = fmt.to_font_hash
    assert_equal "Calibri", f_hash[:name]
    assert_equal 14, f_hash[:sz]
    assert_equal true, f_hash[:bold]
    assert_equal true, f_hash[:italic]
    assert_equal "FF0000", f_hash[:color]
    wb.close
  end

  def test_print_area_and_repeat_titles
    wb = Xlsxrb::Adapters::FastExcel.open
    ws = wb.add_worksheet("FastPrint")
    ws.print_area(0, 0, 19, 5)
    ws.repeat_rows(0, 1)
    ws.repeat_columns(0, 2)

    xlsxrb_ws = ws.to_xlsxrb
    assert_equal "A1:F20", xlsxrb_ws.print_area
    assert_equal({ rows: "1:2", cols: "A:C" }, xlsxrb_ws.print_titles)
    wb.close
  end

  def test_column_style_index_in_fastexcel
    wb = Xlsxrb::Adapters::FastExcel.open
    ws = wb.add_worksheet("Cols")
    fmt = wb.bold_format
    ws.set_column(0, 0, 30.0, fmt)

    xlsxrb_ws = ws.to_xlsxrb
    col = xlsxrb_ws.columns.first
    assert_not_nil col.style_index
    assert_equal 30.0, col.width
    wb.close
  end

  def test_color_to_hex_fallback_to_xlsxrb_colors
    assert_equal 0x000080, Xlsxrb::Adapters::FastExcel.color_to_hex(:navy)
    assert_equal 0xFF0000, Xlsxrb::Adapters::FastExcel.color_to_hex("#FF0000")
  end
end
