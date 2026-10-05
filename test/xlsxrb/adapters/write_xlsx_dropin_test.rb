# frozen_string_literal: true

require "test_helper"
require "roo"
require "stringio"

class XlsxrbAdaptersWriteXLSXDropinTest < Test::Unit::TestCase
  def test_dropin_constants_and_classes
    target = Xlsxrb::Adapters::WriteXLSX

    assert defined?(target::Workbook)
    assert defined?(target::Worksheet)
    assert defined?(target::Format)
    assert defined?(target::Chart)
    assert defined?(target::Colors)
    assert defined?(target::Utility)
    assert defined?(target::ROW_MAX)
    assert defined?(target::COL_MAX)
    assert defined?(target::STR_MAX)
    assert defined?(target::SHEETNAME_MAX)
    assert defined?(target::MAX_URL_LENGTH)
    assert defined?(target::WriteXLSXInsufficientArgumentError)
    assert defined?(target::WriteXLSXDimensionError)
    assert defined?(target::WriteXLSXOptionParameterError)

    assert_equal 1_048_576, target::ROW_MAX
    assert_equal 16_384, target::COL_MAX
    assert_equal 32_767, target::STR_MAX
    assert_equal 31, target::SHEETNAME_MAX
  end

  def test_dropin_full_workflow
    target = Xlsxrb::Adapters::WriteXLSX

    with_tempfile do |filepath|
      wb = target.new(filepath)
      assert_instance_of target, wb

      # Formats
      title_fmt = wb.add_format(bold: 1, size: 16, color: "blue")
      header_fmt = wb.add_format(bold: 1, bg_color: "silver", align: "center", border: 1)
      currency_fmt = wb.add_format(num_format: "$#,##0.00")
      date_fmt = wb.add_format(num_format: "yyyy-mm-dd")

      # Worksheets
      ws = wb.add_worksheet("Sales Summary")
      assert_instance_of target::Worksheet, ws
      assert_equal "Sales Summary", ws.name

      # Writing Title & Header
      ws.write("A1", "Annual Sales Report", title_fmt)
      ws.write_row("A3", %w[Item Category Price Qty Total Date URL], header_fmt)

      # Writing Rows
      items = [
        ["Laptop", "Tech", 1200.50, 5, "=C4*D4", Date.new(2026, 1, 15), "https://example.com/laptop"],
        ["Desk", "Furniture", 350.00, 10, "=C5*D5", Date.new(2026, 2, 20), "https://example.com/desk"],
        ["Chair", "Furniture", 150.25, 25, "=C6*D6", Date.new(2026, 3, 5), "https://example.com/chair"]
      ]

      items.each_with_index do |(name, cat, price, qty, fml, date, url), idx|
        row = idx + 3
        ws.write_string(row, 0, name)
        ws.write_string(row, 1, cat)
        ws.write_number(row, 2, price, currency_fmt)
        ws.write_number(row, 3, qty)
        ws.write_formula(row, 4, fml, currency_fmt, price * qty)
        ws.write_date_time(row, 5, date, date_fmt)
        ws.write_url(row, 6, url, nil, "Link")
      end

      # Total row
      ws.write("B7", "Grand Total:", header_fmt)
      ws.write_formula("E7", "=SUM(E4:E6)", currency_fmt, 13_261.25)

      # Layout adjustments
      ws.set_column(0, 1, 16)
      ws.set_column(2, 4, 12)
      ws.set_column(5, 6, 15)
      ws.freeze_panes(3, 0)
      ws.autofilter("A3:G6")

      # Add chart
      chart = wb.add_chart(type: :column)
      assert_instance_of target::Chart::Column, chart
      chart.add_series(
        categories: "='Sales Summary'!$A$4:$A$6",
        values: "='Sales Summary'!$E$4:$E$6",
        name: "Total Sales"
      )
      chart.set_title(name: "Sales by Item")
      ws.insert_chart("I3", chart)

      wb.close

      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      # Verify content using Roo
      roo = Roo::Excelx.new(filepath)
      assert_equal ["Sales Summary"], roo.sheets
      assert_equal "Annual Sales Report", roo.cell(1, 1)
      assert_equal "Item", roo.cell(3, 1)
      assert_equal "Total", roo.cell(3, 5)
      assert_equal "Laptop", roo.cell(4, 1)
      assert_equal 1200.50, roo.cell(4, 3)
      assert_equal 5, roo.cell(4, 4)
      assert_equal "C4*D4", roo.formula(4, 5)
      assert_equal "SUM(E4:E6)", roo.formula(7, 5)
      assert_equal "Link", roo.cell(4, 7)
    end
  end

  def test_dropin_stringio_and_read_string
    target = Xlsxrb::Adapters::WriteXLSX

    # StringIO target
    io = StringIO.new
    wb = target.new(io)
    ws = wb.add_worksheet("StreamSheet")
    ws.write(0, 0, "Stream Content")
    wb.close

    assert io.string.bytesize.positive?

    # In-memory read_string
    wb_mem = target.new
    ws_mem = wb_mem.add_worksheet("MemorySheet")
    ws_mem.write("A1", "Memory Content")
    data = wb_mem.read_string
    assert data.is_a?(String)
    assert data.bytesize.positive?
  end

  def test_dropin_bridge_methods
    target = Xlsxrb::Adapters::WriteXLSX

    wb = target.new
    ws = wb.add_worksheet("BridgeSheet")
    ws.write(0, 0, "Bridge Value")

    xlsxrb_wb = target.to_xlsxrb(wb)
    assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb

    wb_back = target.from_xlsxrb(xlsxrb_wb)
    assert_instance_of target::Workbook, wb_back
    assert_equal 1, wb_back.sheets.size
    assert_equal "BridgeSheet", wb_back.sheets.first.name
    assert_equal "Bridge Value", wb_back.sheets.first.table[[0, 0]].value
  end
end
