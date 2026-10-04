# frozen_string_literal: true

require "test_helper"
require "roo"

class XlsxrbAdaptersFastExcelDropinTest < Test::Unit::TestCase
  def test_dropin_constant_definitions
    target_mod = Xlsxrb::Adapters::FastExcel

    assert defined?(target_mod::Workbook)
    assert defined?(target_mod::Worksheet)
    assert defined?(target_mod::Format)
    assert defined?(target_mod::Formula)
    assert defined?(target_mod::URL)
    assert defined?(target_mod::ERROR_ENUM)
    assert defined?(target_mod::COLOR_ENUM)
    assert defined?(target_mod::BORDER_ENUM)
    assert defined?(target_mod::ALIGN_ENUM)
    assert defined?(target_mod::EXTRA_COLORS)
    assert_equal 8.43, target_mod::DEF_COL_WIDTH
    assert_equal 86_400.0, target_mod::XLSX_DATE_DAY
    assert_equal 25_569.0, target_mod::XLSX_DATE_EPOCH_DIFF
  end

  def test_dropin_workbook_creation_and_block_open
    target = Xlsxrb::Adapters::FastExcel

    with_tempfile do |filepath|
      target.open(filepath) do |wb|
        ws = wb.add_worksheet("SalesReport")
        header_fmt = wb.add_format(bold: true, bg_color: :navy, font_color: :white, align: :center)
        currency_fmt = wb.number_format("$#,##0.00")

        ws.append_row(%w[Item Qty UnitPrice Subtotal], header_fmt)
        ws.append_row(["Keyboard", 2, 45.0, target::Formula.new("B2*C2")], [nil, nil, currency_fmt, currency_fmt])
        ws.append_row(["Mouse", 5, 20.0, target::Formula.new("B3*C3")], [nil, nil, currency_fmt, currency_fmt])
        ws.append_row(["Monitor", 1, 300.0, target::Formula.new("B4*C4")], [nil, nil, currency_fmt, currency_fmt])
        ws.append_row(["Total", target::Formula.new("SUM(B2:B4)"), nil, target::Formula.new("SUM(D2:D4)")], [header_fmt, nil, nil, currency_fmt])

        ws.set_column(0, 0, 18)
        ws.set_column(1, 1, 10)
        ws.set_column(2, 3, 15)
        ws.freeze_panes(1, 0)
        ws.autofilter(0, 0, 3, 3)
      end

      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      roo = Roo::Excelx.new(filepath)
      assert_equal ["SalesReport"], roo.sheets
      assert_equal "Item", roo.cell(1, 1)
      assert_equal "Keyboard", roo.cell(2, 1)
      assert_equal 2, roo.cell(2, 2)
      assert_equal 45.0, roo.cell(2, 3)
      assert_equal "B2*C2", roo.formula(2, 4)
      assert_equal "SUM(B2:B4)", roo.formula(5, 2)
      assert_equal "SUM(D2:D4)", roo.formula(5, 4)
    end
  end

  def test_dropin_in_memory_read_string_workflow
    target = Xlsxrb::Adapters::FastExcel

    wb = target.open
    ws = wb.add_worksheet
    ws.auto_width = true

    ws << %w[City Population Established]
    ws << ["Tokyo", 14_000_000, Date.new(1603, 1, 1)]
    ws << ["New York", 8_400_000, Date.new(1624, 1, 1)]
    ws << ["London", 9_000_000, Date.new(47, 1, 1)]

    data = wb.read_string
    refute_nil data
    assert data.bytesize.positive?
    assert data.start_with?("PK")

    # Read back binary buffer using Tempfile
    Tempfile.create(["in_memory_report", ".xlsx"]) do |tmp|
      tmp.binmode
      tmp.write(data)
      tmp.flush

      roo = Roo::Excelx.new(tmp.path)
      assert_equal "City", roo.cell(1, 1)
      assert_equal "Tokyo", roo.cell(2, 1)
      assert_equal 14_000_000, roo.cell(2, 2)
      assert_equal "London", roo.cell(4, 1)
    end
  end

  def test_dropin_constant_memory_protection
    target = Xlsxrb::Adapters::FastExcel

    with_tempfile do |filepath|
      wb = target.open(filepath, constant_memory: true)
      assert wb.constant_memory?

      ws = wb.add_worksheet
      ws.write_value(0, 0, "Row 0")
      ws.write_value(2, 0, "Row 2")
      assert_equal 2, ws.last_row_number

      assert_raise(ArgumentError) do
        ws.write_value(1, 0, "Row 1 back-write attempt")
      end

      wb.close
    end
  end

  def test_dropin_roundtrip_xlsxrb_bridge
    target = Xlsxrb::Adapters::FastExcel

    with_tempfile do |filepath|
      wb = target.open(filepath)
      ws = wb.add_worksheet("Inventory")
      ws.append_row(%w[SKU Name Price])
      ws.append_row(["SKU-1", "Gadget", 19.99])
      ws.append_row(["SKU-2", "Gizmo", 29.99])

      xlsxrb_wb = wb.to_xlsxrb
      assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
      assert_equal "Inventory", xlsxrb_wb.sheets.first.name

      roundtrip_wb = target.from_xlsxrb(xlsxrb_wb)
      assert_instance_of Xlsxrb::Adapters::FastExcel::Workbook, roundtrip_wb
      assert_equal 1, roundtrip_wb.sheets.size
      roundtrip_sheet = roundtrip_wb.sheets.first
      assert_equal "Inventory", roundtrip_sheet.name
      assert_equal 2, roundtrip_sheet.last_row_number

      data = roundtrip_wb.read_string
      assert data.start_with?("PK")

      wb.close
    end
  end
end
