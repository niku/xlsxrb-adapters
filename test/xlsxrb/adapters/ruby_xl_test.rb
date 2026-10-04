# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersRubyXLTest < Test::Unit::TestCase
  def test_version
    refute_nil ::Xlsxrb::Adapters::VERSION
  end

  def test_workbook_initialization_defaults
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    assert_equal 1, wb.worksheets.size
    assert_equal "Sheet1", wb[0].sheet_name
    assert_equal wb[0], wb["Sheet1"]
    assert_equal wb[0], wb[:Sheet1]
  end

  def test_workbook_add_worksheet
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws2 = wb.add_worksheet("CustomSheet")
    assert_equal 2, wb.worksheets.size
    assert_equal "CustomSheet", ws2.sheet_name
    assert_equal ws2, wb["CustomSheet"]
  end

  def test_worksheet_add_cell_and_access
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]

    cell = ws.add_cell(1, 2, "Test Value")
    assert_equal 1, cell.row
    assert_equal 2, cell.column
    assert_equal "Test Value", cell.value
    assert_equal "Test Value", cell.raw_value

    row = ws[1]
    refute_nil row
    assert_equal cell, row[2]
    assert_nil row[0]
    assert_nil row[1]
    assert_equal cell, ws[1][2]
    assert_equal cell, ws.sheet_data[1][2]
  end

  def test_worksheet_sparse_and_each
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "Row0")
    ws.add_cell(2, 0, "Row2")

    assert_nil ws[1]
    rows = ws.each.to_a
    assert_equal 3, rows.size
    assert_equal "Row0", rows[0][0].value
    assert_nil rows[1]
    assert_equal "Row2", rows[2][0].value
  end

  def test_cell_change_contents
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    cell = ws.add_cell(0, 0, "Initial")
    assert_equal "Initial", cell.value

    cell.change_contents("Updated")
    assert_equal "Updated", cell.value
    assert_equal "Updated", cell.raw_value

    cell.change_contents(42)
    assert_equal 42, cell.value
    assert_equal 42, cell.raw_value

    cell.change_contents("FormulaVal", "SUM(A1:B1)")
    assert_equal "FormulaVal", cell.value
    assert_equal "SUM(A1:B1)", cell.formula.expression
    assert_equal "SUM(A1:B1)", cell.formula.to_s

    # Changing contents without formula keeps existing formula
    cell.change_contents(99)
    assert_equal 99, cell.value
    assert_equal "SUM(A1:B1)", cell.formula.expression

    # remove_formula
    cell.remove_formula
    assert_nil cell.formula
  end

  def test_cell_types_support
    today = Date.new(2026, 9, 26)
    now = Time.new(2026, 9, 26, 12, 0, 0)

    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]

    c_str = ws.add_cell(0, 0, "String")
    c_int = ws.add_cell(0, 1, 100)
    c_flt = ws.add_cell(0, 2, 99.5)
    c_true = ws.add_cell(0, 3, true)
    c_false = ws.add_cell(0, 4, false)
    c_date = ws.add_cell(0, 5, today)
    c_time = ws.add_cell(0, 6, now)
    c_nil = ws.add_cell(0, 7, nil)
    c_form = ws.add_cell(0, 8, 30, "A1+B1")

    assert_equal "String", c_str.value
    assert_equal 100, c_int.value
    assert_equal 99.5, c_flt.value
    assert_equal true, c_true.value
    assert_equal false, c_false.value
    assert_equal today, c_date.value
    assert_equal now.to_datetime, c_time.value
    assert_nil c_nil.value
    assert_equal 30, c_form.value
    assert_equal "A1+B1", c_form.formula.to_s
  end

  def test_to_xlsxrb_and_from_xlsxrb_roundtrip
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.sheet_name = "Roundtrip"
    ws.add_cell(0, 0, "Greeting")
    ws.add_cell(0, 1, 12_345)
    ws.add_cell(1, 0, "", "SUM(B1:B2)")

    # Convert to xlsxrb Elements::Workbook
    xlsxrb_wb = wb.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
    assert_equal 1, xlsxrb_wb.sheets.size
    assert_equal "Roundtrip", xlsxrb_wb.sheets.first.name

    doc_sheet = xlsxrb_wb.sheets.first
    assert_equal "Greeting", doc_sheet["A1"].value
    assert_equal 12_345, doc_sheet["B1"].value
    assert_equal "SUM(B1:B2)", doc_sheet["A2"].formula.expression

    # Convert back from xlsxrb
    restored_wb = Xlsxrb::Adapters::RubyXL.from_xlsxrb(xlsxrb_wb)
    assert_instance_of Xlsxrb::Adapters::RubyXL::Workbook, restored_wb
    assert_equal 1, restored_wb.worksheets.size
    assert_equal "Roundtrip", restored_wb[0].sheet_name
    assert_equal "Greeting", restored_wb[0][0][0].value
    assert_equal 12_345, restored_wb[0][0][1].value
    assert_equal "SUM(B1:B2)", restored_wb[0][1][0].formula.expression
  end

  def test_write_and_parse_roundtrip
    with_tempfile do |filepath|
      wb = Xlsxrb::Adapters::RubyXL::Workbook.new
      ws = wb[0]
      ws.sheet_name = "SavedSheet"
      ws.add_cell(0, 0, "Hello Saved")
      ws.add_cell(0, 1, 888)
      ws.add_cell(1, 0, "", "SUM(A1:B1)")
      wb.write(filepath)

      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      # Parse via Adapter Parser
      parsed = Xlsxrb::Adapters::RubyXL::Parser.parse(filepath)
      assert_instance_of Xlsxrb::Adapters::RubyXL::Workbook, parsed
      assert_equal 1, parsed.worksheets.size
      assert_equal "SavedSheet", parsed[0].sheet_name
      assert_equal "Hello Saved", parsed[0][0][0].value
      assert_equal 888, parsed[0][0][1].value
      assert_equal "SUM(A1:B1)", parsed[0][1][0].formula.expression
    end
  end

  def test_stream_and_parse_buffer_roundtrip
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.sheet_name = "BufferSheet"
    ws.add_cell(0, 0, "BufferData")
    ws.add_cell(0, 1, 999)

    io = wb.stream
    assert_instance_of StringIO, io
    assert io.string.bytesize.positive?

    parsed = Xlsxrb::Adapters::RubyXL::Parser.parse_buffer(io)
    assert_instance_of Xlsxrb::Adapters::RubyXL::Workbook, parsed
    assert_equal "BufferSheet", parsed[0].sheet_name
    assert_equal "BufferData", parsed[0][0][0].value
    assert_equal 999, parsed[0][0][1].value

    # Also test parse_buffer with binary String directly
    parsed_from_string = Xlsxrb::Adapters::RubyXL::Parser.parse_buffer(io.string)
    assert_equal "BufferData", parsed_from_string[0][0][0].value
  end

  def test_worksheet_add_cell_overwrite_behavior
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    cell1 = ws.add_cell(0, 0, "Original")

    # overwrite: false should return existing cell and leave contents unchanged
    cell2 = ws.add_cell(0, 0, "NewValue", nil, false)
    assert_equal cell1, cell2
    assert_equal "Original", cell1.value

    # overwrite: true (default) should update contents
    cell3 = ws.add_cell(0, 0, "NewValue", nil, true)
    assert_equal "NewValue", cell3.value
  end

  def test_row_cells_sparse_handling
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 3, "Col3")

    row = ws[0]
    assert_equal 4, row.size
    assert_nil row[0]
    assert_nil row[1]
    assert_nil row[2]
    assert_equal "Col3", row[3].value
    assert_equal([nil, nil, nil, "Col3"], row.cells.map { |c| c&.value })
  end

  def test_multiple_worksheets_roundtrip
    with_tempfile do |filepath|
      wb = Xlsxrb::Adapters::RubyXL::Workbook.new
      ws1 = wb[0]
      ws1.sheet_name = "Summary"
      ws1.add_cell(0, 0, "Summary Data")

      ws2 = wb.add_worksheet("Details")
      ws2.add_cell(0, 0, "Details Data")

      wb.write(filepath)

      parsed = Xlsxrb::Adapters::RubyXL::Parser.parse(filepath)
      assert_equal 2, parsed.worksheets.size
      assert_equal "Summary", parsed[0].sheet_name
      assert_equal "Details", parsed[1].sheet_name
      assert_equal "Summary Data", parsed["Summary"][0][0].value
      assert_equal "Details Data", parsed["Details"][0][0].value
    end
  end

  def test_formula_object_behavior
    f1 = Xlsxrb::Adapters::RubyXL::Formula.new("SUM(A1:A10)")
    f2 = Xlsxrb::Adapters::RubyXL::Formula.new("SUM(A1:A10)")
    f3 = Xlsxrb::Adapters::RubyXL::Formula.new("COUNT(A1:A10)")

    assert_equal f1, f2
    refute_equal f1, f3
    assert_equal f1, "SUM(A1:A10)"
    assert_equal "SUM(A1:A10)", f1.to_s
    assert_equal "SUM(A1:A10)", f1.to_str
    assert_equal f1.hash, f2.hash
  end

  def test_worksheet_to_xlsxrb
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.sheet_name = "DirectWorksheet"
    ws.add_cell(0, 0, "DirectCell")

    xlsxrb_ws = ws.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Worksheet, xlsxrb_ws
    assert_equal "DirectWorksheet", xlsxrb_ws.name
    assert_equal "DirectCell", xlsxrb_ws["A1"].value
  end

  def test_worksheet_data_validations_roundtrip
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_validation_list("A1:A10", %w[Apple Banana Orange])

    xlsxrb_wb = wb.to_xlsxrb
    xlsxrb_ws = xlsxrb_wb.sheet(0)
    assert_equal 1, xlsxrb_ws.data_validations.size
    assert_equal "A1:A10", xlsxrb_ws.data_validations.first[:sqref]
    assert_equal "list", xlsxrb_ws.data_validations.first[:type]

    reloaded_wb = Xlsxrb::Adapters::RubyXL::Workbook.from_xlsxrb(xlsxrb_wb)
    assert_equal 1, reloaded_wb[0].data_validations.size
    assert_equal "A1:A10", reloaded_wb[0].data_validations.first[:sqref]
  end

  def test_dimension_and_bounds
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(3, 1, "B4")
    ws.add_cell(4, 1, "B5")

    assert_equal "B4:B5", ws.dimension
    assert_equal 4, ws.first_row
    assert_equal 2, ws.first_column
    assert_equal 2, ws.first_col
    assert_equal 5, ws.last_row
    assert_equal 2, ws.last_column
    assert_equal 2, ws.last_col
    assert_equal false, ws.date1904?
    assert_equal false, wb.date1904?
  end

  def test_date1904_system
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    wb.date1904 = true
    ws = wb[0]
    ws.add_cell(0, 0, Date.new(2026, 1, 1))

    assert_equal true, wb.date1904?
    assert_equal true, ws.date1904?

    xlsxrb_wb = wb.to_xlsxrb
    assert_equal true, xlsxrb_wb.sheets[0].date1904?

    wb_back = Xlsxrb::Adapters::RubyXL::Workbook.from_xlsxrb(xlsxrb_wb)
    assert_equal true, wb_back.date1904?
    assert_equal true, wb_back[0].date1904?
  end

  def test_each_row_values
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "A1")
    ws.add_cell(0, 1, "B1")
    ws.add_cell(1, 0, "A2")

    values = ws.each_row_values.to_a
    assert_equal 2, values.size
    assert_equal %w[A1 B1], values[0]
    assert_equal ["A2"], values[1]
  end
end
