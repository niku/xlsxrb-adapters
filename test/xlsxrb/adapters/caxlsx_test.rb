# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCaxlsxTest < Test::Unit::TestCase
  def test_version_and_module_structure
    refute_nil ::Xlsxrb::Adapters::VERSION
    assert_equal ::Xlsxrb::Adapters::Caxlsx, ::Xlsxrb::Adapters::Axlsx
    assert_equal ::Xlsxrb::Adapters::Caxlsx::Package, ::Xlsxrb::Adapters::Axlsx::Package
  end

  def test_package_initialization_and_workbook
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    refute_nil pkg.workbook
    assert_equal 0, pkg.workbook.worksheets.size
  end

  def test_workbook_add_worksheet_with_name_and_block
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    wb = pkg.workbook

    ws1 = wb.add_worksheet(name: "FirstSheet")
    assert_equal "FirstSheet", ws1.name
    assert_equal 1, wb.worksheets.size
    assert_equal ws1, wb.worksheets.first

    yielded_sheet = nil
    ws2 = wb.add_worksheet(name: "SecondSheet") do |sheet|
      yielded_sheet = sheet
      sheet.add_row(%w[Header1 Header2])
    end

    assert_equal ws2, yielded_sheet
    assert_equal 2, wb.worksheets.size
    assert_equal "SecondSheet", ws2.name
    assert_equal 1, ws2.rows.size
    assert_equal "Header1", ws2.rows[0].cells[0].value
  end

  def test_worksheet_add_row_with_various_types
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    pkg.workbook.escape_formulas = false
    ws = pkg.workbook.add_worksheet(name: "DataTypes")

    now = Time.now
    today = Date.today

    row = ws.add_row(["Text", 12_345, 99.95, true, false, today, now, nil, "=SUM(B1:C1)"])
    assert_equal 1, ws.rows.size
    assert_equal 9, row.cells.size

    c_str, c_int, c_flt, c_true, c_false, c_date, c_time, c_nil, c_form = row.cells

    assert_equal :string, c_str.type
    assert_equal "Text", c_str.value

    assert_equal :integer, c_int.type
    assert_equal 12_345, c_int.value

    assert_equal :float, c_flt.type
    assert_in_delta 99.95, c_flt.value, 0.001

    assert_equal :boolean, c_true.type
    assert_equal 1, c_true.value

    assert_equal :boolean, c_false.type
    assert_equal 0, c_false.value

    assert_equal :date, c_date.type
    assert_equal today, c_date.value

    assert_equal :time, c_time.type
    assert_equal now, c_time.value

    assert_nil c_nil.value

    assert c_form.is_formula?
    assert_equal "=SUM(B1:C1)", c_form.value
    assert_nil c_form.formula_value
  end

  def test_worksheet_indexing_and_coordinate_access
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Coords")

    ws.add_row([10, 20, 30])
    ws.add_row([40, 50, 60])

    # Index by row integer
    assert_equal 10, ws[0][0].value
    assert_equal 20, ws[0][1].value
    assert_equal 60, ws[1][2].value

    # Index by single cell name
    assert_equal 10, ws["A1"].value
    assert_equal 20, ws["B1"].value
    assert_equal 50, ws["B2"].value

    # Index by cell range
    range = ws["A1:B2"]
    assert_equal 4, range.size
    assert_equal [10, 20, 40, 50], range.map(&:value)

    # Coordinate references
    cell = ws["B2"]
    assert_equal "B2", cell.r
    assert_equal "$B$2", cell.r_abs
    assert_equal [1, 1], cell.pos
  end

  def test_cell_formula_options
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    pkg.workbook.escape_formulas = false
    ws = pkg.workbook.add_worksheet

    cell = ws.add_row(["=1+1"]).cells.first
    assert cell.is_formula?
    assert_equal "=1+1", cell.value

    # When escape_formulas is true on worksheet, formula prefix is escaped
    ws.escape_formulas = true
    cell_esc = ws.add_row(["=SUM(A1:B1)"]).cells.first
    refute cell_esc.is_formula?
    assert_equal "=SUM(A1:B1)", cell_esc.value
  end

  def test_package_to_stream
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "StreamTest")
    ws.add_row(["Alpha", "Beta", 100])

    stream = pkg.to_stream
    assert_instance_of StringIO, stream
    bytes = stream.read
    refute_empty bytes
    # Zip magic number PK\x03\x04
    assert_equal "PK\x03\x04", bytes[0, 4]
  end

  def test_package_serialize_to_file
    with_tempfile do |filepath|
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      ws = pkg.workbook.add_worksheet(name: "FileTest")
      ws.add_row(["Row1", "Data", 42])

      assert pkg.serialize(filepath)
      assert File.exist?(filepath)
      assert File.size(filepath).positive?
    end
  end

  def test_native_bridge_to_xlsxrb
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    pkg.workbook.escape_formulas = false
    ws = pkg.workbook.add_worksheet(name: "BridgeSheet")
    ws.add_row(%w[Item Qty Price])
    ws.add_row(["Apple", 5, 1.25])
    ws.add_row(["Banana", 10, 0.75])
    ws.add_row(["Total", "=SUM(B2:B3)", "=SUM(C2:C3)"])

    # Mutable adapter -> Immutable xlsxrb workbook
    xlsxrb_wb = pkg.to_xlsxrb
    assert_instance_of ::Xlsxrb::Elements::Workbook, xlsxrb_wb
    assert_equal 1, xlsxrb_wb.sheets.size
    assert_equal "BridgeSheet", xlsxrb_wb.sheets.first.name

    xlsxrb_ws = xlsxrb_wb.sheets.first
    assert_equal 4, xlsxrb_ws.rows.size
    assert_equal "Apple", xlsxrb_ws.rows[1].cells[0].value
    assert_equal 5, xlsxrb_ws.rows[1].cells[1].value
    assert_in_delta 1.25, xlsxrb_ws.rows[1].cells[2].value, 0.001
    assert_equal "SUM(B2:B3)", xlsxrb_ws.rows[3].cells[1].formula.expression

    # Verify native xlsxrb reads the serialized file cleanly
    with_tempfile do |path|
      assert pkg.serialize(path)
      read_wb = ::Xlsxrb.read(path).load
      assert_equal "BridgeSheet", read_wb.sheets.first.name
      assert_equal 4, read_wb.sheets.first.rows.size
      assert_equal "Apple", read_wb.sheets.first.rows[1].cells[0].value
      assert_equal 5, read_wb.sheets.first.rows[1].cells[1].value
      assert_in_delta 1.25, read_wb.sheets.first.rows[1].cells[2].value, 0.001
    end
  end

  def test_package_validation
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "ValidationSheet")
    ws.add_row(["Valid Data"])

    # Schema validation returns an Array of errors (empty when valid)
    errors = pkg.validate
    assert_instance_of Array, errors
  end
end
