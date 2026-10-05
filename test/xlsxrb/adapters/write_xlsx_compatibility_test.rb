# frozen_string_literal: true

require "test_helper"
require "write_xlsx"
require "roo"

class XlsxrbAdaptersWriteXLSXCompatibilityTest < Test::Unit::TestCase
  def test_side_by_side_primitive_values
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        # 1. Generate with official WriteXLSX
        orig_wb = ::WriteXLSX.new(orig_path)
        orig_ws = orig_wb.add_worksheet("Primitives")
        orig_ws.write_row(0, 0, %w[Title Count Ratio Active Date])
        orig_ws.write_row(1, 0, ["Widget", 100, 3.14159, "Active", "2026-01-15T00:00:00"])
        orig_ws.write_row(2, 0, ["Gadget", 0, -0.5, "Inactive", "2026-12-31T00:00:00"])
        orig_ws.write_boolean(3, 0, true)
        orig_ws.write_boolean(3, 1, false)
        orig_wb.close

        # 2. Generate with Xlsxrb WriteXLSX adapter
        adapt_wb = Xlsxrb::Adapters::WriteXLSX.new(adapt_path)
        adapt_ws = adapt_wb.add_worksheet("Primitives")
        adapt_ws.write_row(0, 0, %w[Title Count Ratio Active Date])
        adapt_ws.write_row(1, 0, ["Widget", 100, 3.14159, "Active", "2026-01-15T00:00:00"])
        adapt_ws.write_row(2, 0, ["Gadget", 0, -0.5, "Inactive", "2026-12-31T00:00:00"])
        adapt_ws.write_boolean(3, 0, true)
        adapt_ws.write_boolean(3, 1, false)
        adapt_wb.close

        # 3. Compare with Roo
        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        assert_equal orig_roo.sheets, adapt_roo.sheets

        (1..3).each do |row|
          (1..5).each do |col|
            orig_val = orig_roo.cell(row, col)
            adapt_val = adapt_roo.cell(row, col)

            if orig_val.is_a?(Float) && adapt_val.is_a?(Float)
              assert_in_delta orig_val, adapt_val, 0.0001, "Mismatch at row #{row}, col #{col}"
            else
              assert_equal orig_val, adapt_val, "Mismatch at row #{row}, col #{col}"
            end
          end
        end
      end
    end
  end

  def test_side_by_side_formulas
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        orig_wb = ::WriteXLSX.new(orig_path)
        orig_ws = orig_wb.add_worksheet("Formulas")
        orig_ws.write_row(0, 0, [10, 20, 30])
        orig_ws.write_formula(1, 0, "=SUM(A1:C1)", nil, 60)
        orig_ws.write_formula(2, 0, "=AVERAGE(A1:C1)", nil, 20.0)
        orig_wb.close

        adapt_wb = Xlsxrb::Adapters::WriteXLSX.new(adapt_path)
        adapt_ws = adapt_wb.add_worksheet("Formulas")
        adapt_ws.write_row(0, 0, [10, 20, 30])
        adapt_ws.write_formula(1, 0, "=SUM(A1:C1)", nil, 60)
        adapt_ws.write_formula(2, 0, "=AVERAGE(A1:C1)", nil, 20.0)
        adapt_wb.close

        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        assert_equal "SUM(A1:C1)", orig_roo.formula(2, 1).sub(/^=/, "")
        assert_equal "SUM(A1:C1)", adapt_roo.formula(2, 1).sub(/^=/, "")
        assert_equal "AVERAGE(A1:C1)", orig_roo.formula(3, 1).sub(/^=/, "")
        assert_equal "AVERAGE(A1:C1)", adapt_roo.formula(3, 1).sub(/^=/, "")
      end
    end
  end

  def test_side_by_side_multiple_sheets_and_merges
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        # Upstream
        orig_wb = ::WriteXLSX.new(orig_path)
        orig_ws1 = orig_wb.add_worksheet("Sheet Uno")
        fmt1 = orig_wb.add_format(bold: 1, align: "center")
        orig_ws1.merge_range("A1:C1", "Merged Title", fmt1)
        orig_ws1.write("A2", "Subdata")
        orig_ws2 = orig_wb.add_worksheet("Sheet Dos")
        orig_ws2.write("B2", "Second Sheet Content")
        orig_wb.close

        # Adapter
        adapt_wb = Xlsxrb::Adapters::WriteXLSX.new(adapt_path)
        adapt_ws1 = adapt_wb.add_worksheet("Sheet Uno")
        fmt2 = adapt_wb.add_format(bold: 1, align: "center")
        adapt_ws1.merge_range("A1:C1", "Merged Title", fmt2)
        adapt_ws1.write("A2", "Subdata")
        adapt_ws2 = adapt_wb.add_worksheet("Sheet Dos")
        adapt_ws2.write("B2", "Second Sheet Content")
        adapt_wb.close

        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        assert_equal ["Sheet Uno", "Sheet Dos"], orig_roo.sheets
        assert_equal ["Sheet Uno", "Sheet Dos"], adapt_roo.sheets

        orig_roo.default_sheet = "Sheet Uno"
        adapt_roo.default_sheet = "Sheet Uno"
        assert_equal "Merged Title", orig_roo.cell(1, 1)
        assert_equal "Merged Title", adapt_roo.cell(1, 1)
        assert_equal "Subdata", orig_roo.cell(2, 1)
        assert_equal "Subdata", adapt_roo.cell(2, 1)

        orig_roo.default_sheet = "Sheet Dos"
        adapt_roo.default_sheet = "Sheet Dos"
        assert_equal "Second Sheet Content", orig_roo.cell(2, 2)
        assert_equal "Second Sheet Content", adapt_roo.cell(2, 2)
      end
    end
  end

  def test_side_by_side_utility_methods
    # Create upstream workbook and adapter workbook to call utility methods
    orig_wb = ::WriteXLSX.new("dummy_orig.xlsx")
    adapt_wb = Xlsxrb::Adapters::WriteXLSX.new("dummy_adapt.xlsx")

    # Cell to rowcol
    assert_equal orig_wb.xl_cell_to_rowcol("A1"), adapt_wb.xl_cell_to_rowcol("A1")
    assert_equal orig_wb.xl_cell_to_rowcol("B3"), adapt_wb.xl_cell_to_rowcol("B3")
    assert_equal orig_wb.xl_cell_to_rowcol("$A$1"), adapt_wb.xl_cell_to_rowcol("$A$1")
    assert_equal orig_wb.xl_cell_to_rowcol("Z100"), adapt_wb.xl_cell_to_rowcol("Z100")
    assert_equal orig_wb.xl_cell_to_rowcol("XFD1048576"), adapt_wb.xl_cell_to_rowcol("XFD1048576")

    # Rowcol to cell
    assert_equal orig_wb.xl_rowcol_to_cell(0, 0), adapt_wb.xl_rowcol_to_cell(0, 0)
    assert_equal orig_wb.xl_rowcol_to_cell(2, 1), adapt_wb.xl_rowcol_to_cell(2, 1)
    assert_equal orig_wb.xl_rowcol_to_cell(0, 0, true, true), adapt_wb.xl_rowcol_to_cell(0, 0, true, true)
    assert_equal orig_wb.xl_rowcol_to_cell(99, 25), adapt_wb.xl_rowcol_to_cell(99, 25)

    # Col to name
    assert_equal orig_wb.xl_col_to_name(0, false), adapt_wb.xl_col_to_name(0)
    assert_equal orig_wb.xl_col_to_name(25, false), adapt_wb.xl_col_to_name(25)
    assert_equal orig_wb.xl_col_to_name(26, false), adapt_wb.xl_col_to_name(26)
    assert_equal orig_wb.xl_col_to_name(16_383, false), adapt_wb.xl_col_to_name(16_383)
    assert_equal orig_wb.xl_col_to_name(0, true), adapt_wb.xl_col_to_name(0, true)

    # Range
    assert_equal orig_wb.xl_range(0, 1, 0, 1), adapt_wb.xl_range(0, 1, 0, 1)
    assert_equal orig_wb.xl_range(0, 10, 0, 5), adapt_wb.xl_range(0, 10, 0, 5)

    # Range formula
    orig_ws = orig_wb.add_worksheet
    adapt_ws = adapt_wb.add_worksheet
    assert_equal orig_ws.xl_range_formula("Sheet1", 0, 1, 0, 1), adapt_ws.xl_range_formula("Sheet1", 0, 1, 0, 1)
    assert_equal orig_ws.xl_range_formula("My Sheet", 0, 1, 0, 1), adapt_ws.xl_range_formula("My Sheet", 0, 1, 0, 1)
  end
end
