# frozen_string_literal: true

require "test_helper"
require "roo"

class XlsxrbAdaptersFastExcelCompatibilityTest < Test::Unit::TestCase
  def test_side_by_side_primitive_values
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        # 1. Generate with official FastExcel
        orig_wb = ::FastExcel.open(orig_path)
        orig_ws = orig_wb.add_worksheet("Primitives")
        orig_ws.append_row(%w[Title Count Ratio Active Date])
        orig_ws.append_row(["Widget", 100, 3.14159, true, Date.new(2026, 1, 15)])
        orig_ws.append_row(["Gadget", 0, -0.5, false, Date.new(2026, 12, 31)])
        orig_wb.close

        # 2. Generate with Xlsxrb adapter
        adapt_wb = Xlsxrb::Adapters::FastExcel.open(adapt_path)
        adapt_ws = adapt_wb.add_worksheet("Primitives")
        adapt_ws.append_row(%w[Title Count Ratio Active Date])
        adapt_ws.append_row(["Widget", 100, 3.14159, true, Date.new(2026, 1, 15)])
        adapt_ws.append_row(["Gadget", 0, -0.5, false, Date.new(2026, 12, 31)])
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
        orig_wb = ::FastExcel.open(orig_path)
        orig_ws = orig_wb.add_worksheet
        orig_ws.append_row([10, 20, 30])
        orig_ws.append_row([::FastExcel::Formula.new("SUM(A1:C1)")])
        orig_ws.write_formula_num(2, 0, "AVERAGE(A1:C1)", nil, 20.0)
        orig_wb.close

        adapt_wb = Xlsxrb::Adapters::FastExcel.open(adapt_path)
        adapt_ws = adapt_wb.add_worksheet
        adapt_ws.append_row([10, 20, 30])
        adapt_ws.append_row([Xlsxrb::Adapters::FastExcel::Formula.new("SUM(A1:C1)")])
        adapt_ws.write_formula_num(2, 0, "AVERAGE(A1:C1)", nil, 20.0)
        adapt_wb.close

        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        assert_equal orig_roo.formula(2, 1), adapt_roo.formula(2, 1)
        assert_equal orig_roo.formula(3, 1), adapt_roo.formula(3, 1)
        assert_equal orig_roo.cell(3, 1), adapt_roo.cell(3, 1)
      end
    end
  end

  def test_side_by_side_auto_width_calculation
    orig_wb = ::FastExcel.open
    orig_ws = orig_wb.add_worksheet
    orig_ws.auto_width = true
    orig_ws.append_row(["Short", "A substantially longer string value", 12_345_678])
    orig_ws.close

    adapt_wb = Xlsxrb::Adapters::FastExcel.open
    adapt_ws = adapt_wb.add_worksheet
    adapt_ws.auto_width = true
    adapt_ws.append_row(["Short", "A substantially longer string value", 12_345_678])
    adapt_ws.close

    assert_equal orig_ws.calculated_column_widths.keys.sort, adapt_ws.calculated_column_widths.keys.sort
    orig_ws.calculated_column_widths.each do |col_idx, width|
      adapt_width = adapt_ws.calculated_column_widths[col_idx]
      assert_in_delta width, adapt_width, 1.0, "Auto width mismatch on col #{col_idx}"
    end
  end

  def test_side_by_side_format_properties
    orig_wb = ::FastExcel.open
    adapt_wb = Xlsxrb::Adapters::FastExcel.open

    orig_fmt = orig_wb.add_format(
      bold: true,
      italic: true,
      font_size: 14,
      font_name: "Arial",
      font_color: :red,
      bg_color: :yellow,
      top: :thin,
      bottom: :thin,
      align: { h: :center, v: :center }
    )

    adapt_fmt = adapt_wb.add_format(
      bold: true,
      italic: true,
      font_size: 14,
      font_name: "Arial",
      font_color: :red,
      bg_color: :yellow,
      top: :thin,
      bottom: :thin,
      align: { h: :center, v: :center }
    )

    assert_equal orig_fmt.bold, adapt_fmt.bold
    assert_equal orig_fmt.italic, adapt_fmt.italic
    assert_equal orig_fmt.font_size, adapt_fmt.font_size
    assert_equal orig_fmt.font_name, adapt_fmt.font_name
    assert_equal orig_fmt.font_color, adapt_fmt.font_color
    assert_equal orig_fmt.bg_color, adapt_fmt.bg_color
    assert_equal orig_fmt.top, adapt_fmt.top
    assert_equal orig_fmt.bottom, adapt_fmt.bottom
    assert_equal orig_fmt.align, adapt_fmt.align
  end

  def test_side_by_side_read_string_output
    orig_wb = ::FastExcel.open
    orig_ws = orig_wb.add_worksheet("StringTest")
    orig_ws.append_row(%w[Alpha Beta Gamma])
    orig_ws.append_row([1, 2, 3])
    orig_data = orig_wb.read_string

    adapt_wb = Xlsxrb::Adapters::FastExcel.open
    adapt_ws = adapt_wb.add_worksheet("StringTest")
    adapt_ws.append_row(%w[Alpha Beta Gamma])
    adapt_ws.append_row([1, 2, 3])
    adapt_data = adapt_wb.read_string

    assert orig_data.start_with?("PK")
    assert adapt_data.start_with?("PK")

    Tempfile.create(["orig_read_string", ".xlsx"]) do |tmp1|
      tmp1.binmode
      tmp1.write(orig_data)
      tmp1.flush

      Tempfile.create(["adapt_read_string", ".xlsx"]) do |tmp2|
        tmp2.binmode
        tmp2.write(adapt_data)
        tmp2.flush

        orig_roo = Roo::Excelx.new(tmp1.path)
        adapt_roo = Roo::Excelx.new(tmp2.path)

        assert_equal orig_roo.sheets, adapt_roo.sheets
        assert_equal orig_roo.cell(1, 1), adapt_roo.cell(1, 1)
        assert_equal orig_roo.cell(2, 3), adapt_roo.cell(2, 3)
      end
    end
  end
end
