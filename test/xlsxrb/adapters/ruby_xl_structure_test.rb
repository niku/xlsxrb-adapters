# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersRubyXLStructureTest < Test::Unit::TestCase
  def test_insert_row_and_shift
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "Row0")
    ws.add_cell(1, 0, "Row1")
    ws.add_cell(2, 0, "Row2")

    # Insert row at 1
    new_row = ws.insert_row(1)
    assert_equal 1, new_row.index
    assert_equal "Row0", ws[0][0].value
    assert_nil ws[1][0].value
    assert_equal "Row1", ws[2][0].value
    assert_equal 2, ws[2][0].row
    assert_equal "Row2", ws[3][0].value
    assert_equal 3, ws[3][0].row
  end

  def test_insert_row_inherits_style_from_above
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    c = ws.add_cell(0, 0, "Styled")
    c.change_fill("FF0000")
    styled_index = c.style_index

    ws.insert_row(1)
    inserted_cell = ws[1][0]
    assert_not_nil inserted_cell
    assert_equal styled_index, inserted_cell.style_index
  end

  def test_delete_row_and_shift
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "Row0")
    ws.add_cell(1, 0, "Row1")
    ws.add_cell(2, 0, "Row2")

    deleted = ws.delete_row(1)
    assert_not_nil deleted
    assert_equal "Row0", ws[0][0].value
    assert_equal "Row2", ws[1][0].value
    assert_equal 1, ws[1][0].row
    assert_nil ws[2]
  end

  def test_insert_and_delete_row_with_merged_cells
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.merge_cells(1, 0, 3, 2) # Row 1..3, Col 0..2

    # Insert row above merged cells (at 0)
    ws.insert_row(0)
    mc = ws.merged_cells.first
    assert_equal 2, mc.ref.first_row
    assert_equal 4, mc.ref.last_row

    # Insert row inside merged cells (at 3)
    ws.insert_row(3)
    mc = ws.merged_cells.first
    assert_equal 2, mc.ref.first_row
    assert_equal 5, mc.ref.last_row

    # Delete row above merged cells (at 0)
    ws.delete_row(0)
    mc = ws.merged_cells.first
    assert_equal 1, mc.ref.first_row
    assert_equal 4, mc.ref.last_row

    # Delete row entirely containing a 1-row merged cell
    ws.merge_cells(10, 0, 10, 2) # single row merged cell
    assert_equal 2, ws.merged_cells.size
    ws.delete_row(10)
    assert_equal 1, ws.merged_cells.size
  end

  def test_insert_column_and_shift
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "A1")
    ws.add_cell(0, 1, "B1")
    ws.add_cell(0, 2, "C1")

    ws.insert_column(1)
    assert_equal "A1", ws[0][0].value
    assert_nil ws[0][1]
    assert_equal "B1", ws[0][2].value
    assert_equal 2, ws[0][2].column
    assert_equal "C1", ws[0][3].value
    assert_equal 3, ws[0][3].column
  end

  def test_delete_column_and_shift
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "A1")
    ws.add_cell(0, 1, "B1")
    ws.add_cell(0, 2, "C1")

    ws.delete_column(1)
    assert_equal "A1", ws[0][0].value
    assert_equal "C1", ws[0][1].value
    assert_equal 1, ws[0][1].column
    assert_nil ws[0][2]
  end

  def test_insert_and_delete_column_with_merged_cells
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.merge_cells(0, 1, 2, 3) # Row 0..2, Col 1..3

    # Insert column to the left (at 0)
    ws.insert_column(0)
    mc = ws.merged_cells.first
    assert_equal 2, mc.ref.first_col
    assert_equal 4, mc.ref.last_col

    # Delete column to the left (at 0)
    ws.delete_column(0)
    mc = ws.merged_cells.first
    assert_equal 1, mc.ref.first_col
    assert_equal 3, mc.ref.last_col
  end

  def test_insert_cell_shift_right_and_down
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "A1")
    ws.add_cell(0, 1, "B1")

    # Shift right
    ws.insert_cell(0, 1, "NewB1", nil, :right)
    assert_equal "A1", ws[0][0].value
    assert_equal "NewB1", ws[0][1].value
    assert_equal "B1", ws[0][2].value

    # Shift down
    ws.add_cell(1, 0, "A2")
    ws.insert_cell(0, 0, "NewA1", nil, :down)
    assert_equal "NewA1", ws[0][0].value
    assert_equal "A1", ws[1][0].value
    assert_equal "A2", ws[2][0].value
  end

  def test_delete_cell_shift_left_and_up
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]
    ws.add_cell(0, 0, "A1")
    ws.add_cell(0, 1, "B1")
    ws.add_cell(0, 2, "C1")

    # Shift left
    deleted = ws.delete_cell(0, 1, :left)
    assert_equal "B1", deleted.value
    assert_equal "A1", ws[0][0].value
    assert_equal "C1", ws[0][1].value

    # Shift up
    ws.add_cell(1, 0, "A2")
    ws.add_cell(2, 0, "A3")
    deleted_up = ws.delete_cell(0, 0, :up)
    assert_equal "A1", deleted_up.value
    assert_equal "A2", ws[0][0].value
    assert_equal "A3", ws[1][0].value
  end

  def test_merge_cells_signatures_and_persistence
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]

    # 4 params
    mc1 = ws.merge_cells(0, 0, 1, 1)
    assert_equal "A1:B2", mc1.ref.to_s

    # string param
    mc2 = ws.merge_cells("C1:D3")
    assert_equal "C1:D3", mc2.ref.to_s

    # hash param
    mc3 = ws.merge_cells(row_from: 2, row_to: 4, col_from: 0, col_to: 2)
    assert_equal "A3:C5", mc3.ref.to_s

    assert_equal 3, ws.merged_cells.size

    # Verify persistence through write & parse
    with_tempfile do |filepath|
      wb.write(filepath)
      parsed_wb = Xlsxrb::Adapters::RubyXL::Parser.parse(filepath)
      parsed_ws = parsed_wb[0]
      assert_not_nil parsed_ws.merged_cells
      refs = parsed_ws.merged_cells.map { |m| m.ref.to_s }
      assert_includes refs, "A1:B2"
      assert_includes refs, "C1:D3"
      assert_includes refs, "A3:C5"
    end
  end
end
