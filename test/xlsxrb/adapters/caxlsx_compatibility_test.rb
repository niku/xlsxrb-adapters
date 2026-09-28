# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCaxlsxCompatibilityTest < Test::Unit::TestCase
  def test_side_by_side_basic_sheet_and_cells
    orig_pkg = ::Axlsx::Package.new
    orig_ws = orig_pkg.workbook.add_worksheet(name: "CompatSheet")

    adapt_pkg = ::Xlsxrb::Adapters::Caxlsx::Package.new
    adapt_ws = adapt_pkg.workbook.add_worksheet(name: "CompatSheet")

    assert_equal orig_ws.name, adapt_ws.name

    today = Date.today
    data = ["Hello World", 42, 3.14159, true, false, today]

    orig_row = orig_ws.add_row(data)
    adapt_row = adapt_ws.add_row(data)

    assert_equal orig_row.cells.size, adapt_row.cells.size

    orig_row.cells.each_with_index do |orig_cell, idx|
      adapt_cell = adapt_row.cells[idx]

      assert_equal orig_cell.type, adapt_cell.type
      assert_equal orig_cell.value, adapt_cell.value
      assert_equal orig_cell.r, adapt_cell.r
      assert_equal orig_cell.r_abs, adapt_cell.r_abs
      assert_equal orig_cell.pos, adapt_cell.pos
    end
  end

  def test_side_by_side_styles
    orig_pkg = ::Axlsx::Package.new
    adapt_pkg = ::Xlsxrb::Adapters::Caxlsx::Package.new

    style_opts = {
      font_name: "Times New Roman",
      sz: 16,
      b: true,
      i: true,
      u: :single,
      strike: true,
      fg_color: "0000FF",
      bg_color: "FFFF00",
      alignment: {
        horizontal: :center,
        vertical: :center,
        wrap_text: true
      },
      border: {
        style: :thin,
        color: "000000",
        edges: %i[top bottom]
      }
    }

    orig_id = orig_pkg.workbook.styles.add_style(style_opts)
    adapt_id = adapt_pkg.workbook.styles.add_style(style_opts)

    assert_equal orig_id, adapt_id

    orig_xf = orig_pkg.workbook.styles.cellXfs[orig_id]
    adapt_xf = adapt_pkg.workbook.styles.cellXfs[adapt_id]

    orig_font = orig_pkg.workbook.styles.fonts[orig_xf.fontId]
    adapt_font = adapt_pkg.workbook.styles.fonts[adapt_xf.fontId]

    assert_equal orig_font.name, adapt_font.name
    assert_equal orig_font.sz, adapt_font.sz
    assert_equal orig_font.b, adapt_font.b
    assert_equal orig_font.i, adapt_font.i
    assert_equal orig_font.u, adapt_font.u
    assert_equal orig_font.strike, adapt_font.strike
    assert_equal orig_font.color.rgb, adapt_font.color.rgb

    assert_equal orig_xf.alignment.horizontal, adapt_xf.alignment.horizontal
    assert_equal orig_xf.alignment.vertical, adapt_xf.alignment.vertical
    assert_equal orig_xf.alignment.wrap_text, adapt_xf.alignment.wrap_text
  end

  def test_side_by_side_worksheet_features
    orig_pkg = ::Axlsx::Package.new
    orig_ws = orig_pkg.workbook.add_worksheet(name: "Features")

    adapt_pkg = ::Xlsxrb::Adapters::Caxlsx::Package.new
    adapt_ws = adapt_pkg.workbook.add_worksheet(name: "Features")

    # Merged cells
    orig_ws.merge_cells "A1:C1"
    adapt_ws.merge_cells "A1:C1"
    assert_equal orig_ws.send(:merged_cells).size, adapt_ws.merged_cells.size

    # Auto filter
    orig_ws.auto_filter = "A1:D10"
    adapt_ws.auto_filter = "A1:D10"
    assert_equal orig_ws.auto_filter.range, adapt_ws.auto_filter.range

    # Page setup
    orig_ws.page_setup.orientation = :landscape
    adapt_ws.page_setup.orientation = :landscape
    assert_equal orig_ws.page_setup.orientation, adapt_ws.page_setup.orientation

    orig_ws.page_margins.left = 1.5
    adapt_ws.page_margins.left = 1.5
    assert_in_delta orig_ws.page_margins.left, adapt_ws.page_margins.left, 0.001

    # Freeze panes
    orig_ws.sheet_view.pane do |p|
      p.state = :frozen
      p.y_split = 2
      p.active_pane = :bottom_left
    end
    adapt_ws.sheet_view.pane do |p|
      p.state = :frozen
      p.y_split = 2
      p.active_pane = :bottom_left
    end
    assert_equal orig_ws.sheet_view.pane.state, adapt_ws.sheet_view.pane.state
    assert_equal orig_ws.sheet_view.pane.y_split, adapt_ws.sheet_view.pane.y_split
    assert_equal orig_ws.sheet_view.pane.active_pane, adapt_ws.sheet_view.pane.active_pane
  end

  def test_interoperability_adapter_output_read_by_rubyxl
    with_tempfile do |filepath|
      adapt_pkg = ::Xlsxrb::Adapters::Caxlsx::Package.new
      adapt_ws = adapt_pkg.workbook.add_worksheet(name: "Report")
      adapt_ws.add_row(%w[Department Headcount Budget])
      adapt_ws.add_row(["R&D", 25, 500_000.0])
      adapt_ws.add_row(["Sales", 40, 750_000.0])
      adapt_ws.add_row(["Total", "=SUM(B2:B3)", "=SUM(C2:C3)"])

      assert adapt_pkg.serialize(filepath)
      assert File.exist?(filepath)

      # Read with official rubyXL
      read_wb = ::RubyXL::Parser.parse(filepath)
      assert_equal 1, read_wb.worksheets.size
      read_ws = read_wb["Report"]
      refute_nil read_ws

      assert_equal "Department", read_ws[0][0].value
      assert_equal "Headcount", read_ws[0][1].value
      assert_equal 25, read_ws[1][1].value
      assert_in_delta 500_000.0, read_ws[1][2].value, 0.001
      assert_equal "SUM(B2:B3)", read_ws[3][1].formula.expression
      assert_equal "SUM(C2:C3)", read_ws[3][2].formula.expression
    end
  end

  def test_interoperability_caxlsx_file_read_by_xlsxrb_bridge
    with_tempfile do |filepath|
      orig_pkg = ::Axlsx::Package.new
      orig_ws = orig_pkg.workbook.add_worksheet(name: "OfficialSheet")
      orig_ws.add_row(%w[Product Price Tax])
      orig_ws.add_row(["Widget", 100, 10])
      orig_pkg.serialize(filepath)

      # Read the file generated by official caxlsx using xlsxrb
      xlsxrb_wb = ::Xlsxrb.read(filepath).load
      assert_equal 1, xlsxrb_wb.sheets.size
      assert_equal "OfficialSheet", xlsxrb_wb.sheets.first.name

      sheet = xlsxrb_wb.sheets.first
      assert_equal 2, sheet.rows.size
      assert_equal "Product", sheet.rows[0].cells[0].value
      assert_equal "Widget", sheet.rows[1].cells[0].value
      assert_equal 100, sheet.rows[1].cells[1].value
    end
  end
end
