# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCaxlsxStylesTest < Test::Unit::TestCase
  def test_styles_initialization_defaults
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    refute_nil styles
    assert styles.fonts.size >= 1
    assert styles.fills.size >= 2 # none and gray125
    assert styles.borders.size >= 1
    assert styles.cellXfs.size >= 1
  end

  def test_add_style_font_properties
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    style_id = styles.add_style(
      font_name: "Arial",
      sz: 14,
      b: true,
      i: true,
      u: :single,
      strike: true,
      fg_color: "FF0000"
    )

    xf = styles.cellXfs[style_id]
    font = styles.fonts[xf.fontId]

    assert_equal "Arial", font.name
    assert_equal 14, font.sz
    assert_equal true, font.b
    assert_equal true, font.i
    assert_equal :single, font.u
    assert_equal true, font.strike
    assert_equal "FFFF0000", font.color.rgb
  end

  def test_add_style_fill_properties
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    style_id = styles.add_style(
      bg_color: "00FF00",
      fg_color: "0000FF",
      pattern_type: :solid
    )

    xf = styles.cellXfs[style_id]
    fill = styles.fills[xf.fillId]
    pf = fill.fill_type

    assert_instance_of Xlsxrb::Adapters::Caxlsx::PatternFill, pf
    assert_equal :solid, pf.patternType
  end

  def test_add_style_border_properties
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    style_id = styles.add_style(
      border: {
        style: :thick,
        color: "FF00FF",
        edges: %i[top bottom left right]
      }
    )

    xf = styles.cellXfs[style_id]
    border = styles.borders[xf.borderId]

    assert_equal :thick, border.prs.find { |p| p.name == :top }.style
    assert_equal :thick, border.prs.find { |p| p.name == :bottom }.style
    assert_equal "FFFF00FF", border.prs.find { |p| p.name == :top }.color.rgb
  end

  def test_add_style_alignment_properties
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    style_id = styles.add_style(
      alignment: {
        horizontal: :center,
        vertical: :center,
        wrap_text: true,
        text_rotation: 45,
        indent: 1
      }
    )

    xf = styles.cellXfs[style_id]
    align = xf.alignment

    refute_nil align
    assert_equal :center, align.horizontal
    assert_equal :center, align.vertical
    assert_equal true, align.wrap_text
    assert_equal 45, align.text_rotation
    assert_equal 1, align.indent
  end

  def test_add_style_num_fmt_properties
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    # Standard built-in number format
    style_currency = styles.add_style(num_fmt: 5)
    xf_curr = styles.cellXfs[style_currency]
    assert_equal 5, xf_curr.numFmtId

    # Custom format string
    style_custom = styles.add_style(format_code: "$#,##0.00;($#,##0.00);\"-\"")
    xf_custom = styles.cellXfs[style_custom]
    assert xf_custom.numFmtId >= 100

    custom_fmt = styles.numFmts.find { |nf| nf.numFmtId == xf_custom.numFmtId }
    refute_nil custom_fmt
    assert_equal "$#,##0.00;($#,##0.00);\"-\"", custom_fmt.formatCode
  end

  def test_add_style_protection_properties
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    style_id = styles.add_style(hidden: true, locked: true)
    xf = styles.cellXfs[style_id]
    protection = xf.protection

    refute_nil protection
    assert_equal true, protection.hidden
    assert_equal true, protection.locked
  end

  def test_dxf_styles_for_conditional_formatting
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles

    dxf_id = styles.add_style(
      type: :dxf,
      b: true,
      fg_color: "FF0000",
      bg_color: "FFFF00"
    )

    dxf = styles.dxfs[dxf_id]
    refute_nil dxf
    assert_equal true, dxf.font.b
  end

  def test_applying_styles_to_cells_and_rows
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "StyledSheet")

    header_style = ws.workbook.styles.add_style(b: true, sz: 12)
    currency_style = ws.workbook.styles.add_style(num_fmt: 5)

    # Style applied to whole row
    ws.add_row(%w[Code Amount], style: header_style)
    assert_equal header_style, ws.rows[0].cells[0].style
    assert_equal header_style, ws.rows[0].cells[1].style

    # Array of styles per cell
    ws.add_row(["USD", 1234.56], style: [nil, currency_style])
    assert_equal 0, ws.rows[1].cells[0].style
    assert_equal currency_style, ws.rows[1].cells[1].style
  end

  def test_styles_to_xml_string
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    styles = pkg.workbook.styles
    styles.add_style(b: true, sz: 14, fg_color: "0000FF")

    xml = styles.to_xml_string
    refute_empty xml
    assert xml.include?("<styleSheet")
    assert xml.include?("<fonts")
    assert xml.include?("<fills")
    assert xml.include?("<borders")
    assert xml.include?("<cellXfs")
    assert xml.include?("</styleSheet>")
  end

  def test_styles_serialization_roundtrip_with_xlsxrb
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Roundtrip")

    bold_red = ws.workbook.styles.add_style(b: true, fg_color: "FF0000")
    ws.add_row(["Important"], style: bold_red)

    # Convert to xlsxrb workbook
    xlsxrb_wb = pkg.to_xlsxrb
    refute_nil xlsxrb_wb.styles
    cell = xlsxrb_wb.sheets.first.rows.first.cells.first
    assert_equal bold_red, cell.style_index

    # Verify serialization
    with_tempfile do |path|
      assert pkg.serialize(path)
      read_wb = ::Xlsxrb.read(path).load
      read_cell = read_wb.sheets.first.rows.first.cells.first
      assert_equal "Important", read_cell.value
      assert_equal bold_red, read_cell.style_index
    end
  end
end
