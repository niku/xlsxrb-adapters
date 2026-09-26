# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersRubyXLStylesTest < Test::Unit::TestCase
  def setup
    @workbook = Xlsxrb::Adapters::RubyXL::Workbook.new
    @worksheet = @workbook[0]
    @cell = @worksheet.add_cell(0, 0, "StyledCell")
  end

  def test_cell_fill_and_validation
    assert_raise(RuntimeError) { @cell.change_fill("G") }
    assert_raise(RuntimeError) { @cell.change_fill("#0f0f0f") }

    @cell.change_fill("0f0f0f")
    assert_equal "0f0f0f", @cell.fill_color
  end

  def test_cell_font_name_and_size
    @cell.change_font_name("Arial")
    assert_equal "Arial", @cell.font_name

    @cell.change_font_size(24)
    assert_equal 24, @cell.font_size

    assert_raise(RuntimeError) { @cell.change_font_size("30") }
  end

  def test_cell_font_color_and_validation
    assert_raise(RuntimeError) { @cell.change_font_color("invalid") }
    assert_raise(RuntimeError) { @cell.change_font_color("#123456") }

    @cell.change_font_color("FF0000")
    assert_equal "FF0000", @cell.font_color
  end

  def test_cell_font_styles
    assert_nil @cell.is_italicized
    @cell.change_font_italics(true)
    assert_equal true, @cell.is_italicized

    assert_nil @cell.is_bolded
    @cell.change_font_bold(true)
    assert_equal true, @cell.is_bolded

    assert_nil @cell.is_underlined
    @cell.change_font_underline(true)
    assert_equal true, @cell.is_underlined

    assert_nil @cell.is_struckthrough
    @cell.change_font_strikethrough(true)
    assert_equal true, @cell.is_struckthrough
  end

  def test_cell_alignment
    assert_nil @cell.horizontal_alignment
    @cell.change_horizontal_alignment("center")
    assert_equal "center", @cell.horizontal_alignment

    assert_nil @cell.vertical_alignment
    @cell.change_vertical_alignment("top")
    assert_equal "top", @cell.vertical_alignment

    assert_nil @cell.text_wrap
    @cell.change_text_wrap(true)
    assert_equal true, @cell.text_wrap

    assert_nil @cell.text_rotation
    @cell.change_text_rotation(90)
    assert_equal 90, @cell.text_rotation

    assert_nil @cell.text_indent
    @cell.change_text_indent(2)
    assert_equal 2, @cell.text_indent
  end

  def test_cell_borders
    %i[top bottom left right].each do |direction|
      assert_nil @cell.get_border(direction)
      @cell.change_border(direction, "thin")
      assert_equal "thin", @cell.get_border(direction)

      assert_nil @cell.get_border_color(direction)
      @cell.change_border_color(direction, "00FF00")
      assert_equal "00FF00", @cell.get_border_color(direction)
    end
  end

  def test_cell_number_format
    @cell.set_number_format("$#,##0.00")
    xf = @cell.get_cell_xf
    assert xf.num_fmt_id >= 164
    assert_equal "$#,##0.00", @workbook.stylesheet.number_formats[xf.num_fmt_id]
  end

  def test_row_styling
    @worksheet.add_cell(1, 0, "R1C0")
    @worksheet.add_cell(1, 1, "R1C1")

    # Height
    @worksheet.change_row_height(1, 40)
    assert_equal 40, @worksheet.get_row_height(1)

    # Fill
    @worksheet.change_row_fill(1, "AABBCC")
    assert_equal "AABBCC", @worksheet.get_row_fill(1)
    assert_equal "AABBCC", @worksheet[1][0].fill_color
    assert_equal "AABBCC", @worksheet[1][1].fill_color

    # Font
    @worksheet.change_row_font_name(1, "Courier")
    assert_equal "Courier", @worksheet.get_row_font_name(1)
    assert_equal "Courier", @worksheet[1][0].font_name

    @worksheet.change_row_font_size(1, 16)
    assert_equal 16, @worksheet.get_row_font_size(1)
    assert_equal 16, @worksheet[1][0].font_size

    @worksheet.change_row_font_color(1, "FF00FF")
    assert_equal "FF00FF", @worksheet.get_row_font_color(1)
    assert_equal "FF00FF", @worksheet[1][0].font_color

    @worksheet.change_row_bold(1, true)
    assert_equal true, @worksheet.is_row_bolded(1)
    assert_equal true, @worksheet[1][0].is_bolded

    @worksheet.change_row_italics(1, true)
    assert_equal true, @worksheet.is_row_italicized(1)
    assert_equal true, @worksheet[1][0].is_italicized

    @worksheet.change_row_underline(1, true)
    assert_equal true, @worksheet.is_row_underlined(1)
    assert_equal true, @worksheet[1][0].is_underlined

    @worksheet.change_row_strikethrough(1, true)
    assert_equal true, @worksheet.is_row_struckthrough(1)
    assert_equal true, @worksheet[1][0].is_struckthrough

    # Alignment
    @worksheet.change_row_horizontal_alignment(1, "right")
    assert_equal "right", @worksheet.get_row_alignment(1, true)

    @worksheet.change_row_vertical_alignment(1, "bottom")
    assert_equal "bottom", @worksheet.get_row_alignment(1, false)

    # Border
    @worksheet.change_row_border(1, :top, "medium")
    assert_equal "medium", @worksheet.get_row_border(1, :top)
    assert_equal "medium", @worksheet[1][0].get_border(:top)

    @worksheet.change_row_border_color(1, :top, "112233")
    assert_equal "112233", @worksheet.get_row_border_color(1, :top)
    assert_equal "112233", @worksheet[1][0].get_border_color(:top)
  end

  def test_column_styling
    @worksheet.add_cell(0, 2, "R0C2")
    @worksheet.add_cell(1, 2, "R1C2")

    # Width
    @worksheet.change_column_width(2, 25)
    assert_equal 25, @worksheet.get_column_width(2)

    @worksheet.change_column_width_raw(2, 30.5)
    assert_equal 30.5, @worksheet.get_column_width_raw(2)

    # Fill
    @worksheet.change_column_fill(2, "CCBBAA")
    assert_equal "CCBBAA", @worksheet.get_column_fill(2)
    assert_equal "CCBBAA", @worksheet[0][2].fill_color
    assert_equal "CCBBAA", @worksheet[1][2].fill_color

    # Font
    @worksheet.change_column_font_name(2, "Helvetica")
    assert_equal "Helvetica", @worksheet.get_column_font_name(2)
    assert_equal "Helvetica", @worksheet[0][2].font_name

    @worksheet.change_column_font_size(2, 14)
    assert_equal 14, @worksheet.get_column_font_size(2)
    assert_equal 14, @worksheet[0][2].font_size

    @worksheet.change_column_font_color(2, "334455")
    assert_equal "334455", @worksheet.get_column_font_color(2)
    assert_equal "334455", @worksheet[0][2].font_color

    @worksheet.change_column_bold(2, true)
    assert_equal true, @worksheet.is_column_bolded(2)
    assert_equal true, @worksheet[0][2].is_bolded

    @worksheet.change_column_italics(2, true)
    assert_equal true, @worksheet.is_column_italicized(2)
    assert_equal true, @worksheet[0][2].is_italicized

    @worksheet.change_column_underline(2, true)
    assert_equal true, @worksheet.is_column_underlined(2)
    assert_equal true, @worksheet[0][2].is_underlined

    @worksheet.change_column_strikethrough(2, true)
    assert_equal true, @worksheet.is_column_struckthrough(2)
    assert_equal true, @worksheet[0][2].is_struckthrough

    # Alignment
    @worksheet.change_column_horizontal_alignment(2, "center")
    assert_equal "center", @worksheet.get_column_alignment(2, :horizontal)

    @worksheet.change_column_vertical_alignment(2, "center")
    assert_equal "center", @worksheet.get_column_alignment(2, :vertical)

    # Border
    @worksheet.change_column_border(2, :left, "thick")
    assert_equal "thick", @worksheet.get_column_border(2, :left)
    assert_equal "thick", @worksheet[0][2].get_border(:left)

    @worksheet.change_column_border_color(2, :left, "998877")
    assert_equal "998877", @worksheet.get_column_border_color(2, :left)
    assert_equal "998877", @worksheet[0][2].get_border_color(:left)
  end

  def test_styles_persistence_on_write_and_parse
    @cell.change_font_name("Georgia")
    @cell.change_font_size(18)
    @cell.change_font_bold(true)
    @cell.change_fill("FFAA00")
    @cell.change_horizontal_alignment("center")
    @cell.change_vertical_alignment("top")
    @cell.change_shrink_to_fit(true)
    @cell.change_text_indent(2)
    @cell.change_border(:bottom, "double")
    @cell.change_border_color(:bottom, "0000FF")

    with_tempfile do |filepath|
      @workbook.write(filepath)

      parsed = Xlsxrb::Adapters::RubyXL::Parser.parse(filepath)
      parsed_cell = parsed[0][0][0]
      assert_equal "Georgia", parsed_cell.font_name
      assert_equal 18, parsed_cell.font_size
      assert_equal true, parsed_cell.is_bolded
      assert_equal "FFAA00", parsed_cell.fill_color
      assert_equal "center", parsed_cell.horizontal_alignment
      assert_equal "top", parsed_cell.vertical_alignment
      assert_equal true, parsed_cell.get_cell_xf.alignment&.shrink_to_fit
      assert_equal 2, parsed_cell.get_cell_xf.alignment&.indent
      assert_equal "double", parsed_cell.get_border(:bottom)
      assert_equal "0000FF", parsed_cell.get_border_color(:bottom)
    end
  end

  def test_new_convenience_methods
    # Cell remove_formula & hyperlink
    c = @worksheet.add_cell(1, 1, 100, "SUM(A1:A2)")
    refute_nil c.formula
    c.remove_formula
    assert_nil c.formula

    c.add_hyperlink("https://example.com", "Example Link")
    assert_equal "https://example.com", c.hyperlink
    assert_equal "Example Link", c.tooltip

    # Workbook password_hash
    hash = @workbook.password_hash("password123")
    assert_kind_of String, hash
    assert_equal hash, @workbook.password_hash("password123")

    # Worksheet get_col_style & add_validation_list
    style_idx = @worksheet.get_col_style(2)
    assert_kind_of Integer, style_idx

    @worksheet.add_validation_list("A1:A10", %w[Apple Banana Orange])
    refute_nil @worksheet.data_validations
    assert_equal 1, @worksheet.data_validations.size
    assert_equal "A1:A10", @worksheet.data_validations.first[:sqref]
    assert_equal '"Apple,Banana,Orange"', @worksheet.data_validations.first[:formula]
  end

  def test_row_and_column_styles_roundtrip_persistence
    @worksheet.change_row_fill(1, "CCDDFF")
    @worksheet.change_row_bold(1, true)
    @worksheet.change_row_italics(1, true)
    @worksheet.change_row_font_color(1, "112233")
    @worksheet.change_column_fill(3, "FFEEAA")
    @worksheet.change_column_bold(3, true)
    @worksheet.change_column_font_size(3, 16)
    @worksheet.change_column_border(3, :left, "medium")
    @worksheet.change_column_border_color(3, :left, "FF0000")

    with_tempfile do |filepath|
      @workbook.write(filepath)

      parsed = Xlsxrb::Adapters::RubyXL::Parser.parse(filepath)
      parsed_ws = parsed[0]

      # Verify row styles persisted
      assert_equal "CCDDFF", parsed_ws.get_row_fill(1)
      assert_equal true, parsed_ws.is_row_bolded(1)
      assert_equal true, parsed_ws.is_row_italicized(1)
      assert_equal "112233", parsed_ws.get_row_font_color(1)

      # Verify column styles persisted
      assert_equal "FFEEAA", parsed_ws.get_column_fill(3)
      assert_equal true, parsed_ws.is_column_bolded(3)
      assert_equal 16, parsed_ws.get_column_font_size(3)
      assert_equal "medium", parsed_ws.get_column_border(3, :left)
      assert_equal "FF0000", parsed_ws.get_column_border_color(3, :left)
    end
  end
end
