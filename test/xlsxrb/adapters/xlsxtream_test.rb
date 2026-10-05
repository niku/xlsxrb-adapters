# frozen_string_literal: true

require "test_helper"
require "roo"
require "stringio"

class XlsxrbAdaptersXlsxtreamTest < Test::Unit::TestCase
  def test_constants_and_module_structure
    mod = Xlsxrb::Adapters::Xlsxtream
    assert_equal "3.1.0", mod::VERSION
    assert mod::Error < StandardError
    assert mod::Deprecation < StandardError

    # Font family IDs
    assert_equal 0, mod::Workbook::FONT_FAMILY_IDS[""]
    assert_equal 1, mod::Workbook::FONT_FAMILY_IDS["roman"]
    assert_equal 2, mod::Workbook::FONT_FAMILY_IDS["swiss"]
    assert_equal 3, mod::Workbook::FONT_FAMILY_IDS["modern"]
    assert_equal 4, mod::Workbook::FONT_FAMILY_IDS["script"]
    assert_equal 5, mod::Workbook::FONT_FAMILY_IDS["decorative"]

    # Row constants
    assert_equal 1, mod::Row::DATE_STYLE
    assert_equal 2, mod::Row::TIME_STYLE
    assert_equal "true", mod::Row::TRUE_STRING
    assert_equal "false", mod::Row::FALSE_STRING
  end

  def test_xml_utilities
    xml = Xlsxrb::Adapters::Xlsxtream::XML
    assert_equal "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n", xml.header
    assert_equal "<a><b>text</b></a>", xml.strip("<a>\n  <b>text</b>\n</a>")

    assert_equal "&lt;tag attr=&quot;val&amp;more&quot;&gt;", xml.escape_attr('<tag attr="val&more">')
    assert_equal "&lt;val &amp; &gt;", xml.escape_value("<val & >")
    # ST_Xstring escaping
    assert_equal "_x005f_x0020_", xml.escape_value("_x0020_")
  end

  def test_shared_string_table
    sst = Xlsxrb::Adapters::Xlsxtream::SharedStringTable.new
    assert_equal 0, sst.references
    assert_equal 0, sst.size

    idx0 = sst["hello"]
    assert_equal 0, idx0
    assert_equal 1, sst.references
    assert_equal 1, sst.size

    idx0_again = sst["hello"]
    assert_equal 0, idx0_again
    assert_equal 2, sst.references
    assert_equal 1, sst.size

    idx1 = sst["world"]
    assert_equal 1, idx1
    assert_equal 3, sst.references
    assert_equal 2, sst.size
  end

  def test_columns_xml_generation
    cols_class = Xlsxrb::Adapters::Xlsxtream::Columns

    # Unspecified widths
    cols1 = cols_class.new([{}])
    assert_equal '<cols><col min="1" max="1"/></cols>', cols1.to_xml

    # Width in characters
    cols2 = cols_class.new([{ width_chars: 10 }])
    xml2 = cols2.to_xml
    assert_include xml2, '<col min="1" max="1" width="'
    assert_include xml2, 'customWidth="1"/>'

    # Width in pixels
    cols3 = cols_class.new([{ width_pixels: 45.5 }])
    assert_equal '<cols><col min="1" max="1" width="45.5" customWidth="1"/></cols>', cols3.to_xml
  end

  def test_row_generation_and_data_types
    row_class = Xlsxrb::Adapters::Xlsxtream::Row

    # Primitives without SST (inlineStr)
    r1 = row_class.new([42, 3.14, true, false, "simple"], 1)
    xml1 = r1.to_xml
    assert_include xml1, '<c r="A1" t="n"><v>42</v></c>'
    assert_include xml1, '<c r="B1" t="n"><v>3.14</v></c>'
    assert_include xml1, '<c r="C1" t="b"><v>1</v></c>'
    assert_include xml1, '<c r="D1" t="b"><v>0</v></c>'
    assert_include xml1, '<c r="E1" t="inlineStr"><is><t>simple</t></is></c>'

    # Empty string omission
    r_empty = row_class.new(["", nil, "val"], 2)
    xml_empty = r_empty.to_xml
    assert_not_include xml_empty, 'r="A2"'
    assert_not_include xml_empty, 'r="B2"'
    assert_include xml_empty, '<c r="C2" t="inlineStr"><is><t>val</t></is></c>'

    # Date and Time
    d = Date.new(2026, 1, 15)
    t = Time.utc(2026, 1, 15, 12, 0, 0)
    dt = DateTime.new(2026, 1, 15, 12, 0, 0)
    r_dates = row_class.new([d, t, dt], 3)
    xml_dates = r_dates.to_xml
    assert_include xml_dates, '<c r="A3" s="1">'
    assert_include xml_dates, '<c r="B3" s="2">'
    assert_include xml_dates, '<c r="C3" s="2">'

    # With SST
    sst = Xlsxrb::Adapters::Xlsxtream::SharedStringTable.new
    r_sst = row_class.new(%w[foo bar foo], 4, sst: sst)
    xml_sst = r_sst.to_xml
    assert_include xml_sst, '<c r="A4" t="s"><v>0</v></c>'
    assert_include xml_sst, '<c r="B4" t="s"><v>1</v></c>'
    assert_include xml_sst, '<c r="C4" t="s"><v>0</v></c>'
    assert_equal 2, sst.size
    assert_equal 3, sst.references
  end

  def test_row_auto_format
    row_class = Xlsxrb::Adapters::Xlsxtream::Row
    r = row_class.new(["true", "false", "123", "45.67", "2026-05-20", "2026-05-20T14:30:00", "regular"], 1, auto_format: true)
    xml = r.to_xml
    assert_include xml, '<c r="A1" t="b"><v>1</v></c>'
    assert_include xml, '<c r="B1" t="b"><v>0</v></c>'
    assert_include xml, '<c r="C1" t="n"><v>123</v></c>'
    assert_include xml, '<c r="D1" t="n"><v>45.67</v></c>'
    assert_include xml, '<c r="E1" s="1">'
    assert_include xml, '<c r="F1" s="2">'
    assert_include xml, '<c r="G1" t="inlineStr"><is><t>regular</t></is></c>'
  end

  def test_workbook_basic_lifecycle
    with_tempfile do |filepath|
      wb = Xlsxrb::Adapters::Xlsxtream::Workbook.new(filepath)
      assert_instance_of Xlsxrb::Adapters::Xlsxtream::Workbook, wb

      ws = wb.add_worksheet("TestSheet")
      assert_instance_of Xlsxrb::Adapters::Xlsxtream::Worksheet, ws
      assert_equal "TestSheet", ws.name
      assert_equal 1, ws.id
      refute ws.closed?

      ws << %w[Col1 Col2]
      ws << [100, 200]
      ws.close
      assert ws.closed?

      wb.close

      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      roo = Roo::Excelx.new(filepath)
      assert_equal ["TestSheet"], roo.sheets
      assert_equal "Col1", roo.cell(1, 1)
      assert_equal 200, roo.cell(2, 2)
    end
  end

  def test_workbook_block_form_and_write_worksheet
    with_tempfile do |filepath|
      Xlsxrb::Adapters::Xlsxtream::Workbook.open(filepath) do |wb|
        wb.write_worksheet("Page1") do |sheet|
          sheet << %w[Alpha Beta]
          sheet.add_row([1, 2])
        end

        wb.write_worksheet("Page2") do |sheet|
          sheet << %w[Gamma Delta]
        end
      end

      roo = Roo::Excelx.new(filepath)
      assert_equal %w[Page1 Page2], roo.sheets
      assert_equal "Alpha", roo.cell(1, 1, "Page1")
      assert_equal 2, roo.cell(2, 2, "Page1")
      assert_equal "Gamma", roo.cell(1, 1, "Page2")
    end
  end

  def test_workbook_invalid_font_family_raises_error
    with_tempfile do |filepath|
      wb = Xlsxrb::Adapters::Xlsxtream::Workbook.new(filepath, font: { family: "comic_sans" })
      wb.write_worksheet("Sheet1") { |ws| ws << [1, 2] }
      assert_raise(Xlsxrb::Adapters::Xlsxtream::Error) do
        wb.close
      end
    end
  end

  def test_workbook_worksheet_close_guard
    with_tempfile do |filepath|
      wb = Xlsxrb::Adapters::Xlsxtream::Workbook.new(filepath)
      wb.add_worksheet("Unclosed")
      assert_raise(Xlsxrb::Adapters::Xlsxtream::Error) do
        wb.add_worksheet("Second")
      end
      wb.worksheets.first.close
      ws2 = wb.add_worksheet("Second")
      ws2.close
      wb.close
    end
  end

  def test_to_xlsxrb_conversion
    sio = StringIO.new
    wb = Xlsxrb::Adapters::Xlsxtream::Workbook.new(
      sio,
      font: { name: "Arial", size: 10, family: "Swiss" },
      columns: [{ width_pixels: 20.0 }]
    )
    ws = wb.add_worksheet("Report")
    ws << %w[Name Score Registered]
    ws << ["Alice", 95, Date.new(2026, 1, 1)]
    ws.close
    wb.close

    xlsxrb_wb = wb.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
    assert_equal 1, xlsxrb_wb.sheets.size

    sheet = xlsxrb_wb.sheets.first
    assert_equal "Report", sheet.name
    assert_equal 2, sheet.rows.size
    assert_equal "Name", sheet.rows[0].cells[0].value
    assert_equal 95, sheet.rows[1].cells[1].value
    assert_equal Date.new(2026, 1, 1), sheet.rows[1].cells[2].value
  end

  def test_from_xlsxrb_conversion
    source_wb = Xlsxrb.build do |b|
      b.sheet("Export") do |s|
        s.row(%w[Product Qty Price])
        s.row(["Desk", 4, 120.0])
      end
    end

    sio = StringIO.new
    adapter = Xlsxrb::Adapters::Xlsxtream.from_xlsxrb(source_wb, sio)
    assert_instance_of Xlsxrb::Adapters::Xlsxtream::Workbook, adapter
    assert sio.string.bytesize.positive?

    with_tempfile do |path|
      File.binwrite(path, sio.string)
      roo = Roo::Excelx.new(path)
      assert_equal ["Export"], roo.sheets
      assert_equal "Product", roo.cell(1, 1)
      assert_equal "Desk", roo.cell(2, 1)
      assert_equal 4, roo.cell(2, 2)
      assert_equal 120.0, roo.cell(2, 3)
    end
  end
end
