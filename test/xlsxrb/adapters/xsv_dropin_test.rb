# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersXsvDropinTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def test_dropin_constant_aliasing
    target_mod = Xlsxrb::Adapters::Xsv

    assert defined?(target_mod::Workbook)
    assert defined?(target_mod::Sheet)
    assert defined?(target_mod::Error)
    assert defined?(target_mod::DuplicateHeaders)
    assert defined?(target_mod::AssertionFailed)
    assert defined?(target_mod::Helpers)

    doc = target_mod.open(fixture("simple_spreadsheet.xlsx"))
    assert_instance_of target_mod::Workbook, doc
    assert_equal %w[Sheet1 Sheet2 Sheet3], doc.sheets.map(&:name)

    sheet0 = doc[0]
    sheet0.row_skip = 2
    sheet0.parse_headers!
    first_row = sheet0.first
    assert_equal Date.new(2007, 5, 7), first_row["Date"]
    assert_equal 9.25, first_row["Start time"]
    doc.close
  end

  def test_dropin_roundtrip_workflow
    with_tempfile do |filepath|
      # 1. Create a spreadsheet using Caxlsx adapter
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      ws = pkg.workbook.add_worksheet(name: "Inventory")
      ws.add_row(%w[SKU Product Quantity Price InStock])
      ws.add_row(["A-101", "Widget A", 100, 2.50, true])
      ws.add_row(["B-202", "Widget B", 50, 15.75, false])
      pkg.serialize(filepath)

      # 2. Read back using Xsv adapter in array mode
      wb1 = Xlsxrb::Adapters::Xsv.open(filepath)
      assert_equal ["Inventory"], wb1.sheets.map(&:name)
      rows = wb1[0].to_a
      assert_equal 3, rows.size
      assert_equal %w[SKU Product Quantity Price InStock], rows[0]
      assert_equal ["A-101", "Widget A", 100, 2.5, true], rows[1]
      wb1.close

      # 3. Read back in hash mode using parse_headers: true
      wb2 = Xlsxrb::Adapters::Xsv.open(filepath, parse_headers: true)
      sheet = wb2[0]
      assert_equal :hash, sheet.mode
      assert_equal %w[SKU Product Quantity Price InStock], sheet.headers

      hash_rows = sheet.to_a
      assert_equal 2, hash_rows.size
      assert_equal({ "SKU" => "A-101", "Product" => "Widget A", "Quantity" => 100, "Price" => 2.5, "InStock" => true }, hash_rows[0])
      assert_equal({ "SKU" => "B-202", "Product" => "Widget B", "Quantity" => 50, "Price" => 15.75, "InStock" => false }, hash_rows[1])

      # 4. Convert to xlsxrb and back
      xlsxrb_wb = wb2.to_xlsxrb
      assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
      assert_equal "Inventory", xlsxrb_wb.sheets[0].name

      wb3 = Xlsxrb::Adapters::Xsv.from_xlsxrb(xlsxrb_wb, parse_headers: true)
      assert_equal hash_rows, wb3[0].to_a

      wb2.close
      wb3.close
    end
  end
end
