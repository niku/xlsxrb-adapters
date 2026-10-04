# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCreekDropinTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/creek/#{name}", __dir__)
  end

  def test_dropin_constant_aliasing
    target_mod = Xlsxrb::Adapters::Creek

    assert defined?(target_mod::Book)
    assert defined?(target_mod::Sheet)
    assert defined?(target_mod::Styles)
    assert defined?(target_mod::Drawing)
    assert defined?(target_mod::SharedStrings)
    assert defined?(target_mod::Utils)
    assert_equal "2.6.3", target_mod::VERSION

    doc = target_mod::Book.new(fixture("sample.xlsx"))
    assert_instance_of target_mod::Book, doc
    assert_equal ["Sheet1"], doc.sheets.map(&:name)

    sheet0 = doc.sheets.first
    rows = sheet0.rows.to_a
    assert_equal 9, rows.size
    assert_equal "Content 1", rows[0]["A1"]
    assert_equal 0.15, rows[8]["A10"]
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

      # 2. Read back using Creek adapter in coordinate rows mode
      book1 = Xlsxrb::Adapters::Creek::Book.new(filepath)
      assert_equal ["Inventory"], book1.sheets.map(&:name)
      rows = book1.sheets[0].rows.to_a
      assert_equal 3, rows.size
      assert_equal({ "A1" => "SKU", "B1" => "Product", "C1" => "Quantity", "D1" => "Price", "E1" => "InStock" }, rows[0])
      assert_equal({ "A2" => "A-101", "B2" => "Widget A", "C2" => "100", "D2" => "2.5", "E2" => true }, rows[1])
      assert_equal({ "A3" => "B-202", "B3" => "Widget B", "C3" => "50", "D3" => "15.75", "E3" => false }, rows[2])
      book1.close

      # 3. Read back in simple_rows mode with headers
      book2 = Xlsxrb::Adapters::Creek::Book.new(filepath, with_headers: true)
      sheet = book2.sheets[0]
      assert_equal true, sheet.with_headers

      simple_rows = sheet.simple_rows.to_a
      assert_equal 3, simple_rows.size
      assert_equal({ "A" => "SKU", "B" => "Product", "C" => "Quantity", "D" => "Price", "E" => "InStock" }, simple_rows[0])
      assert_equal({ "SKU" => "A-101", "Product" => "Widget A", "Quantity" => "100", "Price" => "2.5", "InStock" => true }, simple_rows[1])
      assert_equal({ "SKU" => "B-202", "Product" => "Widget B", "Quantity" => "50", "Price" => "15.75", "InStock" => false }, simple_rows[2])

      # 4. Convert to xlsxrb and back
      xlsxrb_wb = book2.to_xlsxrb
      assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
      assert_equal "Inventory", xlsxrb_wb.sheets[0].name

      book3 = Xlsxrb::Adapters::Creek.from_xlsxrb(xlsxrb_wb, with_headers: true)
      assert_equal simple_rows, book3.sheets[0].simple_rows.to_a

      book2.close
      book3.close
    end
  end
end
