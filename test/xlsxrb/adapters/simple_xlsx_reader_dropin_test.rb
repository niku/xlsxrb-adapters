# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersSimpleXlsxReaderDropinTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/simple_xlsx_reader/#{name}", __dir__)
  end

  def test_dropin_constant_aliasing
    target_mod = Xlsxrb::Adapters::SimpleXlsxReader

    assert defined?(target_mod::VERSION)
    assert defined?(target_mod::Document)
    assert defined?(target_mod::Document::Sheet)
    assert defined?(target_mod::Document::RowsProxy)
    assert defined?(target_mod::Sheet)
    assert defined?(target_mod::Loader)
    assert defined?(target_mod::Hyperlink)
    assert defined?(target_mod::CellLoadError)
    assert defined?(target_mod::Configuration)

    doc = target_mod.open(fixture("misc_numbers.xlsx"))
    assert_instance_of target_mod::Document, doc
    assert_equal ["Sheet1"], doc.sheets.map(&:name)

    sheet = doc.sheets.first
    assert_equal ["Medium number", "Big Number"], sheet.rows.first
    assert_equal [98_070, 1_234_567_890_123], sheet.rows.to_a[1]
  end

  def test_dropin_roundtrip_workflow
    with_tempfile do |filepath|
      # 1. Create a spreadsheet using Caxlsx adapter
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      ws = pkg.workbook.add_worksheet(name: "Products")
      ws.add_row(%w[SKU Name Price Quantity Active])
      ws.add_row(["A-100", "Widget Alpha", 12.50, 100, true])
      ws.add_row(["B-200", "Widget Beta", 99.95, 50, false])
      pkg.serialize(filepath)

      # 2. Read back using SimpleXlsxReader adapter
      doc = Xlsxrb::Adapters::SimpleXlsxReader.open(filepath)
      assert_equal ["Products"], doc.sheets.map(&:name)

      sheet = doc.sheets.first
      rows = sheet.rows.to_a
      assert_equal 3, rows.size
      assert_equal %w[SKU Name Price Quantity Active], rows[0]
      assert_equal ["A-100", "Widget Alpha", 12.5, 100, true], rows[1]
      assert_equal ["B-200", "Widget Beta", 99.95, 50, false], rows[2]

      # 3. Read back with headers: true
      dict_rows = sheet.rows.each(headers: true).to_a
      assert_equal 2, dict_rows.size
      assert_equal({ "SKU" => "A-100", "Name" => "Widget Alpha", "Price" => 12.5, "Quantity" => 100, "Active" => true }, dict_rows[0])
      assert_equal({ "SKU" => "B-200", "Name" => "Widget Beta", "Price" => 99.95, "Quantity" => 50, "Active" => false }, dict_rows[1])

      # 4. to_hash
      hash_data = doc.to_hash
      assert_equal ["Products"], hash_data.keys
      assert_equal 3, hash_data["Products"].size

      # 5. Convert to Xlsxrb::Elements::Workbook and back
      wb = doc.to_xlsxrb
      assert_instance_of Xlsxrb::Elements::Workbook, wb
      assert_equal 1, wb.sheets.size
      assert_equal "Products", wb.sheets.first.name

      doc_from_wb = Xlsxrb::Adapters::SimpleXlsxReader.from_xlsxrb(wb)
      assert_instance_of Xlsxrb::Adapters::SimpleXlsxReader::Document, doc_from_wb
      assert_equal ["Products"], doc_from_wb.sheets.map(&:name)
      assert_equal rows, doc_from_wb.sheets.first.rows.to_a
    end
  end
end
