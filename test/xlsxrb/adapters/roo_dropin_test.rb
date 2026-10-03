# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersRooDropinTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def test_dropin_constant_aliasing
    # Emulate user opting into drop-in replacement by aliasing Roo to adapter module
    target_mod = Xlsxrb::Adapters::Roo

    assert defined?(target_mod::Spreadsheet)
    assert defined?(target_mod::Excelx)
    assert defined?(target_mod::Link)
    assert defined?(target_mod::Font)

    doc = target_mod::Spreadsheet.open(fixture("simple_spreadsheet.xlsx"))
    assert_instance_of target_mod::Excelx, doc
    assert_equal %w[Sheet1 Sheet2 Sheet3], doc.sheets
    assert_equal "Sheet1", doc.default_sheet
    assert_equal 1.75, doc.cell(5, 5)
    assert doc.formula?(5, 5)

    # Test reading with Roo::Excelx.new interface
    doc2 = target_mod::Excelx.new(fixture("style.xlsx"))
    assert_equal true, doc2.font(1, 1).bold?

    # Test reading comments
    doc3 = target_mod::Excelx.new(fixture("comments.xlsx"))
    assert_equal "Kommentar fuer B4", doc3.comment(4, 2)
  end

  def test_dropin_roundtrip_workflow
    with_tempfile do |filepath|
      # 1. Create a spreadsheet using Caxlsx adapter
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      ws = pkg.workbook.add_worksheet(name: "SalesData")
      ws.add_row(%w[Product Quantity Price Total])
      ws.add_row(["Widget A", 10, 2.5, "=B2*C2"])
      ws.add_row(["Widget B", 20, 5.0, "=B3*C3"])
      pkg.serialize(filepath)

      # 2. Read back using Roo adapter
      roo = Xlsxrb::Adapters::Roo::Spreadsheet.open(filepath)
      assert_equal ["SalesData"], roo.sheets
      assert_equal "Product", roo.cell(1, 1)
      assert_equal 10, roo.cell(2, 2)
      assert_equal 2.5, roo.cell(2, 3)
      assert_equal "B2*C2", roo.formula(2, 4)

      # 3. Stream rows
      streamed = []
      roo.each_row_streaming { |r| streamed << r.map(&:value) }
      assert_equal 3, streamed.size
      assert_equal %w[Product Quantity Price Total], streamed[0]

      # 4. Parse headers
      parsed = roo.parse(headers: true)
      assert_equal 3, parsed.size
    end
  end
end
