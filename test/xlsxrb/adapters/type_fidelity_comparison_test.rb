# frozen_string_literal: true

require "test_helper"
require "date"
require "roo"
require "simple_xlsx_reader"
require "xsv"

# Tests verifying type fidelity across official gems and xlsxrb adapters,
# ensuring strings that look like numbers ("0001", "202502"), booleans,
# dates, floats, and integers retain their exact values and classes.
class XlsxrbAdaptersTypeFidelityComparisonTest < Test::Unit::TestCase
  def generate_type_fixture(path)
    Xlsxrb.write(path) do |w|
      w.sheet("Types") do |s|
        s.row([
                "0001",
                "202502",
                "Hello World",
                42,
                3.14,
                Date.new(2026, 2, 1),
                true,
                false,
                ""
              ])
      end
    end
  end

  def test_roo_type_and_class_fidelity_against_official
    with_tempfile do |path|
      generate_type_fixture(path)

      official_roo = ::Roo::Excelx.new(path)
      adapter_roo = ::Xlsxrb::Adapters::Roo::Excelx.new(path)

      (1..9).each do |col|
        off_val = official_roo.cell(1, col)
        adapt_val = adapter_roo.cell(1, col)

        assert_equal off_val, adapt_val, "Roo value mismatch at col #{col}"
        assert_equal off_val.class, adapt_val.class, "Roo class mismatch at col #{col}: expected #{off_val.class}, got #{adapt_val.class}"
      end
    end
  end

  def test_simple_xlsx_reader_type_and_class_fidelity_against_official
    with_tempfile do |path|
      generate_type_fixture(path)

      official_doc = ::SimpleXlsxReader.open(path)
      adapter_doc = ::Xlsxrb::Adapters::SimpleXlsxReader.open(path)

      off_row = official_doc.sheets.first.rows.to_a[0]
      adapt_row = adapter_doc.sheets.first.rows.to_a[0]

      assert_equal off_row.size, adapt_row.size
      off_row.each_with_index do |off_val, idx|
        adapt_val = adapt_row[idx]
        assert_equal off_val, adapt_val, "SimpleXlsxReader value mismatch at index #{idx}"
        assert_equal off_val.class, adapt_val.class, "SimpleXlsxReader class mismatch at index #{idx}"
      end
    end
  end

  def test_xsv_type_and_class_fidelity_against_official
    with_tempfile do |path|
      generate_type_fixture(path)

      official_xsv = ::Xsv::Workbook.open(path)
      adapter_xsv = ::Xlsxrb::Adapters::Xsv::Workbook.open(path)

      off_row = official_xsv.sheets.first.to_a[0]
      adapt_row = adapter_xsv.sheets.first.to_a[0]

      assert_equal off_row.size, adapt_row.size
      off_row.each_with_index do |off_val, idx|
        adapt_val = adapt_row[idx]
        assert_equal off_val, adapt_val, "Xsv value mismatch at index #{idx}"
        assert_equal off_val.class, adapt_val.class, "Xsv class mismatch at index #{idx}"
      end
    end
  end

  def test_ruby_xl_numeric_text_type_preservation
    with_tempfile do |path|
      generate_type_fixture(path)

      adapter_wb = ::Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      cells = adapter_wb["Types"][0].cells.map(&:value)

      # String values that look like numbers must remain String
      assert_equal "0001", cells[0]
      assert_equal String, cells[0].class

      assert_equal "202502", cells[1]
      assert_equal String, cells[1].class

      assert_equal "Hello World", cells[2]
      assert_equal String, cells[2].class

      assert_equal 42, cells[3]
      assert_equal Integer, cells[3].class

      assert_in_delta 3.14, cells[4], 1e-4
      assert_equal Float, cells[4].class

      # Round-trip check: write back and inspect with official rubyXL
      adapter_wb.write(path)
      official_wb = ::RubyXL::Parser.parse(path)
      official_cells = official_wb["Types"][0].cells.map(&:value)

      assert_equal "0001", official_cells[0]
      assert_equal String, official_cells[0].class

      assert_equal "202502", official_cells[1]
      assert_equal String, official_cells[1].class
    end
  end
end
