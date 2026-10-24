# frozen_string_literal: true

require "test_helper"
require "date"
require "roo"
require "simple_xlsx_reader"
require "xsv"
require "creek"

# Comprehensive differential test suite comparing official gems vs xlsxrb adapters
# across primitive types, numeric-like strings, dates, blank cells, and sparse layouts.
class XlsxrbAdaptersDifferentialFuzzingTest < Test::Unit::TestCase
  def build_matrix_values
    [
      # Numeric edge cases
      0, 1, -1, 42, -42, 2_147_483_647, -2_147_483_648, 9_007_199_254_740_991,
      0.0, -0.0, 3.1415926535, 1.23e-4, 1e8,
      # Booleans
      true, false,
      # Strings and special cases
      "", "   ", "  trimmed  ", "0001", "012345", "202502", "2026-10-24", "1e5", "true", "FALSE",
      "line1\nline2", "A & B", "<foo>bar</foo>", "\"quoted\"", "日本語テスト", "🚀 Rocket 🎉",
      # Dates
      Date.new(1900, 1, 1), Date.new(1900, 2, 28), Date.new(2026, 10, 24), Date.new(2099, 12, 31)
    ]
  end

  def test_differential_roo_against_official
    values = build_matrix_values

    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("Sheet1") do |s|
          values.each_with_index do |v, idx|
            s.row(["k_#{idx}", v])
          end
        end
      end

      official_roo = ::Roo::Excelx.new(path)
      adapter_roo = ::Xlsxrb::Adapters::Roo::Excelx.new(path)

      values.each_with_index do |_v, idx|
        row = idx + 1
        off_val = official_roo.cell(row, 2)
        ad_val = adapter_roo.cell(row, 2)
        assert_equal off_val, ad_val, "Roo value mismatch at row #{row}"
        assert_equal official_roo.celltype(row, 2), adapter_roo.celltype(row, 2), "Roo celltype mismatch at row #{row}"
      end
    end
  end

  def test_differential_simple_xlsx_reader_against_official
    values = build_matrix_values

    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("Sheet1") do |s|
          values.each_with_index do |v, idx|
            s.row(["k_#{idx}", v])
          end
        end
      end

      official_doc = ::SimpleXlsxReader.open(path)
      adapter_doc = ::Xlsxrb::Adapters::SimpleXlsxReader.open(path)

      off_rows = official_doc.sheets.first.rows.to_a
      ad_rows = adapter_doc.sheets.first.rows.to_a

      assert_equal off_rows.size, ad_rows.size
      values.each_with_index do |_v, idx|
        assert_equal off_rows[idx][1], ad_rows[idx][1], "SimpleXlsxReader mismatch at row #{idx}"
        assert_equal off_rows[idx][1].class, ad_rows[idx][1].class, "SimpleXlsxReader class mismatch at row #{idx}"
      end
    end
  end

  def test_differential_creek_against_official
    values = build_matrix_values

    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("Sheet1") do |s|
          values.each_with_index do |v, idx|
            s.row(["k_#{idx}", v])
          end
        end
      end

      official_book = ::Creek::Book.new(path)
      adapter_book = ::Xlsxrb::Adapters::Creek::Book.new(path)

      off_rows = official_book.sheets.first.rows.to_a
      ad_rows = adapter_book.sheets.first.rows.to_a

      assert_equal off_rows.size, ad_rows.size
      values.each_with_index do |_v, idx|
        off_v = off_rows[idx] ? off_rows[idx].values[1] : nil
        ad_v = ad_rows[idx] ? ad_rows[idx].values[1] : nil
        assert_equal off_v, ad_v, "Creek mismatch at row #{idx}"
      end
    end
  end

  def test_differential_sparse_and_skipped_cells
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("Sparse") do |s|
          s.row(["Row 1", nil, "Col C"])
          s.row(["Row 2", "Col B", nil, "Col D"])
        end
      end

      # Roo
      off_r = ::Roo::Excelx.new(path)
      ad_r = ::Xlsxrb::Adapters::Roo::Excelx.new(path)
      assert_equal off_r.last_row, ad_r.last_row
      assert_equal off_r.cell(1, 2), ad_r.cell(1, 2)
      assert_equal off_r.cell(1, 3), ad_r.cell(1, 3)

      # SimpleXlsxReader
      off_s = ::SimpleXlsxReader.open(path).sheets.first.rows.to_a
      ad_s = ::Xlsxrb::Adapters::SimpleXlsxReader.open(path).sheets.first.rows.to_a
      assert_equal off_s, ad_s
    end
  end
end
