# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersXsvCompatibilityTest < Test::Unit::TestCase
  FIXTURES = %w[
    comments.xlsx
    formula.xlsx
    hidden_sheets.xlsx
    link.xlsx
    named_cells.xlsx
    numbers1.xlsx
    paragraph.xlsx
    richtext_example.xlsx
    simple_spreadsheet.xlsx
    style.xlsx
    whitespace.xlsx
  ].freeze

  def fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def test_side_by_side_array_mode_all_fixtures
    FIXTURES.each do |filename|
      path = fixture(filename)
      orig_wb = Xsv.open(path)
      adapt_wb = Xlsxrb::Adapters::Xsv.open(path)

      assert_equal orig_wb.sheets.size, adapt_wb.sheets.size, "#{filename}: sheet count mismatch"
      assert_equal orig_wb.sheets.map(&:name), adapt_wb.sheets.map(&:name), "#{filename}: sheet names mismatch"

      orig_wb.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_wb[idx]
        assert_equal orig_sheet.name, adapt_sheet.name
        assert_equal orig_sheet.hidden?, adapt_sheet.hidden?, "#{filename} [#{orig_sheet.name}]: hidden mismatch"
        assert_equal orig_sheet.id, adapt_sheet.id, "#{filename} [#{orig_sheet.name}]: id mismatch"

        orig_rows = orig_sheet.to_a
        adapt_rows = adapt_sheet.to_a

        assert_equal orig_rows.size, adapt_rows.size, "#{filename} [#{orig_sheet.name}]: row count mismatch"
        orig_rows.each_with_index do |orig_row, r_idx|
          adapt_row = adapt_rows[r_idx]
          assert_equal orig_row, adapt_row, "#{filename} [#{orig_sheet.name}] row #{r_idx} mismatch"
        end
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_hash_mode_all_fixtures
    FIXTURES.each do |filename|
      path = fixture(filename)

      # Attempt opening in hash mode; skip if duplicate headers are present in the fixture
      orig_error = nil
      adapt_error = nil

      begin
        orig_wb = Xsv.open(path, parse_headers: true)
      rescue StandardError => e
        orig_error = e.class
      end

      begin
        adapt_wb = Xlsxrb::Adapters::Xsv.open(path, parse_headers: true)
      rescue StandardError => e
        adapt_error = e.class
      end

      orig_err_name = orig_error ? orig_error.name.split("::").last : nil
      adapt_err_name = adapt_error ? adapt_error.name.split("::").last : nil
      assert_equal orig_err_name, adapt_err_name, "#{filename}: parse_headers error class mismatch"

      next unless orig_error.nil? && adapt_error.nil?

      orig_wb.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_wb[idx]
        assert_equal orig_sheet.headers, adapt_sheet.headers, "#{filename} [#{orig_sheet.name}]: headers mismatch"

        orig_rows = orig_sheet.to_a
        adapt_rows = adapt_sheet.to_a

        assert_equal orig_rows.size, adapt_rows.size, "#{filename} [#{orig_sheet.name}]: hash row count mismatch"
        orig_rows.each_with_index do |orig_row, r_idx|
          adapt_row = adapt_rows[r_idx]
          assert_equal orig_row, adapt_row, "#{filename} [#{orig_sheet.name}] hash row #{r_idx} mismatch"
        end
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_trim_empty_rows
    FIXTURES.each do |filename|
      path = fixture(filename)
      orig_wb = Xsv.open(path, trim_empty_rows: true)
      adapt_wb = Xlsxrb::Adapters::Xsv.open(path, trim_empty_rows: true)

      orig_wb.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_wb[idx]
        orig_rows = orig_sheet.to_a
        adapt_rows = adapt_sheet.to_a

        assert_equal orig_rows.size, adapt_rows.size, "#{filename} [#{orig_sheet.name}] (trimmed): row count mismatch"
        orig_rows.each_with_index do |orig_row, r_idx|
          adapt_row = adapt_rows[r_idx]
          assert_equal orig_row, adapt_row, "#{filename} [#{orig_sheet.name}] (trimmed) row #{r_idx} mismatch"
        end
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_row_skip
    path = fixture("simple_spreadsheet.xlsx")
    orig_wb = Xsv.open(path)
    adapt_wb = Xlsxrb::Adapters::Xsv.open(path)

    orig_sheet = orig_wb[0]
    adapt_sheet = adapt_wb[0]

    orig_sheet.row_skip = 2
    adapt_sheet.row_skip = 2

    assert_equal orig_sheet.to_a, adapt_sheet.to_a

    orig_sheet.parse_headers!
    adapt_sheet.parse_headers!

    assert_equal orig_sheet.headers, adapt_sheet.headers
    assert_equal orig_sheet.to_a, adapt_sheet.to_a

    orig_wb.close
    adapt_wb.close
  end

  def test_side_by_side_indexing_and_ranges
    path = fixture("numbers1.xlsx")
    orig_wb = Xsv.open(path)
    adapt_wb = Xlsxrb::Adapters::Xsv.open(path)

    orig_sheet = orig_wb[0]
    adapt_sheet = adapt_wb[0]

    # Row access by integer
    [0, 1, 2, 4, 17, 18, 99, -1].each do |idx|
      assert_equal orig_sheet[idx], adapt_sheet[idx], "Row index #{idx} mismatch"
    end

    # Row access by range
    [0..1, 2..5, 0..17, 10..20].each do |range|
      assert_equal orig_sheet[range], adapt_sheet[range], "Row range #{range} mismatch"
    end

    orig_wb.close
    adapt_wb.close
  end
end
