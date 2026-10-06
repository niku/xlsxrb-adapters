# frozen_string_literal: true

require "test_helper"
require "simple_xlsx_reader"

class XlsxrbAdaptersSimpleXlsxReaderCompatibilityTest < Test::Unit::TestCase
  OFFICIAL_FIXTURES = %w[
    chunky_utf8.xlsx
    date1904.xlsx
    datetimes.xlsx
    gdocs_sheet.xlsx
    lower_case_sharedstrings.xlsx
    misc_numbers.xlsx
    percentages_n_currencies.xlsx
    sesame_street_blog.xlsx
  ].freeze

  ROO_FIXTURES = %w[
    comments.xlsx
    formula.xlsx
    hidden_sheets.xlsx
    link.xlsx
    named_cells.xlsx
    numbers1.xlsx
    paragraph.xlsx
    simple_spreadsheet.xlsx
    style.xlsx
    whitespace.xlsx
  ].freeze

  def fixture(name)
    File.expand_path("../../fixtures/simple_xlsx_reader/#{name}", __dir__)
  end

  def roo_fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def test_side_by_side_official_fixtures
    OFFICIAL_FIXTURES.each do |filename|
      path = fixture(filename)
      orig_doc = SimpleXlsxReader.open(path)
      adapt_doc = Xlsxrb::Adapters::SimpleXlsxReader.open(path)

      assert_equal orig_doc.sheets.size, adapt_doc.sheets.size, "#{filename}: sheet count mismatch"
      assert_equal orig_doc.sheets.map(&:name), adapt_doc.sheets.map(&:name), "#{filename}: sheet names mismatch"

      orig_doc.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_doc.sheets[idx]
        assert_equal orig_sheet.name, adapt_sheet.name

        orig_rows = orig_sheet.rows.to_a
        adapt_rows = adapt_sheet.rows.to_a

        assert_equal orig_rows.size, adapt_rows.size, "#{filename} [#{orig_sheet.name}]: row count mismatch"
        orig_rows.each_with_index do |orig_row, r_idx|
          adapt_row = adapt_rows[r_idx]
          assert_equal orig_row, adapt_row, "#{filename} [#{orig_sheet.name}] row #{r_idx} mismatch"
        end
      end

      assert_equal orig_doc.to_hash, adapt_doc.to_hash, "#{filename}: to_hash mismatch"
    end
  end

  def test_side_by_side_roo_fixtures
    ROO_FIXTURES.each do |filename|
      path = roo_fixture(filename)
      orig_doc = SimpleXlsxReader.open(path)
      adapt_doc = Xlsxrb::Adapters::SimpleXlsxReader.open(path)

      assert_equal orig_doc.sheets.size, adapt_doc.sheets.size, "#{filename}: sheet count mismatch"
      assert_equal orig_doc.sheets.map(&:name), adapt_doc.sheets.map(&:name), "#{filename}: sheet names mismatch"

      orig_doc.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_doc.sheets[idx]
        assert_equal orig_sheet.name, adapt_sheet.name

        orig_rows = orig_sheet.rows.to_a
        adapt_rows = adapt_sheet.rows.to_a

        assert_equal orig_rows.size, adapt_rows.size, "#{filename} [#{orig_sheet.name}]: row count mismatch"
        orig_rows.each_with_index do |orig_row, r_idx|
          adapt_row = adapt_rows[r_idx]
          assert_equal orig_row, adapt_row, "#{filename} [#{orig_sheet.name}] row #{r_idx} mismatch"
        end
      end

      assert_equal orig_doc.to_hash, adapt_doc.to_hash, "#{filename}: to_hash mismatch"
    end
  end

  def test_side_by_side_headers_and_slurp
    OFFICIAL_FIXTURES.each do |filename|
      path = fixture(filename)
      orig_doc = SimpleXlsxReader.open(path)
      adapt_doc = Xlsxrb::Adapters::SimpleXlsxReader.open(path)

      orig_doc.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_doc.sheets[idx]

        orig_slurped = orig_sheet.slurp
        adapt_slurped = adapt_sheet.slurp
        assert_equal orig_slurped, adapt_slurped, "#{filename} [#{orig_sheet.name}]: slurp mismatch"

        assert_equal orig_sheet.headers, adapt_sheet.headers, "#{filename} [#{orig_sheet.name}]: headers mismatch"
        assert_equal orig_sheet.data, adapt_sheet.data, "#{filename} [#{orig_sheet.name}]: data mismatch"
      end
    end
  end

  def test_side_by_side_each_headers_true
    %w[percentages_n_currencies.xlsx misc_numbers.xlsx].each do |filename|
      path = fixture(filename)
      orig_doc = SimpleXlsxReader.open(path)
      adapt_doc = Xlsxrb::Adapters::SimpleXlsxReader.open(path)

      orig_rows = orig_doc.sheets.first.rows.each(headers: true).to_a
      adapt_rows = adapt_doc.sheets.first.rows.each(headers: true).to_a
      assert_equal orig_rows, adapt_rows, "#{filename}: each(headers: true) mismatch"
    end
  end

  def test_side_by_side_parse_buffer_and_io
    content = File.binread(fixture("percentages_n_currencies.xlsx"))

    orig_buf = SimpleXlsxReader.parse(content)
    adapt_buf = Xlsxrb::Adapters::SimpleXlsxReader.parse(content)
    assert_equal orig_buf.sheets.map(&:name), adapt_buf.sheets.map(&:name)
    assert_equal orig_buf.sheets.first.rows.to_a, adapt_buf.sheets.first.rows.to_a

    orig_io = SimpleXlsxReader.parse(StringIO.new(content))
    adapt_io = Xlsxrb::Adapters::SimpleXlsxReader.parse(StringIO.new(content))
    assert_equal orig_io.sheets.map(&:name), adapt_io.sheets.map(&:name)
    assert_equal orig_io.sheets.first.rows.to_a, adapt_io.sheets.first.rows.to_a
  end

  # Official simple_xlsx_reader has a known bug in its SAX </t> handler for inlineStr cells:
  # it resets @captured on each </t>, overwriting earlier runs in multi-run inline rich text.
  # Xlsxrb adapter preserves and concatenates all text runs, correctly recovering "Example richtext".
  def test_inline_richtext_handling
    path = roo_fixture("richtext_example.xlsx")
    orig_doc = SimpleXlsxReader.open(path)
    adapt_doc = Xlsxrb::Adapters::SimpleXlsxReader.open(path)

    # First cell is single-run rich text
    assert_equal orig_doc.sheets.first.rows.to_a[0][0], adapt_doc.sheets.first.rows.to_a[0][0]
    assert_equal "Example richtext", adapt_doc.sheets.first.rows.to_a[0][0]

    # Second cell has two runs: <r><t>Example</t></r><r><t> richtext</t></r>
    # Official gem erroneously clobbers first run to " richtext", adapter properly yields "Example richtext"
    assert_equal " richtext", orig_doc.sheets.first.rows.to_a[0][1]
    assert_equal "Example richtext", adapt_doc.sheets.first.rows.to_a[0][1]
  end
end
