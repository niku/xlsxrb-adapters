# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersSimpleXlsxReaderTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/simple_xlsx_reader/#{name}", __dir__)
  end

  def roo_fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def teardown
    Xlsxrb::Adapters::SimpleXlsxReader.configuration.auto_slurp = false
    Xlsxrb::Adapters::SimpleXlsxReader.configuration.catch_cell_load_errors = false
  end

  def test_version_and_constants
    assert_equal "5.1.0", Xlsxrb::Adapters::SimpleXlsxReader::VERSION
    assert_equal StandardError, Xlsxrb::Adapters::SimpleXlsxReader::CellLoadError.superclass
    assert defined?(Xlsxrb::Adapters::SimpleXlsxReader::Document)
    assert defined?(Xlsxrb::Adapters::SimpleXlsxReader::Document::Sheet)
    assert defined?(Xlsxrb::Adapters::SimpleXlsxReader::Sheet)
    assert defined?(Xlsxrb::Adapters::SimpleXlsxReader::Loader)
    assert defined?(Xlsxrb::Adapters::SimpleXlsxReader::Hyperlink)
  end

  def test_configuration
    cfg = Xlsxrb::Adapters::SimpleXlsxReader.configuration
    assert_equal false, cfg.auto_slurp
    assert_equal false, cfg.catch_cell_load_errors

    cfg.auto_slurp = true
    assert_equal true, cfg.auto_slurp

    cfg.catch_cell_load_errors = true
    assert_equal true, cfg.catch_cell_load_errors
  ensure
    Xlsxrb::Adapters::SimpleXlsxReader.configuration.auto_slurp = false
    Xlsxrb::Adapters::SimpleXlsxReader.configuration.catch_cell_load_errors = false
  end

  def test_open_with_filename
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("misc_numbers.xlsx"))
    assert_instance_of Xlsxrb::Adapters::SimpleXlsxReader::Document, doc
    assert_equal 1, doc.sheets.size
    sheet = doc.sheets.first
    assert_equal "Sheet1", sheet.name

    rows = sheet.rows.to_a
    assert_equal 2, rows.size
    assert_equal ["Medium number", "Big Number"], rows[0]
    assert_equal [98_070, 1_234_567_890_123], rows[1]
  end

  def test_open_with_block
    sheet_count = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("misc_numbers.xlsx")) do |doc|
      assert_instance_of Xlsxrb::Adapters::SimpleXlsxReader::Document, doc
      doc.sheets.size
    end
    assert_equal 1, sheet_count
  end

  def test_parse_with_string_and_io
    content = File.binread(fixture("misc_numbers.xlsx"))

    # Buffer string
    doc1 = Xlsxrb::Adapters::SimpleXlsxReader.parse(content)
    assert_equal 1, doc1.sheets.size
    assert_equal ["Medium number", "Big Number"], doc1.sheets.first.rows.first

    # StringIO
    doc2 = Xlsxrb::Adapters::SimpleXlsxReader.parse(StringIO.new(content))
    assert_equal 1, doc2.sheets.size
    assert_equal ["Medium number", "Big Number"], doc2.sheets.first.rows.first

    # File IO
    File.open(fixture("misc_numbers.xlsx"), "rb") do |f|
      doc3 = Xlsxrb::Adapters::SimpleXlsxReader.parse(f)
      assert_equal 1, doc3.sheets.size
      assert_equal ["Medium number", "Big Number"], doc3.sheets.first.rows.first
    end
  end

  def test_open_error_cases
    assert_raise(ArgumentError) do
      Xlsxrb::Adapters::SimpleXlsxReader::Document.new
    end

    assert_raise(Errno::ENOENT) do
      Xlsxrb::Adapters::SimpleXlsxReader.open("non_existent_file.xlsx")
    end
  end

  def test_to_hash
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("misc_numbers.xlsx"))
    h = doc.to_hash
    assert_instance_of Hash, h
    assert_equal ["Sheet1"], h.keys
    assert_equal 2, h["Sheet1"].size
    assert_equal ["Medium number", "Big Number"], h["Sheet1"][0]
  end

  def test_sheet_slurp_and_slurped
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("misc_numbers.xlsx"))
    sheet = doc.sheets.first
    assert_equal false, sheet.rows.slurped?
    rows = sheet.slurp
    assert_equal true, sheet.rows.slurped?
    assert_equal 2, rows.size
    assert_equal rows, sheet.rows.to_a
  end

  def test_auto_slurp_configuration
    Xlsxrb::Adapters::SimpleXlsxReader.configuration.auto_slurp = true

    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("misc_numbers.xlsx"))
    sheet = doc.sheets.first
    assert_equal false, sheet.rows.slurped?
    assert_equal ["Medium number", "Big Number"], sheet.headers
    assert_equal true, sheet.rows.slurped?
    assert_equal 2, sheet.rows.slurped.size
  end

  def test_headers_access
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("percentages_n_currencies.xlsx"))
    sheet = doc.sheets.first
    assert_raise(RuntimeError) { sheet.headers }
    sheet.slurp
    assert_equal ["Name", "Class", "Score % ", "Score Actual", "Salary", "My vals", nil], sheet.headers
  end

  def test_each_with_headers_true
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("percentages_n_currencies.xlsx"))
    sheet = doc.sheets.first

    rows = []
    sheet.rows.each(headers: true) { |row| rows << row }

    assert_equal 3, rows.size
    assert_equal "Ben", rows[0]["Name"]
    assert_equal 0.87, rows[0]["Score % "]
  end

  def test_each_with_headers_proc
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("percentages_n_currencies.xlsx"))
    sheet = doc.sheets.first

    finder = ->(row) { row.include?("Class") }
    rows = sheet.rows.each(headers: finder).to_a

    assert_equal 3, rows.size
    assert_equal "Ben", rows[0]["Name"]
  end

  def test_each_with_headers_hash_mapping
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("percentages_n_currencies.xlsx"))
    sheet = doc.sheets.first

    header_map = { name: "Name", score: /Score %/ }
    rows = sheet.rows.each(headers: header_map).to_a

    assert_equal 3, rows.size
    assert_equal "Ben", rows[0][:name]
    assert_equal 0.87, rows[0][:score]
    assert_equal "A", rows[0]["Class"]
  end

  def test_date_systems
    # 1900 date system
    doc1900 = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("datetimes.xlsx"))
    sheet1900 = doc1900.sheets.first
    row1900 = sheet1900.rows.to_a[1]
    assert_equal Time.utc(2013, 8, 19, 18, 30, 0), row1900[0]

    # 1904 date system
    doc1904 = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("date1904.xlsx"))
    sheet1904 = doc1904.sheets.first
    row1904 = sheet1904.rows.to_a[0]
    assert_equal Date.new(2014, 5, 1), row1904[0]
  end

  def test_hyperlinks
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(roo_fixture("link.xlsx"))
    sheet = doc.sheets.first
    rows = sheet.rows.to_a

    link_cell = rows[0][0]
    assert_kind_of String, link_cell
    assert_equal "Google", link_cell.to_s
    assert_equal "http://www.google.com", link_cell.url
  end

  def test_hyperlink_class
    hl = Xlsxrb::Adapters::SimpleXlsxReader::Hyperlink.new("https://example.com", "Example")
    assert_kind_of String, hl
    assert_equal "Example", hl.to_s
    assert_equal "Example", hl.friendly_name
    assert_equal "https://example.com", hl.url

    hl2 = Xlsxrb::Adapters::SimpleXlsxReader::Hyperlink.new("https://example.com")
    assert_equal "https://example.com", hl2.to_s
    assert_nil hl2.friendly_name
    assert_equal "https://example.com", hl2.url
  end

  def test_catch_cell_load_errors
    Xlsxrb::Adapters::SimpleXlsxReader.configuration.catch_cell_load_errors = true

    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("misc_numbers.xlsx"))
    sheet = doc.sheets.first
    rows = sheet.rows.to_a
    assert_equal 2, rows.size
    assert_equal({}, sheet.rows.load_errors)
  end

  def test_to_xlsxrb_roundtrip
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("misc_numbers.xlsx"))
    wb = doc.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Workbook, wb
    assert_equal 1, wb.sheets.size
    assert_equal "Sheet1", wb.sheets.first.name

    # Check reading cells from Xlsxrb::Elements::Workbook
    ws = wb.sheets.first
    assert_equal "Medium number", ws.cell(0, 0)&.value
  end

  def test_chunky_utf8_fixture
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("chunky_utf8.xlsx"))
    sheet = doc.sheets.first
    rows = sheet.rows.to_a
    assert_equal 2, rows.size
    assert_equal ["Company Name", "City", "Street Name"], rows[0]
    assert_equal ["sample-company-1", "Korntal-Münchingen", "Bronholmer straße"], rows[1]
  end

  def test_lower_case_sharedstrings_with_backslash_paths
    doc = Xlsxrb::Adapters::SimpleXlsxReader.open(fixture("lower_case_sharedstrings.xlsx"))
    assert_equal 2, doc.sheets.size
    assert_equal ["0", "Run Information"], doc.sheets.map(&:name)
    sheet1 = doc.sheets[0]
    assert sheet1.rows.to_a.size > 10
    sheet2 = doc.sheets[1]
    assert_equal 13, sheet2.rows.to_a.size
    assert_equal "File Name", sheet2.rows.to_a[0][0]
  end
end
