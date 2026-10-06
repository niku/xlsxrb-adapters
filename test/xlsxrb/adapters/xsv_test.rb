# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersXsvTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def test_version_and_constants
    assert_equal "1.4.1", Xlsxrb::Adapters::Xsv::VERSION
    assert_equal StandardError, Xlsxrb::Adapters::Xsv::Error.superclass
    assert_equal StandardError, Xlsxrb::Adapters::Xsv::DuplicateHeaders.superclass
    assert_equal StandardError, Xlsxrb::Adapters::Xsv::AssertionFailed.superclass
  end

  def test_open_with_filename
    wb = Xlsxrb::Adapters::Xsv.open(fixture("simple_spreadsheet.xlsx"))
    assert_instance_of Xlsxrb::Adapters::Xsv::Workbook, wb
    assert_equal 3, wb.sheets.size
    assert_equal %w[Sheet1 Sheet2 Sheet3], wb.sheets.map(&:name)
    wb.close
  end

  def test_open_with_block
    closed = false
    result = Xlsxrb::Adapters::Xsv.open(fixture("simple_spreadsheet.xlsx")) do |wb|
      assert_instance_of Xlsxrb::Adapters::Xsv::Workbook, wb
      closed = true
      wb.sheets.size
    end
    assert closed
    assert_equal 3, result
  end

  def test_open_with_io_and_buffer
    bin = File.binread(fixture("simple_spreadsheet.xlsx"))

    # Buffer string
    wb1 = Xlsxrb::Adapters::Xsv.open(bin)
    assert_equal 3, wb1.sheets.size
    wb1.close

    # StringIO
    wb2 = Xlsxrb::Adapters::Xsv.open(StringIO.new(bin))
    assert_equal 3, wb2.sheets.size
    wb2.close

    # File IO
    File.open(fixture("simple_spreadsheet.xlsx"), "rb") do |f|
      wb3 = Xlsxrb::Adapters::Xsv.open(f)
      assert_equal 3, wb3.sheets.size
      wb3.close
    end
  end

  def test_open_error_cases
    err_class = defined?(::Zip::Error) ? ::Zip::Error : Xlsxrb::Adapters::Xsv::Error

    assert_raise(err_class) do
      Xlsxrb::Adapters::Xsv.open("non_existent_file.xlsx")
    end

    with_tempfile do |filepath|
      File.write(filepath, "")
      assert_raise(err_class) do
        Xlsxrb::Adapters::Xsv.open(filepath)
      end
    end
  end

  def test_workbook_accessors_and_indexing
    wb = Xlsxrb::Adapters::Xsv.open(fixture("simple_spreadsheet.xlsx"))

    # Access by integer index
    sheet0 = wb[0]
    assert_not_nil sheet0
    assert_equal "Sheet1", sheet0.name
    assert_nil wb[99]

    # Access by string name
    sheet_named = wb["Sheet1"]
    assert_not_nil sheet_named
    assert_equal "Sheet1", sheet_named.name
    assert_nil wb["NonExistent"]

    # Invalid type
    err = assert_raise(ArgumentError) do
      wb[:invalid_sym]
    end
    assert_equal "Sheets can be accessed by Integer of String only", err.message

    # sheets_by_name
    assert_equal 1, wb.sheets_by_name("Sheet1").size
    assert_equal 0, wb.sheets_by_name("NonExistent").size

    # Enumerable
    names = wb.map(&:name)
    assert_equal %w[Sheet1 Sheet2 Sheet3], names
    assert_equal "Sheet1", wb.first.name

    # Inspect
    assert_match(/#<Xlsxrb::Adapters::Xsv::Workbook:\d+ sheets=3 trim_empty_rows=false>/, wb.inspect)

    # Close
    assert wb.close
    assert_nil wb.sheets
  end

  def test_sheet_attributes_and_hidden
    wb = Xlsxrb::Adapters::Xsv.open(fixture("hidden_sheets.xlsx"))
    assert_equal 3, wb.sheets.size

    sheet0 = wb[0]
    sheet1 = wb[1]
    sheet2 = wb[2]

    assert_equal "HiddenSheet1", sheet0.name
    assert_equal true, sheet0.hidden?
    assert_equal :array, sheet0.mode
    assert_equal 0, sheet0.row_skip

    assert_equal "VisibleSheet1", sheet1.name
    assert_equal false, sheet1.hidden?

    assert_equal "HiddenSheet2", sheet2.name
    assert_equal true, sheet2.hidden?

    assert_match(/#<Xlsxrb::Adapters::Xsv::Sheet:\d+ mode=array>/, sheet0.inspect)
    wb.close
  end

  def test_sheet_array_mode_iteration
    wb = Xlsxrb::Adapters::Xsv.open(fixture("numbers1.xlsx"))
    sheet = wb[0]

    rows = sheet.to_a
    assert_equal 18, rows.size
    assert_equal [1, 2, 3, 4, 10, nil, nil], rows[0]
    assert_equal [5, 6, 7, 8, 9, "test", 11], rows[1]
    assert_equal [nil, nil, nil, nil, nil, nil, nil], rows[2] # padded empty row

    # Date conversion via format code
    assert_equal Date.new(1961, 11, 21), rows[4][0]
    assert_equal Date.new(2007, 5, 31), rows[17][0]

    wb.close
  end

  def test_sheet_hash_mode_iteration
    wb = Xlsxrb::Adapters::Xsv.open(fixture("numbers1.xlsx"))
    sheet = wb[0]

    sheet.parse_headers!
    assert_equal :hash, sheet.mode
    assert_equal [1, 2, 3, 4, 10, nil, nil], sheet.headers

    data_rows = sheet.to_a
    assert_equal 17, data_rows.size # 18 total rows - 1 header row = 17

    expected_first = { 1 => 5, 2 => 6, 3 => 7, 4 => 8, 10 => 9 }
    assert_equal expected_first, data_rows[0]

    expected_empty = { 1 => nil, 2 => nil, 3 => nil, 4 => nil, 10 => nil }
    assert_equal expected_empty, data_rows[1]

    wb.close
  end

  def test_sheet_row_skip
    wb = Xlsxrb::Adapters::Xsv.open(fixture("simple_spreadsheet.xlsx"))
    sheet = wb[0]

    sheet.row_skip = 2
    sheet.parse_headers!

    assert_equal ["Date", "Start time", "End time", "Pause", "Sum", "Comment"], sheet.headers.compact

    first_row = sheet.first
    assert_equal Date.new(2007, 5, 7), first_row["Date"]
    assert_equal 9.25, first_row["Start time"]
    assert_equal 10.25, first_row["End time"]
    assert_equal 0, first_row["Pause"]
    assert_equal 1, first_row["Sum"]
    assert_equal "Task 1", first_row["Comment"]

    wb.close
  end

  def test_sheet_duplicate_headers_error
    with_tempfile do |filepath|
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      pkg.workbook.add_worksheet(name: "Dup") do |s|
        s.add_row %w[ID Name Name Age]
        s.add_row [1, "Alice", "Bob", 30]
      end
      pkg.serialize(filepath)

      wb = Xlsxrb::Adapters::Xsv.open(filepath)
      sheet = wb[0]

      err = assert_raise(Xlsxrb::Adapters::Xsv::DuplicateHeaders) do
        sheet.parse_headers!
      end
      assert_match(/Duplicate header 'Name' found/, err.message)
      wb.close

      # Also test parse_headers: true on open
      assert_raise(Xlsxrb::Adapters::Xsv::DuplicateHeaders) do
        Xlsxrb::Adapters::Xsv.open(filepath, parse_headers: true)
      end
    end
  end

  def test_sheet_square_brackets_indexing
    wb = Xlsxrb::Adapters::Xsv.open(fixture("comments.xlsx"))
    sheet = wb[0]

    # Indexing by integer
    row0 = sheet[0]
    assert_equal [nil, nil], row0

    row3 = sheet[3]
    assert_equal [nil, "B4 (mit Kommentar)"], row3

    # Out of bounds returns empty row
    out_of_bounds = sheet[99]
    assert_equal [nil, nil], out_of_bounds

    # Range indexing
    range_rows = sheet[2..4]
    assert_equal 3, range_rows.size
    assert_equal [nil, nil], range_rows[0]
    assert_equal [nil, "B4 (mit Kommentar)"], range_rows[1]
    assert_equal [nil, "B5 (mit Kommentar)"], range_rows[2]

    # Invalid type
    assert_raise(ArgumentError) do
      sheet["invalid"]
    end

    wb.close
  end

  def test_helpers_calculations
    dummy = Class.new { include Xlsxrb::Adapters::Xsv::Helpers }.new

    # column_index
    assert_equal 0, dummy.column_index("A")
    assert_equal 0, dummy.column_index("A1")
    assert_equal 1, dummy.column_index("B4")
    assert_equal 25, dummy.column_index("Z100")
    assert_equal 26, dummy.column_index("AA1")
    assert_equal 255, dummy.column_index("IV19")

    # parse_date
    assert_equal Date.new(2007, 5, 7), dummy.parse_date(39_209)

    # parse_time
    assert_equal "12:00", dummy.parse_time(0.5)
    assert_equal "09:30", dummy.parse_time(0.3958333333333333)

    # parse_datetime
    dt = dummy.parse_datetime(39_209.5)
    assert_instance_of Time, dt
    assert_equal 2007, dt.year
    assert_equal 5, dt.month
    assert_equal 7, dt.day
    assert_equal 12, dt.hour
    assert_equal 0, dt.min

    # parse_number
    assert_equal 42, dummy.parse_number("42")
    assert_equal 3.14, dummy.parse_number("3.14")
    assert_equal 1000.0, dummy.parse_number("1E+03")
    assert_equal 10, dummy.parse_number(10)

    # parse_number_format
    assert_equal Date.new(2007, 5, 7), dummy.parse_number_format(39_209, "yyyy-mm-dd")
    assert_equal "12:00", dummy.parse_number_format(0.5, "hh:mm")
    assert_equal 123.45, dummy.parse_number_format("123.45", "#,##0.00")
  end

  def test_roundtrip_to_and_from_xlsxrb
    wb_xsv = Xlsxrb::Adapters::Xsv.open(fixture("comments.xlsx"))

    # Convert to xlsxrb Elements::Workbook
    xlsxrb_wb = wb_xsv.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
    assert_equal 3, xlsxrb_wb.sheets.size
    assert_equal "Sheet1", xlsxrb_wb.sheets[0].name

    # Convert back from xlsxrb Elements::Workbook
    wb_back = Xlsxrb::Adapters::Xsv.from_xlsxrb(xlsxrb_wb)
    assert_instance_of Xlsxrb::Adapters::Xsv::Workbook, wb_back
    assert_equal 3, wb_back.sheets.size
    assert_equal [nil, "B4 (mit Kommentar)"], wb_back[0][3]

    wb_xsv.close
    wb_back.close
  end

  def test_dimension_and_bounds
    wb = Xlsxrb::Adapters::Xsv.open(fixture("comments.xlsx"))
    sheet = wb[0]

    assert_equal "B4:B5", sheet.dimension
    assert_equal 4, sheet.first_row
    assert_equal 2, sheet.first_column
    assert_equal 2, sheet.first_col
    assert_equal 5, sheet.last_row
    assert_equal 2, sheet.last_column
    assert_equal 2, sheet.last_col
    assert_equal false, sheet.date1904?
    assert_equal false, wb.date1904?
    assert_equal false, sheet.trim_empty_rows?
    assert_equal false, wb.trim_empty_rows?
    wb.close

    wb2 = Xlsxrb::Adapters::Xsv.open(fixture("simple_spreadsheet.xlsx"))
    sheet2 = wb2[0]
    assert_equal "A1:IV19", sheet2.dimension
    assert_equal 1, sheet2.first_row
    assert_equal 1, sheet2.first_col
    assert_equal 19, sheet2.last_row
    assert_equal 256, sheet2.last_col
    wb2.close
  end

  def test_date1904_system
    with_tempfile do |filepath|
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      pkg.workbook.date1904 = true
      pkg.workbook.add_worksheet(name: "Dates") do |s|
        s.add_row [Date.new(2026, 1, 1)]
      end
      pkg.serialize(filepath)

      wb = Xlsxrb::Adapters::Xsv.open(filepath)
      assert_equal true, wb.date1904?
      assert_equal true, wb[0].date1904?
      row0 = wb[0][0]
      assert_equal Date.new(2026, 1, 1), row0[0]

      wb_xlsxrb = wb.to_xlsxrb
      assert_equal true, wb_xlsxrb.sheets[0].date1904?
      wb.close
    end
  end

  def test_each_row_values
    wb = Xlsxrb::Adapters::Xsv.open(fixture("comments.xlsx"))
    sheet = wb[0]

    values = sheet.each_row_values.to_a
    assert_equal 2, values.size
    assert_equal [nil, "B4 (mit Kommentar)"], values[0]
    assert_equal [nil, "B5 (mit Kommentar)"], values[1]
    wb.close
  end

  def test_trim_empty_rows
    wb_no_trim = Xlsxrb::Adapters::Xsv.open(fixture("simple_spreadsheet.xlsx"), trim_empty_rows: false)
    assert_equal false, wb_no_trim.trim_empty_rows?
    assert_equal 19, wb_no_trim[0].last_row
    wb_no_trim.close

    wb_trim = Xlsxrb::Adapters::Xsv.open(fixture("simple_spreadsheet.xlsx"), trim_empty_rows: true)
    assert_equal true, wb_trim.trim_empty_rows?
    assert_equal 13, wb_trim[0].last_row
    wb_trim.close
  end

  def test_whitespace_and_newlines_in_cell_tags
    sheet_xml = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row r="1">
            <c r="A1" t="str">
              <f>SUM(1, 2)</f>
              <v>
                3
              </v>
            </c>
            <c r="B1" t="inlineStr">
              <is>
                <t>Hello World</t>
              </is>
            </c>
          </row>
        </sheetData>
      </worksheet>
    XML

    sheet = Xlsxrb::Adapters::Xsv::Sheet.new(nil, sheet_xml, { sheet_id: 1, name: "Test" })
    rows = sheet.to_a
    assert_equal 1, rows.size
    assert_equal ["3", "Hello World"], rows[0]
  end

  def test_omitted_row_r_attribute_sequential_fallback
    sheet_xml = <<~XML
      <?xml version="1.0" encoding="UTF-8" standalone="yes"?>
      <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
        <sheetData>
          <row>
            <c r="A1" t="str"><v>Row 1</v></c>
          </row>
          <row>
            <c r="A2" t="str"><v>Row 2</v></c>
          </row>
          <row r="5">
            <c r="A5" t="str"><v>Row 5</v></c>
          </row>
          <row>
            <c r="A6" t="str"><v>Row 6</v></c>
          </row>
        </sheetData>
      </worksheet>
    XML

    sheet = Xlsxrb::Adapters::Xsv::Sheet.new(nil, sheet_xml, { sheet_id: 1, name: "Test" })
    rows = sheet.to_a
    assert_equal 6, rows.size
    assert_equal ["Row 1"], rows[0]
    assert_equal ["Row 2"], rows[1]
    assert_equal [nil], rows[2]
    assert_equal [nil], rows[3]
    assert_equal ["Row 5"], rows[4]
    assert_equal ["Row 6"], rows[5]
  end
end
