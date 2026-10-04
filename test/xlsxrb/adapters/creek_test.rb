# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCreekTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/creek/#{name}", __dir__)
  end

  def test_version_and_constants
    assert_equal "2.6.3", Xlsxrb::Adapters::Creek::VERSION
    assert defined?(Xlsxrb::Adapters::Creek::Book)
    assert defined?(Xlsxrb::Adapters::Creek::Sheet)
    assert defined?(Xlsxrb::Adapters::Creek::Styles)
    assert defined?(Xlsxrb::Adapters::Creek::Styles::Constants)
    assert defined?(Xlsxrb::Adapters::Creek::Styles::StyleTypes)
    assert defined?(Xlsxrb::Adapters::Creek::Styles::Converter)
    assert defined?(Xlsxrb::Adapters::Creek::Drawing)
    assert defined?(Xlsxrb::Adapters::Creek::SharedStrings)
    assert defined?(Xlsxrb::Adapters::Creek::Utils)
  end

  def test_open_with_filename
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample.xlsx"))
    assert_instance_of Xlsxrb::Adapters::Creek::Book, book
    assert_equal 1, book.sheets.size
    sheet = book.sheets.first
    assert_equal "Sheet1", sheet.name
    assert_equal "rId1", sheet.rid
    assert_nil sheet.state
    assert_nil sheet.sheetid
    book.close
  end

  def test_open_with_io_and_stringio
    bin = File.binread(fixture("sample.xlsx"))

    # StringIO with original_filename
    sio = StringIO.new(bin)
    book1 = Xlsxrb::Adapters::Creek::Book.new(sio, original_filename: "sample.xlsx")
    assert_equal 1, book1.sheets.size
    assert_equal "Sheet1", book1.sheets.first.name
    book1.close

    # File IO with check_file_extension: false
    File.open(fixture("sample.xlsx"), "rb") do |f|
      book2 = Xlsxrb::Adapters::Creek::Book.new(f, check_file_extension: false)
      assert_equal 1, book2.sheets.size
      book2.close
    end
  end

  def test_file_extension_validation
    assert_raise(RuntimeError, "Not a valid file format.") do
      Xlsxrb::Adapters::Creek::Book.new(fixture("invalid.xls"))
    end

    assert_raise(RuntimeError, "Not a valid file format.") do
      Xlsxrb::Adapters::Creek::Book.new(fixture("invalid.xls"), check_file_extension: true)
    end

    # Can circumvent check with check_file_extension: false
    assert_nothing_raised do
      book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample-as-zip.zip"), check_file_extension: false)
      book.close
    end

    # Check original_filename option
    path = fixture("temp_string_io_file_path_with_no_extension")
    assert_raise(RuntimeError, "Not a valid file format.") do
      Xlsxrb::Adapters::Creek::Book.new(path, original_filename: "invalid.xls")
    end

    assert_nothing_raised do
      book = Xlsxrb::Adapters::Creek::Book.new(path, original_filename: "valid.xlsx")
      book.close
    end
  end

  def test_rows_iteration
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample.xlsx"))
    sheet = book.sheets.first
    rows = sheet.rows.to_a
    assert_equal 9, rows.size

    # Row 0 (A1..E1)
    assert_equal({ "A1" => "Content 1", "B1" => nil, "C1" => "Content 2", "D1" => nil, "E1" => "Content 3" }, rows[0])

    # Row 1 (A2..F2)
    assert_equal({ "A2" => nil, "B2" => "Content 4", "C2" => nil, "D2" => "Content 5", "E2" => nil, "F2" => "Content 6" }, rows[1])

    # Row 2 is an empty self-closing row in the XML: {}
    assert_equal({}, rows[2])

    # Row 3 (A4..F4)
    assert_equal({ "A4" => "Content 7", "B4" => "Content 8", "C4" => "Content 9", "D4" => "Content 10", "E4" => "Content 11", "F4" => "Content 12" }, rows[3])

    # Row 8 (row 10 in sheet)
    assert_equal({ "A10" => 0.15, "B10" => 0.15 }, rows[8])

    book.close
  end

  def test_simple_rows_iteration
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample.xlsx"))
    sheet = book.sheets.first
    rows = sheet.simple_rows.to_a
    assert_equal 9, rows.size

    assert_equal({ "A" => "Content 1", "B" => nil, "C" => "Content 2", "D" => nil, "E" => "Content 3" }, rows[0])
    assert_equal({ "A" => nil, "B" => "Content 4", "C" => nil, "D" => "Content 5", "E" => nil, "F" => "Content 6" }, rows[1])
    assert_equal({}, rows[2])
    assert_equal({ "A" => "Content 7", "B" => "Content 8", "C" => "Content 9", "D" => "Content 10", "E" => "Content 11", "F" => "Content 12" }, rows[3])
    assert_equal({ "A" => 0.15, "B" => 0.15 }, rows[8])

    book.close
  end

  def test_rows_with_meta_data
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample.xlsx"))
    sheet = book.sheets.first
    meta_rows = sheet.rows_with_meta_data.to_a
    assert_equal 9, meta_rows.size

    row0 = meta_rows[0]
    assert_equal "1", row0["r"]
    assert_equal "12", row0["ht"]
    assert_equal "1", row0["customHeight"]
    assert_equal({ "A1" => "Content 1", "B1" => nil, "C1" => "Content 2", "D1" => nil, "E1" => "Content 3" }, row0["cells"])

    row2 = meta_rows[2]
    assert_equal "3", row2["r"]
    assert_equal({}, row2["cells"])

    book.close
  end

  def test_simple_rows_with_meta_data
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample.xlsx"))
    sheet = book.sheets.first
    meta_rows = sheet.simple_rows_with_meta_data.to_a
    assert_equal 9, meta_rows.size

    row0 = meta_rows[0]
    assert_equal "1", row0["r"]
    assert_equal({ "A" => "Content 1", "B" => nil, "C" => "Content 2", "D" => nil, "E" => "Content 3" }, row0["cells"])

    book.close
  end

  def test_with_headers_option
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample-with-headers.xlsx"), with_headers: true)
    sheet = book.sheets.first
    assert_equal true, sheet.with_headers

    rows = sheet.simple_rows.to_a
    assert_equal 15, rows.size

    # Row 0 has raw column letters as header keys
    assert_equal({ "A" => "HeaderA", "B" => "HeaderB", "C" => "HeaderC" }, rows[0])

    # Row 1 maps values to header names
    assert_equal({ "HeaderA" => "value1", "HeaderB" => "value2", "HeaderC" => "value3" }, rows[1])

    # Multiple calls to simple_rows return identical headers mapping
    rows_again = sheet.simple_rows.to_a
    assert_equal rows[1], rows_again[1]

    book.close
  end

  def test_date_and_time_parsing
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample_dates.xlsx"))
    sheet = book.sheets.first
    rows = sheet.rows.to_a

    assert_equal Date.new(2018, 1, 1), rows[2]["B3"]
    assert_equal Time.new(2018, 1, 1, 0, 0, 0), rows[3]["B4"]
    assert_equal Time.new(2018, 1, 1, 23, 59, 59), rows[4]["B5"]

    book.close
  end

  def test_large_numbers_parsing
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("large_numbers.xlsx"))
    sheet = book.sheets.first
    first_row = sheet.simple_rows.first

    assert_equal "7.83294732E8", first_row["A"]
    assert_equal "783294732", first_row["B"]
    assert_equal 783_294_732.0, first_row["C"]

    book.close
  end

  def test_escaped_characters
    book1 = Xlsxrb::Adapters::Creek::Book.new(fixture("escaped.xlsx"))
    rows1 = book1.sheets[0].rows.map(&:values)
    assert_equal [%w[abc def], ["ghi", "j&k"]], rows1
    book1.close

    book2 = Xlsxrb::Adapters::Creek::Book.new(fixture("escaped2.xlsx"))
    rows2 = book2.sheets[0].rows.map(&:values)
    assert_equal [%w[abc def], ["ghi", "j&k"]], rows2
    book2.close
  end

  def test_images_and_drawings
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample-with-images.xlsx"))
    sheet = book.sheets.first
    sheet.with_images

    # images_at A2
    imgs_a2 = sheet.images_at("A2")
    assert_not_nil imgs_a2
    assert_equal 1, imgs_a2.size
    assert_instance_of Pathname, imgs_a2[0]
    assert imgs_a2[0].exist?
    assert_equal "image1.jpeg", imgs_a2[0].basename.to_s

    # Cell without images returns nil
    assert_nil sheet.images_at("B3")

    # Cell with multiple images (A10)
    imgs_a10 = sheet.images_at("A10")
    assert_not_nil imgs_a10
    assert_equal 2, imgs_a10.size

    # Embedded images inside rows
    rows = sheet.rows.to_a
    assert_equal 1, rows[1]["A2"].size
    assert_equal "Fluffy", rows[1]["B2"]

    book.close
  end

  def test_one_cell_anchored_images
    book = Xlsxrb::Adapters::Creek::Book.new(fixture("sample-with-one-cell-anchored-images.xlsx"))
    sheet = book.sheets.first.with_images
    rows = sheet.rows.to_a

    assert_equal 1, rows[1]["A2"].size
    assert_equal 1, rows[2]["A3"].size
    assert_nil sheet.images_at("A4")

    book.close
  end

  def test_styles_converter_standalone
    conv = Xlsxrb::Adapters::Creek::Styles::Converter

    assert_equal Date.new(2013, 1, 1), conv.call("41275", "n", :date)
    assert_equal Time.new(2013, 1, 1, 0, 0, 0), conv.call("41275", "n", :date_time)
    assert_equal 123, conv.call("123", nil, :fixnum)
    assert_equal 12.34, conv.call("12.34", nil, :float)
    assert_equal true, conv.call("1", "b", nil)
    assert_equal false, conv.call("0", "b", nil)
    assert_nil conv.call("", "s", nil)
  end

  def test_styles_styletypes_standalone
    xml_content = File.read(fixture("styles/first.xml"))
    style_types = Xlsxrb::Adapters::Creek::Styles::StyleTypes.new(xml_content).call
    assert_equal 8, style_types.size
    assert_equal :date_time, style_types[3]
    assert_equal %i[unsupported unsupported unsupported date_time unsupported unsupported unsupported unsupported], style_types
  end

  def test_shared_strings_standalone
    xml_content = File.read(fixture("sst.xml"))
    dict = Xlsxrb::Adapters::Creek::SharedStrings.parse_shared_string_from_document(xml_content)
    assert_equal 7, dict.keys.size
    assert_equal "Cell A1", dict[0]
    assert_equal "Cell B1", dict[1]
    assert_equal "My Cell", dict[2]
    assert_equal "Cell A2", dict[3]
    assert_equal "Cell B2", dict[4]
    assert_equal "Cell with\rescaped\rcharacters", dict[5]
    assert_equal "吉田兼好", dict[6]
  end

  def test_base_date_and_date1904
    book1900 = Xlsxrb::Adapters::Creek::Book.new(fixture("sample.xlsx"))
    assert_equal Date.new(1899, 12, 30), book1900.base_date
    book1900.close

    with_tempfile do |filepath|
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      pkg.workbook.date1904 = true
      pkg.workbook.add_worksheet(name: "Dates") do |s|
        s.add_row [Date.new(2026, 1, 1)]
      end
      pkg.serialize(filepath)

      book1904 = Xlsxrb::Adapters::Creek::Book.new(filepath)
      assert_equal Date.new(1904, 1, 1), book1904.base_date
      book1904.close
    end
  end

  def test_to_xlsxrb_and_from_xlsxrb_roundtrip
    with_tempfile do |filepath|
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      ws = pkg.workbook.add_worksheet(name: "Inventory")
      ws.add_row(%w[SKU Product Quantity Price InStock])
      ws.add_row(["A-101", "Widget A", 100, 2.50, true])
      ws.add_row(["B-202", "Widget B", 50, 15.75, false])
      pkg.serialize(filepath)

      book = Xlsxrb::Adapters::Creek::Book.new(filepath)
      rows_orig = book.sheets.first.rows.to_a
      assert_equal 3, rows_orig.size

      # Convert to xlsxrb Elements::Workbook
      xlsxrb_wb = book.to_xlsxrb
      assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
      assert_equal "Inventory", xlsxrb_wb.sheets[0].name

      # Convert back from xlsxrb Elements::Workbook
      book_back = Xlsxrb::Adapters::Creek.from_xlsxrb(xlsxrb_wb)
      assert_instance_of Xlsxrb::Adapters::Creek::Book, book_back
      rows_back = book_back.sheets.first.rows.to_a

      assert_equal rows_orig, rows_back

      book.close
      book_back.close
    end
  end

  def test_to_xlsxrb_with_images_and_date1904
    path = fixture("sample-with-images.xlsx")
    book = Xlsxrb::Adapters::Creek::Book.new(path)
    sheet = book.sheets.first.with_images
    ws = sheet.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Worksheet, ws
    assert_operator ws.images.size, :>, 0
    img = ws.images.first
    assert_instance_of Xlsxrb::Elements::Image, img
    assert_equal "A2", img.cell_ref
    book.close
  end
end
