# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersRubyXLCoreTest < Test::Unit::TestCase
  def test_reference_ind2ref_and_ref2ind
    ref_cls = Xlsxrb::Adapters::RubyXL::Reference

    assert_equal "A1", ref_cls.ind2ref(0, 0)
    assert_equal "B2", ref_cls.ind2ref(1, 1)
    assert_equal "Z26", ref_cls.ind2ref(25, 25)
    assert_equal "AA27", ref_cls.ind2ref(26, 26)
    assert_equal "XFD1048576", ref_cls.ind2ref(1_048_575, 16_383)

    # Absolute references
    assert_equal "$A$1", ref_cls.ind2ref(0, 0, true, true)
    assert_equal "A$1", ref_cls.ind2ref(0, 0, true, false)
    assert_equal "$A1", ref_cls.ind2ref(0, 0, false, true)

    # ref2ind
    assert_equal [0, 0, false, false], ref_cls.ref2ind("A1")
    assert_equal [1, 1, false, false], ref_cls.ref2ind("B2")
    assert_equal [0, 0, true, true], ref_cls.ref2ind("$A$1")
    assert_equal [0, 0, true, false], ref_cls.ref2ind("A$1")
    assert_equal [0, 0, false, true], ref_cls.ref2ind("$A1")
    assert_equal [-1, -1], ref_cls.ref2ind("invalid")
  end

  def test_reference_object_initialization_and_properties
    ref_cls = Xlsxrb::Adapters::RubyXL::Reference

    r1 = ref_cls.new(0, 0)
    assert r1.single_cell?
    assert r1.valid?
    assert_equal 0, r1.first_row
    assert_equal 0, r1.last_row
    assert_equal 0, r1.first_col
    assert_equal 0, r1.last_col
    assert_equal "A1", r1.to_s

    r2 = ref_cls.new(0, 2, 0, 3)
    refute r2.single_cell?
    assert_equal "A1:D3", r2.to_s

    r3 = ref_cls.new("B2:E5")
    assert_equal 1, r3.first_row
    assert_equal 4, r3.last_row
    assert_equal 1, r3.first_col
    assert_equal 4, r3.last_col
    assert_equal "B2:E5", r3.to_s

    # With sheet name
    r_sheet = ref_cls.new("'Sheet 1'!A1:B2")
    assert_equal "Sheet 1", r_sheet.sheet_name
    assert_equal "'Sheet 1'!A1:B2", r_sheet.to_s

    # Cover
    assert r2.cover?(r1)
    refute r1.cover?(r2)
  end

  def test_sqref
    sqref_cls = Xlsxrb::Adapters::RubyXL::Sqref
    sq = sqref_cls.new("A1:B2 C3:D4")
    assert_equal 2, sq.size
    assert_equal "A1:B2 C3:D4", sq.to_s
  end

  def test_workbook_core_properties_metadata
    now = Time.at(Time.now.to_i)
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new(
      creator: "Test Author",
      modifier: "Test Modifier",
      created_at: now,
      modified_at: now + 3600
    )
    assert_equal "Test Author", wb.creator
    assert_equal "Test Modifier", wb.modifier
    assert_equal now, wb.created_at
    assert_equal now + 3600, wb.modified_at

    wb.title = "Document Title"
    assert_equal "Document Title", wb.title

    wb.company = "Ruby Enterprise"
    assert_equal "Ruby Enterprise", wb.company

    # Date 1904
    refute wb.date1904
    wb.date1904 = true
    assert wb.date1904
  end

  def test_workbook_defined_names
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    dn = wb.define_new_name("TaxRate", "Sheet1!$A$1")
    assert_equal "TaxRate", dn.name
    assert_equal "Sheet1!$A$1", dn.reference
    assert_equal dn, wb.get_defined_name("TaxRate")
    assert_nil wb.get_defined_name("NonExistent")
  end

  def test_workbook_sheet_name_validation_and_trim
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new

    # Invalid characters
    %w[[ ] * ? / \\].each do |char|
      assert_raise(RuntimeError) { wb.add_worksheet("Bad#{char}Sheet") }
    end

    # Forbidden reserved name
    assert_raise(RuntimeError) { wb.add_worksheet("History") }
    assert_raise(RuntimeError) { wb.add_worksheet("history") }

    # Trimming over 31 chars on write
    long_name = "This is a very long sheet name that exceeds 31 chars"
    ws = wb.add_worksheet(long_name)
    assert_equal long_name, ws.sheet_name

    with_tempfile do |filepath|
      wb.write(filepath)
      parsed_wb = Xlsxrb::Adapters::RubyXL::Parser.parse(filepath)
      assert_equal long_name[0..30], parsed_wb.worksheets[1].sheet_name
    end
  end

  def test_workbook_enumerable_each
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    wb.add_worksheet("Second")
    count = 0
    wb.each { count += 1 }
    assert_equal 2, count
    assert_equal %w[Sheet1 Second], wb.map(&:sheet_name)
  end

  def test_cell_string_length_limit
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]

    valid_str = "A" * 32_767
    invalid_str = "A" * 32_768

    ws.add_cell(0, 0, valid_str)
    assert_equal 32_767, ws[0][0].value.length

    assert_raise(ArgumentError) do
      ws.add_cell(0, 1, invalid_str)
    end

    assert_raise(ArgumentError) do
      ws[0][0].value = invalid_str
    end

    assert_raise(ArgumentError) do
      ws[0][0].change_contents(invalid_str)
    end
  end

  def test_cell_date_serial_conversion
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]

    date = Date.new(2026, 9, 26)
    cell = ws.add_cell(0, 0, date)
    serial = cell.raw_value
    assert_in_delta 46_291.0, serial, 1.0

    # Num to date
    recovered_date = wb.num_to_date(serial)
    assert_equal date.year, recovered_date.year
    assert_equal date.month, recovered_date.month
    assert_equal date.day, recovered_date.day
  end

  def test_cell_rich_text_and_datatypes
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    ws = wb[0]

    # Numeric
    c_num = ws.add_cell(0, 0, 123)
    assert_nil c_num.datatype

    # Raw String
    c_str = ws.add_cell(0, 1, "Hello")
    assert_equal Xlsxrb::Adapters::RubyXL::DataType::RAW_STRING, c_str.datatype

    # Rich Text
    rich_text = Xlsxrb::Adapters::RubyXL::RichText.new(
      t: Xlsxrb::Adapters::RubyXL::Text.new(value: "Styled Text")
    )
    c_rich = ws.add_cell(0, 2, rich_text)
    assert_equal Xlsxrb::Adapters::RubyXL::DataType::INLINE_STRING, c_rich.datatype
    assert_equal "Styled Text", c_rich.to_s

    # Shared String
    c_shared = ws.add_cell(0, 3)
    c_shared.add_shared_string("Shared Entry")
    assert_equal Xlsxrb::Adapters::RubyXL::DataType::SHARED_STRING, c_shared.datatype
    assert_equal "Shared Entry", wb.shared_strings[c_shared.raw_value]
  end

  def test_parser_with_buffer_and_properties
    wb = Xlsxrb::Adapters::RubyXL::Workbook.new
    wb.creator = "Parser Tester"
    ws = wb[0]
    ws.add_cell(0, 0, "& < > \"")

    stream = wb.stream
    parsed_wb = Xlsxrb::Adapters::RubyXL::Parser.parse_buffer(stream)
    assert_equal 1, parsed_wb.worksheets.size
    assert_equal "& < > \"", parsed_wb[0][0][0].value
    assert_equal "Parser Tester", parsed_wb.creator
  end
end
