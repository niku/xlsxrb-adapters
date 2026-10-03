# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersRooTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def test_module_and_class_structure
    assert_equal Xlsxrb::Adapters::Roo::Excelx, Xlsxrb::Adapters::Roo.const_get(:Excelx)
    assert Xlsxrb::Adapters::Roo::Excelx < Xlsxrb::Adapters::Roo::Base
    assert defined?(Xlsxrb::Adapters::Roo::Spreadsheet)
    assert defined?(Xlsxrb::Adapters::Roo::Excelx::Coordinate)
    assert defined?(Xlsxrb::Adapters::Roo::Link)
    assert defined?(Xlsxrb::Adapters::Roo::Font)
  end

  def test_spreadsheet_open_and_block
    path = fixture("simple_spreadsheet.xlsx")
    doc = Xlsxrb::Adapters::Roo::Spreadsheet.open(path)
    assert_instance_of Xlsxrb::Adapters::Roo::Excelx, doc
    assert_equal %w[Sheet1 Sheet2 Sheet3], doc.sheets

    yielded = nil
    Xlsxrb::Adapters::Roo::Spreadsheet.open(path) do |s|
      yielded = s
      assert_equal 3, s.sheets.size
    end
    assert_not_nil yielded

    assert_raise(ArgumentError) do
      Xlsxrb::Adapters::Roo::Spreadsheet.open("unknown.invalid_ext")
    end
  end

  def test_sheets_and_default_sheet_switching
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("simple_spreadsheet.xlsx"))
    assert_equal "Sheet1", doc.default_sheet
    assert_equal %w[Sheet1 Sheet2 Sheet3], doc.sheets

    doc.default_sheet = "Sheet2"
    assert_equal "Sheet2", doc.default_sheet

    # 0-based integer sheet index
    doc.default_sheet = 2
    assert_equal "Sheet3", doc.default_sheet

    assert_raise(RangeError) do
      doc.default_sheet = "NonExistent"
    end

    assert_raise(RangeError) do
      doc.default_sheet = 99
    end

    sheet_names = []
    doc.each_with_pagename do |name, sheet|
      sheet_names << name
      assert_instance_of Xlsxrb::Adapters::Roo::Excelx, sheet
    end
    assert_equal %w[Sheet1 Sheet2 Sheet3], sheet_names
  end

  def test_cell_values_and_types
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))
    assert_equal 1, doc.cell(1, 1)
    assert_equal 1, doc.cell("A", 1)
    assert_equal 1, doc.cell(1, "A")
    assert_equal :float, doc.celltype(1, 1)

    assert_equal 2, doc.cell(1, 2)
    assert_equal "test", doc.cell(2, 6)
    assert_equal :string, doc.celltype(2, 6)

    assert_equal Date.new(1961, 11, 21), doc.cell(5, 1)
    assert_equal :date, doc.celltype(5, 1)
  end

  def test_formulas
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("formula.xlsx"))
    assert_equal 1, doc.cell(1, 1)
    assert_equal 2, doc.cell(2, 1)
    assert_equal 21, doc.cell(7, 1)
    assert_equal "SUM(A1:A6)", doc.formula(7, 1)
    assert doc.formula?(7, 1)
    assert !doc.formula?(1, 1)
    assert_equal :formula, doc.celltype(7, 1)
  end

  def test_fonts
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("style.xlsx"))
    f1 = doc.font(1, 1)
    assert_not_nil f1
    assert f1.bold?
    assert !f1.italic?
    assert !f1.underline?

    f2 = doc.font(2, 1)
    assert_not_nil f2
    assert !f2.bold?
    assert f2.italic?

    f4 = doc.font(4, 1)
    assert_not_nil f4
    assert f4.underline?

    f5 = doc.font(5, 1)
    assert_not_nil f5
    assert f5.bold?
    assert f5.italic?
  end

  def test_comments
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("comments.xlsx"))
    assert doc.comment?(4, 2)
    assert_equal "Kommentar fuer B4", doc.comment(4, 2)
    assert doc.comment?(5, 2)
    assert_equal "Kommentar fuer B5", doc.comment(5, 2)
    assert !doc.comment?(1, 1)
    assert_nil doc.comment(1, 1)

    all_comments = doc.comments
    assert_equal 2, all_comments.size
    assert_include all_comments, [4, 2, "Kommentar fuer B4"]
    assert_include all_comments, [5, 2, "Kommentar fuer B5"]
  end

  def test_hyperlinks
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("link.xlsx"))
    assert doc.hyperlink?(1, 1)
    assert_equal "http://www.google.com", doc.hyperlink(1, 1)
    c = doc.cell(1, 1)
    assert_instance_of Xlsxrb::Adapters::Roo::Link, c
    assert_equal "http://www.google.com", c.url
    assert_equal "http://www.google.com", c.href
    assert_equal "http://www.google.com", c.to_uri.to_s
    assert_equal "Google", c.to_s
  end

  def test_bounds_and_rows_columns
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))
    assert_equal 1, doc.first_row
    assert_equal 18, doc.last_row
    assert_equal 1, doc.first_column
    assert_equal 7, doc.last_column

    r1 = doc.row(1)
    assert_equal [1.0, 2.0, 3.0, 4.0, 10.0, nil, nil], r1

    c1 = doc.column(1)
    assert_equal 18, c1.size
    assert_equal 1.0, c1[0]

    assert !doc.empty?(1, 1)
    assert doc.empty?(3, 1)
  end

  def test_defined_names_and_labels
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("named_cells.xlsx"))
    lbls = doc.labels
    assert_equal 3, lbls.size
    assert_equal [5, 3, "Sheet1"], doc.label("anton")
    assert_equal "Anton", doc.cell(*doc.label("anton"))

    # Method missing dispatch
    assert_equal "Anton", doc.anton
    assert_equal "Bertha", doc.berta
    assert_equal "Cäsar", doc.caesar

    # Coordinate method missing
    assert_equal "Anton", doc.c5
    assert_equal "Bertha", doc.b4
  end

  def test_each_row_streaming
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))
    rows = []
    doc.each_row_streaming do |row|
      rows << row
    end
    assert_equal 18, rows.size
    assert_kind_of Xlsxrb::Adapters::Roo::Excelx::Cell::Base, rows[0][0]
    assert_equal 1.0, rows[0][0].value

    # Test pad_cells, offset, max_rows
    padded_rows = []
    doc.each_row_streaming(pad_cells: true, offset: 1, max_rows: 2) do |row|
      padded_rows << row
    end
    assert_equal 3, padded_rows.size
    assert_equal 5.0, padded_rows[0][0].value
  end

  def test_parse_and_row_with
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))
    parsed = doc.parse(headers: true)
    assert_equal 18, parsed.size
    assert_equal 1, parsed[0][1]
    assert_equal 5, parsed[1][1]

    row_found = doc.row_with(["test"])
    assert_equal 2, row_found

    parsed_search = doc.parse(header_search: ["test"])
    assert_equal 16, parsed_search.size
  end

  def test_exports
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))

    csv = doc.to_csv
    assert_include csv, "1,2,3,4,10"
    assert_include csv, "5,6,7,8,9,\"test\",11"

    mat = doc.to_matrix
    assert_equal 18, mat.row_count
    assert_equal 7, mat.column_count
    assert_equal 1.0, mat[0, 0]

    xml = doc.to_xml
    assert_include xml, "<cell"
    assert_include xml, "1"

    yaml = doc.to_yaml
    assert_include yaml, "test"
  end

  def test_set_cell_value
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))
    assert_equal 1.0, doc.cell(1, 1)
    doc.set(1, 1, 999.0)
    assert_equal 999.0, doc.cell(1, 1)
  end

  def test_options_only_visible_sheets
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("hidden_sheets.xlsx"), only_visible_sheets: true)
    assert_equal ["VisibleSheet1"], doc.sheets
  end

  def test_options_disable_html_wrapper
    doc_html = Xlsxrb::Adapters::Roo::Excelx.new(fixture("paragraph.xlsx"))
    doc_plain = Xlsxrb::Adapters::Roo::Excelx.new(fixture("paragraph.xlsx"), disable_html_wrapper: true)
    assert_equal doc_plain.cell(1, 1), doc_html.cell(1, 1)
  end

  def test_xlsxrb_roundtrip
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))
    xlsxrb_wb = doc.to_xlsxrb
    assert_instance_of Xlsxrb::Elements::Workbook, xlsxrb_wb
    assert_equal 5, xlsxrb_wb.sheets.size
    assert_equal 12, xlsxrb_wb.sheets[0].rows.size

    restored = Xlsxrb::Adapters::Roo::Excelx.from_xlsxrb(xlsxrb_wb)
    assert_equal doc.sheets, restored.sheets
    assert_equal doc.cell(1, 1), restored.cell(1, 1)
    assert_equal doc.cell(1, 2), restored.cell(1, 2)
  end

  def test_sheet_visibility_methods
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("hidden_sheets.xlsx"))
    assert doc.sheet_hidden?("HiddenSheet1")
    assert !doc.sheet_visible?("HiddenSheet1")
    assert doc.sheet_visible?("VisibleSheet1")
    assert !doc.sheet_hidden?("VisibleSheet1")

    hidden_sheet = doc.sheet_for("HiddenSheet1")
    assert_not_nil hidden_sheet
    assert hidden_sheet.hidden?
    assert !hidden_sheet.visible?

    visible_sheet = doc.sheet_for("VisibleSheet1")
    assert_not_nil visible_sheet
    assert visible_sheet.visible?
    assert !visible_sheet.hidden?
  end

  def test_date_formatted_value_with_number_formatter
    doc = Xlsxrb::Adapters::Roo::Excelx.new(fixture("numbers1.xlsx"))
    assert_equal "21/11/61", doc.formatted_value(5, 1)
  end

  def test_xlsxrb_roundtrip_with_extended_metadata
    doc_link = Xlsxrb::Adapters::Roo::Excelx.new(fixture("link.xlsx"))
    wb_link = doc_link.to_xlsxrb
    assert_equal 1, wb_link.sheets[0].hyperlinks.size
    cell_a1 = wb_link.sheets[0]["A1"]
    assert cell_a1.link?
    assert_equal "http://www.google.com", cell_a1.url

    doc_comm = Xlsxrb::Adapters::Roo::Excelx.new(fixture("comments.xlsx"))
    wb_comm = doc_comm.to_xlsxrb
    assert_equal 2, wb_comm.sheets[0].comments.size
    cell_b4 = wb_comm.sheets[0]["B4"]
    assert cell_b4.comment?
    assert_equal "Kommentar fuer B4", cell_b4.comment_text
    assert_equal 1, wb_comm.defined_names.size
    assert_equal "befuenf", wb_comm.defined_names[0][:name]

    doc_form = Xlsxrb::Adapters::Roo::Excelx.new(fixture("formula.xlsx"))
    wb_form = doc_form.to_xlsxrb
    cell_a7 = wb_form.sheets[0]["A7"]
    assert cell_a7.formula?
    assert_equal "SUM(A1:A6)", cell_a7.formula_expression
  end
end
