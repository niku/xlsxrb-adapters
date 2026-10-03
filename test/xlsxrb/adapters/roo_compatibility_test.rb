# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersRooCompatibilityTest < Test::Unit::TestCase
  def fixture(name)
    File.expand_path("../../fixtures/roo/#{name}", __dir__)
  end

  def assert_spreadsheets_match(filename, options = {})
    path = fixture(filename)
    roo_doc = ::Roo::Excelx.new(path, options)
    our_doc = ::Xlsxrb::Adapters::Roo::Excelx.new(path, options)

    assert_equal roo_doc.sheets, our_doc.sheets, "Sheets mismatch in #{filename}"

    roo_doc.sheets.each do |s|
      roo_doc.default_sheet = s
      our_doc.default_sheet = s

      next if roo_doc.first_row.nil? && our_doc.first_row.nil?

      assert_equal roo_doc.first_row, our_doc.first_row, "first_row mismatch in #{filename} sheet #{s}"
      assert_equal roo_doc.last_row, our_doc.last_row, "last_row mismatch in #{filename} sheet #{s}"
      assert_equal roo_doc.first_column, our_doc.first_column, "first_column mismatch in #{filename} sheet #{s}"
      assert_equal roo_doc.last_column, our_doc.last_column, "last_column mismatch in #{filename} sheet #{s}"

      (roo_doc.first_row..roo_doc.last_row).each do |r|
        (roo_doc.first_column..roo_doc.last_column).each do |c|
          rc = roo_doc.cell(r, c)
          oc = our_doc.cell(r, c)
          rt = roo_doc.celltype(r, c)
          ot = our_doc.celltype(r, c)
          rf = roo_doc.formula(r, c)
          of = our_doc.formula(r, c)

          assert_equal rc, oc, "Cell mismatch in #{filename} [#{s}](#{r},#{c})"
          assert_equal rt, ot, "Cell type mismatch in #{filename} [#{s}](#{r},#{c})"
          assert_equal rf, of, "Formula mismatch in #{filename} [#{s}](#{r},#{c})"
        end
      end
    end
  end

  def test_numbers1_compatibility
    assert_spreadsheets_match("numbers1.xlsx")
  end

  def test_formula_compatibility
    assert_spreadsheets_match("formula.xlsx")
  end

  def test_style_compatibility
    assert_spreadsheets_match("style.xlsx")

    roo_doc = ::Roo::Excelx.new(fixture("style.xlsx"))
    our_doc = ::Xlsxrb::Adapters::Roo::Excelx.new(fixture("style.xlsx"))

    (1..5).each do |r|
      (1..5).each do |c|
        rf = roo_doc.font(r, c)
        of = our_doc.font(r, c)
        if rf
          assert_not_nil of, "Font should exist at (#{r},#{c})"
          assert_equal rf.bold?, of.bold?, "Font bold mismatch at (#{r},#{c})"
          assert_equal rf.italic?, of.italic?, "Font italic mismatch at (#{r},#{c})"
          assert_equal rf.underline?, of.underline?, "Font underline mismatch at (#{r},#{c})"
        else
          assert_nil of, "Font should be nil at (#{r},#{c})"
        end
      end
    end
  end

  def test_comments_compatibility
    assert_spreadsheets_match("comments.xlsx")

    roo_doc = ::Roo::Excelx.new(fixture("comments.xlsx"))
    our_doc = ::Xlsxrb::Adapters::Roo::Excelx.new(fixture("comments.xlsx"))

    assert_equal roo_doc.comments, our_doc.comments
    (1..5).each do |r|
      (1..5).each do |c|
        assert_equal roo_doc.comment(r, c), our_doc.comment(r, c)
        assert_equal roo_doc.comment?(r, c), our_doc.comment?(r, c)
      end
    end
  end

  def test_link_compatibility
    assert_spreadsheets_match("link.xlsx")

    roo_doc = ::Roo::Excelx.new(fixture("link.xlsx"))
    our_doc = ::Xlsxrb::Adapters::Roo::Excelx.new(fixture("link.xlsx"))

    assert_equal roo_doc.hyperlink?(1, 1), our_doc.hyperlink?(1, 1)
    assert_equal roo_doc.hyperlink(1, 1), our_doc.hyperlink(1, 1)
  end

  def test_named_cells_compatibility
    assert_spreadsheets_match("named_cells.xlsx")

    roo_doc = ::Roo::Excelx.new(fixture("named_cells.xlsx"))
    our_doc = ::Xlsxrb::Adapters::Roo::Excelx.new(fixture("named_cells.xlsx"))

    assert_equal roo_doc.labels, our_doc.labels
    roo_doc.labels.map(&:first).each do |name|
      assert_equal roo_doc.label(name), our_doc.label(name)
      assert_equal roo_doc.cell(*roo_doc.label(name)), our_doc.cell(*our_doc.label(name))
    end
  end

  def test_simple_spreadsheet_compatibility
    assert_spreadsheets_match("simple_spreadsheet.xlsx")
  end

  def test_hidden_sheets_compatibility
    assert_spreadsheets_match("hidden_sheets.xlsx")
    assert_spreadsheets_match("hidden_sheets.xlsx", only_visible_sheets: true)
  end

  def test_paragraph_compatibility
    assert_spreadsheets_match("paragraph.xlsx")
  end

  def test_pfand_utf16_compatibility
    assert_spreadsheets_match("Pfand_from_windows_phone.xlsx")
  end

  def test_whitespace_compatibility
    assert_spreadsheets_match("whitespace.xlsx")
  end

  def test_richtext_compatibility
    assert_spreadsheets_match("richtext_example.xlsx")
  end
end
