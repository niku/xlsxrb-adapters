# frozen_string_literal: true

require "test_helper"

# Parsing a file, editing it and writing it back must keep what was not edited,
# as a template-based export with rubyXL does.
class XlsxrbAdaptersRubyXLRoundTripTest < Test::Unit::TestCase
  def write_template(path)
    Xlsxrb.write(path) do |w|
      w.sheet("Template") do |s|
        s.row(%w[202502 202503])
        s.conditional_format("A1:B1", type: :cell_is, operator: :greaterThan, formula: "15", priority: 1)
        s.sheet_format(default_row_height: 21.0)
      end
    end
  end

  def test_text_that_looks_like_a_number_stays_text
    with_tempfile do |path|
      write_template(path)

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal %w[202502 202503], wb["Template"][0].cells.map(&:value)
      assert_equal %w[202502 202503], ::RubyXL::Parser.parse(path)["Template"][0].cells.map(&:value)

      wb.write(path)
      assert_equal %w[202502 202503], ::RubyXL::Parser.parse(path)["Template"][0].cells.map(&:value)
    end
  end

  def test_sheet_state_is_read_and_written_in_rubyxl_form
    with_tempfile do |path|
      source = ::RubyXL::Workbook.new
      source[0].sheet_name = "Visible"
      source.add_worksheet("Hidden").state = "hidden"
      source.write(path)

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_nil wb["Visible"].state
      assert_equal "hidden", wb["Hidden"].state

      wb["Visible"].state = "hidden"
      wb["Hidden"].state = nil
      wb.write(path)
      official = ::RubyXL::Parser.parse(path)
      assert_equal "hidden", official["Visible"].state
      assert_nil official["Hidden"].state
    end
  end

  def test_conditional_formatting_and_sheet_format_can_be_copied_to_another_sheet
    with_tempfile do |path|
      write_template(path)

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      template = wb["Template"]
      copy = wb.add_worksheet("Copy")
      copy.conditional_formatting = template.conditional_formatting.dup
      copy.sheet_format_pr = template.sheet_format_pr
      wb.write(path)

      official = ::RubyXL::Parser.parse(path)["Copy"]
      assert_equal(["A1:B1"], official.conditional_formatting.map { |cf| cf.sqref.to_s })
      assert_in_delta 21.0, official.sheet_format_pr.default_row_height, 1e-6
    end
  end

  def test_active_tab_and_full_calc_on_load_are_written
    with_tempfile do |path|
      write_template(path)

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      wb.add_worksheet("Second")
      wb.book_views.first.active_tab = 1
      wb.calc_pr.full_calc_on_load = true
      wb.company = "Example"
      wb.write(path)

      official = ::RubyXL::Parser.parse(path)
      assert_equal 1, official.book_views.first.active_tab
      assert_equal true, official.calc_pr.full_calc_on_load
      assert_equal "Example", official.company

      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 1, reparsed.book_views.first.active_tab
      assert_equal true, reparsed.calc_pr.full_calc_on_load
    end
  end

  def test_sheet_views_can_be_dropped
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("A") do |s|
          s.row(["a"])
          s.sheet_view(:zoom_scale, 150)
        end
      end

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 150, wb["A"].sheet_views[:zoom_scale].to_i
      wb["A"].sheet_views = nil
      wb.write(path)

      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_nil reparsed["A"].sheet_views[:zoom_scale]
    end
  end

  def test_data_validation_object_and_hash_access
    ws = Xlsxrb::Adapters::RubyXL::Worksheet.new
    ws.data_validations = [{ type: "list", sqref: "A1:A10", formula1: '"Option1,Option2"' }]
    val = ws.data_validations.first

    # Object-style access
    assert_equal "list", val.type
    assert_equal "A1:A10", val.sqref
    assert_equal '"Option1,Option2"', val.formula1

    # Hash-style access
    assert_equal "list", val[:type]
    assert_equal "A1:A10", val[:sqref]

    # Modification via setter
    val.type = "whole"
    assert_equal "whole", val.type
    assert_equal "whole", val[:type]
  end

  def test_round_trip_preserves_merged_cells_and_defined_names
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("Sheet1") do |s|
          s.row(%w[A1 B1])
          s.row(%w[A2 B2])
          s.merge("A1:B1")
        end
        w.defined_name("MyRange", "Sheet1!$A$1:$B$1", sheet: "Sheet1", hidden: false)
      end

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 1, wb[0].merged_cells.size
      assert_equal "A1:B1", wb[0].merged_cells.first.ref.to_s
      assert_equal 1, wb.defined_names.size
      assert_equal "MyRange", wb.defined_names.first.name
      assert_equal 0, wb.defined_names.first.local_sheet_id
      assert_equal false, wb.defined_names.first.hidden

      # Write and verify reparsed
      wb.write(path)
      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 1, reparsed[0].merged_cells.size
      assert_equal "A1:B1", reparsed[0].merged_cells.first.ref.to_s
      assert_equal 1, reparsed.defined_names.size
      assert_equal "MyRange", reparsed.defined_names.first.name
    end
  end

  def test_column_ranges_min_max_preservation_roundtrip
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("Cols") do |s|
          s.row([10])
          s.column(0, width: 18.5)
        end
      end

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      ws = wb[0]
      cr = ws.cols.first
      cr.min = 1
      cr.max = 5
      cr.width = 18.5
      wb.write(path)

      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      re_cr = reparsed[0].cols.first
      assert_equal 1, re_cr.min
      assert_equal 5, re_cr.max
      assert_in_delta 18.5, re_cr.width, 0.1
    end
  end

  def test_data_validation_show_drop_down_and_ime_mode_roundtrip
    with_tempfile do |path|
      wb = Xlsxrb::Adapters::RubyXL::Workbook.new
      ws = wb[0]
      ws.sheet_name = "DVTest"
      ws.add_cell(0, 0, "Test")
      ws.data_validations = [
        {
          type: "list",
          sqref: "A1:A10",
          formula1: '"Yes,No"',
          show_drop_down: true,
          ime_mode: "hiragana"
        }
      ]
      wb.write(path)

      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      dv = reparsed[0].data_validations.first
      assert_equal true, dv.show_drop_down
      assert_equal "hiragana", dv.ime_mode
    end
  end

  def test_dxf_preservation_roundtrip
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("CF") do |s|
          s.row([10, 20])
          s.conditional_format("A1:B1", type: :cell_is, operator: :greaterThan, formula: "15", priority: 1, fill_color: "FFBFBFBF")
        end
      end

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 1, wb.stylesheet.dxfs.size
      wb.write(path)

      re_styles = Xlsxrb::Ooxml::StylesParser.parse(Xlsxrb::Ooxml::ZipReader.open(path) { |z| z.read_entry("xl/styles.xml") })
      assert_equal 1, re_styles.dxfs.size
    end
  end

  def test_code_name_and_sheet_pr_roundtrip
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.code_name("WorkbookVBA")
        w.sheet("Sheet1", code_name: "WorksheetVBA") do |s|
          s.row(["Hello"])
          s.tab_color("FF00AA")
          s.fit_to_page(true)
        end
      end

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal "WorkbookVBA", wb.workbook_properties.code_name
      ws = wb["Sheet1"]
      assert_equal "WorksheetVBA", ws.sheet_pr.code_name
      assert_equal "FF00AA", ws.sheet_pr.tab_color
      assert_equal true, ws.sheet_pr.page_set_up_pr.fit_to_page

      # Modify and write back
      ws.sheet_pr.code_name = "ModifiedSheetVBA"
      wb.workbook_properties.code_name = "ModifiedWbVBA"
      wb.write(path)

      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal "ModifiedWbVBA", reparsed.workbook_properties.code_name
      assert_equal "ModifiedSheetVBA", reparsed["Sheet1"].sheet_pr.code_name
      assert_equal "FF00AA", reparsed["Sheet1"].sheet_pr.tab_color
      assert_equal true, reparsed["Sheet1"].sheet_pr.page_set_up_pr.fit_to_page
    end
  end

  def test_ignored_errors_roundtrip
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.sheet("Sheet1") do |s|
          s.row(["0123"])
          s.ignore_error("A1:C10", number_stored_as_text: true)
        end
      end

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 1, wb["Sheet1"].ignored_errors.size
      assert_equal "A1:C10", wb["Sheet1"].ignored_errors.first[:sqref]
      assert_equal true, wb["Sheet1"].ignored_errors.first[:number_stored_as_text]

      wb["Sheet1"].ignored_errors << { sqref: "D1:D5", eval_error: true }
      wb.write(path)

      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 2, reparsed["Sheet1"].ignored_errors.size
    end
  end

  def test_workbook_views_roundtrip
    with_tempfile do |path|
      Xlsxrb.write(path) do |w|
        w.tab_ratio(600)
        w.show_sheet_tabs(true)
        w.sheet("First") { |s| s.row([1]) }
        w.sheet("Second", active: true) { |s| s.row([2]) }
        w.first_sheet(0)
      end

      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      view = wb.book_views.first
      assert_equal 1, view.active_tab
      assert_equal 600, view.tab_ratio
      assert_equal 0, view.first_sheet
      assert_equal true, view.show_sheet_tabs

      view.tab_ratio = 700
      wb.write(path)

      reparsed = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      assert_equal 1, reparsed.book_views.first.active_tab
      assert_equal 700, reparsed.book_views.first.tab_ratio
    end
  end

  def test_non_finite_float_raises_argument_error
    with_tempfile do |path|
      wb = Xlsxrb::Adapters::RubyXL::Workbook.new
      ws = wb[0]
      ws.add_cell(0, 0, Float::INFINITY)
      assert_raises(ArgumentError) do
        wb.write(path)
      end
    end

    with_tempfile do |path|
      wb = Xlsxrb::Adapters::RubyXL::Workbook.new
      ws = wb[0]
      ws.add_cell(0, 0, Float::NAN)
      assert_raises(ArgumentError) do
        wb.write(path)
      end
    end
  end
end
