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

      assert_nil Xlsxrb::Adapters::RubyXL::Parser.parse(path)["A"].sheet_views
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
end
