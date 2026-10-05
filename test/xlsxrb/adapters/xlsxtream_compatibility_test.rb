# frozen_string_literal: true

require "test_helper"
require "xlsxtream"
require "roo"

class XlsxrbAdaptersXlsxtreamCompatibilityTest < Test::Unit::TestCase
  def test_side_by_side_primitive_values
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        # 1. Generate with official Xlsxtream
        ::Xlsxtream::Workbook.open(orig_path) do |wb|
          wb.write_worksheet("Primitives") do |ws|
            ws << %w[Name Age Balance Active Registered LoggedAt]
            ws << ["Alice", 30, 1500.50, true, Date.new(2026, 1, 15), Time.utc(2026, 1, 15, 10, 30, 0)]
            ws << ["Bob", 25, 0.0, false, Date.new(2026, 5, 20), Time.utc(2026, 5, 20, 18, 45, 0)]
          end
        end

        # 2. Generate with Xlsxrb Xlsxtream adapter
        Xlsxrb::Adapters::Xlsxtream::Workbook.open(adapt_path) do |wb|
          wb.write_worksheet("Primitives") do |ws|
            ws << %w[Name Age Balance Active Registered LoggedAt]
            ws << ["Alice", 30, 1500.50, true, Date.new(2026, 1, 15), Time.utc(2026, 1, 15, 10, 30, 0)]
            ws << ["Bob", 25, 0.0, false, Date.new(2026, 5, 20), Time.utc(2026, 5, 20, 18, 45, 0)]
          end
        end

        # 3. Compare with Roo
        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        assert_equal orig_roo.sheets, adapt_roo.sheets

        (1..3).each do |row|
          (1..6).each do |col|
            orig_val = orig_roo.cell(row, col)
            adapt_val = adapt_roo.cell(row, col)

            if orig_val.is_a?(Float) && adapt_val.is_a?(Float)
              assert_in_delta orig_val, adapt_val, 0.0001, "Mismatch at row #{row}, col #{col}"
            else
              assert_equal orig_val, adapt_val, "Mismatch at row #{row}, col #{col}"
            end
          end
        end
      end
    end
  end

  def test_side_by_side_shared_strings_table
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        # 1. Upstream with SST
        ::Xlsxtream::Workbook.open(orig_path, use_shared_strings: true) do |wb|
          wb.write_worksheet("Shared") do |ws|
            ws << %w[alpha beta gamma alpha]
            ws << %w[beta alpha delta beta]
          end
        end

        # 2. Adapter with SST
        Xlsxrb::Adapters::Xlsxtream::Workbook.open(adapt_path, use_shared_strings: true) do |wb|
          wb.write_worksheet("Shared") do |ws|
            ws << %w[alpha beta gamma alpha]
            ws << %w[beta alpha delta beta]
          end
        end

        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        (1..2).each do |row|
          (1..4).each do |col|
            assert_equal orig_roo.cell(row, col), adapt_roo.cell(row, col)
          end
        end
      end
    end
  end

  def test_side_by_side_auto_format
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        input_row = ["true", "false", "100", "25.75", "2026-10-05"]

        ::Xlsxtream::Workbook.open(orig_path) do |wb|
          wb.write_worksheet("Auto", auto_format: true) do |ws|
            ws << input_row
          end
        end

        Xlsxrb::Adapters::Xlsxtream::Workbook.open(adapt_path) do |wb|
          wb.write_worksheet("Auto", auto_format: true) do |ws|
            ws << input_row
          end
        end

        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        (1..5).each do |col|
          orig_val = orig_roo.cell(1, col)
          adapt_val = adapt_roo.cell(1, col)
          if orig_val.is_a?(Float) && adapt_val.is_a?(Float)
            assert_in_delta orig_val, adapt_val, 0.0001
          else
            assert_equal orig_val, adapt_val
          end
        end
      end
    end
  end

  def test_side_by_side_multiple_sheets_and_columns
    with_tempfile do |orig_path|
      with_tempfile do |adapt_path|
        col_settings = [{ width_pixels: 40 }, { width_chars: 12 }]

        ::Xlsxtream::Workbook.open(orig_path, columns: col_settings) do |wb|
          wb.write_worksheet("Sheet1") { |ws| ws << %w[First Sheet] }
          wb.write_worksheet("Sheet2") { |ws| ws << %w[Second Sheet] }
        end

        Xlsxrb::Adapters::Xlsxtream::Workbook.open(adapt_path, columns: col_settings) do |wb|
          wb.write_worksheet("Sheet1") { |ws| ws << %w[First Sheet] }
          wb.write_worksheet("Sheet2") { |ws| ws << %w[Second Sheet] }
        end

        orig_roo = Roo::Excelx.new(orig_path)
        adapt_roo = Roo::Excelx.new(adapt_path)

        assert_equal orig_roo.sheets, adapt_roo.sheets
        assert_equal orig_roo.cell(1, 1, "Sheet1"), adapt_roo.cell(1, 1, "Sheet1")
        assert_equal orig_roo.cell(1, 1, "Sheet2"), adapt_roo.cell(1, 1, "Sheet2")
      end
    end
  end
end
