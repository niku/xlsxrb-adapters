# frozen_string_literal: true

require "test_helper"
require "roo"
require "stringio"

class XlsxrbAdaptersXlsxtreamDropinTest < Test::Unit::TestCase
  def test_dropin_classes_and_constants
    target = Xlsxrb::Adapters::Xlsxtream

    assert defined?(target::Workbook)
    assert defined?(target::Worksheet)
    assert defined?(target::Row)
    assert defined?(target::Columns)
    assert defined?(target::SharedStringTable)
    assert defined?(target::ZipKitWriter)
    assert defined?(target::XML)
    assert defined?(target::Error)
    assert defined?(target::Deprecation)
    assert defined?(target::VERSION)
  end

  def test_dropin_streaming_file_export
    target = Xlsxrb::Adapters::Xlsxtream

    with_tempfile do |filepath|
      target::Workbook.open(filepath) do |xlsx|
        xlsx.write_worksheet("Orders") do |sheet|
          sheet << %w[ID Customer Total Shipped Date]
          sheet << [1001, "Acme Corp", 1234.56, true, Date.new(2026, 3, 1)]
          sheet << [1002, "Globex Inc", 789.10, false, Date.new(2026, 3, 2)]
        end

        xlsx.write_worksheet("Summary") do |sheet|
          sheet << %w[Metric Value]
          sheet << ["Total Orders", 2]
        end
      end

      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      roo = Roo::Excelx.new(filepath)
      assert_equal %w[Orders Summary], roo.sheets
      assert_equal "Acme Corp", roo.cell(2, 2, "Orders")
      assert_equal 1234.56, roo.cell(2, 3, "Orders")
      assert_equal "Total Orders", roo.cell(2, 1, "Summary")
    end
  end

  def test_dropin_in_memory_stringio_export
    target = Xlsxrb::Adapters::Xlsxtream
    sio = StringIO.new

    xlsx = target::Workbook.new(sio, use_shared_strings: true)
    xlsx.write_worksheet("Repeated") do |sheet|
      sheet << %w[repeat repeat repeat]
      sheet << %w[repeat again repeat]
    end
    xlsx.close

    assert sio.string.bytesize.positive?
    assert sio.string.start_with?("PK")

    with_tempfile do |path|
      File.binwrite(path, sio.string)
      roo = Roo::Excelx.new(path)
      assert_equal ["Repeated"], roo.sheets
      assert_equal "repeat", roo.cell(1, 1)
      assert_equal "again", roo.cell(2, 2)
    end
  end

  def test_dropin_auto_format_feature
    target = Xlsxrb::Adapters::Xlsxtream

    with_tempfile do |filepath|
      target::Workbook.open(filepath) do |xlsx|
        xlsx.write_worksheet(name: "Data", auto_format: true) do |sheet|
          sheet << ["true", "42", "19.99", "2026-06-15"]
        end
      end

      roo = Roo::Excelx.new(filepath)
      assert_equal true, roo.cell(1, 1)
      assert_equal 42, roo.cell(1, 2)
      assert_in_delta 19.99, roo.cell(1, 3), 0.001
      assert_equal Date.new(2026, 6, 15), roo.cell(1, 4)
    end
  end

  def test_dropin_column_widths_configuration
    target = Xlsxrb::Adapters::Xlsxtream

    with_tempfile do |filepath|
      target::Workbook.open(filepath, columns: [
                              { width_pixels: 50 },
                              { width_chars: 15 }
                            ]) do |xlsx|
        xlsx.write_worksheet("Cols") do |sheet|
          sheet << ["Col A", "Col B"]
        end
      end

      roo = Roo::Excelx.new(filepath)
      assert_equal "Col A", roo.cell(1, 1)
      assert_equal "Col B", roo.cell(1, 2)
    end
  end
end
