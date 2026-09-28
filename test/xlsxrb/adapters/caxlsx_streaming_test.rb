# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCaxlsxStreamingTest < Test::Unit::TestCase
  def test_streaming_package_open_with_block
    with_tempfile do |path|
      Xlsxrb::Adapters::Caxlsx::StreamingPackage.open(path) do |pkg|
        wb = pkg.workbook
        header_style = wb.styles.add_style(b: true, sz: 12)
        row_style = wb.styles.add_style(num_fmt: 1)

        wb.add_worksheet(name: "StreamingSheet") do |ws|
          ws.column_widths(15, 25, 20)
          ws.add_row(%w[ID Name Score], style: header_style, height: 25)
          ws.add_row([1, "Alice", 95.5], style: [row_style, nil, nil])
          ws.add_row([2, "Bob", 88.0], style: [row_style, nil, nil])
          ws.merge_cells "A4:C4"
          ws.auto_filter = "A1:C3"
          ws.add_conditional_formatting("C2:C3", { type: :cellIs, operator: :greaterThan, formula: "90" })
          ws.add_data_validation("B2:B3", { type: :whole, operator: :between, formula1: "1", formula2: "100" })
        end
      end

      assert File.exist?(path)
      assert File.size(path).positive?

      # Verify with Xlsxrb reader
      read_wb = Xlsxrb.read(path).load
      sheet = read_wb.sheet("StreamingSheet")
      assert_not_nil sheet
      assert_equal 3, sheet.rows.size
      assert_equal "ID", sheet["A1"].value
      assert_equal "Name", sheet["B1"].value
      assert_equal "Score", sheet["C1"].value
      assert_equal 1, sheet["A2"].value
      assert_equal "Alice", sheet["B2"].value
      assert_equal 95.5, sheet["C2"].value
      assert_equal 2, sheet["A3"].value
      assert_equal "Bob", sheet["B3"].value
      assert_equal 88.0, sheet["C3"].value

      assert_equal 1, sheet.conditional_formatting.size
      assert_equal "C2:C3", sheet.conditional_formatting.first[:sqref]
      assert_equal 1, sheet.data_validations.size
      assert_equal "B2:B3", sheet.data_validations.first[:sqref]
    end
  end

  def test_streaming_package_unspecified_target_and_to_stream
    pkg = Xlsxrb::Adapters::Caxlsx::StreamingPackage.new
    wb = pkg.workbook
    ws = wb.add_worksheet(name: "StreamTest")
    ws.add_row(%w[Product Price Quantity])
    ws.add_row(["Widget", 19.99, 100])
    ws.add_row(["Gadget", 29.99, 50], offset: 1)

    stream = pkg.to_stream
    assert_instance_of StringIO, stream
    assert stream.size.positive?

    with_tempfile do |path|
      assert pkg.serialize(path)
      read_wb = Xlsxrb.read(path).load
      sheet = read_wb.sheet(0)
      assert_equal "Widget", sheet["A2"].value
      assert_equal 19.99, sheet["B2"].value
      assert_equal 100, sheet["C2"].value
      # Offset: empty first cell
      assert_nil sheet["A3"]
      assert_equal "Gadget", sheet["B3"].value
    end
  end

  def test_streaming_package_multiple_sheets
    with_tempfile do |path|
      Xlsxrb::Adapters::Caxlsx::StreamingPackage.open(path) do |pkg|
        wb = pkg.workbook

        wb.add_worksheet(name: "Sheet1") do |s1|
          s1.add_row(["Sheet 1 Row 1"])
        end

        wb.add_worksheet(name: "Sheet2") do |s2|
          s2.add_row(["Sheet 2 Row 1"])
          s2.add_row(["Sheet 2 Row 2"])
        end
      end

      read_wb = Xlsxrb.read(path).load
      assert_equal 2, read_wb.sheet_names.size
      assert_equal "Sheet1", read_wb.sheet(0).name
      assert_equal "Sheet2", read_wb.sheet(1).name
      assert_equal "Sheet 1 Row 1", read_wb.sheet(0)["A1"].value
      assert_equal "Sheet 2 Row 2", read_wb.sheet(1)["A2"].value
    end
  end

  def test_streaming_package_date1904
    with_tempfile do |path|
      Xlsxrb::Adapters::Caxlsx::StreamingPackage.open(path) do |pkg|
        pkg.workbook.date1904 = true
        assert pkg.workbook.date1904

        pkg.workbook.add_worksheet(name: "Calendar") do |ws|
          ws.add_row(["Today", Date.new(2026, 9, 27)])
        end
      end

      read_wb = Xlsxrb.read(path)
      assert read_wb.date1904?
    end
  end

  def test_streaming_package_alias_namespace
    assert_equal Xlsxrb::Adapters::Caxlsx::StreamingPackage, Xlsxrb::Adapters::Axlsx::StreamingPackage
  end
end
