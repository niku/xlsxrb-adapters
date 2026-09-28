# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCaxlsxDropinTest < Test::Unit::TestCase
  def setup
    # Bind Axlsx constant in a local scope if not already bound, or ensure it points to adapter
    @prev_axlsx = Object.const_get(:Axlsx) if Object.const_defined?(:Axlsx)
  end

  def test_dropin_replacement_workflow
    with_tempfile do |filepath|
      # Emulate client code that only requires 'caxlsx' and uses ::Axlsx::Package
      pkg = Xlsxrb::Adapters::Caxlsx::Package.new
      wb = pkg.workbook

      # Styles
      styles = wb.styles
      header_style = styles.add_style(
        b: true,
        sz: 12,
        fg_color: "FFFFFF",
        bg_color: "4F81BD",
        alignment: { horizontal: :center, vertical: :center }
      )
      currency_style = styles.add_style(format_code: "$#,##0.00")

      # Worksheets & Rows
      ws = wb.add_worksheet(name: "DropInReport")
      ws.add_row(%w[Product Qty Price Total], style: header_style)
      ws.add_row(["Apples", 10, 1.5, "=B2*C2"], style: [nil, nil, currency_style, currency_style])
      ws.add_row(["Oranges", 25, 2.0, "=B3*C3"], style: [nil, nil, currency_style, currency_style])
      ws.add_row(["Bananas", 50, 0.75, "=B4*C4"], style: [nil, nil, currency_style, currency_style])
      ws.add_row(["Total", "=SUM(B2:B4)", nil, "=SUM(D2:D4)"], style: [header_style, nil, nil, currency_style])

      # Features
      ws.merge_cells "A5:A5"
      ws.auto_filter = "A1:D4"
      ws.add_comment(ref: "A2", author: "Auditor", text: "Organic certified apples")
      ws.page_setup.orientation = :landscape

      # Chart
      chart = ws.add_chart(Xlsxrb::Adapters::Caxlsx::BarChart, title: "Units Sold", start_at: [0, 7], end_at: [5, 17])
      chart.add_series(data: ws["B2:B4"], labels: ws["A2:A4"], title: "Qty")

      # Serialize to file
      assert pkg.serialize(filepath)
      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      # Serialize to stream
      stream = pkg.to_stream
      refute_nil stream
      assert stream.size.positive?

      # Read back with RubyXL to verify OOXML structure and values
      read_wb = RubyXL::Parser.parse(filepath)
      assert_equal 1, read_wb.worksheets.size
      read_ws = read_wb["DropInReport"]
      refute_nil read_ws

      assert_equal "Product", read_ws[0][0].value
      assert_equal "Qty", read_ws[0][1].value
      assert_equal "Apples", read_ws[1][0].value
      assert_equal 10, read_ws[1][1].value
      assert_in_delta 1.5, read_ws[1][2].value, 0.001
      assert_equal "B2*C2", read_ws[1][3].formula.expression
      assert_equal "SUM(B2:B4)", read_ws[4][1].formula.expression
      assert_equal "SUM(D2:D4)", read_ws[4][3].formula.expression
    end
  end

  def test_streaming_dropin_workflow
    with_tempfile do |filepath|
      stream_pkg = Xlsxrb::Adapters::Caxlsx::StreamingPackage.new(filepath)
      wb = stream_pkg.workbook
      ws = wb.add_worksheet(name: "StreamData")

      100.times do |i|
        ws.add_row(["Row #{i + 1}", i * 10, (i * 1.25).round(2)])
      end

      stream_pkg.close
      assert File.exist?(filepath)
      assert File.size(filepath).positive?

      # Verify contents with Xlsxrb reader
      reader = Xlsxrb.read(filepath).load
      sheet = reader.sheets.first
      assert_equal "StreamData", sheet.name
      assert_equal 100, sheet.rows.size
      assert_equal "Row 1", sheet.rows[0].cells[0].value
      assert_equal 0, sheet.rows[0].cells[1].value
      assert_equal "Row 100", sheet.rows[99].cells[0].value
      assert_equal 990, sheet.rows[99].cells[1].value
    end
  end
end
