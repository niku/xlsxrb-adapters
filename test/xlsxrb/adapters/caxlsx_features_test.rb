# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCaxlsxFeaturesTest < Test::Unit::TestCase
  def test_merged_cells
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Merged")
    ws.add_row(["Header 1", "Header 2", "Header 3"])
    ws.add_row(["Data 1", "Data 2", "Data 3"])

    ws.merge_cells "A1:C1"
    assert_equal 1, ws.merged_cells.size
    assert_equal "A1:C1", ws.merged_cells.first

    # Multiple merges
    ws.merge_cells "A2:B2"
    assert_equal 2, ws.merged_cells.size
    assert_equal ["A1:C1", "A2:B2"], ws.merged_cells
  end

  def test_auto_filter
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Filtered")
    ws.add_row(%w[Name Department Salary])
    ws.add_row(["Alice", "Engineering", 120_000])
    ws.add_row(["Bob", "Marketing", 90_000])

    ws.auto_filter = "A1:C3"
    assert_equal "A1:C3", ws.auto_filter.range

    ws.auto_filter.add_column(1, :filters, filter_val: ["Engineering"])

    assert_equal 1, ws.auto_filter.columns.size
    assert_equal 1, ws.auto_filter.columns.first.col_id
  end

  def test_page_setup_and_margins
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "PageSetup")

    ws.page_setup.orientation = :landscape
    ws.page_setup.paper_size = 9 # A4
    ws.page_setup.fit_to_width = 1
    ws.page_setup.fit_to_height = 0

    assert_equal :landscape, ws.page_setup.orientation
    assert_equal 9, ws.page_setup.paper_size
    assert_equal 1, ws.page_setup.fit_to_width
    assert_equal 0, ws.page_setup.fit_to_height

    ws.page_margins.left = 1.0
    ws.page_margins.right = 1.0
    ws.page_margins.top = 1.2
    ws.page_margins.bottom = 1.2

    assert_in_delta 1.0, ws.page_margins.left, 0.001
    assert_in_delta 1.2, ws.page_margins.top, 0.001
  end

  def test_sheet_views_and_panes
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "FreezePanes")
    ws.add_row(["Header 1", "Header 2"])
    10.times { |i| ws.add_row(["Val #{i}", i * 10]) }

    ws.sheet_view.pane do |pane|
      pane.top_left_cell = "A2"
      pane.state = :frozen
      pane.y_split = 1
      pane.active_pane = :bottom_left
    end

    pane = ws.sheet_view.pane
    refute_nil pane
    assert_equal "A2", pane.top_left_cell
    assert_equal "frozen", pane.state
    assert_equal 1, pane.y_split
    assert_equal "bottomLeft", pane.active_pane
  end

  def test_sheet_protection
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Protected")

    ws.sheet_protection.password = "secure_pwd"
    ws.sheet_protection.sheet = true
    ws.sheet_protection.objects = true
    ws.sheet_protection.scenarios = true
    ws.sheet_protection.format_cells = false

    assert ws.sheet_protection.sheet
    assert ws.sheet_protection.objects
    assert ws.sheet_protection.scenarios
    refute ws.sheet_protection.format_cells
    refute_nil ws.sheet_protection.password
  end

  def test_comments
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Comments")
    ws.add_row(["Data with comment"])

    comment = ws.add_comment(ref: "A1", text: "Check this figure", author: "Reviewer")
    assert_equal 1, ws.comments.size
    assert_equal "A1", comment.ref
    assert_equal "Check this figure", comment.text
    assert_equal "Reviewer", comment.author
  end

  def test_data_validation
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Validation")
    ws.add_row([10])

    ws.add_data_validation(
      "A1:A10",
      type: :whole,
      operator: :between,
      formula1: "1",
      formula2: "100",
      showErrorMessage: true,
      errorTitle: "Input Error",
      error: "Value must be between 1 and 100"
    )

    assert_equal 1, ws.data_validations.size
    dv = ws.data_validations.first
    assert_equal "A1:A10", dv.sqref
    assert_equal :whole, dv.type
    assert_equal :between, dv.operator
  end

  def test_conditional_formatting
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "CondFormat")
    ws.add_row([50])
    ws.add_row([150])

    dxf_id = ws.workbook.styles.add_style(type: :dxf, b: true, fg_color: "FF0000")

    ws.add_conditional_formatting(
      "A1:A2",
      type: :cellIs,
      operator: :greaterThan,
      formula: "100",
      dxfId: dxf_id,
      priority: 1
    )

    assert_equal 1, ws.conditional_formattings.size
    cf = ws.conditional_formattings.first
    assert_equal "A1:A2", cf.sqref
    assert_equal 1, cf.rules.size
    rule = cf.rules.first
    assert_equal :cellIs, rule.type
    assert_equal :greaterThan, rule.operator
    assert_equal dxf_id, rule.dxfId
  end

  def test_tables
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "TableSheet")
    ws.add_row(%w[Product Qty Price])
    ws.add_row(["Keyboard", 2, 50.0])
    ws.add_row(["Mouse", 5, 25.0])

    table = ws.add_table("A1:C3", name: "ProductTable", style_info: { name: "TableStyleLight1" })
    assert_equal 1, ws.tables.size
    assert_equal "ProductTable", table.name
    assert_equal "A1:C3", table.ref
    assert_equal "TableStyleLight1", table.style_info.name
  end

  def test_pivot_tables
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "DataSheet")
    ws.add_row(%w[Year Quarter Revenue])
    ws.add_row(["2025", "Q1", 10_000])
    ws.add_row(["2025", "Q2", 15_000])
    ws.add_row(["2026", "Q1", 20_000])

    pivot_ws = pkg.workbook.add_worksheet(name: "PivotSheet")
    pt = pivot_ws.add_pivot_table("A3:D10", "DataSheet!A1:C4") do |pivot|
      pivot.rows = ["Year"]
      pivot.columns = ["Quarter"]
      pivot.data = [{ ref: "Revenue", subtotal: "sum" }]
    end

    assert_equal 1, pivot_ws.pivot_tables.size
    assert_equal "DataSheet!A1:C4", pt.range
    assert_equal ["Year"], pt.rows
    assert_equal ["Quarter"], pt.columns
  end

  def test_charts_creation_and_series
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "ChartSheet")
    ws.add_row(%w[Quarter Revenue])
    ws.add_row(["Q1", 100])
    ws.add_row(["Q2", 150])
    ws.add_row(["Q3", 200])

    # BarChart
    chart = ws.add_chart(Xlsxrb::Adapters::Caxlsx::BarChart, title: "Quarterly Revenue", start_at: [0, 5], end_at: [5, 20])
    chart.add_series(data: ws["B2:B4"], titles: ws["A2:A4"], title: "Rev")

    assert_instance_of Xlsxrb::Adapters::Caxlsx::BarChart, chart
    assert_equal "Quarterly Revenue", chart.title.text
    assert_equal 1, chart.series.size
    assert_equal 1, ws.drawing.charts.size

    # LineChart
    line_chart = ws.add_chart(Xlsxrb::Adapters::Caxlsx::LineChart, title: "Trend", start_at: [6, 5], end_at: [11, 20])
    line_chart.add_series(data: ws["B2:B4"])
    assert_instance_of Xlsxrb::Adapters::Caxlsx::LineChart, line_chart
    assert_equal 2, ws.drawing.charts.size

    # PieChart
    pie_chart = ws.add_chart(Xlsxrb::Adapters::Caxlsx::PieChart, title: "Share", start_at: [0, 22], end_at: [5, 37])
    pie_chart.add_series(data: ws["B2:B4"], labels: ws["A2:A4"])
    assert_instance_of Xlsxrb::Adapters::Caxlsx::PieChart, pie_chart
    assert_equal 3, ws.drawing.charts.size

    # Verify chart serialization
    stream = pkg.to_stream
    refute_nil stream
    assert stream.size.positive?
  end

  def test_hyperlinks
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Links")
    ws.add_row(["Visit Web", "Documentation"])

    ws.add_hyperlink(location: "https://github.com", ref: "A1")
    ws.add_hyperlink(location: "https://example.org", ref: "B1", display: "Docs")

    assert_equal 2, ws.hyperlinks.size
    assert_equal "https://github.com", ws.hyperlinks[0].location
    assert_equal "A1", ws.hyperlinks[0].ref
    assert_equal "Docs", ws.hyperlinks[1].display
  end

  def test_full_features_serialization_to_stream
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "Complete")
    ws.add_row(%w[Name Value])
    ws.add_row(["A", 10])
    ws.add_row(["B", 20])

    ws.merge_cells "A4:B4"
    ws.auto_filter = "A1:B3"
    ws.add_comment(ref: "A2", text: "Sample note", author: "Author")
    ws.add_hyperlink(location: "https://example.com", ref: "A1")

    stream = pkg.to_stream
    assert_instance_of StringIO, stream
    assert stream.size.positive?

    # Ensure it can be written to disk and is a valid zip
    with_tempfile do |path|
      assert pkg.serialize(path)
      assert File.exist?(path)
      assert File.size(path).positive?
    end
  end

  def test_rich_text_roundtrip_with_rich_text_run
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "RichTextSheet")

    rt = Xlsxrb::Adapters::Caxlsx::RichText.new
    rt.add_run("Normal ")
    rt.add_run("Bold", b: true, sz: 14, color: "FF0000")
    rt.add_run(" Italic", i: true)

    ws.add_row([rt])

    xlsxrb_wb = pkg.to_xlsxrb
    xlsxrb_sheet = xlsxrb_wb.sheet("RichTextSheet")
    cell_val = xlsxrb_sheet.rows[0].cells[0].value
    assert_instance_of Xlsxrb::Elements::RichText, cell_val
    assert_equal 3, cell_val.runs.size
    assert_instance_of Xlsxrb::Elements::RichTextRun, cell_val.runs[0]
    assert_equal "Normal ", cell_val.runs[0].text
    assert_nil cell_val.runs[0].font
    assert_instance_of Xlsxrb::Elements::RichTextRun, cell_val.runs[1]
    assert_equal "Bold", cell_val.runs[1].text
    assert_equal({ bold: true, sz: 14, color: "FFFF0000" }, cell_val.runs[1].font)
    assert_equal "Normal Bold Italic", cell_val.to_s

    # Verify serialization and reading via xlsxrb
    with_tempfile do |path|
      assert pkg.serialize(path)
      read_wb = Xlsxrb.read(path).load
      read_cell = read_wb.sheet(0).rows[0].cells[0]
      assert_equal "Normal Bold Italic", read_cell.value
    end
  end

  def test_date1904_calendar_system
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    pkg.workbook.date1904 = true
    ws = pkg.workbook.add_worksheet(name: "Dates")
    test_date = Date.new(2026, 9, 27)
    ws.add_row([test_date])

    xlsxrb_wb = pkg.to_xlsxrb
    assert xlsxrb_wb.date1904?
    assert_equal({ date1904: true }, xlsxrb_wb.unmapped_data[:workbook_properties])

    with_tempfile do |path|
      assert pkg.serialize(path)
      read_wb = Xlsxrb.read(path)
      assert read_wb.date1904?
    end
  end

  def test_first_class_conditional_formatting_and_data_validations
    pkg = Xlsxrb::Adapters::Caxlsx::Package.new
    ws = pkg.workbook.add_worksheet(name: "ValidationSheet")
    ws.add_row([10, 20, "Active"])

    ws.add_conditional_formatting("A1:A10", { type: :cellIs, operator: :greaterThan, formula: "15" })
    ws.add_data_validation("C1:C10", { type: :list, formula1: '"Active,Inactive"', allow_blank: true })

    xlsxrb_wb = pkg.to_xlsxrb
    xlsxrb_sheet = xlsxrb_wb.sheet(0)

    # First-class DOM assertions on Worksheet
    assert_equal 1, xlsxrb_sheet.conditional_formatting.size
    assert_equal "A1:A10", xlsxrb_sheet.conditional_formatting.first[:sqref]
    assert_equal :cellIs, xlsxrb_sheet.conditional_formatting.first[:type]

    assert_equal 1, xlsxrb_sheet.data_validations.size
    assert_equal "C1:C10", xlsxrb_sheet.data_validations.first[:sqref]
    assert_equal :list, xlsxrb_sheet.data_validations.first[:type]

    # Verify serialization and reading via xlsxrb
    with_tempfile do |path|
      assert pkg.serialize(path)
      read_wb = Xlsxrb.read(path).load
      assert_equal 1, read_wb.sheet(0).conditional_formatting.size
      assert_equal "A1:A10", read_wb.sheet(0).conditional_formatting.first[:sqref]
      assert_equal 1, read_wb.sheet(0).data_validations.size
      assert_equal "C1:C10", read_wb.sheet(0).data_validations.first[:sqref]
    end
  end
end
