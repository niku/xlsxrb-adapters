# frozen_string_literal: true

require "test_helper"
require "date"

# Cross-engine round-trip verification across 5 Writers x 5 Readers (25 pipelines)
class XlsxrbAdaptersCrossAdapterRoundTripTest < Test::Unit::TestCase
  WRITERS = %i[caxlsx write_xlsx fast_excel xlsxtream ruby_xl].freeze
  READERS = %i[roo creek xsv simple_xlsx_reader ruby_xl].freeze

  TEST_DATA = [
    ["int_0", 0],
    ["int_pos", 42],
    ["int_neg", -100],
    ["float_pi", 3.14],
    ["str_simple", "Hello"],
    ["str_amp", "A & B"],
    ["str_quote", "He said \"Hi\""],
    ["str_unicode", "日本語テスト"],
    ["date_standard", Date.new(2026, 10, 24)]
  ].freeze

  def write_with(writer_sym, path)
    case writer_sym
    when :caxlsx
      p = Xlsxrb::Adapters::Caxlsx::Package.new
      p.workbook.add_worksheet(name: "Data") do |sheet|
        TEST_DATA.each { |row| sheet.add_row(row) }
      end
      p.serialize(path)
    when :write_xlsx
      wb = Xlsxrb::Adapters::Writexlsx::Workbook.new(path)
      date_fmt = wb.add_format(num_format: "yyyy-mm-dd")
      ws = wb.add_worksheet("Data")
      TEST_DATA.each_with_index do |row, r|
        row.each_with_index do |val, c|
          if val.is_a?(Date)
            ws.write_date_time(r, c, "#{val}T", date_fmt)
          else
            ws.write(r, c, val)
          end
        end
      end
      wb.close
    when :fast_excel
      wb = Xlsxrb::Adapters::FastExcel.open(path, constant_memory: false)
      ws = wb.add_worksheet("Data")
      TEST_DATA.each { |row| ws.append_row(row) }
      wb.close
    when :xlsxtream
      Xlsxrb::Adapters::Xlsxtream::Workbook.open(path) do |wb|
        wb.write_worksheet("Data") do |ws|
          TEST_DATA.each { |row| ws << row }
        end
      end
    when :ruby_xl
      wb = Xlsxrb::Adapters::RubyXL::Workbook.new
      ws = wb[0]
      ws.sheet_name = "Data"
      TEST_DATA.each_with_index do |row, r|
        row.each_with_index do |val, c|
          c_obj = ws.add_cell(r, c, val)
          c_obj.set_number_format("yyyy-mm-dd") if val.is_a?(Date)
        end
      end
      wb.write(path)
    end
  end

  def read_with(reader_sym, path)
    case reader_sym
    when :roo
      doc = Xlsxrb::Adapters::Roo::Excelx.new(path)
      (1..TEST_DATA.size).map do |r|
        [doc.cell(r, 1), doc.cell(r, 2)]
      end
    when :creek
      doc = Xlsxrb::Adapters::Creek::Book.new(path)
      doc.sheets.first.rows.map do |row|
        [row.values[0], row.values[1]]
      end
    when :xsv
      doc = Xlsxrb::Adapters::Xsv::Workbook.open(path)
      doc.sheets.first.to_a.map do |row|
        [row[0], row[1]]
      end
    when :simple_xlsx_reader
      doc = Xlsxrb::Adapters::SimpleXlsxReader.open(path)
      doc.sheets.first.rows.to_a.map do |row|
        [row[0], row[1]]
      end
    when :ruby_xl
      doc = Xlsxrb::Adapters::RubyXL::Parser.parse(path)
      ws = doc[0]
      (0...TEST_DATA.size).map do |r|
        row = ws[r]
        [row ? row[0]&.value : nil, row ? row[1]&.value : nil]
      end
    end
  end

  WRITERS.each do |w|
    READERS.each do |r|
      define_method("test_round_trip_#{w}_to_#{r}") do
        with_tempfile do |path|
          write_with(w, path)
          rows = read_with(r, path)

          assert_equal TEST_DATA.size, rows.size, "Row count mismatch for #{w} -> #{r}"

          TEST_DATA.each_with_index do |(expected_key, expected_val), idx|
            actual_key, actual_val = rows[idx]
            assert_equal expected_key, actual_key, "#{w} -> #{r} row #{idx} key mismatch"

            if expected_val.is_a?(Numeric)
              assert_in_delta expected_val.to_f, actual_val.to_f, 0.001, "#{w} -> #{r} row #{idx} numeric mismatch"
            elsif expected_val.is_a?(Date)
              date_matched = actual_val.to_s.start_with?("2026-10-24") || actual_val.to_f.round == 46319
              assert date_matched, "#{w} -> #{r} row #{idx} date mismatch: expected #{expected_val}, got #{actual_val}"
            else
              assert_equal expected_val.to_s, actual_val.to_s, "#{w} -> #{r} row #{idx} string mismatch"
            end
          end
        end
      end
    end
  end
end
