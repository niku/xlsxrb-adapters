# frozen_string_literal: true

# rbs_inline: enabled

require "stringio"

module Xlsxrb
  module Adapters
    module Roo
      module Formatters
        # CSV export formatter matching Roo::Formatters::CSV.
        module CSV
          # Exports the spreadsheet to CSV format.
          #
          # @param filename [String, File, nil]
          # @param old_separator [String, nil]
          # @param old_sheet [String, Integer, nil]
          # @param separator [String]
          # @param sheet [String, Integer]
          # @return [String, true]
          #: (?String? filename, ?String? old_separator, ?untyped old_sheet, ?separator: String, ?sheet: untyped) -> (String | bool)
          def to_csv(filename = nil, old_separator = nil, old_sheet = nil, separator: ",", sheet: default_sheet)
            sep = old_separator || separator
            target_sheet = old_sheet || sheet

            if filename
              if filename.respond_to?(:write)
                write_csv_content(filename, target_sheet, sep)
              else
                File.open(filename.to_s, "w") do |file|
                  write_csv_content(file, target_sheet, sep)
                end
              end
              true
            else
              sio = StringIO.new
              write_csv_content(sio, target_sheet, sep)
              sio.string
            end
          end

          private

          #: (untyped stream, untyped sheet, String separator) -> void
          def write_csv_content(stream, sheet, separator)
            return unless first_row(sheet)

            1.upto(last_row(sheet)) do |r|
              1.upto(last_column(sheet)) do |c|
                stream.print(separator) if c > 1
                stream.print(cell_to_csv(r, c, sheet))
              end
              stream.print("\n")
            end
          end

          #: (Integer row, Integer col, untyped sheet) -> String
          def cell_to_csv(row, col, sheet)
            return "" if empty?(row, col, sheet)

            val = cell(row, col, sheet)
            type = celltype(row, col, sheet)

            case type
            when :string
              val_str = val.to_s
              val_str.empty? ? "" : %("#{val_str.gsub('"', '""')}")
            when :boolean
              b_str = val ? "true" : "false"
              %("#{b_str}")
            when :float, :percentage
              num = val.to_f
              num == num.to_i ? num.to_i.to_s : num.to_s
            when :formula
              case val
              when String
                val.empty? ? "" : %("#{val.gsub('"', '""')}")
              when Float
                val == val.to_i ? val.to_i.to_s : val.to_s
              else
                val.to_s
              end
            when :time
              integer_to_timestring(val)
            when :link
              href = val.respond_to?(:url) ? val.url : val.to_s
              %("#{href.gsub('"', '""')}")
            else
              val.to_s
            end || ""
          end
        end
      end
    end
  end
end
