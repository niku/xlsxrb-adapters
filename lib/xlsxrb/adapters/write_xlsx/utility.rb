# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require_relative "constants"

module Xlsxrb
  module Adapters
    module Writexlsx
      # Collection of utilities for cell references, coordinates, and conversions
      module Utility
        # Cell coordinate and reference calculation utilities
        module CellReference
          extend self

          # Converts 0-based row and column to an A1 cell reference string
          #
          # @param row_or_name [Integer, String]
          # @param col [Integer]
          # @param row_absolute [Boolean]
          # @param col_absolute [Boolean]
          # @return [String]
          #: (Integer | String row_or_name, Integer col, ?bool row_absolute, ?bool col_absolute) -> String
          def xl_rowcol_to_cell(row_or_name, col, row_absolute = false, col_absolute = false)
            row_str = row_or_name.is_a?(Integer) ? (row_or_name + 1).to_s : row_or_name.to_s
            col_str = xl_col_to_name(col, col_absolute)
            "#{col_str}#{absolute_char(row_absolute)}#{row_str}"
          end

          # Converts an A1 cell reference into 0-based [row, col, row_abs, col_abs]
          #
          # @param cell [String]
          # @return [Array<Integer, Integer, Boolean, Boolean>]
          #: (String cell) -> [Integer, Integer, bool, bool]
          def xl_cell_to_rowcol(cell)
            cell.to_s.strip.upcase =~ /(\$?)([A-Z]{1,3})(\$?)(\d+)/
            raise WriteXLSXDimensionError, "Invalid cell coordinate: #{cell}" unless ::Regexp.last_match

            col_abs = ::Regexp.last_match(1) == "$"
            col_letters = ::Regexp.last_match(2)
            row_abs = ::Regexp.last_match(3) == "$"
            row_num = ::Regexp.last_match(4).to_i

            col = 0
            col_letters.each_char do |ch|
              col = (col * 26) + (ch.ord - "A".ord + 1)
            end
            col -= 1
            row = row_num - 1

            [row, col, row_abs, col_abs]
          end

          # Converts 0-based column index to Excel column name (e.g. 0 -> "A", 27 -> "AB")
          #
          # @param col [Integer]
          # @param col_absolute [Boolean]
          # @return [String]
          #: (Integer col, ?bool col_absolute) -> String
          def xl_col_to_name(col, col_absolute = false)
            col_num = col.to_i
            res = +""
            loop do
              res.prepend(("A".ord + (col_num % 26)).chr)
              col_num = (col_num / 26) - 1
              break if col_num.negative?
            end
            Common.ptrue?(col_absolute) ? "$#{res}" : res
          end

          # Returns an A1 range string for two cell coordinates
          #
          # @param row_1 [Integer]
          # @param row_2 [Integer]
          # @param col_1 [Integer]
          # @param col_2 [Integer]
          # @param row_abs_1 [Boolean]
          # @param row_abs_2 [Boolean]
          # @param col_abs_1 [Boolean]
          # @param col_abs_2 [Boolean]
          # @return [String]
          #: (Integer row_1, Integer row_2, Integer col_1, Integer col_2, ?bool row_abs_1, ?bool row_abs_2, ?bool col_abs_1, ?bool col_abs_2) -> String
          def xl_range(row_1, row_2, col_1, col_2,
                       row_abs_1 = false, row_abs_2 = false, col_abs_1 = false, col_abs_2 = false)
            range1 = xl_rowcol_to_cell(row_1, col_1, row_abs_1, col_abs_1)
            range2 = xl_rowcol_to_cell(row_2, col_2, row_abs_2, col_abs_2)
            range1 == range2 ? range1 : "#{range1}:#{range2}"
          end

          # Returns a formula string pointing to a sheet and cell range
          #
          # @param sheetname [String]
          # @param row_1 [Integer]
          # @param row_2 [Integer]
          # @param col_1 [Integer]
          # @param col_2 [Integer]
          # @return [String]
          #: (String sheetname, Integer row_1, Integer row_2, Integer col_1, Integer col_2) -> String
          def xl_range_formula(sheetname, row_1, row_2, col_1, col_2)
            sheet = quote_sheetname(sheetname)
            range1 = xl_rowcol_to_cell(row_1, col_1, true, true)
            range2 = xl_rowcol_to_cell(row_2, col_2, true, true)
            "=#{sheet}!#{range1}:#{range2}"
          end

          # Quotes a worksheet name with single quotes if required
          #
          # @param sheetname [String]
          # @return [String]
          #: (String sheetname) -> String
          def quote_sheetname(sheetname)
            name = sheetname.to_s
            return name if name.start_with?("'") && name.end_with?("'")

            if name =~ /\A[a-zA-Z_]\w*\z/
              name
            else
              "'#{name.gsub("'", "''")}'"
            end
          end

          # Returns "$" if absolute is true, else empty string
          #
          # @param absolute [Boolean]
          # @return [String]
          #: (bool absolute) -> String
          def absolute_char(absolute)
            Common.ptrue?(absolute) ? "$" : ""
          end

          # Substitute an A1 reference into 0-based row/col arguments
          #
          # @param cell [String]
          # @param args [Array<untyped>]
          # @return [Array<untyped>]
          def substitute_cellref(cell, *args)
            normalized = cell.to_s.strip.upcase
            case normalized
            when /\$?([A-Z]{1,3}):\$?([A-Z]{1,3})/
              r1, c1 = xl_cell_to_rowcol("#{::Regexp.last_match(1)}1")
              r2, c2 = xl_cell_to_rowcol("#{::Regexp.last_match(2)}#{ROW_MAX}")
              [r1, c1, r2, c2, *args]
            when /\$?([A-Z]{1,3}\$?\d+):\$?([A-Z]{1,3}\$?\d+)/
              r1, c1 = xl_cell_to_rowcol(::Regexp.last_match(1))
              r2, c2 = xl_cell_to_rowcol(::Regexp.last_match(2))
              [r1, c1, r2, c2, *args]
            when /\$?([A-Z]{1,3}\$?\d+)/
              r1, c1 = xl_cell_to_rowcol(::Regexp.last_match(1))
              [r1, c1, *args]
            else
              raise "Unknown cell reference #{normalized}"
            end
          end

          # Checks if row parameter is in A1 notation and substitutes row/col
          #
          # @param row_or_a1 [Object]
          # @return [Array<untyped>, nil]
          def row_col_notation(row_or_a1)
            return nil unless row_or_a1.is_a?(String)

            str = row_or_a1.to_s.strip
            substitute_cellref(str) if str =~ /^\D/
          end
        end

        # General utility helpers
        module Common
          extend self

          # Tests if a value is truthy under Perl/WriteXLSX conventions
          #
          # @param val [Object]
          # @return [Boolean]
          #: (untyped val) -> bool
          def ptrue?(val)
            !val.nil? && val != false && val != 0 && val != "0" && val != ""
          end

          # Emits deprecation warning message
          #
          # @param msg [String]
          # @return [void]
          #: (String msg) -> void
          def put_deprecate_message(msg)
            warn "WriteXLSX Deprecation: #{msg}"
          end

          # Converts a float to an Excel-compatible string
          #
          # @param val [Numeric]
          # @return [String]
          #: (Numeric val) -> String
          def float_to_str(val)
            val.to_s
          end
        end

        # Date and time parsing / serial conversion helpers
        module DateTime
          extend self

          # Converts an ISO 8601 date/time string or Ruby date/time to Excel serial number
          #
          # @param value [String, Date, Time, DateTime]
          # @param date1904 [Boolean]
          # @return [Float, nil]
          #: (untyped value, ?bool date1904) -> Float?
          def convert_date_time(value, date1904 = false)
            return nil if value.nil?

            if value.is_a?(Time) || (defined?(DateTime) && value.is_a?(DateTime))
              return Xlsxrb::Ooxml::Utils.datetime_to_serial(value, date1904: date1904)
            elsif value.is_a?(Date)
              return Xlsxrb::Ooxml::Utils.date_to_serial(value, date1904: date1904).to_f
            end

            str = value.to_s.strip.sub(/Z$/, "")
            return nil if str =~ /[^0-9T:\-.Z]/

            date_part, time_part = str.split("T")
            time_part = nil if time_part && time_part.empty?

            seconds_fraction = 0.0
            if time_part
              return nil unless time_part =~ /^(\d\d):(\d\d)(:(\d\d(?:\.\d+)?))?/

              h = ::Regexp.last_match(1).to_i
              m = ::Regexp.last_match(2).to_i
              s = ::Regexp.last_match(4).to_f
              return nil if h >= 24 || m >= 60 || s >= 60

              seconds_fraction = ((h * 3600.0) + (m * 60.0) + s) / 86_400.0

            end

            return seconds_fraction if date_part.nil? || date_part.empty?

            return unless date_part =~ /^(\d{4})-(\d{2})-(\d{2})$/

            year = ::Regexp.last_match(1).to_i
            month = ::Regexp.last_match(2).to_i
            day = ::Regexp.last_match(3).to_i
            begin
              d = ::Date.new(year, month, day)
              days = Xlsxrb::Ooxml::Utils.date_to_serial(d, date1904: date1904)
              days.to_f + seconds_fraction
            rescue ArgumentError
              nil
            end
          end
        end

        include CellReference
        include Common
        include DateTime
        extend CellReference
        extend Common
        extend DateTime
      end
    end
  end
end
