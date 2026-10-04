# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"

module Xlsxrb
  module Adapters
    module Xsv
      # Helper methods and date/time/number formatting routines matching Xsv::Helpers.
      module Helpers
        # The default OOXML Spreadsheet number formats according to the ECMA standard
        # User formats are appended from index 174 onward
        BUILT_IN_NUMBER_FORMATS = {
          1 => "0",
          2 => "0.00",
          3 => "#, ##0",
          4 => "#, ##0.00",
          5 => "$#, ##0_);($#, ##0)",
          6 => "$#, ##0_);[Red]($#, ##0)",
          7 => "$#, ##0.00_);($#, ##0.00)",
          8 => "$#, ##0.00_);[Red]($#, ##0.00)",
          9 => "0%",
          10 => "0.00%",
          11 => "0.00E+00",
          12 => "# ?/?",
          13 => "# ??/??",
          14 => "m/d/yyyy",
          15 => "d-mmm-yy",
          16 => "d-mmm",
          17 => "mmm-yy",
          18 => "h:mm AM/PM",
          19 => "h:mm:ss AM/PM",
          20 => "h:mm",
          21 => "h:mm:ss",
          22 => "m/d/yyyy h:mm",
          37 => "#, ##0_);(#, ##0)",
          38 => "#, ##0_);[Red](#, ##0)",
          39 => "#, ##0.00_);(#, ##0.00)",
          40 => "#, ##0.00_);[Red](#, ##0.00)",
          45 => "mm:ss",
          46 => "[h]:mm:ss",
          47 => "mm:ss.0",
          48 => "##0.0E+0",
          49 => "@"
        }.freeze

        MINUTE = 60
        HOUR = 3600
        A_CODEPOINT = 65 # "A".ord
        # The epoch for all dates in OOXML Spreadsheet documents (1900 date system)
        EPOCH = Date.new(1899, 12, 30).freeze
        # The epoch for the 1904 date system (Mac Excel standard)
        EPOCH_1904 = Date.new(1904, 1, 1).freeze

        # Return the index number for the given Excel column name (i.e. "A1" => 0)
        # @param col [String] Column name in A1 notation
        # @return [Integer]
        #: (String col) -> Integer
        def column_index(col)
          coords = Xlsxrb::Utils.ref_to_row_col(col)
          return coords[1] if coords

          Xlsxrb::Utils.col_name_to_index(col)
        end

        # Return a Date for the given Excel date value
        # @param number [Numeric]
        # @param date1904 [Boolean, nil]
        # @return [Date]
        #: (Numeric number, ?date1904: bool?) -> Date
        def parse_date(number, date1904: nil)
          is1904 = if date1904.nil?
                     defined?(@workbook) && @workbook ? @workbook.date1904? : false
                   else
                     date1904
                   end
          base = is1904 ? EPOCH_1904 : EPOCH
          base + number
        end

        # Return a time as a string for the given Excel time value
        # @param number [Numeric]
        # @return [String]
        #: (Numeric number) -> String
        def parse_time(number)
          # Disregard date part
          num = number.positive? ? number - number.truncate : number

          base = num * 24

          hours = base.truncate
          minutes = ((base - hours) * 60).round

          # Compensate for rounding errors
          if minutes >= 60
            hours += (minutes / 60)
            minutes %= 60
          end

          format("%<hours>02d:%<minutes>02d", hours: hours, minutes: minutes)
        end

        # Returns a time including a date as a {Time} object
        # @param number [Numeric]
        # @param date1904 [Boolean, nil]
        # @return [Time]
        #: (Numeric number, ?date1904: bool?) -> Time
        def parse_datetime(number, date1904: nil)
          date_base = number.truncate
          time = parse_date(date_base, date1904: date1904).to_time

          time_base = (number - date_base) * 24

          hours = time_base.truncate
          minutes = (time_base - hours) * 60

          time + (hours * HOUR) + (minutes.round * MINUTE)
        end

        # Returns a number as either Integer or Float
        # @param string_or_num [String, Numeric]
        # @return [Numeric]
        #: (untyped string_or_num) -> Numeric
        def parse_number(string_or_num)
          return string_or_num if string_or_num.is_a?(Numeric)

          string = string_or_num.to_s
          if string.include? "."
            string.to_f
          elsif string.include?("E") || string.include?("e")
            Complex(string).to_f
          else
            string.to_i
          end
        end

        # Apply date or time number formats, if applicable
        # @param number [String, Numeric]
        # @param format [String, nil]
        # @param style_id [Integer, nil]
        # @return [Object]
        #: (untyped number, String? format, ?style_id: Integer?) -> untyped
        def parse_number_format(number, format, style_id: nil)
          num = parse_number(number)
          return num if format.nil? && style_id.nil?

          if style_id && defined?(@workbook) && @workbook&.styles.is_a?(Xlsxrb::Elements::Styles)
            case @workbook.styles.format_type(style_id)
            when :datetime
              return parse_datetime(num)
            when :date
              return parse_date(num)
            when :time
              return parse_time(num)
            when :number, :text, :general
              return num
            end
          end

          return num if format.nil?

          is_date_format = format.scan(/[dmy]+/).length > 1
          is_time_format = format.scan(/[hms]+/).length > 1

          if !is_date_format && !is_time_format
            num
          elsif is_date_format && is_time_format
            parse_datetime(num)
          elsif is_date_format
            parse_date(num)
          elsif is_time_format
            parse_time(num)
          end
        end
      end
    end
  end
end
