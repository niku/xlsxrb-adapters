# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      class Excelx
        # Number and date format parsing utility matching Roo::Excelx::Format.
        module Format
          module_function

          EXCEPTIONAL_FORMATS = {
            "h:mm am/pm" => :date,
            "h:mm:ss am/pm" => :date
          }.freeze #: Hash[String, Symbol]

          STANDARD_FORMATS = {
            0 => "General",
            1 => "0",
            2 => "0.00",
            3 => "#,##0",
            4 => "#,##0.00",
            9 => "0%",
            10 => "0.00%",
            11 => "0.00E+00",
            12 => "# ?/?",
            13 => "# ??/??",
            14 => "mm-dd-yy",
            15 => "d-mmm-yy",
            16 => "d-mmm",
            17 => "mmm-yy",
            18 => "h:mm AM/PM",
            19 => "h:mm:ss AM/PM",
            20 => "h:mm",
            21 => "h:mm:ss",
            22 => "m/d/yy h:mm",
            37 => "#,##0 ;(#,##0)",
            38 => "#,##0 ;[Red](#,##0)",
            39 => "#,##0.00;(#,##0.00)",
            40 => "#,##0.00;[Red](#,##0.00)",
            45 => "mm:ss",
            46 => "[h]:mm:ss",
            47 => "mmss.0",
            48 => "##0.0E+0",
            49 => "@"
          }.freeze #: Hash[Integer, String]

          # Determines data type Symbol from Excel format code string.
          #
          # @param format [String, nil]
          # @return [Symbol]
          #: (untyped format) -> Symbol
          def to_type(format)
            fmt = format.to_s.downcase
            exceptional = EXCEPTIONAL_FORMATS[fmt]
            return exceptional if exceptional

            return :float if fmt.include?("#")

            if fmt.include?("y") || fmt.match?(/d+(?!\])/)
              if fmt.include?("h") || fmt.include?("s")
                :datetime
              else
                :date
              end
            elsif fmt.include?("h") || fmt.include?("s")
              :time
            elsif fmt.include?("%")
              :percentage
            else
              :float
            end
          end
        end
      end
    end
  end
end
