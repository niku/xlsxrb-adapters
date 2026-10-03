# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require_relative "constants"
require_relative "coordinate"
require_relative "link"
require_relative "format"

module Xlsxrb
  module Adapters
    module Roo
      class Excelx
        # Container for Excel cell objects matching Roo::Excelx::Cell.
        class Cell
          # Base class for all Excelx cell types.
          class Base
            attr_reader :cell_value, :cell_type, :formula, :coordinate, :format
            attr_accessor :value

            # @param value [Object]
            # @param formula [String, nil]
            # @param excelx_type [Object]
            # @param style [Integer, nil]
            # @param link [String, nil]
            # @param coordinate [Coordinate, nil]
            #: (untyped value, ::String? formula, untyped excelx_type, Integer? style, ::String? link, Coordinate? coordinate) -> void
            def initialize(value, formula, excelx_type, style, link, coordinate)
              @cell_value = value
              @cell_type = excelx_type if excelx_type
              @formula = formula if formula
              @style = style unless style == 1
              @coordinate = coordinate
              @format = excelx_type.is_a?(Array) ? excelx_type.last : excelx_type
              @value = link ? Link.new(link, value.to_s) : value
            end

            # @return [Integer, nil]
            #: () -> Integer?
            def style
              defined?(@style) ? @style : 1
            end

            # @return [Symbol]
            #: () -> Symbol?
            def type
              if formula?
                :formula
              elsif link?
                :link
              else
                default_type
              end
            end

            # @return [Symbol]
            #: () -> Symbol?
            def default_type
              :base
            end

            # @return [Boolean]
            #: () -> bool
            def formula?
              !!(defined?(@formula) && @formula)
            end

            # @return [Boolean]
            #: () -> bool
            def link?
              @value.is_a?(Link)
            end

            # @return [Boolean]
            #: () -> bool
            def empty?
              false
            end

            # @return [String, Object]
            #: () -> untyped
            def formatted_value
              @value
            end

            # @return [String]
            #: () -> ::String
            def to_s
              formatted_value.to_s
            end

            # @return [Boolean]
            #: () -> bool
            def hyperlink
              link?
            end

            # @return [Boolean]
            #: () -> bool
            def link
              link?
            end

            # @return [Object]
            #: () -> untyped
            def excelx_value
              cell_value
            end

            # @return [Object]
            #: () -> untyped
            def excelx_type
              cell_type
            end
          end

          # Number cell representation.
          class Number < Base
            attr_reader :format

            # @param value [Object]
            # @param formula [String, nil]
            # @param excelx_type [Object]
            # @param style [Integer, nil]
            # @param link [String, nil]
            # @param coordinate [Coordinate, nil]
            #: (untyped value, ::String? formula, untyped excelx_type, Integer? style, ::String? link, Coordinate? coordinate) -> void
            def initialize(value, formula, excelx_type, style, link, coordinate)
              super(value, formula, excelx_type, style, nil, coordinate)
              @format = excelx_type.is_a?(Array) ? excelx_type.last.to_s : excelx_type.to_s
              @value = link ? Link.new(link, value.to_s) : create_numeric(value)
            end

            # @return [Symbol]
            #: () -> Symbol
            def default_type
              :float
            end

            # Converts raw value string to Integer or Float.
            #
            # @param number [Object]
            # @return [Numeric, String, nil]
            #: (untyped number) -> untyped
            def create_numeric(number)
              return number if Roo::ERROR_VALUES.include?(number)

              str = number.to_s.strip
              return nil if str.empty?

              case @format
              when /%/, /\.0/
                Float(str)
              else
                if str.include?(".") || (/\A[-+]?\d+E[-+]?\d+\z/i =~ str)
                  Float(str)
                else
                  Integer(str, 10)
                end
              end
            rescue StandardError
              number
            end

            # Returns formatted number representation matching format code.
            #
            # @return [String]
            #: () -> ::String
            def formatted_value
              return @cell_value.to_s if Roo::ERROR_VALUES.include?(@cell_value)

              formatter = generate_formatter(@format)
              if formatter.is_a?(Proc)
                formatter.call(@cell_value)
              else
                Kernel.format(formatter, @cell_value)
              end
            rescue StandardError
              @cell_value.to_s
            end

            private

            # Generates a formatter string or Proc for the given format code.
            #: (::String format) -> untyped
            def generate_formatter(format)
              case format
              when /^General$/i then general_formatter
              when "0" then "%.0f"
              when /^(0+)$/ then "%0#{::Regexp.last_match(1).size}d"
              when /^0\.(0+)$/ then "%.#{::Regexp.last_match(1).size}f"
              when "#,##0" then number_format("%.0f")
              when /^#,##0\.(0+)$/ then number_format("%.#{::Regexp.last_match(1).size}f")
              when "0%"
                proc { |num| Kernel.format("%.0f%%", num.to_f * 100) }
              when "0.00%"
                proc { |num| Kernel.format("%.2f%%", num.to_f * 100) }
              when "0.00E+00" then "%.2E"
              when "#,##0 ;(#,##0)" then number_format("%.0f", "(%.0f)")
              when "#,##0 ;[Red](#,##0)" then number_format("%.0f", "[Red](%.0f)")
              when "#,##0.00;(#,##0.00)" then number_format("%.2f", "(%.2f)")
              when "#,##0.00;[Red](#,##0.00)" then number_format("%.2f", "[Red](%.2f)")
              when "##0.0E+0" then "%.1E"
              when "_-* #,##0.00\\ _€_-;\\-* #,##0.00\\ _€_-;_-* \"-\"??\\ _€_-;_-@_-"
                number_format("%.2f", "-%.2f")
              when /^(?:_\()?"([^"]*)"(?:\* )?([^_]+)/, /^_[- (]\[\$([^-]*)[^#@]+([^_]+)/
                prefix = ::Regexp.last_match(1)
                sub_fmt = ::Regexp.last_match(2)
                proc do |num|
                  formatted_num = generate_formatter(sub_fmt).call(num)
                  "#{prefix}#{formatted_num}"
                end
              when /\A([$€£])([#0,.]+)\z/
                sym = ::Regexp.last_match(1)
                sub_fmt = ::Regexp.last_match(2)
                proc do |num|
                  sub_res = if sub_fmt.include?(".")
                              number_format("%.#{sub_fmt.split(".").last.size}f").call(num)
                            else
                              number_format("%.0f").call(num)
                            end
                  "#{sym}#{sub_res}"
                end
              else
                proc { |num| Xlsxrb::NumberFormatter.format(num, format) }
              end
            end

            #: () -> Proc
            def general_formatter
              proc do |num|
                val = num.is_a?(Numeric) ? num : Float(num.to_s)
                val == val.to_i ? val.to_i.to_s : val.to_s
              rescue StandardError
                num.to_s
              end
            end

            #: (::String formatter, ?::String? negative_formatter) -> Proc
            def number_format(formatter, negative_formatter = nil)
              proc do |num|
                f = formatter
                n = num.to_f
                if negative_formatter && n.negative?
                  f = negative_formatter
                  n = n.abs
                end

                formatted = Kernel.format(f, n)
                formatted.reverse.gsub(/(\d{3})(?=\d)/, '\1,').reverse
              end
            end
          end

          # Date and DateTime base cell representation.
          class DateTime < Base
            SECONDS_IN_DAY = 86_400

            DATE_FORMATS = {
              "yyyy" => "%Y",
              "yy" => "%y",
              "mmmm" => "%B",
              "mmm" => "%^b",
              "mm" => "%m",
              "m" => "%-m",
              "dddd" => "%A",
              "ddd" => "%^a",
              "dd" => "%d",
              "d" => "%-d"
            }.freeze #: Hash[::String, ::String]

            TIME_FORMATS = {
              "hh" => "%H",
              "h" => "%-k",
              "mm" => "%M",
              "m" => "%-M",
              "ss" => "%S",
              "s" => "%-S",
              "am/pm" => "%p",
              "000" => "%3N",
              "00" => "%2N",
              "0" => "%1N"
            }.freeze #: Hash[::String, ::String]

            attr_reader :format

            # @param value [Object]
            # @param formula [String, nil]
            # @param excelx_type [Object]
            # @param style [Integer, nil]
            # @param link [String, nil]
            # @param base_timestamp [Integer, nil]
            # @param coordinate [Coordinate, nil]
            #: (untyped value, ::String? formula, untyped excelx_type, Integer? style, ::String? link, Integer? base_timestamp, Coordinate? coordinate) -> void
            def initialize(value, formula, excelx_type, style, link, base_timestamp, coordinate)
              super(value, formula, excelx_type, style, nil, coordinate)
              @format = excelx_type.is_a?(Array) ? excelx_type.last.to_s : excelx_type.to_s
              @value = link ? Link.new(link, value.to_s) : create_datetime(base_timestamp, value)
            end

            # @return [Symbol]
            #: () -> Symbol
            def default_type
              :datetime
            end

            # @return [String]
            #: () -> ::String
            def formatted_value
              return @value.strftime("%F %T") unless @value.respond_to?(:strftime)

              parts = @format.to_s.downcase.split
              parsed_parts = parts.map { |part| parse_date_or_time_format(part) }
              return @value.strftime("%F %T") if parsed_parts.include?(false)

              formatter = parsed_parts.join(" ")
              @value.strftime(formatter)
            rescue StandardError
              @value.strftime("%F %T")
            end

            private

            #: (Integer? base_timestamp, untyped value) -> untyped
            def create_datetime(base_timestamp, value)
              return value if value.is_a?(::DateTime)
              return value.to_datetime if value.is_a?(::Time)
              return value.to_datetime if value.is_a?(::Date)

              ts = base_timestamp || 0
              timestamp = (ts + (value.to_f.round(6) * SECONDS_IN_DAY)).round(0)
              ::Time.at(timestamp).utc.to_datetime
            end

            #: (::String part) -> (::String | false)
            def parse_date_or_time_format(part)
              date_regex = %r{(?<date>[dmy]+[-/][dmy]+(?:[-/][dmy]+)?)}
              time_regex = /(?<time>(?:\[?h\]?:)?m+(?::?ss|:?s)?)/

              if part[date_regex] == part
                part.gsub(/#{DATE_FORMATS.keys.join("|")}/, DATE_FORMATS)
              elsif part[time_regex]
                part.gsub(/#{TIME_FORMATS.keys.join("|")}/, TIME_FORMATS)
              else
                false
              end
            end
          end

          # Date cell representation.
          class Date < DateTime
            # @param value [Object]
            # @param formula [String, nil]
            # @param excelx_type [Object]
            # @param style [Integer, nil]
            # @param link [String, nil]
            # @param base_date [::Date, nil]
            # @param coordinate [Coordinate, nil]
            #: (untyped value, ::String? formula, untyped excelx_type, Integer? style, ::String? link, ::Date? base_date, Coordinate? coordinate) -> void
            def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
              super(value, formula, excelx_type, style, nil, nil, coordinate)
              @format = excelx_type.is_a?(Array) ? excelx_type.last.to_s : excelx_type.to_s
              @value = link ? Link.new(link, value.to_s) : create_date(base_date, value)
            end

            # @return [Symbol]
            #: () -> Symbol
            def default_type
              :date
            end

            # @return [String]
            #: () -> ::String
            def formatted_value
              return @value.to_s unless @value.is_a?(::Date) || @value.respond_to?(:strftime)

              if @format && !@format.empty?
                formatted = Xlsxrb::NumberFormatter.format(@value, @format)
                return formatted unless formatted.empty?
              end

              @value.strftime("%Y-%m-%d")
            rescue StandardError
              @value.respond_to?(:strftime) ? @value.strftime("%Y-%m-%d") : @value.to_s
            end

            private

            #: (::Date? base_date, untyped value) -> untyped
            def create_date(base_date, value)
              return value if value.is_a?(::Date)

              base = base_date || ::Date.new(1899, 12, 30)
              base + value.to_i
            end
          end

          # Time cell representation.
          class Time < DateTime
            # @param value [Object]
            # @param formula [String, nil]
            # @param excelx_type [Object]
            # @param style [Integer, nil]
            # @param link [String, nil]
            # @param base_date [::Date, nil]
            # @param coordinate [Coordinate, nil]
            #: (untyped value, ::String? formula, untyped excelx_type, Integer? style, ::String? link, ::Date? base_date, Coordinate? coordinate) -> void
            def initialize(value, formula, excelx_type, style, link, base_date, coordinate)
              super(value, formula, excelx_type, style, nil, nil, coordinate)
              @format = excelx_type.is_a?(Array) ? excelx_type.last.to_s : excelx_type.to_s
              base_ts = base_date ? base_date.to_datetime.to_time.to_i : 0
              @datetime = create_datetime(base_ts, value)
              @value = link ? Link.new(link, value.to_s) : (value.to_f * 86_400).round.to_i
            end

            # @return [Symbol]
            #: () -> Symbol
            def default_type
              :time
            end

            # @return [String]
            #: () -> ::String
            def formatted_value
              formatter = @format.gsub(/#{TIME_FORMATS.keys.join("|")}/, TIME_FORMATS)
              @datetime.respond_to?(:strftime) ? @datetime.strftime(formatter) : @value.to_s
            rescue StandardError
              @datetime.to_s
            end

            alias to_s formatted_value
          end

          # Boolean cell representation.
          class Boolean < Base
            # @param value [Object]
            # @param formula [String, nil]
            # @param style [Integer, nil]
            # @param link [String, nil]
            # @param coordinate [Coordinate, nil]
            #: (untyped value, ::String? formula, Integer? style, ::String? link, Coordinate? coordinate) -> void
            def initialize(value, formula, style, link, coordinate)
              super(value, formula, nil, style, nil, coordinate)
              @cell_type = :boolean
              @value = link ? Link.new(link, value.to_s) : create_boolean(value)
            end

            # @return [Symbol]
            #: () -> Symbol
            def default_type
              :boolean
            end

            # @return [String]
            #: () -> ::String
            def formatted_value
              @value ? "TRUE" : "FALSE"
            end

            private

            #: (untyped value) -> bool
            def create_boolean(value)
              value == true || value.to_s == "1" || value.to_s.casecmp?("true")
            end
          end

          # String cell representation.
          class String < Base
            # @param value [Object]
            # @param formula [String, nil]
            # @param style [Integer, nil]
            # @param link [String, nil]
            # @param coordinate [Coordinate, nil]
            #: (untyped value, ::String? formula, Integer? style, ::String? link, Coordinate? coordinate) -> void
            def initialize(value, formula, style, link, coordinate)
              super(value, formula, nil, style, link, coordinate)
              @cell_type = :string
            end

            # @return [Symbol]
            #: () -> Symbol
            def default_type
              :string
            end

            # @return [Boolean]
            #: () -> bool
            def empty?
              @value.to_s.empty?
            end
          end

          # Empty cell representation.
          class Empty < Base
            # @param coordinate [Coordinate, nil]
            #: (?Coordinate? coordinate) -> void
            def initialize(coordinate = nil)
              super(nil, nil, nil, nil, nil, coordinate)
              @cell_type = nil
              @style = nil
            end

            # @return [nil]
            #: () -> nil
            def default_type
              nil
            end

            # @return [true]
            #: () -> true
            def empty?
              true
            end
          end

          # Factory method to create cell by type.
          #
          # @param type [Symbol]
          # @param values [Array<Object>]
          # @return [Base, nil]
          #: (Symbol type, *untyped values) -> Base?
          def self.create_cell(type, *values)
            klass = cell_class(type)
            klass&.new(*values)
          end

          # Returns cell class for type Symbol.
          #
          # @param type [Symbol]
          # @return [Class, nil]
          #: (Symbol type) -> singleton(Base)?
          def self.cell_class(type)
            case type
            when :string then Cell::String
            when :boolean then Cell::Boolean
            when :number, :float, :percentage then Cell::Number
            when :date then Cell::Date
            when :datetime then Cell::DateTime
            when :time then Cell::Time
            when :empty then Cell::Empty
            end
          end
        end
      end
    end
  end
end
