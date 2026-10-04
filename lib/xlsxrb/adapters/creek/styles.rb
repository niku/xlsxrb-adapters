# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "bigdecimal"

module Xlsxrb
  module Adapters
    module Creek
      # Represents spreadsheet styles and cell formatting types matching Creek::Styles.
      class Styles
        # Map of non-custom numFmtId to casting symbol.
        module Constants
          NumFmtMap = {
            0 => :string, # General
            1 => :fixnum, # 0
            2 => :float, # 0.00
            3 => :fixnum, # #,##0
            4 => :float, # #,##0.00
            5 => :unsupported, # $#,##0_);($#,##0)
            6 => :unsupported, # $#,##0_);[Red]($#,##0)
            7 => :unsupported, # $#,##0.00_);($#,##0.00)
            8 => :unsupported, # $#,##0.00_);[Red]($#,##0.00)
            9 => :percentage, # 0%
            10 => :percentage, # 0.00%
            11 => :bignum, # 0.00E+00
            12 => :unsupported, # # ?/?
            13 => :unsupported, # # ??/??
            14 => :date, # mm-dd-yy
            15 => :date, # d-mmm-yy
            16 => :date, # d-mmm
            17 => :date, # mmm-yy
            18 => :time, # h:mm AM/PM
            19 => :time, # h:mm:ss AM/PM
            20 => :time, # h:mm
            21 => :time, # h:mm:ss
            22 => :date_time, # m/d/yy h:mm
            37 => :unsupported, # #,##0 ;(#,##0)
            38 => :unsupported, # #,##0 ;[Red](#,##0)
            39 => :unsupported, # #,##0.00;(#,##0.00)
            40 => :unsupported, # #,##0.00;[Red](#,##0.00)
            45 => :time, # mm:ss
            46 => :time, # [h]:mm:ss
            47 => :time, # mmss.0
            48 => :bignum, # ##0.0E+0
            49 => :unsupported # @
          }.freeze
        end

        # Back-calculates ruby type symbols from cell style formatting definitions.
        class StyleTypes
          include Constants

          # @return [untyped]
          attr_accessor :styles_xml_doc

          # @param styles_xml_doc [untyped] Nokogiri XML document, XML string, or IO
          #: (untyped styles_xml_doc) -> void
          def initialize(styles_xml_doc)
            @styles_xml_doc = styles_xml_doc
          end

          # Returns an Array of style type symbols corresponding to the spreadsheet cellXfs list.
          #
          # @return [Array<Symbol?>]
          #: () -> Array[Symbol?]
          def call
            @style_types ||= if @styles_xml_doc.respond_to?(:css)
                               @styles_xml_doc.css("styleSheet cellXfs xf").map do |xstyle|
                                 a = num_fmt_id(xstyle)
                                 style_type_by_num_fmt_id(a)
                               end
                             else
                               raw_xml = if @styles_xml_doc.respond_to?(:read)
                                           @styles_xml_doc.rewind if @styles_xml_doc.respond_to?(:rewind)
                                           @styles_xml_doc.read
                                         else
                                           @styles_xml_doc.to_s
                                         end
                               extract_style_types_from_xml(raw_xml)
                             end
          end

          # Returns the numFmtId attribute value from an xstyle node.
          #
          # @param xstyle [untyped]
          # @return [String, nil]
          #: (untyped xstyle) -> String?
          def num_fmt_id(xstyle)
            return nil unless xstyle.attributes["numFmtId"]

            xstyle.attributes["numFmtId"].value
          end

          # Finds the type symbol for a given number format ID.
          #
          # @param id [Integer, String, nil]
          # @return [Symbol, nil]
          #: (Integer | String | nil id) -> Symbol?
          def style_type_by_num_fmt_id(id)
            return nil unless id

            int_id = id.to_i
            NumFmtMap[int_id] || custom_style_types[int_id]
          end

          # Returns a mapping of custom numFmtId (>= 164) to inferred type symbols.
          #
          # @return [Hash<Integer, Symbol>]
          #: () -> Hash[Integer, Symbol]
          def custom_style_types
            @custom_style_types ||= if @styles_xml_doc.respond_to?(:css)
                                      @styles_xml_doc.css("styleSheet numFmts numFmt").each_with_object({}) do |xstyle, acc|
                                        index = xstyle.attributes["numFmtId"].value.to_i
                                        value = xstyle.attributes["formatCode"].value
                                        acc[index] = determine_custom_style_type(value)
                                      end
                                    else
                                      raw_xml = if @styles_xml_doc.respond_to?(:read)
                                                  @styles_xml_doc.rewind if @styles_xml_doc.respond_to?(:rewind)
                                                  @styles_xml_doc.read
                                                else
                                                  @styles_xml_doc.to_s
                                                end
                                      extract_custom_style_types_from_xml(raw_xml)
                                    end
          end

          # Inferred custom style type based on formatting pattern.
          #
          # @param string [String]
          # @return [Symbol]
          #: (String string) -> Symbol
          def determine_custom_style_type(string)
            return :float if string[0] == "_"
            return :float if string[0] == " 0"
            return :date_time if string =~ /(^|\])[^\[]*[ymdhis]/i

            :unsupported
          end

          private

          #: (String raw_xml) -> Array[Symbol?]
          def extract_style_types_from_xml(raw_xml)
            cell_xfs_part = raw_xml[%r{<cellXfs\b[^>]*>(.*?)</cellXfs>}m, 1]
            return [] unless cell_xfs_part

            types = []
            cell_xfs_part.scan(%r{<xf\b([^>]*?)(?:/>|>.*?</xf>)}m) do |attrs_str,|
              num_fmt_match = attrs_str[/numFmtId="(\d+)"/, 1]
              types << style_type_by_num_fmt_id(num_fmt_match)
            end
            types
          end

          #: (String raw_xml) -> Hash[Integer, Symbol]
          def extract_custom_style_types_from_xml(raw_xml)
            num_fmts_part = raw_xml[%r{<numFmts\b[^>]*>(.*?)</numFmts>}m, 1]
            return {} unless num_fmts_part

            res = {}
            num_fmts_part.scan(%r{<numFmt\b([^>]*?)(?:/>|>.*?</numFmt>)}m) do |attrs_str,|
              idx = attrs_str[/numFmtId="(\d+)"/, 1]
              val = attrs_str[/formatCode="([^"]*)"/, 1]
              res[idx.to_i] = determine_custom_style_type(val || "") if idx
            end
            res
          end
        end

        # Heart of typecasting in Creek: converts cell values based on type or style symbol.
        class Converter
          include Constants

          # Excel non-printable character escape sequence.
          HEX_ESCAPE_REGEXP = /_x[0-9A-Fa-f]{4}_/

          DATE_TYPES = %i[date time date_time].to_set

          # Casts a cell value based on explicit cell type and style formatting type.
          #
          # @param value [String, nil]
          # @param type [String, Symbol, nil]
          # @param style [Symbol, nil]
          # @param options [Hash]
          # @return [Object]
          #: (untyped value, untyped type, Symbol? style, ?Hash[Symbol, untyped] options) -> untyped
          def self.call(value, type, style, options = {})
            return nil if value.nil? || value.empty?

            # Sometimes the type is dictated by the style alone
            type = style if type.nil? || (type == "n" && DATE_TYPES.include?(style))

            case type
            when "s"
              strings = options[:shared_strings]
              strings ? strings[value.to_i] : value
            when "n", :float, :percentage
              value.to_f
            when "b"
              value.to_i == 1
            when "str", "inlineStr"
              unescape_string(value)
            when :string
              value
            when :fixnum
              value.to_i
            when :date
              convert_date(value, options)
            when :time, :date_time
              convert_datetime(value, options)
            when :bignum
              convert_bignum(value)
            else
              convert_unknown(value)
            end
          end

          # Best-effort type deduction for unformatted numeric or string values.
          #
          # @param value [untyped]
          # @return [untyped]
          #: (untyped value) -> untyped
          def self.convert_unknown(value)
            return value if value.nil? || (value.respond_to?(:empty?) && value.empty?)
            return value.to_i if value.to_i.to_s == value.to_s
            return value.to_f if value.to_f.to_s == value.to_s

            value
          rescue StandardError
            value
          end

          # Converts serial numeric date to Date object.
          #
          # @param value [Numeric, String]
          # @param options [Hash]
          # @return [Date]
          #: (untyped value, ?Hash[Symbol, untyped] options) -> Date
          def self.convert_date(value, options = {})
            date = base_date(options) + value.to_i
            yyyy, mm, dd = date.strftime("%Y-%m-%d").split("-")

            ::Date.new(yyyy.to_i, mm.to_i, dd.to_i)
          end

          # Converts serial numeric date/time to Time object.
          #
          # @param value [Numeric, String]
          # @param options [Hash]
          # @return [Time]
          #: (untyped value, ?Hash[Symbol, untyped] options) -> Time
          def self.convert_datetime(value, options = {})
            date = base_date(options) + value.to_f.round(6)

            round_datetime(date.strftime("%Y-%m-%d %H:%M:%S.%N"))
          end

          # Converts big numeric values to BigDecimal or Float.
          #
          # @param value [Numeric, String]
          # @return [BigDecimal, Float]
          #: (untyped value) -> (BigDecimal | Float)
          def self.convert_bignum(value)
            BigDecimal(value.to_s)
          rescue StandardError
            value.to_f
          end

          # Unescapes Excel hex escape sequences (e.g. _x000D_ -> \r).
          #
          # @param value [String]
          # @return [String]
          #: (String value) -> String
          def self.unescape_string(value)
            value.gsub(HEX_ESCAPE_REGEXP) { |match| match[2, 4].to_i(16).chr(Encoding::UTF_8) }
          end

          # Returns the base date for date serial arithmetic (1899-12-30 or 1904-01-01).
          #
          # @param options [Hash]
          # @return [Date]
          #: (?Hash[Symbol, untyped] options) -> Date
          def self.base_date(options = {})
            options[:base_date] || Date.new(1899, 12, 30)
          end

          # Rounds a datetime string to nearest whole second.
          #
          # @param datetime_string [String]
          # @return [Time]
          #: (String datetime_string) -> Time
          def self.round_datetime(datetime_string)
            /(?<yyyy>\d+)-(?<mm>\d+)-(?<dd>\d+) (?<hh>\d+):(?<mi>\d+):(?<ss>\d+\.\d+)/ =~ datetime_string

            ::Time.new(yyyy.to_i, mm.to_i, dd.to_i, hh.to_i, mi.to_i, ss.to_r).round(0)
          end
        end

        # @return [untyped]
        attr_accessor :book

        # @param book [untyped]
        #: (untyped book) -> void
        def initialize(book)
          @book = book
        end

        # Path to styles entry in workbook.
        #
        # @return [String]
        #: () -> String
        def path
          "xl/styles.xml"
        end

        # Returns the parsed styles XML document or string.
        #
        # @return [untyped]
        #: () -> untyped
        def styles_xml
          @styles_xml ||= if @book.files.file.exist?(path)
                            doc = @book.files.file.open(path)
                            if defined?(::Nokogiri::XML::Document)
                              ::Nokogiri::XML::Document.parse(doc)
                            else
                              doc.read
                            end
                          end
        end

        # Returns the array of style types for the workbook.
        #
        # @return [Array<Symbol?>]
        #: () -> Array[Symbol?]
        def style_types
          @style_types ||= StyleTypes.new(styles_xml).call
        end
      end
    end
  end
end
