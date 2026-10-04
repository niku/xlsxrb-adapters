# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a cell or range reference in an Excel worksheet.
      class Reference
        ROW_MAX = 1024 * 1024
        COL_MAX = 16_384

        attr_reader :row_range #: Range[Integer]
        attr_reader :col_range #: Range[Integer]
        attr_reader :sheet_name #: String?
        attr_reader :row_from_absolute #: bool
        attr_reader :row_to_absolute #: bool
        attr_reader :col_from_absolute #: bool
        attr_reader :col_to_absolute #: bool

        # Creates a new Reference.
        # Can be initialized with:
        # - Reference.new(row, col)
        # - Reference.new(row_from, row_to, col_from, col_to)
        # - Reference.new("A1") or Reference.new("A1:B2") or Reference.new("'Sheet 1'!A1:B2")
        # - Reference.new(row_from: 0, row_to: 1, col_from: 0, col_to: 1)
        #
        # @param params [Array<untyped>]
        #: (*untyped params) -> void
        def initialize(*params)
          row_from = nil
          row_to = nil
          col_from = nil
          col_to = nil
          @row_from_absolute = false
          @row_to_absolute = false
          @col_from_absolute = false
          @col_to_absolute = false
          @sheet_name = nil

          case params.size
          when 4
            row_from, row_to, col_from, col_to = params
          when 2
            row_from, col_from = params
          when 1
            first = params.first
            case first
            when Hash
              h = first
              row_from = h[:row_from]
              row_to = h[:row_to]
              col_from = h[:col_from]
              col_to = h[:col_to]
            when String
              str = first
              match_data = str.match(/^(?:'(?<sheet_name1>[^']+)'|(?<sheet_name2>[^']+))!/)
              if match_data
                @sheet_name = match_data["sheet_name1"] || match_data["sheet_name2"]
                str = str[match_data[0].size..] || ""
              end

              from, to = str.split(":")
              row_from, col_from, @row_from_absolute, @col_from_absolute = self.class.ref2ind(from)
              row_to, col_to, @row_to_absolute, @col_to_absolute = self.class.ref2ind(to) if to
            else
              raise ArgumentError, "invalid value for #{self.class}: #{first.inspect}"
            end
          else
            raise ArgumentError, "wrong number of arguments (given #{params.size}, expected 1, 2, or 4)"
          end

          @row_range = Range.new((row_from || 0).to_i, (row_to || row_from || ROW_MAX).to_i)
          @col_range = Range.new((col_from || 0).to_i, (col_to || col_from || COL_MAX).to_i)
        end

        # @return [bool]
        #: () -> bool
        def single_cell?
          (@row_range.begin == @row_range.end) && (@col_range.begin == @col_range.end)
        end

        # @return [bool]
        #: () -> bool
        def valid?
          !(@row_range.begin&.negative? || @col_range.begin&.negative?)
        end

        # @return [Integer]
        #: () -> Integer
        def first_row
          @row_range.begin || 0
        end

        # @return [Integer]
        #: () -> Integer
        def last_row
          @row_range.end || 0
        end

        # @return [Integer]
        #: () -> Integer
        def first_col
          @col_range.begin || 0
        end

        # @return [Integer]
        #: () -> Integer
        def last_col
          @col_range.end || 0
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false if other.nil?
          return false unless other.is_a?(Reference)

          (@sheet_name == other.sheet_name) &&
            (@row_range == other.row_range) && (@col_range == other.col_range)
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def cover?(other)
          return false if other.nil?
          return false unless other.is_a?(Reference)

          other_r_begin = other.row_range.begin
          other_r_end = other.row_range.end
          other_c_begin = other.col_range.begin
          other_c_end = other.col_range.end

          return false if other_r_begin.nil? || other_r_end.nil? || other_c_begin.nil? || other_c_end.nil?

          @row_range.cover?(other_r_begin) &&
            @row_range.cover?(other_r_end) &&
            @col_range.cover?(other_c_begin) &&
            @col_range.cover?(other_c_end)
        end

        # @return [String]
        #: () -> String
        def to_s
          result = +""

          if @sheet_name
            result << if @sheet_name.include?(" ")
                        "'#{@sheet_name}'"
                      else
                        @sheet_name
                      end
            result << "!"
          end

          r_begin = @row_range.begin || 0
          c_begin = @col_range.begin || 0
          r_end = @row_range.end || 0
          c_end = @col_range.end || 0

          result << self.class.ind2ref(r_begin, c_begin, @row_from_absolute, @col_from_absolute)
          unless single_cell?
            result << ":"
            result << self.class.ind2ref(r_end, c_end, @row_to_absolute, @col_to_absolute)
          end

          result
        end

        # @return [String]
        #: () -> String
        def inspect
          if single_cell?
            "#<#{self.class} @sheet_name=#{@sheet_name.inspect} @row=#{@row_range.begin} @col=#{@col_range.begin}>"
          else
            "#<#{self.class} @sheet_name=#{@sheet_name.inspect} @row_range=#{@row_range} @col_range=#{@col_range}>"
          end
        end

        # Converts 0-indexed row and col to Excel reference (e.g. 0, 0 -> "A1")
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param row_abs [Boolean]
        # @param col_abs [Boolean]
        # @return [String]
        #: (?Integer row, ?Integer col, ?bool row_abs, ?bool col_abs) -> String
        def self.ind2ref(row = 0, col = 0, row_abs = false, col_abs = false)
          col_ref = Xlsxrb::Utils.col_index_to_name(col)
          "#{"$" if col_abs}#{col_ref}#{"$" if row_abs}#{row + 1}"
        end

        # Converts Excel reference to 0-indexed row and col (e.g. "A1" -> [0, 0, false, false])
        #
        # @param str [String]
        # @return [Array(Integer, Integer, bool, bool)]
        #: (String str) -> ([Integer, Integer] | [Integer, Integer, bool, bool])
        def self.ref2ind(str)
          matchdata = str.match(/\A(?<cabs>\$?)(?<col>[A-Za-z]+)(?<rabs>\$?)(?<row>\d+)\Z/)
          return [-1, -1] unless matchdata

          row_num = matchdata[:row].to_i - 1
          col_str = matchdata[:col].upcase
          col_num = Xlsxrb::Utils.col_name_to_index(col_str)
          rabs = !matchdata[:rabs].to_s.empty?
          cabs = !matchdata[:cabs].to_s.empty?

          [row_num, col_num, rabs, cabs]
        end
      end

      # Space-separated list of references
      # @rbs inherits Array[Reference]
      class Sqref < Array
        # @param str [String]
        #: (String str) -> void
        def initialize(str)
          super()
          str.split.each { |ref_str| self << Reference.new(ref_str) }
        end

        # @return [String]
        #: () -> String
        def to_s
          join(" ")
        end
      end
    end
  end
end
