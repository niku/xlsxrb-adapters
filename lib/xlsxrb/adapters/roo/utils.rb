# frozen_string_literal: true

# rbs_inline: enabled

require "uri"
require_relative "coordinate"

module Xlsxrb
  module Adapters
    module Roo
      # Coordinate conversion and range utility module matching Roo::Utils.
      module Utils
        module_function

        LETTERS = ("A".."Z").to_a.freeze #: Array[String]
        URI_PARSER = defined?(::URI::RFC2396_PARSER) ? ::URI::RFC2396_PARSER : ::URI::DEFAULT_PARSER

        # Extracts coordinate as Excelx::Coordinate from a reference string like "A1".
        #
        # @param str [String, Symbol]
        # @return [Excelx::Coordinate]
        #: (String | Symbol str) -> Excelx::Coordinate
        def extract_coordinate(str)
          s = str.to_s
          num = 0
          letter_num = 0
          num_only = false

          s.each_byte do |b|
            if !num_only && (b.between?(65, 90) || b.between?(97, 122))
              idx = b > 96 ? b - 96 : b - 64
              letter_num = (letter_num * 26) + idx
            elsif b.between?(48, 57)
              num_only = true
              num = (num * 10) + (b - 48)
            else
              raise ArgumentError, "invalid coordinate string: #{s.inspect}"
            end
          end
          raise ArgumentError, "invalid coordinate string: #{s.inspect}" if letter_num.zero? || !num_only

          Excelx::Coordinate.new(num, letter_num)
        end

        # @param str [String]
        # @return [Excelx::Coordinate]
        #: (String | Symbol str) -> Excelx::Coordinate
        def split_coordinate(str)
          extract_coordinate(str)
        end

        # @param str [String]
        # @return [Excelx::Coordinate]
        #: (String | Symbol str) -> Excelx::Coordinate
        def ref_to_key(str)
          extract_coordinate(str)
        end

        # Splits a coordinate into [column_letter, row_number].
        #
        # @param str [String]
        # @return [Array(String, Integer)]
        #: (String | Symbol str) -> [String, Integer]
        def split_coord(str)
          coord = extract_coordinate(str)
          [number_to_letter(coord.column), coord.row]
        end

        # Converts a 1-based column number to letter representation (e.g. 1 => "A", 27 => "AA").
        #
        # @param num [Integer, nil]
        # @return [String]
        #: (Integer? num) -> String
        def number_to_letter(num)
          return "" if num.nil? || num <= 0

          Xlsxrb::Utils.col_index_to_name(num - 1)
        end

        # Converts column letter(s) to 1-based number (e.g. "A" => 1, "AA" => 27).
        #
        # @param letters [String, Symbol, nil]
        # @return [Integer]
        #: (String | Symbol letters) -> Integer
        def letter_to_number(letters)
          str = letters.to_s
          return 0 if str.empty?

          Xlsxrb::Utils.col_name_to_index(str) + 1
        end

        # Computes total cell count in a range like "A1:C10".
        #
        # @param str [String, nil]
        # @return [Integer]
        #: (String? str) -> Integer
        def num_cells_in_range(str)
          return 0 if str.nil? || str.empty?

          cells = str.split(":")
          return 1 if cells.one?
          raise ArgumentError, "invalid range string: #{str}. Supported range format 'A1:B2'" if cells.count != 2

          c1 = extract_coordinate(cells[0])
          c2 = extract_coordinate(cells[1])
          (c2.row - (c1.row - 1)) * (c2.column - (c1.column - 1))
        end

        # Yields or returns enumerator of coordinates in range.
        #
        # @param str [String]
        # @yield [coord]
        # @yieldparam coord [Excelx::Coordinate]
        # @return [Enumerator, void]
        #: (String str) { (Excelx::Coordinate) -> void } -> void
        #: (String str) -> Enumerator[Excelx::Coordinate, void]
        def coordinates_in_range(str, &block)
          return to_enum(:coordinates_in_range, str) unless block

          cells = str.to_s.split(":", 2).map { |s| extract_coordinate(s) }
          case cells.size
          when 1
            block.call(cells[0])
          when 2
            tl = cells[0]
            br = cells[1]
            rows = tl.row..br.row
            cols = tl.column..br.column
            rows.each do |row|
              cols.each do |col|
                block.call(Excelx::Coordinate.new(row, col))
              end
            end
          end
        end
      end
    end
  end
end
