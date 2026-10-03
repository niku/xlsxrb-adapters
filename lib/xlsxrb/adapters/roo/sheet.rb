# frozen_string_literal: true

# rbs_inline: enabled

require_relative "coordinate"
require_relative "cell"

module Xlsxrb
  module Adapters
    module Roo
      class Excelx
        # In-memory sheet representation matching Roo::Excelx::Sheet.
        class Sheet
          attr_reader :name, :cells, :hyperlinks, :comments, :images, :dimensions, :styles, :formats, :state

          # @param name [String]
          # @param cells [Hash[Array[Integer], Cell::Base]]
          # @param hyperlinks [Hash[Array[Integer], String]]
          # @param comments [Hash[Array[Integer], String]]
          # @param images [Array[String]]
          # @param dimensions [String, nil]
          # @param formats [Hash[Array[Integer], String]]
          # @param styles [Hash[untyped, untyped]]
          # @param state [Symbol, String]
          #: (String name, ?cells: Hash[untyped, untyped], ?hyperlinks: Hash[untyped, untyped], ?comments: Hash[untyped, untyped], ?images: Array[String], ?dimensions: String?, ?formats: Hash[untyped, untyped], ?styles: Hash[untyped, untyped], ?state: (Symbol | String)) -> void
          def initialize(name, cells: {}, hyperlinks: {}, comments: {}, images: [], dimensions: nil, formats: {}, styles: {}, state: :visible)
            @name = name
            @cells = cells
            @hyperlinks = hyperlinks
            @comments = comments
            @images = images
            @dimensions = dimensions
            @formats = formats
            @styles = styles
            @state = state ? state.to_sym : :visible
            @first_last = nil
          end

          # Returns whether sheet is hidden.
          #
          # @return [Boolean]
          #: () -> bool
          def hidden?
            @state == :hidden || @state == :very_hidden
          end

          # Returns whether sheet is visible.
          #
          # @return [Boolean]
          #: () -> bool
          def visible?
            !hidden?
          end

          # Returns cell object at coordinate [row, col] or nil.
          #
          # @param coord [Array<Integer>]
          # @return [Cell::Base, nil]
          #: (Array[Integer] coord) -> Cell::Base?
          def cell_at(coord)
            @cells[coord]
          end

          # Returns row cell values as Array.
          #
          # @param row_number [Integer] 1-based row index.
          # @return [Array<Object>]
          #: (Integer row_number) -> Array[untyped]
          def row(row_number)
            f_col = first_column
            l_col = last_column
            return [] unless f_col && l_col

            f_col.upto(l_col).map do |c|
              @cells[[row_number, c]]&.value
            end
          end

          # Returns column cell values as Array.
          #
          # @param col_number [Integer] 1-based column index.
          # @return [Array<Object>]
          #: (Integer col_number) -> Array[untyped]
          def column(col_number)
            f_row = first_row
            l_row = last_row
            return [] unless f_row && l_row

            f_row.upto(l_row).map do |r|
              @cells[[r, col_number]]&.value
            end
          end

          # Returns first non-empty 1-based row index or nil.
          #
          # @return [Integer, nil]
          #: () -> Integer?
          def first_row
            first_last[:first_row]
          end

          # Returns last non-empty 1-based row index or nil.
          #
          # @return [Integer, nil]
          #: () -> Integer?
          def last_row
            first_last[:last_row]
          end

          # Returns first non-empty 1-based column index or nil.
          #
          # @return [Integer, nil]
          #: () -> Integer?
          def first_column
            first_last[:first_column]
          end

          # Returns last non-empty 1-based column index or nil.
          #
          # @return [Integer, nil]
          #: () -> Integer?
          def last_column
            first_last[:last_column]
          end

          # Returns format code for cell coordinate.
          #
          # @param key [Array(Integer, Integer)]
          # @return [String, nil]
          #: (untyped key) -> String?
          def excelx_format(key)
            @formats[key] || @cells[key]&.format
          end

          # Iterates over rows yielding Array of Cell objects.
          #
          # @param options [Hash]
          # @yield [row]
          # @yieldparam row [Array<Cell::Base, nil>]
          # @return [Enumerator, void]
          #: (?Hash[Symbol, untyped] options) { (Array[untyped]) -> void } -> void
          #: (?Hash[Symbol, untyped] options) -> Enumerator[Array[untyped], void]
          def each_row(options = {}, &block)
            return to_enum(:each_row, options) unless block

            offset = options[:offset] || 0
            max_rows = options[:max_rows]
            pad_cells = options[:pad_cells] || false

            r_hash = rows_by_index
            return if r_hash.empty?

            min_r = r_hash.keys.min || 1
            max_r = r_hash.keys.max || 1

            row_count = 0
            (min_r..max_r).each do |r_idx|
              break if max_rows && row_count == (max_rows + offset + 1)

              if row_count >= offset
                row_cells = row_elements(r_idx, pad_cells)
                block.call(row_cells)
              end
              row_count += 1
            end
          end

          # Computes first/last row and column coordinates.
          #: () -> Hash[Symbol, Integer?]
          def first_last
            @first_last ||= begin
              f_row = nil
              l_row = nil
              f_col = nil
              l_col = nil

              @cells.each do |(r, c), cell|
                next if cell.nil? || cell.empty? || cell.value.nil?

                f_row ||= r
                l_row ||= r
                f_col ||= c
                l_col ||= c

                l_row = r if r > l_row
                f_row = r if r < f_row
                l_col = c if c > l_col
                f_col = c if c < f_col
              end

              { first_row: f_row, last_row: l_row, first_column: f_col, last_column: l_col }
            end
          end

          # Resets cached first/last coordinates when cells change.
          #: () -> void
          def reset_bounds!
            @first_last = nil
            @rows_by_index = nil
          end

          private

          #: () -> Hash[Integer, Array[[Integer, untyped]]]
          def rows_by_index
            @rows_by_index ||= begin
              hash = {}
              @cells.each do |k, cell|
                r = k[0]
                c = k[1]
                (hash[r] ||= []) << [c, cell]
              end
              hash.each_value { |c_list| c_list.sort_by!(&:first) }
              hash
            end
          end

          #: (Integer r_idx, bool pad_cells) -> Array[untyped]
          def row_elements(r_idx, pad_cells)
            sorted_cells = rows_by_index[r_idx]
            if sorted_cells.nil? || sorted_cells.empty?
              return pad_cells ? [nil] : []
            end

            return sorted_cells.map(&:last) unless pad_cells

            cells_arr = []
            cur_col = 1
            sorted_cells.each do |col, cell|
              while cur_col < col
                cells_arr << nil
                cur_col += 1
              end
              cells_arr << cell
              cur_col = col + 1
            end
            cells_arr
          end
        end
      end
    end
  end
end
