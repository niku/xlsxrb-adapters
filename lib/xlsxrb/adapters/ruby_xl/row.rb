# frozen_string_literal: true

# rbs_inline: enabled

require_relative "cell"

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a mutable row wrapper compatible with RubyXL::Row.
      class Row
        [Enumerable].each { |m| include m }

        DEFAULT_HEIGHT = 13

        attr_accessor :worksheet #: Worksheet?
        attr_accessor :index #: Integer
        attr_accessor :cells #: Array[Cell?]
        attr_accessor :height #: (Float | Integer)?
        attr_accessor :hidden #: bool
        attr_accessor :custom_height #: bool
        attr_accessor :outline_level #: Integer?
        attr_accessor :style_index #: Integer?

        alias ht height
        alias ht= height=

        # @param worksheet [Worksheet, nil] Parent worksheet.
        # @param index [Integer] 0-based row index.
        # @param cells [Array<Cell, nil>] Initial list of cells.
        # @param height [Float, Integer, nil] Row height.
        # @param hidden [Boolean] Whether row is hidden.
        # @param custom_height [Boolean] Whether row has custom height.
        # @param outline_level [Integer, nil] Outline level.
        # @param style_index [Integer, nil] Style index.
        #: (?worksheet: Worksheet?, ?index: Integer, ?cells: Array[Cell?], ?height: (Float | Integer)?, ?hidden: bool, ?custom_height: bool, ?outline_level: Integer?, ?style_index: Integer?) -> void
        def initialize(worksheet: nil, index: 0, cells: [], height: nil, hidden: false, custom_height: false, outline_level: nil, style_index: 0)
          @worksheet = worksheet
          @index = index
          @cells = cells.is_a?(Array) ? cells.dup : []
          @height = height
          @hidden = hidden
          @custom_height = custom_height
          @outline_level = outline_level
          @style_index = style_index.to_i
        end

        # Accesses a cell by 0-based column index.
        #
        # @param col_idx [Integer]
        # @return [Cell, nil]
        #: (Integer col_idx) -> Cell?
        def [](col_idx)
          @cells[col_idx]
        end

        # Sets a cell at 0-based column index.
        #
        # @param col_idx [Integer]
        # @param cell [Cell, nil]
        # @return [Cell, nil]
        #: (Integer col_idx, Cell? cell) -> Cell?
        def []=(col_idx, cell)
          @cells[col_idx] = cell
        end

        # Iterates over each cell in the row (including nil for sparse entries).
        #
        # @yield [cell]
        # @yieldparam cell [Cell, nil]
        # @return [Enumerator, void]
        #: () { (Cell?) -> void } -> void
        #: () -> Enumerator[Cell?, void]
        def each(&)
          return to_enum(:each) unless block_given?

          @cells.each(&)
        end

        # Returns the number of cells up to the highest column index.
        #
        # @return [Integer]
        #: () -> Integer
        def size
          @cells.size
        end
        alias length size

        # Inserts a cell and shifts following cells to the right.
        #
        # @param cell [Cell, nil]
        # @param col_index [Integer]
        # @return [void]
        #: (Cell? cell, Integer col_index) -> void
        def insert_cell_shift_right(cell, col_index)
          # Pad cells if necessary
          @cells << nil while @cells.size < col_index
          @cells.insert(col_index, cell)
          col_index.upto(@cells.size - 1) do |ci|
            c = @cells[ci]
            c.column = ci if c.is_a?(Cell)
          end
        end

        # Deletes a cell and shifts following cells to the left.
        #
        # @param col_index [Integer]
        # @return [Cell, nil]
        #: (Integer col_index) -> Cell?
        def delete_cell_shift_left(col_index)
          deleted = @cells.delete_at(col_index)
          col_index.upto(@cells.size - 1) do |ci|
            c = @cells[ci]
            c.column = ci if c.is_a?(Cell)
          end
          deleted
        end

        # @return [XF, nil]
        #: () -> XF?
        def get_row_xf
          return nil unless @worksheet&.workbook

          wb = @worksheet.workbook
          wb.cell_xfs[@style_index || 0] || wb.cell_xfs[0]
        end

        # @return [Font, nil]
        #: () -> Font?
        def get_font
          xf = get_row_xf
          return nil unless xf && @worksheet&.workbook

          wb = @worksheet.workbook
          wb.fonts[xf.font_id || 0] || wb.fonts[0]
        end

        # @return [String]
        #: () -> String
        def get_fill_color
          xf = get_row_xf
          if xf && @worksheet&.workbook
            @worksheet.workbook.get_fill_color(xf)
          else
            "ffffff"
          end
        end

        # Converts row to an immutable Xlsxrb::Elements::Row.
        #
        # @return [Xlsxrb::Elements::Row]
        #: () -> Xlsxrb::Elements::Row
        def to_xlsxrb
          sorted_cells = @cells.compact.sort_by(&:column).map(&:to_xlsxrb)
          row_unmapped = {}
          row_unmapped[:style_index] = @style_index if @style_index&.positive?
          Xlsxrb::Elements::Row.new(
            index: @index,
            cells: sorted_cells,
            height: @height,
            hidden: @hidden,
            custom_height: @custom_height,
            outline_level: @outline_level,
            unmapped_data: row_unmapped
          )
        end

        # Inspect representation.
        #
        # @return [String]
        #: () -> String
        def inspect
          "#<#{self.class.name} index=#{@index} cells_count=#{@cells.compact.size}>"
        end
      end
    end
  end
end
