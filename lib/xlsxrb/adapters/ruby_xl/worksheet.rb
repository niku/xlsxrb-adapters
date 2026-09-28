# frozen_string_literal: true

# rbs_inline: enabled

require_relative "row"
require_relative "cell"
require_relative "color"
require_relative "reference"
require_relative "column_range"
require_relative "merged_cells"

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a mutable worksheet wrapper compatible with RubyXL::Worksheet.
      class Worksheet
        [Enumerable].each { |m| include m }

        NAME = 0
        SIZE = 1
        COLOR = 2
        ITALICS = 3
        BOLD = 4
        UNDERLINE = 5
        STRIKETHROUGH = 6

        attr_accessor :workbook #: Workbook?
        attr_accessor :sheet_name #: String
        attr_accessor :rows #: Array[Row?]
        attr_accessor :columns #: Array[untyped]
        attr_accessor :charts #: Array[untyped]
        attr_accessor :unmapped_data #: Hash[untyped, untyped]
        attr_accessor :cols #: ColumnRanges
        attr_accessor :merged_cells #: MergedCells?
        attr_accessor :data_validations #: Array[untyped]?
        attr_accessor :sheet_id #: Integer?

        # Proxy to support RubyXL's `worksheet.sheet_data[r][c]` and `sheet_data.rows` idiom.
        class SheetDataProxy
          [Enumerable].each { |m| include m }

          # @param worksheet [Worksheet]
          #: (Worksheet worksheet) -> void
          def initialize(worksheet)
            @worksheet = worksheet
          end

          #: (Integer row_index) -> Row?
          def [](row_index)
            @worksheet[row_index]
          end

          #: () -> Array[Row?]
          def rows
            @worksheet.rows
          end

          #: () -> Integer
          def size
            @worksheet.rows.size
          end

          #: () { (Row?) -> void } -> void
          def each(&)
            @worksheet.rows.each(&)
          end
        end

        # @param workbook [Workbook, nil] Parent workbook.
        # @param sheet_name [String] Worksheet name.
        # @param rows [Array<Row, nil>] Initial list of rows.
        # @param columns [Array] Column metadata.
        # @param charts [Array] Charts metadata.
        # @param unmapped_data [Hash] Additional metadata.
        #: (?workbook: Workbook?, ?sheet_name: String, ?rows: Array[Row?], ?columns: Array[untyped], ?charts: Array[untyped], ?unmapped_data: Hash[untyped, untyped]) -> void
        def initialize(workbook: nil, sheet_name: "Sheet1", rows: [], columns: [], charts: [], unmapped_data: {})
          @workbook = workbook
          @sheet_name = sheet_name.to_s
          @rows = rows.is_a?(Array) ? rows.dup : []
          @columns = columns.dup
          @charts = charts.dup
          @unmapped_data = unmapped_data.dup
          @cols = ColumnRanges.new
          @merged_cells = nil
          @data_validations = nil
        end

        alias name sheet_name
        alias name= sheet_name=

        # Accesses a row by 0-based row index.
        #
        # @param row_idx [Integer]
        # @return [Row, nil]
        #: (Integer row_idx) -> Row?
        def [](row_idx)
          @rows[row_idx]
        end

        # Sets a row at 0-based row index.
        #
        # @param row_idx [Integer]
        # @param row [Row, nil]
        # @return [Row, nil]
        #: (Integer row_idx, Row? row) -> Row?
        def []=(row_idx, row)
          @rows[row_idx] = row
        end

        # Accesses a cell by Excel reference string (e.g. "A1").
        #
        # @param ref [String] Excel reference string.
        # @return [Cell, nil]
        #: (String ref) -> Cell?
        def cell_at(ref)
          reference = Reference.new(ref)
          raise "Invalid reference: #{ref}" unless reference.valid? && reference.single_cell?

          self[reference.first_row]&.[](reference.first_col)
        end

        # Adds a new row at the given 0-based row index.
        #
        # @param row_index [Integer]
        # @param params [Hash]
        # @return [Row]
        #: (?Integer row_index, **untyped params) -> Row
        def add_row(row_index = 0, **params)
          new_row = Row.new(worksheet: self, index: row_index, **params)
          @rows[row_index] = new_row
          new_row
        end

        # Adds a cell to the worksheet at specified (row, col) coordinates.
        # Compatible with RubyXL::Worksheet#add_cell.
        #
        # @param row_index [Integer] 0-based row index.
        # @param column_index [Integer] 0-based column index.
        # @param data [Object] Cell value.
        # @param formula [String, Formula, nil] Optional formula.
        # @param overwrite [Boolean] Whether to overwrite existing cell if present.
        # @return [Cell] The created or updated Cell.
        #: (?Integer row_index, ?Integer column_index, ?untyped data, ?(String | Formula)? formula, ?bool overwrite) -> Cell
        def add_cell(row_index = 0, column_index = 0, data = "", formula = nil, overwrite = true)
          validate_nonnegative(row_index)
          validate_nonnegative(column_index)

          row = @rows[row_index] ||= Row.new(worksheet: self, index: row_index)
          existing_cell = row[column_index]

          return existing_cell if !overwrite && existing_cell

          cell = existing_cell || Cell.new(worksheet: self, row: row_index, column: column_index)
          cell.worksheet = self
          cell.row = row_index
          cell.column = column_index

          if formula
            cell.formula = formula
            cell.change_contents(data, formula)
          else
            cell.change_contents(data)
          end

          row[column_index] = cell
          cell
        end

        # Iterates through rows in the worksheet.
        #
        # @yield [row]
        # @yieldparam row [Row, nil]
        # @return [Enumerator, void]
        #: () { (Row?) -> void } -> void
        #: () -> Enumerator[Row?, void]
        def each(&)
          return to_enum(:each) unless block_given?

          @rows.each(&)
        end

        # SheetData compatibility proxy.
        #
        # @return [SheetDataProxy]
        #: () -> SheetDataProxy
        def sheet_data
          @sheet_data ||= SheetDataProxy.new(self)
        end

        # --- Validation & Structure Helpers ---

        # @return [void]
        #: () -> void
        def validate_workbook
          raise "Workbook is not associated with this worksheet" if @workbook.nil?
        end

        # @param index [Integer]
        # @return [void]
        #: (Integer index) -> void
        def validate_nonnegative(index)
          raise "Row and Column arguments must be nonnegative" if index.negative?
        end

        # Ensures that a row (and optionally column) exists in the table.
        #
        # @param row_index [Integer]
        # @param column_index [Integer]
        # @return [void]
        #: (Integer row_index, ?Integer column_index) -> void
        def ensure_cell_exists(row_index, column_index = 0)
          validate_nonnegative(row_index)
          validate_nonnegative(column_index)

          @rows << nil while @rows.size <= row_index

          row = @rows[row_index] ||= Row.new(worksheet: self, index: row_index)
          row.cells << nil while row.cells.size <= column_index
        end

        # --- Structural Operations: Rows, Columns, Cells ---

        # Inserts a row at row_index, pushing down subsequent rows and copying styles from the row above.
        #
        # @param row_index [Integer]
        # @return [Row]
        #: (?Integer row_index) -> Row
        def insert_row(row_index = 0)
          validate_workbook
          validate_nonnegative(row_index)
          ensure_cell_exists(row_index)

          old_row = nil
          new_cells = nil

          if row_index.positive?
            old_row = @rows[row_index - 1]
            if old_row
              new_cells = old_row.cells.map do |c|
                if c.nil?
                  nil
                else
                  Cell.new(worksheet: self, row: row_index, column: c.column, style_index: c.style_index)
                end
              end
            end
          end

          row0 = @rows[0]
          new_cells ||= Array.new(row0 ? row0.cells.size : 0)

          @rows.insert(row_index, nil)
          new_row = Row.new(worksheet: self, index: row_index, cells: new_cells, style_index: old_row&.style_index || 0)
          @rows[row_index] = new_row

          # Update row values for all rows at and below the insertion point
          row_index.upto(@rows.size - 1) do |r|
            row = @rows[r]
            next if row.nil?

            row.index = r
            row.cells.each_with_index do |cell, _c|
              next if cell.nil?

              cell.row = r
            end
          end

          # Update merged cells
          @merged_cells&.each do |mc|
            next if mc.ref.last_row < row_index

            in_merged_cell = mc.ref.first_row < row_index
            new_first = mc.ref.first_row + (in_merged_cell ? 0 : 1)
            new_last = mc.ref.last_row + 1
            mc.ref = Reference.new(new_first, new_last, mc.ref.first_col, mc.ref.last_col)
          end

          new_row
        end

        # Deletes a row at row_index, shifting following rows up.
        #
        # @param row_index [Integer]
        # @return [Row, nil]
        #: (?Integer row_index) -> Row?
        def delete_row(row_index = 0)
          validate_workbook
          validate_nonnegative(row_index)

          deleted = @rows.delete_at(row_index)

          # Update row numbers of following cells
          row_index.upto(@rows.size - 1) do |idx|
            row = @rows[idx]
            next if row.nil?

            row.index = idx
            row.cells.each do |c|
              c.row = idx if c.is_a?(Cell)
            end
          end

          # Update merged cells
          if @merged_cells
            @merged_cells.delete_if { |mc| mc.ref.row_range == (row_index..row_index) }
            @merged_cells.each do |mc|
              next if mc.ref.last_row < row_index

              in_merged_cell = mc.ref.first_row <= row_index
              new_first = mc.ref.first_row - (in_merged_cell ? 0 : 1)
              new_last = mc.ref.last_row - 1
              mc.ref = Reference.new(new_first, new_last, mc.ref.first_col, mc.ref.last_col)
            end
            @merged_cells.delete_if { |mc| mc.ref.single_cell? }
          end

          deleted
        end

        # Inserts a column at column_index, shifting following columns to the right.
        #
        # @param column_index [Integer]
        # @return [void]
        #: (?Integer column_index) -> void
        def insert_column(column_index = 0)
          validate_workbook
          validate_nonnegative(column_index)
          ensure_cell_exists(0, column_index)

          old_range = @cols.locate_range(column_index)

          @rows.each_with_index do |row, row_idx|
            next if row.nil?

            old_cell = row[column_index]
            c = nil
            c = Cell.new(worksheet: self, row: row_idx, column: column_index, style_index: old_cell.style_index) if old_cell && old_cell.style_index != 0 && old_range && old_range.style_index != old_cell.style_index

            row.insert_cell_shift_right(c, column_index)
          end

          @cols.insert_column(column_index)

          # Update merged cells
          return unless @merged_cells

          @merged_cells.each do |mc|
            next if mc.ref.last_col < column_index

            in_merged_cell = mc.ref.first_col < column_index
            new_first_col = mc.ref.first_col + (in_merged_cell ? 0 : 1)
            new_last_col = mc.ref.last_col + 1
            mc.ref = Reference.new(mc.ref.first_row, mc.ref.last_row, new_first_col, new_last_col)
          end
        end

        # Deletes a column at column_index, shifting following columns to the left.
        #
        # @param column_index [Integer]
        # @return [void]
        #: (?Integer column_index) -> void
        def delete_column(column_index = 0)
          validate_workbook
          validate_nonnegative(column_index)

          @rows.each do |row|
            row&.delete_cell_shift_left(column_index)
          end

          @cols.each { |r| r.delete_column(column_index) }

          # Update merged cells
          return unless @merged_cells

          @merged_cells.delete_if { |mc| mc.ref.col_range == (column_index..column_index) }
          @merged_cells.each do |mc|
            next if mc.ref.last_col < column_index

            in_merged_cell = mc.ref.first_col <= column_index
            new_first = mc.ref.first_col - (in_merged_cell ? 0 : 1)
            new_last = mc.ref.last_col - 1
            mc.ref = Reference.new(mc.ref.first_row, mc.ref.last_row, new_first, new_last)
          end
          @merged_cells.delete_if { |mc| mc.ref.single_cell? }
        end

        # Inserts a cell with optional shifting (:right or :down).
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param data [Object]
        # @param formula [String, nil]
        # @param shift [Symbol, nil] :right or :down
        # @return [Cell]
        #: (?Integer row, ?Integer col, ?untyped data, ?(String | Formula)? formula, ?Symbol? shift) -> Cell
        def insert_cell(row = 0, col = 0, data = nil, formula = nil, shift = nil)
          validate_workbook
          ensure_cell_exists(row, col)

          case shift
          when nil
            # No shifting
          when :right
            @rows[row]&.insert_cell_shift_right(nil, col)
          when :down
            ensure_cell_exists(@rows.size, col)
            (@rows.size - 1).downto(row + 1) do |idx|
              old_row = @rows[idx - 1]
              if old_row.nil?
                @rows[idx] = nil
              else
                new_row = @rows[idx] ||= Row.new(worksheet: self, index: idx)
                c = old_row.cells[col]
                new_row.cells[col] = c
                c.row = idx if c.is_a?(Cell)
              end
            end
            @rows[row]&.[]=(col, nil)
          else
            raise "invalid shift option: #{shift.inspect}"
          end

          add_cell(row, col, data, formula)
        end

        # Deletes a cell with optional shifting (:left or :up).
        #
        # @param row_index [Integer]
        # @param column_index [Integer]
        # @param shift [Symbol, nil] :left or :up
        # @return [Cell, nil]
        #: (?Integer row_index, ?Integer column_index, ?Symbol? shift) -> Cell?
        def delete_cell(row_index = 0, column_index = 0, shift = nil)
          validate_workbook
          validate_nonnegative(row_index)
          validate_nonnegative(column_index)

          row = @rows[row_index]
          old_cell = row&.[](column_index)

          case shift
          when nil
            row&.[]=(column_index, nil)
          when :left
            row&.delete_cell_shift_left(column_index)
          when :up
            (row_index...(@rows.size - 1)).each do |idx|
              old_r = @rows[idx + 1]
              if old_r.nil?
                @rows[idx] = nil
              else
                target_r = @rows[idx] ||= Row.new(worksheet: self, index: idx)
                c = old_r.cells[column_index]
                target_r.cells[column_index] = c
                c.row = idx if c.is_a?(Cell)
              end
            end
            last_r = @rows.last
            last_r&.[]=(column_index, nil)
          else
            raise "invalid shift option: #{shift.inspect}"
          end

          old_cell
        end

        # Merges cells within a rectangular area.
        #
        # @param params [Array<untyped>]
        # @return [MergedCell]
        #: (*untyped params) -> MergedCell
        def merge_cells(*params)
          validate_workbook

          row_from = nil
          col_from = nil
          row_to = nil
          col_to = nil

          case params.size
          when 4
            row_from, col_from, row_to, col_to = params
          when 1
            first = params.first
            case first
            when Hash
              row_from, row_to, col_from, col_to = first.fetch_values(:row_from, :row_to, :col_from, :col_to)
            when String
              from, to = first.split(":")
              raise ArgumentError, "reference for merging cells must be a range" if to.nil?

              row_from, col_from = Reference.ref2ind(from)
              row_to, col_to = Reference.ref2ind(to)
            else
              raise ArgumentError, "invalid value for #{self.class}: #{first.inspect}"
            end
          else
            raise ArgumentError, "wrong number of arguments (given #{params.size}, expected 1 or 4)"
          end

          @merged_cells ||= MergedCells.new
          mc = MergedCell.new(ref: Reference.new(row_from, row_to, col_from, col_to))
          @merged_cells << mc
          mc
        end

        # --- Row Styling Methods ---

        # @param row_index [Integer]
        # @return [Integer]
        #: (Integer row_index) -> Integer
        def get_row_style(row_index)
          row = @rows[row_index]
          row&.style_index || 0
        end

        # @param row_index [Integer]
        # @return [XF]
        #: (Integer row_index) -> XF
        def get_row_xf(row_index)
          validate_workbook
          wb = @workbook
          raise "Workbook not associated" unless wb

          idx = get_row_style(row_index)
          wb.cell_xfs[idx] || wb.cell_xfs[0]
        end

        # @param row [Integer]
        # @return [String, nil]
        #: (?Integer row) -> String?
        def get_row_fill(row = 0)
          validate_nonnegative(row)
          @rows[row]&.get_fill_color
        end

        # @param row [Integer]
        # @param color_code [String]
        # @return [void]
        #: (Integer row, ?String color_code) -> void
        def change_row_fill(row = 0, color_code = "ffffff")
          validate_workbook
          Color.validate_color(color_code)
          ensure_cell_exists(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          r = @rows[row]
          return unless r

          r.style_index = wb.modify_fill(get_row_style(row), color_code)
          r.cells.each do |c|
            c&.change_fill(color_code)
          end
        end

        # @param row [Integer]
        # @return [Font, nil]
        #: (?Integer row) -> Font?
        def row_font(row = 0)
          validate_nonnegative(row)
          @rows[row]&.get_font
        end

        # @param row [Integer]
        # @return [String, nil]
        #: (?Integer row) -> String?
        def get_row_font_name(row = 0)
          row_font(row)&.get_name
        end

        # @param row [Integer]
        # @return [Numeric, nil]
        #: (?Integer row) -> (Float | Integer)?
        def get_row_font_size(row = 0)
          row_font(row)&.get_size
        end

        # @param row [Integer]
        # @return [String, nil]
        #: (?Integer row) -> String?
        def get_row_font_color(row = 0)
          validate_nonnegative(row)
          font = row_font(row)
          font&.get_rgb_color || (font ? "000000" : nil)
        end

        # @param row [Integer]
        # @return [bool?]
        #: (?Integer row) -> bool?
        def is_row_italicized(row = 0)
          row_font(row)&.is_italic
        end

        # @param row [Integer]
        # @return [bool?]
        #: (?Integer row) -> bool?
        def is_row_bolded(row = 0)
          row_font(row)&.is_bold
        end

        # @param row [Integer]
        # @return [bool?]
        #: (?Integer row) -> bool?
        def is_row_underlined(row = 0)
          row_font(row)&.is_underlined
        end

        # @param row [Integer]
        # @return [bool?]
        #: (?Integer row) -> bool?
        def is_row_struckthrough(row = 0)
          row_font(row)&.is_strikethrough
        end

        # @param row [Integer]
        # @return [Numeric]
        #: (?Integer row) -> (Float | Integer)
        def get_row_height(row = 0)
          validate_workbook
          validate_nonnegative(row)
          r = @rows[row]
          r&.ht || Row::DEFAULT_HEIGHT
        end

        # @param row [Integer]
        # @param height [Numeric]
        # @return [void]
        #: (Integer row, Numeric height) -> void
        def change_row_height(row = 0, height = 10)
          validate_workbook
          ensure_cell_exists(row)
          r = @rows[row]
          return unless r

          r.ht = height
          r.custom_height = true
        end

        # @param row [Integer]
        # @param direction [Symbol, String]
        # @return [String, nil]
        #: (Integer row, Symbol | String direction) -> String?
        def get_row_border(row, direction)
          validate_workbook
          wb = @workbook
          raise "Workbook not associated" unless wb

          xf = get_row_xf(row)
          border = wb.borders[xf.border_id || 0]
          border&.get_edge_style(direction)
        end

        # @param row [Integer]
        # @param direction [Symbol, String]
        # @return [String, nil]
        #: (Integer row, Symbol | String direction) -> String?
        def get_row_border_color(row, direction)
          validate_workbook
          wb = @workbook
          raise "Workbook not associated" unless wb

          xf = get_row_xf(row)
          border = wb.borders[xf.border_id || 0]
          border&.get_edge_color(direction)
        end

        # @param row [Integer]
        # @param is_horizontal [Boolean]
        # @return [String, nil]
        #: (Integer row, bool is_horizontal) -> String?
        def get_row_alignment(row, is_horizontal)
          validate_workbook
          xf = get_row_xf(row)
          return nil if xf.alignment.nil?

          if is_horizontal
            xf.alignment.horizontal
          else
            xf.alignment.vertical
          end
        end

        # @param row [Integer]
        # @yieldparam alignment [Alignment]
        # @return [void]
        #: (Integer row) { (Alignment) -> void } -> void
        def change_row_alignment(row, &)
          validate_workbook
          validate_nonnegative(row)
          ensure_cell_exists(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          r = @rows[row]
          return unless r

          r.style_index = wb.modify_alignment(get_row_style(row), &)
          r.cells.each do |c|
            next if c.nil?

            c.style_index = wb.modify_alignment(c.style_index, &)
          end
        end

        # @param row [Integer]
        # @param alignment [String]
        # @return [void]
        #: (?Integer row, ?String alignment) -> void
        def change_row_horizontal_alignment(row = 0, alignment = "center")
          change_row_alignment(row) { |a| a.horizontal = alignment }
        end

        # @param row [Integer]
        # @param alignment [String]
        # @return [void]
        #: (?Integer row, ?String alignment) -> void
        def change_row_vertical_alignment(row = 0, alignment = "center")
          change_row_alignment(row) { |a| a.vertical = alignment }
        end

        # @param row [Integer]
        # @param direction [Symbol, String]
        # @param weight [String]
        # @return [void]
        #: (Integer row, Symbol | String direction, String weight) -> void
        def change_row_border(row, direction, weight)
          validate_workbook
          ensure_cell_exists(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          r = @rows[row]
          return unless r

          r.style_index = wb.modify_border(get_row_style(row), direction, weight)
          r.cells.each do |c|
            c&.change_border(direction, weight)
          end
        end

        # @param row [Integer]
        # @param direction [Symbol, String]
        # @param color [String]
        # @return [void]
        #: (Integer row, Symbol | String direction, String color) -> void
        def change_row_border_color(row, direction, color)
          validate_workbook
          ensure_cell_exists(row)
          Color.validate_color(color)
          wb = @workbook
          raise "Workbook not associated" unless wb

          r = @rows[row]
          return unless r

          r.style_index = wb.modify_border_color(get_row_style(row), direction, color)
          r.cells.each do |c|
            c&.change_border_color(direction, color)
          end
        end

        # @param row [Integer]
        # @param change_type [Integer]
        # @param arg [untyped]
        # @param font [Font]
        # @param xf [XF]
        # @return [void]
        #: (Integer row, Integer change_type, untyped arg, Font font, XF xf) -> void
        def change_row_font(row, change_type, arg, font, xf)
          validate_workbook
          ensure_cell_exists(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          new_xf = wb.register_new_font(font, xf)
          r = @rows[row]
          return unless r

          r.style_index = wb.register_new_xf(new_xf)
          r.cells.each do |c|
            c&.font_switch(change_type, arg)
          end
        end

        # @param row [Integer]
        # @param font_name [String]
        # @return [void]
        #: (?Integer row, ?String font_name) -> void
        def change_row_font_name(row = 0, font_name = "Verdana")
          xf = get_row_xf(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_name(font_name)
          change_row_font(row, NAME, font_name, font, xf)
        end

        # @param row [Integer]
        # @param font_size [Numeric]
        # @return [void]
        #: (?Integer row, ?Numeric font_size) -> void
        def change_row_font_size(row = 0, font_size = 10)
          xf = get_row_xf(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_size(font_size)
          change_row_font(row, SIZE, font_size, font, xf)
        end

        # @param row [Integer]
        # @param font_color [String]
        # @return [void]
        #: (?Integer row, ?String font_color) -> void
        def change_row_font_color(row = 0, font_color = "000000")
          Color.validate_color(font_color)
          xf = get_row_xf(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_rgb_color(font_color)
          change_row_font(row, COLOR, font_color, font, xf)
        end

        # @param row [Integer]
        # @param italicized [Boolean]
        # @return [void]
        #: (?Integer row, ?bool italicized) -> void
        def change_row_italics(row = 0, italicized = false)
          xf = get_row_xf(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_italic(italicized)
          change_row_font(row, ITALICS, italicized, font, xf)
        end

        # @param row [Integer]
        # @param bolded [Boolean]
        # @return [void]
        #: (?Integer row, ?bool bolded) -> void
        def change_row_bold(row = 0, bolded = false)
          xf = get_row_xf(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_bold(bolded)
          change_row_font(row, BOLD, bolded, font, xf)
        end

        # @param row [Integer]
        # @param underlined [Boolean]
        # @return [void]
        #: (?Integer row, ?bool underlined) -> void
        def change_row_underline(row = 0, underlined = false)
          xf = get_row_xf(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_underline(underlined)
          change_row_font(row, UNDERLINE, underlined, font, xf)
        end

        # @param row [Integer]
        # @param struckthrough [Boolean]
        # @return [void]
        #: (?Integer row, ?bool struckthrough) -> void
        def change_row_strikethrough(row = 0, struckthrough = false)
          xf = get_row_xf(row)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_strikethrough(struckthrough)
          change_row_font(row, STRIKETHROUGH, struckthrough, font, xf)
        end

        # --- Column Styling Methods ---

        # @param column_index [Integer]
        # @return [Integer]
        #: (Integer column_index) -> Integer
        def get_cols_style_index(column_index)
          validate_nonnegative(column_index)
          range = @cols.locate_range(column_index)
          range&.style_index || 0
        end

        # @param col [Integer]
        # @return [XF]
        #: (Integer col) -> XF
        def get_col_xf(col)
          validate_workbook
          wb = @workbook
          raise "Workbook not associated" unless wb

          idx = get_cols_style_index(col)
          wb.cell_xfs[idx] || wb.cell_xfs[0]
        end

        # @param col [Integer]
        # @return [Font]
        #: (Integer col) -> Font
        def column_font(col)
          validate_workbook
          wb = @workbook
          raise "Workbook not associated" unless wb

          xf = get_col_xf(col)
          wb.fonts[xf.font_id || 0] || wb.fonts[0]
        end

        # @param col [Integer]
        # @return [String]
        #: (?Integer col) -> String
        def get_column_fill(col = 0)
          validate_workbook
          validate_nonnegative(col)
          wb = @workbook
          raise "Workbook not associated" unless wb

          wb.get_fill_color(get_col_xf(col))
        end

        # @param col [Integer]
        # @return [String, nil]
        #: (?Integer col) -> String?
        def get_column_font_name(col = 0)
          column_font(col)&.get_name
        end

        # @param col [Integer]
        # @return [Numeric, nil]
        #: (?Integer col) -> (Float | Integer)?
        def get_column_font_size(col = 0)
          column_font(col)&.get_size
        end

        # @param col [Integer]
        # @return [String]
        #: (?Integer col) -> String
        def get_column_font_color(col = 0)
          column_font(col)&.get_rgb_color || "000000"
        end

        # @param col [Integer]
        # @return [bool?]
        #: (?Integer col) -> bool?
        def is_column_italicized(col = 0)
          column_font(col)&.is_italic
        end

        # @param col [Integer]
        # @return [bool?]
        #: (?Integer col) -> bool?
        def is_column_bolded(col = 0)
          column_font(col)&.is_bold
        end

        # @param col [Integer]
        # @return [bool?]
        #: (?Integer col) -> bool?
        def is_column_underlined(col = 0)
          column_font(col)&.is_underlined
        end

        # @param col [Integer]
        # @return [bool?]
        #: (?Integer col) -> bool?
        def is_column_struckthrough(col = 0)
          column_font(col)&.is_strikethrough
        end

        # @param column_index [Integer]
        # @return [Float, nil]
        #: (?Integer column_index) -> Float?
        def get_column_width_raw(column_index = 0)
          validate_workbook
          validate_nonnegative(column_index)

          range = @cols.locate_range(column_index)
          range&.width
        end

        # @param column_index [Integer]
        # @return [Integer]
        #: (?Integer column_index) -> Integer
        def get_column_width(column_index = 0)
          width = get_column_width_raw(column_index)
          return ColumnRange::DEFAULT_WIDTH if width.nil?

          (width - (5.0 / ColumnRange::MAX_DIGIT_WIDTH)).round
        end

        # @param column_index [Integer]
        # @param width [Numeric]
        # @return [void]
        #: (Integer column_index, Numeric width) -> void
        def change_column_width_raw(column_index, width)
          validate_workbook
          ensure_cell_exists(0, column_index)
          range = @cols.get_range(column_index)
          range.width = width.to_f
          range.custom_width = true
        end

        # @param column_index [Integer]
        # @param width_in_chars [Numeric]
        # @return [void]
        #: (Integer column_index, ?Numeric width_in_chars) -> void
        def change_column_width(column_index, width_in_chars = ColumnRange::DEFAULT_WIDTH)
          change_column_width_raw(column_index, ColumnRange.chars2raw(width_in_chars))
        end

        # @param column_index [Integer]
        # @param color_code [String]
        # @return [void]
        #: (Integer column_index, ?String color_code) -> void
        def change_column_fill(column_index, color_code = "ffffff")
          validate_workbook
          Color.validate_color(color_code)
          ensure_cell_exists(0, column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          @cols.get_range(column_index).style_index = wb.modify_fill(get_cols_style_index(column_index), color_code)

          @rows.each do |row|
            next if row.nil?

            c = row[column_index]
            c&.change_fill(color_code)
          end
        end

        # @param column_index [Integer]
        # @param change_type [Integer]
        # @param arg [untyped]
        # @param font [Font]
        # @param xf [XF]
        # @return [void]
        #: (Integer column_index, Integer change_type, untyped arg, Font font, XF xf) -> void
        def change_column_font(column_index, change_type, arg, font, xf)
          validate_workbook
          ensure_cell_exists(0, column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          new_xf = wb.register_new_font(font, xf)
          @cols.get_range(column_index).style_index = wb.register_new_xf(new_xf)

          @rows.each do |row|
            c = row&.[](column_index)
            c&.font_switch(change_type, arg)
          end
        end

        # @param column_index [Integer]
        # @param font_name [String]
        # @return [void]
        #: (?Integer column_index, ?String font_name) -> void
        def change_column_font_name(column_index = 0, font_name = "Verdana")
          xf = get_col_xf(column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_name(font_name)
          change_column_font(column_index, NAME, font_name, font, xf)
        end

        # @param column_index [Integer]
        # @param font_size [Numeric]
        # @return [void]
        #: (Integer column_index, ?Numeric font_size) -> void
        def change_column_font_size(column_index, font_size = 10)
          xf = get_col_xf(column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_size(font_size)
          change_column_font(column_index, SIZE, font_size, font, xf)
        end

        # @param column_index [Integer]
        # @param font_color [String]
        # @return [void]
        #: (Integer column_index, ?String font_color) -> void
        def change_column_font_color(column_index, font_color = "000000")
          Color.validate_color(font_color)
          xf = get_col_xf(column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_rgb_color(font_color)
          change_column_font(column_index, COLOR, font_color, font, xf)
        end

        # @param column_index [Integer]
        # @param italicized [Boolean]
        # @return [void]
        #: (Integer column_index, ?bool italicized) -> void
        def change_column_italics(column_index, italicized = false)
          xf = get_col_xf(column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_italic(italicized)
          change_column_font(column_index, ITALICS, italicized, font, xf)
        end

        # @param column_index [Integer]
        # @param bolded [Boolean]
        # @return [void]
        #: (Integer column_index, ?bool bolded) -> void
        def change_column_bold(column_index, bolded = false)
          xf = get_col_xf(column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_bold(bolded)
          change_column_font(column_index, BOLD, bolded, font, xf)
        end

        # @param column_index [Integer]
        # @param underlined [Boolean]
        # @return [void]
        #: (Integer column_index, ?bool underlined) -> void
        def change_column_underline(column_index, underlined = false)
          xf = get_col_xf(column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_underline(underlined)
          change_column_font(column_index, UNDERLINE, underlined, font, xf)
        end

        # @param column_index [Integer]
        # @param struckthrough [Boolean]
        # @return [void]
        #: (Integer column_index, ?bool struckthrough) -> void
        def change_column_strikethrough(column_index, struckthrough = false)
          xf = get_col_xf(column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          font = (wb.fonts[xf.font_id || 0] || wb.fonts[0]).dup
          font.set_strikethrough(struckthrough)
          change_column_font(column_index, STRIKETHROUGH, struckthrough, font, xf)
        end

        # @param column_index [Integer]
        # @yieldparam alignment [Alignment]
        # @return [void]
        #: (Integer column_index) { (Alignment) -> void } -> void
        def change_column_alignment(column_index, &)
          validate_workbook
          ensure_cell_exists(0, column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          @cols.get_range(column_index).style_index = wb.modify_alignment(get_cols_style_index(column_index), &)
          change_column_width(column_index) if get_column_width_raw(column_index).nil?

          @rows.each do |row|
            next if row.nil?

            c = row[column_index]
            next if c.nil?

            c.style_index = wb.modify_alignment(c.style_index, &)
          end
        end

        # @param column_index [Integer]
        # @param alignment [String]
        # @return [void]
        #: (Integer column_index, ?String alignment) -> void
        def change_column_horizontal_alignment(column_index, alignment = "center")
          change_column_alignment(column_index) { |a| a.horizontal = alignment }
        end

        # @param column_index [Integer]
        # @param alignment [String]
        # @return [void]
        #: (Integer column_index, ?String alignment) -> void
        def change_column_vertical_alignment(column_index, alignment = "center")
          change_column_alignment(column_index) { |a| a.vertical = alignment }
        end

        # @param column_index [Integer]
        # @param direction [Symbol, String]
        # @param weight [String]
        # @return [void]
        #: (Integer column_index, Symbol | String direction, String weight) -> void
        def change_column_border(column_index, direction, weight)
          validate_workbook
          ensure_cell_exists(0, column_index)
          wb = @workbook
          raise "Workbook not associated" unless wb

          @cols.get_range(column_index).style_index = wb.modify_border(get_cols_style_index(column_index), direction, weight)

          @rows.each do |row|
            next if row.nil?

            c = row[column_index]
            c&.change_border(direction, weight)
          end
        end

        # @param column_index [Integer]
        # @param direction [Symbol, String]
        # @param color [String]
        # @return [void]
        #: (Integer column_index, Symbol | String direction, String color) -> void
        def change_column_border_color(column_index, direction, color)
          validate_workbook
          ensure_cell_exists(0, column_index)
          Color.validate_color(color)
          wb = @workbook
          raise "Workbook not associated" unless wb

          @cols.get_range(column_index).style_index = wb.modify_border_color(get_cols_style_index(column_index), direction, color)

          @rows.each do |row|
            next if row.nil?

            c = row[column_index]
            c&.change_border_color(direction, color)
          end
        end

        # @param col [Integer]
        # @param border_direction [Symbol, String]
        # @return [String, nil]
        #: (Integer col, Symbol | String border_direction) -> String?
        def get_column_border(col, border_direction)
          validate_workbook
          wb = @workbook
          raise "Workbook not associated" unless wb

          xf = get_col_xf(col)
          border = wb.borders[xf.border_id || 0]
          border&.get_edge_style(border_direction)
        end

        # @param col [Integer]
        # @param border_direction [Symbol, String]
        # @return [String, nil]
        #: (Integer col, Symbol | String border_direction) -> String?
        def get_column_border_color(col, border_direction)
          validate_workbook
          wb = @workbook
          raise "Workbook not associated" unless wb

          xf = get_col_xf(col)
          border = wb.borders[xf.border_id || 0]
          border&.get_edge_color(border_direction)
        end

        # @param col [Integer]
        # @param type [Symbol, String]
        # @return [untyped]
        #: (Integer col, Symbol | String type) -> untyped
        def get_column_alignment(col, type)
          validate_workbook
          xf = get_col_xf(col)
          return nil if xf.alignment.nil?

          xf.alignment.send(type)
        end

        # Helper method to get the style index for a column.
        #
        # @param column_index [Integer]
        # @return [Integer]
        #: (Integer column_index) -> Integer
        def get_col_style(column_index)
          get_cols_style_index(column_index)
        end

        # Adds data validation list.
        #
        # @param ref [String, Reference]
        # @param list_arr [Array<String>]
        # @return [void]
        #: (String | Reference ref, Array[String] list_arr) -> void
        def add_validation_list(ref, list_arr)
          validate_workbook
          expr = "\"#{list_arr.map { |str| str.gsub('"', '""') }.join(",")}\""
          @data_validations ||= []
          @data_validations << {
            sqref: ref.to_s,
            formula: expr,
            type: "list"
          }
        end

        # Converts worksheet to an immutable Xlsxrb::Elements::Worksheet.
        #
        # @return [Xlsxrb::Elements::Worksheet]
        #: () -> Xlsxrb::Elements::Worksheet
        def to_xlsxrb
          sorted_rows = @rows.compact.sort_by(&:index).map(&:to_xlsxrb)

          merged_unmapped = @unmapped_data.dup
          facade_meta = (merged_unmapped[:facade] || {}).dup
          facade_meta[:merge_cells] = @merged_cells.map { |mc| mc.ref.to_s } if @merged_cells && !@merged_cells.empty?
          merged_unmapped[:facade] = facade_meta unless facade_meta.empty?

          # Convert cols to xlsxrb columns if present
          cols_data = if @cols.empty?
                        @columns
                      else
                        @cols.map do |cr|
                          col_unmapped = {}
                          col_unmapped[:style_index] = cr.style_index if cr.style_index&.positive?
                          Xlsxrb::Elements::Column.new(
                            index: [(cr.min || 1) - 1, 0].max,
                            width: cr.width,
                            hidden: cr.hidden || false,
                            custom_width: cr.custom_width || false,
                            unmapped_data: col_unmapped
                          )
                        end
                      end

          Xlsxrb::Elements::Worksheet.new(
            name: @sheet_name,
            rows: sorted_rows,
            columns: cols_data,
            charts: @charts,
            data_validations: @data_validations || [],
            unmapped_data: merged_unmapped
          )
        end

        # Inspect representation.
        #
        # @return [String]
        #: () -> String
        def inspect
          "#<#{self.class.name} sheet_name=#{@sheet_name.inspect} rows_count=#{@rows.compact.size}>"
        end
      end
    end
  end
end
