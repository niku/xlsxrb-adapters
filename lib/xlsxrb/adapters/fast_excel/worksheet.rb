# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "xlsxrb"
require_relative "constants"
require_relative "formula"
require_relative "url"
require_relative "format"

module Xlsxrb
  module Adapters
    module FastExcel
      # Represents a mutable worksheet compatible with FastExcel::WorksheetExt and Libxlsxwriter::Worksheet
      class Worksheet
        include AttributeHelper

        # Represents a single cell entry in the worksheet
        CellRecord = Struct.new(
          :row,
          :col,
          :value,
          :format,
          :formula,
          :cached_value,
          :url,
          :url_display,
          :url_tooltip,
          :comment
        )

        # @return [Workbook]
        attr_accessor :workbook

        # @return [String]
        attr_accessor :name

        # @return [Integer]
        attr_accessor :index

        # @param workbook [Workbook]
        # @param name [String]
        # @param index [Integer]
        #: (Workbook workbook, String name, ?Integer index) -> void
        def initialize(workbook, name, index = 0)
          @workbook = workbook
          @name = name
          @index = index
          @is_open = true
          @col_formats = {}
          @col_widths = {}
          @row_heights = {}
          @row_formats = {}
          @last_row_number = -1
          @auto_width = false
          @column_widths = {}
          @rows = {}
          @merged_ranges = []
          @autofilter_range = nil
          @freeze_panes = nil
          @split_panes = nil
          @comments = []
          @hyperlinks = {}
          @images = []
          @charts = []
          @state = :visible
          @print_area = nil
          @print_titles = nil

          @struct_fields = {
            name: name,
            right_to_left: 0,
            vcenter: 0,
            print_headers: 0,
            print_options_changed: 0,
            margin_left: 0.7,
            margin_right: 0.7,
            margin_top: 0.75,
            margin_bottom: 0.75,
            vbreaks_count: 0,
            hbreaks_count: 0,
            filter_on: 0,
            active: 0,
            selected: 0,
            hidden: 0
          }
        end

        # @return [Array<Symbol>]
        #: () -> Array[Symbol]
        def members
          @struct_fields.keys
        end

        # Struct-like field access matching FastExcel / libxlsxwriter struct members
        #
        # @param key [Symbol, String]
        # @return [untyped]
        #: (Symbol | String key) -> untyped
        def [](key)
          k = key.to_sym
          return @name if k == :name

          @struct_fields[k]
        end

        # Struct-like field assignment
        #
        # @param key [Symbol, String]
        # @param value [untyped]
        # @return [untyped]
        #: (Symbol | String key, untyped value) -> untyped
        def []=(key, value)
          k = key.to_sym
          @name = value.to_s if k == :name
          @struct_fields[k] = value
        end

        # Returns the 0-based index of the last row written.
        #
        # @return [Integer]
        #: () -> Integer
        attr_reader :last_row_number

        # Writes a single value to a specific cell with automatic type detection.
        #
        # @param row_number [Integer] 0-based row index
        # @param cell_number [Integer] 0-based column index
        # @param value [Object] Value to write
        # @param format [Format, nil] Cell format
        # @return [void]
        #: (Integer row_number, Integer cell_number, untyped value, ?Format? format) -> void
        def write_value(row_number, cell_number, value, format = nil)
          raise ArgumentError, "Can not write to saved row in constant_memory mode (attempted row: #{row_number}, last saved row: #{last_row_number})" if workbook.constant_memory? && row_number < @last_row_number

          if value.is_a?(Numeric)
            write_number(row_number, cell_number, value, format)
          elsif defined?(Date) && value.is_a?(Date) && !(defined?(DateTime) && value.is_a?(DateTime))
            write_datetime(row_number, cell_number, FastExcel.lxw_datetime(value.to_datetime), format)
          elsif value.is_a?(Time) || (defined?(DateTime) && value.is_a?(DateTime))
            write_number(row_number, cell_number, FastExcel.date_num(value), format)
          elsif value.is_a?(TrueClass) || value.is_a?(FalseClass)
            write_boolean(row_number, cell_number, value ? 1 : 0, format)
          elsif value.is_a?(FastExcel::Formula)
            write_formula(row_number, cell_number, value.fml, format)
          elsif value.is_a?(FastExcel::URL)
            write_url(row_number, cell_number, value.url, format)
            add_text_width(value.url, format, cell_number) if auto_width?
          elsif value.nil?
            write_blank(row_number, cell_number, format)
          else
            str = value.to_s
            write_string(row_number, cell_number, str, format)
            add_text_width(str, format, cell_number) if auto_width?
          end

          @last_row_number = [row_number, @last_row_number].max
        end

        # Writes a row of values to a specified row index.
        #
        # @param row_number [Integer] 0-based row index
        # @param values [Array<Object>] Values to write
        # @param formats [Format, Array<Format>, nil] Format or array of formats
        # @return [void]
        #: (Integer row_number, Array[untyped] values, ?(Format | Array[Format])? formats) -> void
        def write_row(row_number, values, formats = nil)
          values.each_with_index do |value, index|
            format = if formats
                       formats.is_a?(Array) ? formats[index] : formats
                     end
            write_value(row_number, index, value, format)
          end
        end

        # Appends a row of values to the bottom of the worksheet.
        #
        # @param values [Array<Object>] Values to write
        # @param formats [Format, Array<Format>, nil] Format or array of formats
        # @return [void]
        #: (Array[untyped] values, ?(Format | Array[Format])? formats) -> void
        def append_row(values, formats = nil)
          @last_row_number += 1
          write_row(last_row_number, values, formats)
        end

        # Appends a row of values to the bottom of the worksheet.
        #
        # @param values [Array<Object>] Values to write
        # @return [self]
        #: (Array[untyped] values) -> self
        def <<(values)
          append_row(values)
          self
        end

        # Writes a numeric cell value.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param number [Numeric]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, Numeric number, ?Format? format) -> Symbol
        def write_number(row, col, number, format = nil)
          record_cell(row, col, value: number, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Writes a string cell value.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param string [String]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, String string, ?Format? format) -> Symbol
        def write_string(row, col, string, format = nil)
          record_cell(row, col, value: string.to_s, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Writes a datetime / date cell value.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param datetime [Datetime, Date, Time, DateTime]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, untyped datetime, ?Format? format) -> Symbol
        def write_datetime(row, col, datetime, format = nil)
          record_cell(row, col, value: datetime, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Writes a boolean cell value.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param value [Boolean, Integer]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, bool | Integer value, ?Format? format) -> Symbol
        def write_boolean(row, col, value, format = nil)
          bool_val = [true, 1].include?(value)
          record_cell(row, col, value: bool_val, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Writes a formula cell value.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param formula [String]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, String formula, ?Format? format) -> Symbol
        def write_formula(row, col, formula, format = nil)
          record_cell(row, col, formula: formula.to_s, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Writes a formula cell value with precalculated numeric result.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param formula [String]
        # @param format [Format, nil]
        # @param result [Numeric]
        # @return [Symbol]
        #: (Integer row, Integer col, String formula, Format? format, Numeric result) -> Symbol
        def write_formula_num(row, col, formula, format, result)
          record_cell(row, col, formula: formula.to_s, cached_value: result, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Writes an array formula.
        #
        # @param first_row [Integer]
        # @param first_col [Integer]
        # @param last_row [Integer]
        # @param last_col [Integer]
        # @param formula [String]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer first_row, Integer first_col, Integer last_row, Integer last_col, String formula, ?Format? format) -> Symbol
        def write_array_formula(first_row, first_col, last_row, _last_col, formula, format = nil)
          record_cell(first_row, first_col, formula: formula.to_s, format: format)
          @last_row_number = last_row if last_row > @last_row_number
          :no_error
        end

        # Writes an array formula with precalculated numeric result.
        #
        # @param first_row [Integer]
        # @param first_col [Integer]
        # @param last_row [Integer]
        # @param last_col [Integer]
        # @param formula [String]
        # @param format [Format, nil]
        # @param result [Numeric]
        # @return [Symbol]
        #: (Integer first_row, Integer first_col, Integer last_row, Integer last_col, String formula, Format? format, Numeric result) -> Symbol
        def write_array_formula_num(first_row, first_col, last_row, _last_col, formula, format, result)
          record_cell(first_row, first_col, formula: formula.to_s, cached_value: result, format: format)
          @last_row_number = last_row if last_row > @last_row_number
          :no_error
        end

        # Writes a hyperlink URL to a cell.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param url [String]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, String url, ?Format? format) -> Symbol
        def write_url(row, col, url, format = nil)
          write_url_opt(row, col, url, format, nil, nil)
        end

        # Writes a hyperlink URL to a cell with optional display text and tooltip.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param url [String]
        # @param format [Format, nil]
        # @param string [String, nil]
        # @param tooltip [String, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, String url, ?Format? format, ?String? string, ?String? tooltip) -> Symbol
        def write_url_opt(row, col, url, format = nil, string = nil, tooltip = nil)
          val = string || url
          ref = Xlsxrb::Utils.row_col_to_ref(row, col)
          @hyperlinks[ref] = { cell: ref, ref: ref, url: url, display: string, tooltip: tooltip }
          record_cell(row, col, value: val, url: url, url_display: string, url_tooltip: tooltip, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Writes a comment / note to a cell.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param string [String]
        # @return [Symbol]
        #: (Integer row, Integer col, String string) -> Symbol
        def write_comment(row, col, string)
          ref = Xlsxrb::Utils.row_col_to_ref(row, col)
          @comments << { ref: ref, text: string.to_s }
          :no_error
        end

        # Writes a blank cell with formatting.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, Integer col, ?Format? format) -> Symbol
        def write_blank(row, col, format = nil)
          record_cell(row, col, value: nil, format: format)
          @last_row_number = row if row > @last_row_number
          :no_error
        end

        # Sets column width and format for a range of columns.
        #
        # @param start_col [Integer]
        # @param end_col [Integer]
        # @param width [Numeric, nil]
        # @param format [Format, nil]
        # @return [void]
        #: (Integer start_col, Integer end_col, ?Numeric? width, ?Format? format) -> void
        def set_column(start_col, end_col, width = nil, format = nil)
          col_w = width || DEF_COL_WIDTH
          start_col.upto(end_col) do |i|
            @col_widths[i] = col_w
            @col_formats[i] = format if format
          end
        end

        # Sets the column width for a single column.
        #
        # @param col [Integer]
        # @param width [Numeric]
        # @return [void]
        #: (Integer col, ?Numeric width) -> void
        def set_column_width(col, width = 60)
          set_column(col, col, width, @col_formats[col])
        end

        # Sets the column width for a range of columns.
        #
        # @param start_col [Integer]
        # @param end_col [Integer]
        # @param width [Numeric]
        # @return [void]
        #: (Integer start_col, Integer end_col, ?Numeric width) -> void
        def set_columns_width(start_col, end_col, width = 60)
          start_col.upto(end_col) do |i|
            set_column_width(i, width)
          end
        end

        # Sets options for a column range.
        #
        # @param first_col [Integer]
        # @param last_col [Integer]
        # @param width [Numeric]
        # @param format [Format, nil]
        # @param _options [Hash, untyped]
        # @return [Symbol]
        def set_column_opt(first_col, last_col, width, format, _options = nil)
          set_column(first_col, last_col, width, format)
          :no_error
        end

        # Sets the height of a specific row.
        #
        # @param row [Integer]
        # @param height [Numeric]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer row, ?Numeric height, ?Format? format) -> Symbol
        def set_row(row, height = 30, format = nil)
          @row_heights[row] = height
          @row_formats[row] = format if format
          :no_error
        end

        # Sets row options.
        #
        # @param row [Integer]
        # @param height [Numeric]
        # @param format [Format, nil]
        # @param _options [Hash, untyped]
        # @return [Symbol]
        def set_row_opt(row, height, format, _options)
          set_row(row, height, format)
          :no_error
        end

        # Merges a range of cells into a single cell, placing text in the top-left cell.
        #
        # @param first_row [Integer]
        # @param first_col [Integer]
        # @param last_row [Integer]
        # @param last_col [Integer]
        # @param string [String]
        # @param format [Format, nil]
        # @return [Symbol]
        #: (Integer first_row, Integer first_col, Integer last_row, Integer last_col, String string, ?Format? format) -> Symbol
        def merge_range(first_row, first_col, last_row, last_col, string, format = nil)
          ref1 = Xlsxrb::Utils.row_col_to_ref(first_row, first_col)
          ref2 = Xlsxrb::Utils.row_col_to_ref(last_row, last_col)
          @merged_ranges << "#{ref1}:#{ref2}"

          write_string(first_row, first_col, string, format)

          first_row.upto(last_row) do |r|
            first_col.upto(last_col) do |c|
              next if r == first_row && c == first_col

              write_blank(r, c, format)
            end
          end

          :no_error
        end

        # Defines an autofilter range.
        #
        # @param first_row [Integer]
        # @param first_col [Integer]
        # @param last_row [Integer]
        # @param last_col [Integer]
        # @return [Symbol]
        #: (Integer first_row, Integer first_col, Integer last_row, Integer last_col) -> Symbol
        def autofilter(first_row, first_col, last_row, last_col)
          ref1 = Xlsxrb::Utils.row_col_to_ref(first_row, first_col)
          ref2 = Xlsxrb::Utils.row_col_to_ref(last_row, last_col)
          @autofilter_range = "#{ref1}:#{ref2}"
          @struct_fields[:filter_on] = 1
          :no_error
        end

        # Enables autofilter on the worksheet headers.
        #
        # @param start_col [Integer]
        # @param end_col [Integer]
        # @return [Symbol]
        #: (?start_col: Integer, end_col: Integer) -> Symbol
        def enable_filters!(end_col:, start_col: 0)
          autofilter(start_col, 0, @last_row_number, end_col)
        end

        # @return [Boolean]
        #: () -> bool
        def auto_width?
          @auto_width ? true : false
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def auto_width=(v)
          @auto_width = v ? true : false
          @column_widths = {}
        end

        # Returns hash of calculated column widths from auto-width estimation.
        #
        # @return [Hash{Integer => Float}]
        #: () -> Hash[Integer, Float]
        def calculated_column_widths
          @column_widths || {}
        end

        # Tracks text width for auto-width estimation.
        #
        # @param value [Object]
        # @param format [Format, nil]
        # @param cell_number [Integer]
        # @return [void]
        #: (untyped value, Format? format, Integer cell_number) -> void
        def add_text_width(value, format, cell_number)
          font_size = 0
          font_size = format.font_size if format
          font_size = @col_formats[cell_number].font_size if font_size.zero? && @col_formats[cell_number]&.font_size
          font_size = workbook.default_format.font_size if font_size.zero? && workbook&.default_format
          font_size = 11 if font_size.nil? || font_size.zero?

          scale = 0.08
          new_width = scale * font_size * value.to_s.length
          @column_widths[cell_number] = if new_width > (@column_widths[cell_number] || 0)
                                          new_width
                                        else
                                          @column_widths[cell_number]
                                        end
        end

        # Sets worksheet display to right-to-left.
        # @return [void]
        def set_right_to_left
          @struct_fields[:right_to_left] = 1
        end

        # Centers worksheet vertically on printed page.
        # @return [void]
        def center_vertically
          @struct_fields[:print_options_changed] = 1
          @struct_fields[:vcenter] = 1
        end

        # Configures printing of row and column headers.
        # @return [void]
        def print_row_col_headers
          @struct_fields[:print_headers] = 1
          @struct_fields[:print_options_changed] = 1
        end

        # Sets print margins in inches.
        #
        # @param left [Float]
        # @param right [Float]
        # @param top [Float]
        # @param bottom [Float]
        # @return [void]
        #: (Float left, Float right, Float top, Float bottom) -> void
        def set_margins(left, right, top, bottom)
          @struct_fields[:margin_left] = left
          @struct_fields[:margin_right] = right
          @struct_fields[:margin_top] = top
          @struct_fields[:margin_bottom] = bottom
        end

        # Sets print area.
        #
        # @param r1 [Integer, String] First row (or cell range string like "A1:H50")
        # @param c1 [Integer, nil] First column
        # @param r2 [Integer, nil] Last row
        # @param c2 [Integer, nil] Last column
        # @return [Symbol]
        #: (Integer | String r1, ?Integer? c1, ?Integer? r2, ?Integer? c2) -> Symbol
        def print_area(r1, c1 = nil, r2 = nil, c2 = nil)
          if r1.is_a?(String) && c1.nil?
            @print_area = r1
          elsif r1.is_a?(Integer) && c1 && r2 && c2
            ref1 = Xlsxrb::Utils.row_col_to_ref(r1, c1)
            ref2 = Xlsxrb::Utils.row_col_to_ref(r2, c2)
            @print_area = "#{ref1}:#{ref2}"
          end
          :no_error
        end
        alias set_print_area print_area

        # Sets repeating rows for print titles.
        #
        # @param r1 [Integer] First row (0-based)
        # @param r2 [Integer] Last row (0-based)
        # @return [Symbol]
        #: (Integer r1, Integer r2) -> Symbol
        def repeat_rows(r1, r2)
          @print_titles ||= {}
          @print_titles[:rows] = "#{r1 + 1}:#{r2 + 1}"
          :no_error
        end

        # Sets repeating columns for print titles.
        #
        # @param c1 [Integer, String] First col (0-based or column letter)
        # @param c2 [Integer, String] Last col (0-based or column letter)
        # @return [Symbol]
        #: (Integer | String c1, Integer | String c2) -> Symbol
        def repeat_columns(c1, c2)
          c1_name = c1.is_a?(Integer) ? Xlsxrb::Utils.col_index_to_name(c1) : c1.to_s
          c2_name = c2.is_a?(Integer) ? Xlsxrb::Utils.col_index_to_name(c2) : c2.to_s
          @print_titles ||= {}
          @print_titles[:cols] = "#{c1_name}:#{c2_name}"
          :no_error
        end

        # Sets vertical page breaks from buffer or array.
        #
        # @param breaks [Array<Integer>, untyped]
        # @return [void]
        def set_v_pagebreaks(breaks)
          count = count_pagebreaks(breaks)
          @struct_fields[:vbreaks_count] = count
        end

        # Sets horizontal page breaks from buffer or array.
        #
        # @param breaks [Array<Integer>, untyped]
        # @return [void]
        def set_h_pagebreaks(breaks)
          count = count_pagebreaks(breaks)
          @struct_fields[:hbreaks_count] = count
        end

        # @return [void]
        def activate
          @struct_fields[:active] = 1
        end

        # @return [void]
        def select
          @struct_fields[:selected] = 1
        end

        # @return [void]
        def hide
          @struct_fields[:hidden] = 1
          @state = :hidden
        end

        # @return [void]
        def set_first_sheet
          # Compatibility stub
        end

        # Freezes rows and columns at specified split coordinate.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @return [void]
        #: (Integer row, Integer col) -> void
        def freeze_panes(row, col)
          @freeze_panes = { row: row, col: col }
        end

        # Splits panes vertically and horizontally.
        #
        # @param vertical [Float]
        # @param horizontal [Float]
        # @return [void]
        #: (Float vertical, Float horizontal) -> void
        def split_panes(vertical, horizontal)
          @split_panes = { vertical: vertical, horizontal: horizontal }
        end

        # @param first_row [Integer]
        # @param first_col [Integer]
        # @param _top_row [Integer]
        # @param _left_col [Integer]
        # @param _type [Integer]
        # @return [void]
        def freeze_panes_opt(first_row, first_col, _top_row, _left_col, _type)
          freeze_panes(first_row, first_col)
        end

        # @param vertical [Float]
        # @param horizontal [Float]
        # @param _top_row [Integer]
        # @param _left_col [Integer]
        # @return [void]
        def split_panes_opt(vertical, horizontal, _top_row, _left_col)
          split_panes(vertical, horizontal)
        end

        # Inserts an image into the worksheet.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param filename [String]
        # @return [Symbol]
        #: (Integer row, Integer col, String filename) -> Symbol
        def insert_image(row, col, filename)
          insert_image_opt(row, col, filename, nil)
        end

        # Inserts an image with options.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param filename [String]
        # @param options [Hash, nil]
        # @return [Symbol]
        def insert_image_opt(row, col, filename, options = nil)
          @images << { row: row, col: col, filename: filename, options: options }
          :no_error
        end

        # Inserts a chart into the worksheet.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param chart [Object]
        # @return [Symbol]
        def insert_chart(row, col, chart)
          insert_chart_opt(row, col, chart, nil)
        end

        # Inserts a chart with options.
        #
        # @param row [Integer]
        # @param col [Integer]
        # @param chart [Object]
        # @param user_options [Hash, nil]
        # @return [Symbol]
        def insert_chart_opt(row, col, chart, user_options = nil)
          @charts << { row: row, col: col, chart: chart, options: user_options }
          :no_error
        end

        # Finalizes worksheet settings and applies column auto widths.
        #
        # @return [void]
        def close
          return unless auto_width?

          @column_widths.transform_values { |w| w || DEF_COL_WIDTH }.each do |num, width|
            set_column_width(num, width + 0.2)
          end
        end

        # Converts Worksheet to an immutable Xlsxrb::Elements::Worksheet.
        #
        # @param style_map [Hash{Format => Integer}, nil]
        # @return [Xlsxrb::Elements::Worksheet]
        #: (?Hash[Format, Integer]? style_map) -> Xlsxrb::Elements::Worksheet
        def to_xlsxrb(style_map = nil)
          style_map ||= if @workbook.respond_to?(:compile_styles, true)
                          @workbook.send(:compile_styles)[1]
                        else
                          {}
                        end
          close if auto_width?

          elements_rows = []
          sorted_row_indices = @rows.keys.sort

          sorted_row_indices.each do |r_idx|
            cells_hash = @rows[r_idx] || {}
            max_c_idx = cells_hash.keys.max || -1
            row_cells = []

            0.upto(max_c_idx) do |c_idx|
              rec = cells_hash[c_idx]
              if rec
                val = rec.value
                val = format_datetime_value(val) if val.is_a?(Datetime)

                fml = rec.formula ? Xlsxrb.formula(rec.formula, cached_value: rec.cached_value || 0) : nil
                fmt = rec.format || @col_formats[c_idx] || @row_formats[r_idx]
                s_idx = fmt ? style_map[fmt] : nil

                cell_val = rec.formula ? (rec.value || rec.cached_value || 0) : val

                row_cells << Xlsxrb::Elements::Cell.fast_create(
                  r_idx,
                  c_idx,
                  cell_val,
                  s_idx,
                  fml,
                  cell_val&.to_s
                )
              else
                # Empty cell: inherit column format if present
                col_fmt = @col_formats[c_idx]
                s_idx = col_fmt ? style_map[col_fmt] : nil
                row_cells << Xlsxrb::Elements::Cell.fast_create(
                  r_idx,
                  c_idx,
                  nil,
                  s_idx
                )
              end
            end

            elements_rows << Xlsxrb::Elements::Row.new(
              index: r_idx,
              cells: row_cells,
              height: @row_heights[r_idx]
            )
          end

          cols_data = []
          all_col_indices = (@col_widths.keys + @col_formats.keys).uniq.sort
          all_col_indices.each do |c_idx|
            w = @col_widths[c_idx] || DEF_COL_WIDTH
            fmt = @col_formats[c_idx]
            s_idx = fmt ? style_map[fmt] : nil
            col_unmapped = {}
            col_unmapped[:style_index] = s_idx if s_idx
            cols_data << Xlsxrb::Elements::Column.new(
              index: c_idx,
              width: w,
              custom_width: true,
              style_index: s_idx,
              unmapped_data: col_unmapped
            )
          end

          facade_meta = {}
          facade_meta[:merge_cells] = @merged_ranges unless @merged_ranges.empty?
          facade_meta[:auto_filter] = @autofilter_range if @autofilter_range

          if @freeze_panes
            facade_meta[:freeze_pane] = {
              x_split: @freeze_panes[:col],
              y_split: @freeze_panes[:row],
              state: :frozen
            }
          elsif @split_panes
            facade_meta[:split_pane] = {
              vertical: @split_panes[:vertical],
              horizontal: @split_panes[:horizontal]
            }
          end

          facade_meta[:page_margins] = {
            left: @struct_fields[:margin_left],
            right: @struct_fields[:margin_right],
            top: @struct_fields[:margin_top],
            bottom: @struct_fields[:margin_bottom]
          }

          if @struct_fields[:print_options_changed] == 1
            facade_meta[:print_options] = {
              vertical_centered: @struct_fields[:vcenter] == 1,
              headings: @struct_fields[:print_headers] == 1
            }
          end

          facade_meta[:sheet_properties] = { right_to_left: true } if @struct_fields[:right_to_left] == 1

          unmapped = { facade: facade_meta }

          Xlsxrb::Elements::Worksheet.new(
            name: @name,
            rows: elements_rows,
            columns: cols_data,
            unmapped_data: unmapped,
            state: @state,
            comments: @comments,
            hyperlinks: @hyperlinks,
            print_area: @print_area,
            print_titles: @print_titles
          )
        end

        private

        #: (Integer row, Integer col, **untyped kwargs) -> void
        def record_cell(row, col, **kwargs)
          @rows[row] ||= {}
          existing = @rows[row][col]
          if existing
            kwargs.each { |k, v| existing[k] = v unless v.nil? }
          else
            @rows[row][col] = CellRecord.new(row: row, col: col, **kwargs)
          end
        end

        #: (Datetime dt) -> (Date | DateTime)
        def format_datetime_value(dt)
          if dt.hour.zero? && dt.min.zero? && dt.sec.zero?
            Date.new(dt.year, dt.month, dt.day)
          else
            DateTime.new(dt.year, dt.month, dt.day, dt.hour, dt.min, dt.sec)
          end
        end

        #: (untyped breaks) -> Integer
        def count_pagebreaks(breaks)
          if breaks.is_a?(Array)
            idx = breaks.index(0)
            idx || breaks.size
          elsif breaks.respond_to?(:get_uint16)
            cnt = 0
            loop do
              val = begin
                breaks.get_uint16(cnt * 2)
              rescue StandardError
                0
              end
              break if val.nil? || val.zero?

              cnt += 1
            end
            cnt
          else
            0
          end
        end
      end
    end
  end
end
