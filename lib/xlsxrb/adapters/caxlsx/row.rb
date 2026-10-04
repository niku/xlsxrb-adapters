# frozen_string_literal: true

# rbs_inline: enabled

require_relative "util"
require_relative "cell"

module Xlsxrb
  module Adapters
    module Caxlsx
      # A single row in a worksheet
      class Row < SimpleTypedList
        include SerializedAttributes
        include Accessors

        attr_reader :worksheet #: untyped
        attr_reader :outline_level #: Integer?
        attr_reader :s #: Integer?

        alias outlineLevel outline_level

        serializable_attributes :hidden, :outline_level, :collapsed, :custom_format, :s, :ph, :custom_height, :ht

        boolean_attr_accessor :hidden, :collapsed, :custom_format, :ph, :custom_height

        # @param worksheet [Worksheet]
        # @param values [Array]
        # @param options [Hash]
        #: (untyped worksheet, ?Array[untyped] values, ?Hash[Symbol, untyped] options) -> void
        def initialize(worksheet, values = [], options = {})
          @worksheet = worksheet
          super(Cell, nil, values.size + options[:offset].to_i)
          self.height = options.delete(:height)
          worksheet.rows << self
          array_to_cells(values, options)
        end

        #: () -> Float?
        def height
          defined?(@ht) && @ht ? @ht.to_f : nil
        end

        #: (untyped v) -> void
        def height=(v)
          return if v.nil?

          Caxlsx.validate_unsigned_numeric(v)
          @custom_height = true
          @ht = v.to_f
        end

        #: (untyped v) -> void
        def outline_level=(v)
          Caxlsx.validate_unsigned_numeric(v)
          @outline_level = v.to_i
        end
        alias outlineLevel= outline_level=

        #: (untyped v) -> void
        def s=(v)
          Caxlsx.validate_unsigned_numeric(v)
          @custom_format = true
          @s = v.to_i
        end

        #: () -> Integer
        def row_index
          @worksheet.rows.index(self) || 0
        end

        #: () -> self
        def cells
          self
        end

        # Adds a single cell to this row
        # @param value [Object]
        # @param options [Hash]
        # @return [Cell]
        #: (?untyped value, ?Hash[Symbol, untyped] options) -> Cell
        def add_cell(value = "", options = {})
          c = Cell.new(self, value, options)
          self << c
          @worksheet.send(:update_column_info, self, []) if @worksheet.respond_to?(:update_column_info, true)
          c
        end

        #: (untyped color) -> void
        def color=(color)
          each_with_index do |cell, index|
            cell.color = color.is_a?(Array) ? color[index] : color
          end
        end

        #: (untyped style) -> void
        def style=(style)
          each_with_index do |cell, index|
            cell.style = style.is_a?(Array) ? style[index] : style
          end
        end

        #: (untyped val) -> void
        def escape_formulas=(val)
          each_with_index do |cell, index|
            cell.escape_formulas = val.is_a?(Array) ? val[index] : val
          end
        end

        #: (untyped val) -> void
        def secure_formulas=(val)
          each_with_index do |cell, index|
            cell.secure_formulas = val.is_a?(Array) ? val[index] : val
          end
        end

        #: (Integer r_idx, ?String str) -> String
        def to_xml_string(r_idx, str = +"")
          serialized_tag("row", str, r: Caxlsx.row_ref(r_idx)) do
            each_with_index { |cell, c_idx| cell.to_xml_string(r_idx, c_idx, str) }
          end
          str
        end

        # Converts to native Xlsxrb::Elements::Row
        # @param r_idx [Integer, nil]
        # @return [Xlsxrb::Elements::Row]
        #: (?Integer? r_idx) -> Xlsxrb::Elements::Row
        def to_xlsxrb(r_idx = nil)
          effective_row_index = r_idx || row_index
          elements_cells = each_with_index.map do |cell, c_idx|
            cell.to_xlsxrb(effective_row_index, c_idx)
          end
          row_unmapped = {}
          row_unmapped[:style_index] = @s if @s

          Xlsxrb::Elements::Row.new(
            index: effective_row_index,
            cells: elements_cells,
            height: height,
            hidden: hidden || false,
            collapsed: collapsed || false,
            custom_height: custom_height || false,
            outline_level: @outline_level,
            unmapped_data: row_unmapped
          )
        end

        private

        def array_to_cells(values, options = {})
          DataTypeValidator.validate :array_to_cells, Array, values
          opts = options.dup
          types = opts.delete(:types)
          style = opts.delete(:style)
          formula_values = opts.delete(:formula_values)
          escape_formulas = opts.delete(:escape_formulas)
          secure_formulas = opts.delete(:secure_formulas)
          offset = opts.delete(:offset)

          offset&.to_i&.times { |idx| self[idx] = Cell.new(self) }

          has_cols = !@worksheet.column_info.empty?
          values.each_with_index do |val, idx|
            cell_opts = opts.dup
            st = style.is_a?(Array) ? style[idx] : style
            cell_opts[:style] = st || (has_cols ? @worksheet.column_info[idx]&.style : nil)
            cell_opts[:type] = types.is_a?(Array) ? types[idx] : types if types
            cell_opts[:escape_formulas] = escape_formulas.is_a?(Array) ? escape_formulas[idx] : escape_formulas unless escape_formulas.nil?
            cell_opts[:secure_formulas] = secure_formulas.is_a?(Array) ? secure_formulas[idx] : secure_formulas unless secure_formulas.nil?
            cell_opts[:formula_value] = formula_values[idx] if formula_values.is_a?(Array)

            self[idx + offset.to_i] = Cell.new(self, val, cell_opts)
          end
        end
      end
    end
  end
end
