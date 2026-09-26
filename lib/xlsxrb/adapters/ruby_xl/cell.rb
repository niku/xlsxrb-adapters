# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require_relative "formula"
require_relative "color"
require_relative "styles"
require_relative "reference"
require_relative "data_type"
require_relative "text"

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a mutable cell wrapper compatible with RubyXL::Cell.
      class Cell
        MAX_STRING_LENGTH = 32_767
        NUMBER_REGEXP = /\A-?\d+((?:\.\d+)?(?:e[+-]?\d+)?)?\Z/i

        attr_accessor :worksheet, :is #: Worksheet?
        attr_accessor :row #: Integer
        attr_accessor :column #: Integer
        attr_accessor :style_index #: Integer?
        attr_accessor :datatype #: String?
        attr_accessor :hyperlink #: String?
        attr_accessor :tooltip #: String?

        alias row_index row
        alias row_index= row=
        alias column_index column
        alias column_index= column= #: untyped
        attr_reader :formula #: Formula?

        # @param worksheet [Worksheet, nil] Parent worksheet.
        # @param row [Integer] 0-based row index.
        # @param column [Integer] 0-based column index.
        # @param value [Object, nil] Ruby-native cell value.
        # @param formula [Formula, String, Xlsxrb::Elements::Formula, nil] Formula for the cell.
        # @param style_index [Integer, nil] Style index.
        # @param datatype [String, nil] OOXML datatype.
        # @param raw_value [Object, nil] Raw cell value.
        # @param is [Object, nil] Inline string / RichText object.
        #: (?worksheet: Worksheet?, ?row: Integer, ?column: Integer, ?value: untyped, ?formula: (Formula | String | Xlsxrb::Elements::Formula)?, ?style_index: Integer?, ?datatype: String?, ?raw_value: untyped, ?is: untyped) -> void
        # rubocop:disable Naming/MethodParameterName
        def initialize(worksheet: nil, row: 0, column: 0, value: nil, formula: nil, style_index: 0, datatype: nil, raw_value: nil, is: nil)
          # rubocop:enable Naming/MethodParameterName
          @worksheet = worksheet
          @row = row
          @column = column
          @style_index = style_index.to_i
          @datatype = datatype
          @formula = normalize_formula(formula)
          @is = is
          @value = value || is&.to_s
          @raw_value = raw_value
        end

        # Returns the number format object for this cell.
        #
        # @return [NumberFormat, nil]
        #: () -> NumberFormat?
        def number_format
          return nil unless workbook&.stylesheet

          workbook.stylesheet.get_number_format_by_id(get_cell_xf.num_fmt_id)
        rescue StandardError
          nil
        end

        # Checks if the cell value represents a date or time.
        #
        # @return [bool]
        #: () -> bool
        def is_date?
          return false unless case raw_value
                              when Numeric then true
                              when String then raw_value =~ NUMBER_REGEXP
                              else false
                              end

          return true if @value.is_a?(Date) || @value.is_a?(Time) || @value.is_a?(DateTime)

          number_format&.is_date_format? || false
        end

        # Returns the cell value.
        #
        # @param _args [Object] Optional arguments.
        # @return [Object, nil]
        #: (?untyped _args) -> untyped
        def value(_args = {})
          r = raw_value

          case datatype
          when DataType::SHARED_STRING
            wb = workbook
            wb&.shared_strings_container ? wb.shared_strings_container[r.to_i].to_s : r.to_s
          when DataType::INLINE_STRING
            @is ? @is.to_s : r.to_s
          when DataType::RAW_STRING
            r
          when DataType::DATE
            if r
              begin
                DateTime.parse(r.to_s)
              rescue StandardError
                r
              end
            end
          when DataType::ERROR
            nil
          else
            if @is
              @is.to_s
            elsif is_date?
              wb = workbook
              if wb
                res = wb.num_to_date(r.to_f)
                if @value.is_a?(Date) && !@value.is_a?(DateTime)
                  res&.to_date
                else
                  res
                end
              else
                @value
              end
            elsif r.is_a?(String) && (r =~ NUMBER_REGEXP)
              Regexp.last_match(1) == "" ? r.to_i : r.to_f
            else
              r.nil? ? @value : r
            end
          end
        end

        # @return [Workbook, nil]
        #: () -> Workbook?
        def workbook
          @worksheet&.workbook
        end

        # Validates parent worksheet exists.
        # @return [void]
        #: () -> void
        def validate_worksheet
          raise "Worksheet is not associated with this cell" if @worksheet.nil?
        end

        # Returns Excel reference string like "A1".
        # @return [String]
        #: () -> String
        def r
          Reference.ind2ref(@row, @column)
        end

        # Sets the cell value.
        #
        # @param val [Object, nil]
        #: (untyped val) -> void
        def value=(val)
          validate_string_length!(val)
          @value = val
          @raw_value = nil
        end

        # Returns the raw underlying cell value.
        #
        # @return [Object, nil]
        #: () -> untyped
        def raw_value
          return @raw_value unless @raw_value.nil?

          case @value
          when Date, Time, DateTime
            if workbook
              workbook.date_to_num(@value)
            else
              Xlsxrb::Ooxml::Utils.datetime_to_serial(@value)
            end
          when RichText
            @value.to_s
          else
            @value
          end
        end

        # Sets the raw value of the cell.
        #
        # @param val [Object, nil]
        #: (untyped val) -> void
        def raw_value=(val)
          @raw_value = val
          @value = nil
        end

        # Sets the cell's formula.
        #
        # @param new_formula [Formula, String, Xlsxrb::Elements::Formula, nil]
        #: ((Formula | String | Xlsxrb::Elements::Formula)? new_formula) -> void
        def formula=(new_formula)
          @formula = normalize_formula(new_formula)
        end

        # Removes any formula associated with this cell.
        #
        # @return [void]
        #: () -> void
        def remove_formula
          @formula = nil
        end

        # Modifies the cell value and optionally updates its formula,
        # preserving styles and compatibility with RubyXL::CellConvenienceMethods#change_contents.
        #
        # @param data [Object] The new cell data.
        # @param formula_expression [String, Formula, nil] Optional new formula expression.
        # @return [Object] The assigned raw value.
        #: (?untyped data, ?(String | Formula)? formula_expression) -> untyped
        def change_contents(data = "", formula_expression = nil)
          validate_worksheet
          validate_string_length!(data)

          if formula_expression
            self.datatype = nil
            self.formula = formula_expression
          else
            self.datatype = case data
                            when Date, Time, DateTime, Numeric then nil
                            when RichText then DataType::INLINE_STRING
                            else DataType::RAW_STRING
                            end
          end

          self.is = data if data.is_a?(RichText)
          data = data.to_datetime if data.is_a?(Time)

          @value = data
          @raw_value = case data
                       when Date, Time, DateTime
                         workbook ? workbook.date_to_num(data) : Xlsxrb::Ooxml::Utils.datetime_to_serial(data)
                       when RichText
                         data.to_s
                       else
                         data
                       end

          @raw_value
        end

        # Adds a shared string and sets cell datatype to SHARED_STRING.
        #
        # @param str [String]
        # @return [Integer]
        #: (String str) -> Integer
        def add_shared_string(str)
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          wb.shared_strings_container ||= SharedStringsTable.new
          wb.shared_strings << str unless wb.shared_strings.include?(str)
          self.datatype = DataType::SHARED_STRING
          self.raw_value = wb.shared_strings_container.add(str)
        end

        # Validates string length does not exceed Excel limitation (32767 chars).
        #
        # @param val [Object]
        #: (untyped val) -> void
        def validate_string_length!(val)
          str = case val
                when String then val
                when RichText then val.to_s
                end
          return unless str && str.length > MAX_STRING_LENGTH

          raise ArgumentError, "String length (#{str.length}) exceeds Excel limit of #{MAX_STRING_LENGTH} characters"
        end
        private :validate_string_length!

        # --- Styles Convenience Methods ---

        # @return [XF]
        #: () -> XF
        def get_cell_xf
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          idx = @style_index || 0
          wb.cell_xfs[idx] || wb.cell_xfs[0]
        end

        # @return [Font]
        #: () -> Font
        def get_cell_font
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          xf = get_cell_xf
          wb.fonts[xf.font_id || 0] || wb.fonts[0]
        end

        # @return [Border]
        #: () -> Border
        def get_cell_border
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          xf = get_cell_xf
          wb.borders[xf.border_id || 0] || wb.borders[0]
        end

        # Helper to update font references in workbook stylesheet.
        #
        # @param modified_font [Font]
        #: (Font modified_font) -> void
        def update_font_references(modified_font)
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          new_xf = wb.register_new_font(modified_font, get_cell_xf)
          @style_index = wb.register_new_xf(new_xf)
        end
        private :update_font_references

        # Sets number format for the cell.
        #
        # @param format_code [String]
        # @return [Integer]
        #: (String format_code) -> Integer
        def set_number_format(format_code)
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          new_xf = get_cell_xf.dup
          new_xf.num_fmt_id = wb.stylesheet.register_number_format(format_code)
          new_xf.apply_number_format = true
          @style_index = wb.register_new_xf(new_xf)
        end

        # Changes fill color of cell.
        #
        # @param rgb [String]
        # @return [Integer]
        #: (?String rgb) -> Integer
        def change_fill(rgb = "ffffff")
          validate_worksheet
          Color.validate_color(rgb)
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_fill(@style_index, rgb)
        end

        # @return [String]
        #: () -> String
        def fill_color
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          wb.get_fill_color(get_cell_xf)
        end

        # Changes font name of cell.
        #
        # @param new_font_name [String]
        #: (?String new_font_name) -> void
        def change_font_name(new_font_name = "Verdana")
          validate_worksheet
          font = get_cell_font.dup
          font.set_name(new_font_name)
          update_font_references(font)
        end

        # Changes font size of cell.
        #
        # @param font_size [Numeric]
        #: (?(Float | Integer) font_size) -> void
        def change_font_size(font_size = 10)
          validate_worksheet
          raise "Argument must be a number" unless font_size.is_a?(Numeric)

          font = get_cell_font.dup
          font.set_size(font_size)
          update_font_references(font)
        end

        # Changes font color of cell.
        #
        # @param font_color [String]
        #: (?String font_color) -> void
        def change_font_color(font_color = "000000")
          validate_worksheet
          Color.validate_color(font_color)

          font = get_cell_font.dup
          font.set_rgb_color(font_color)
          update_font_references(font)
        end

        # Changes font italics setting.
        #
        # @param italicized [Boolean]
        #: (?bool italicized) -> void
        def change_font_italics(italicized = false)
          validate_worksheet
          font = get_cell_font.dup
          font.set_italic(italicized)
          update_font_references(font)
        end

        # Changes font bold setting.
        #
        # @param bolded [Boolean]
        #: (?bool bolded) -> void
        def change_font_bold(bolded = false)
          validate_worksheet
          font = get_cell_font.dup
          font.set_bold(bolded)
          update_font_references(font)
        end

        # Changes font underline setting.
        #
        # @param underlined [Boolean]
        #: (?bool underlined) -> void
        def change_font_underline(underlined = false)
          validate_worksheet
          font = get_cell_font.dup
          font.set_underline(underlined)
          update_font_references(font)
        end

        # Changes font strikethrough setting.
        #
        # @param struckthrough [Boolean]
        #: (?bool struckthrough) -> void
        def change_font_strikethrough(struckthrough = false)
          validate_worksheet
          font = get_cell_font.dup
          font.set_strikethrough(struckthrough)
          update_font_references(font)
        end

        # @return [bool?]
        #: () -> bool?
        def is_italicized
          validate_worksheet
          get_cell_font.is_italic
        end

        # @return [bool?]
        #: () -> bool?
        def is_bolded
          validate_worksheet
          get_cell_font.is_bold
        end

        # @return [bool?]
        #: () -> bool?
        def is_underlined
          validate_worksheet
          get_cell_font.is_underlined
        end

        # @return [bool?]
        #: () -> bool?
        def is_struckthrough
          validate_worksheet
          get_cell_font.is_strikethrough
        end

        # @return [String]
        #: () -> String
        def font_name
          validate_worksheet
          get_cell_font.get_name
        end

        # @return [Float | Integer]
        #: () -> (Float | Integer)
        def font_size
          validate_worksheet
          get_cell_font.get_size
        end

        # @return [String]
        #: () -> String
        def font_color
          validate_worksheet
          get_cell_font.get_rgb_color || "000000"
        end

        # Changes horizontal alignment.
        #
        # @param alignment [String]
        # @return [Integer]
        #: (?String alignment) -> Integer
        def change_horizontal_alignment(alignment = "center")
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_alignment(@style_index) { |a| a.horizontal = alignment }
        end

        # Changes vertical alignment.
        #
        # @param alignment [String]
        # @return [Integer]
        #: (?String alignment) -> Integer
        def change_vertical_alignment(alignment = "center")
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_alignment(@style_index) { |a| a.vertical = alignment }
        end

        # Changes text wrap.
        #
        # @param wrap [Boolean]
        # @return [Integer]
        #: (?bool wrap) -> Integer
        def change_text_wrap(wrap = false)
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_alignment(@style_index) { |a| a.wrap_text = wrap }
        end

        # Changes shrink to fit.
        #
        # @param shrink [Boolean]
        # @return [Integer]
        #: (?bool shrink) -> Integer
        def change_shrink_to_fit(shrink = false)
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_alignment(@style_index) { |a| a.shrink_to_fit = shrink }
        end

        # Changes text rotation.
        #
        # @param rot [Integer]
        # @return [Integer]
        #: (Integer rot) -> Integer
        def change_text_rotation(rot)
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_alignment(@style_index) { |a| a.text_rotation = rot }
        end

        # Changes text indent.
        #
        # @param indent [Integer]
        # @return [Integer]
        #: (Integer indent) -> Integer
        def change_text_indent(indent)
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_alignment(@style_index) { |a| a.indent = indent }
        end

        # @return [String, nil]
        #: () -> String?
        def horizontal_alignment
          validate_worksheet
          get_cell_xf.alignment&.horizontal
        end

        # @return [String, nil]
        #: () -> String?
        def vertical_alignment
          validate_worksheet
          get_cell_xf.alignment&.vertical
        end

        # @return [Boolean, nil]
        #: () -> bool?
        def text_wrap
          validate_worksheet
          get_cell_xf.alignment&.wrap_text
        end

        # @return [Integer, nil]
        #: () -> Integer?
        def text_rotation
          validate_worksheet
          get_cell_xf.alignment&.text_rotation
        end

        # @return [Integer, nil]
        #: () -> Integer?
        def text_indent
          validate_worksheet
          get_cell_xf.alignment&.indent
        end

        # Changes border style.
        #
        # @param direction [Symbol, String]
        # @param weight [String]
        # @param diagonals [untyped]
        # @return [Integer]
        #: (Symbol | String direction, String weight, ?untyped diagonals) -> Integer
        def change_border(direction, weight, diagonals = nil)
          validate_worksheet
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_border(@style_index, direction, weight, diagonals)
        end

        # Changes border color.
        #
        # @param direction [Symbol, String]
        # @param color [String]
        # @return [Integer]
        #: (Symbol | String direction, String color) -> Integer
        def change_border_color(direction, color)
          validate_worksheet
          Color.validate_color(color)
          wb = workbook
          raise "Workbook not associated with worksheet" unless wb

          @style_index = wb.modify_border_color(@style_index, direction, color)
        end

        # @param direction [Symbol, String]
        # @return [String, nil]
        #: (Symbol | String direction) -> String?
        def get_border(direction)
          validate_worksheet
          get_cell_border.get_edge_style(direction)
        end

        # @param direction [Symbol, String]
        # @return [String, nil]
        #: (Symbol | String direction) -> String?
        def get_border_color(direction)
          validate_worksheet
          get_cell_border.get_edge_color(direction)
        end

        # @return [Border::Edge]
        #: () -> Border::Edge
        def border_top
          get_cell_border.top
        end

        # @return [Border::Edge]
        #: () -> Border::Edge
        def border_bottom
          get_cell_border.bottom
        end

        # @return [Border::Edge]
        #: () -> Border::Edge
        def border_left
          get_cell_border.left
        end

        # @return [Border::Edge]
        #: () -> Border::Edge
        def border_right
          get_cell_border.right
        end

        # @return [Border::Edge]
        #: () -> Border::Edge
        def border_diagonal
          get_cell_border.diagonal
        end

        # Changes font property based on Worksheet change_type constant.
        #
        # @param change_type [Integer]
        # @param arg [untyped]
        # @return [void]
        #: (Integer change_type, untyped arg) -> void
        def font_switch(change_type, arg)
          case change_type
          when 0 then change_font_name(arg.to_s)
          when 1 then change_font_size(arg)
          when 2 then change_font_color(arg.to_s)
          when 3 then change_font_italics(arg)
          when 4 then change_font_bold(arg)
          when 5 then change_font_underline(arg)
          when 6 then change_font_strikethrough(arg)
          else raise "Invalid change_type: #{change_type}"
          end
        end

        # Adds hyperlink to cell.
        #
        # @param url [String]
        # @param tooltip [String, nil]
        # @return [void]
        #: (String url, ?String? tooltip) -> void
        def add_hyperlink(url, tooltip = nil)
          validate_worksheet
          @hyperlink = url
          @tooltip = tooltip
          @datatype = DataType::RAW_STRING
        end

        # Converts cell to an immutable Xlsxrb::Elements::Cell.
        #
        # @return [Xlsxrb::Elements::Cell]
        #: () -> Xlsxrb::Elements::Cell
        def to_xlsxrb
          formula_obj = if formula
                          Xlsxrb::Elements::Formula.new(
                            expression: formula.to_s,
                            cached_value: value
                          )
                        end

          Xlsxrb::Elements::Cell.new(
            row_index: row,
            column_index: column,
            value: value,
            formula: formula_obj,
            style_index: style_index
          )
        end

        # String representation of cell value.
        #
        # @return [String]
        #: () -> String
        def to_s
          @value.to_s
        end

        # Inspect representation.
        #
        # @return [String]
        #: () -> String
        def inspect
          "#<#{self.class.name} row=#{@row} column=#{@column} value=#{@value.inspect} formula=#{@formula.inspect}>"
        end

        private

        #: (untyped formula_expr) -> Formula?
        def normalize_formula(formula_expr)
          case formula_expr
          when nil
            nil
          when Formula
            formula_expr
          when Xlsxrb::Elements::Formula
            Formula.new(formula_expr.expression)
          when Hash
            Formula.new(formula_expr[:expression] || formula_expr["expression"])
          else
            Formula.new(formula_expr.to_s)
          end
        end
      end
    end
  end
end
