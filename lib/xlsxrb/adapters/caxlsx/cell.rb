# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "cgi/escape"
require_relative "constants"
require_relative "util"
require_relative "color"
require_relative "rich_text"

module Xlsxrb
  module Adapters
    module Caxlsx
      # A single cell in a worksheet
      class Cell
        include OptionsParser

        attr_reader :row #: untyped
        attr_reader :value #: untyped
        attr_accessor :formula_value #: untyped
        attr_accessor :raw_style #: Hash[Symbol, untyped]?
        attr_reader :ssti #: Integer?
        attr_reader :name #: String?

        # Inline style accessors
        attr_reader :font_name #: String?
        attr_reader :charset #: Integer?
        attr_reader :family #: Integer?
        attr_reader :b #: bool?
        attr_reader :i #: bool?
        attr_reader :strike #: bool?
        attr_reader :outline #: bool?
        attr_reader :shadow #: bool?
        attr_reader :condense #: bool?
        attr_reader :extend #: bool?
        attr_reader :u #: Symbol?
        attr_reader :color #: Color?
        attr_reader :sz #: Integer?
        attr_reader :vertAlign #: Symbol?
        attr_reader :scheme #: Symbol?

        INLINE_STYLES = %i[
          value type font_name charset family b i strike outline
          shadow condense extend u vertAlign sz color scheme
        ].freeze

        CELL_TYPES = %i[date time float integer richtext string boolean iso_8601 text].freeze
        CELL_REFERENCE_REGEX = /([A-Z]+)([0-9]+)/

        # @param row [Row]
        # @param value [Object, nil]
        # @param options [Hash]
        #: (untyped row, ?untyped value, ?Hash[Symbol, untyped] options) -> void
        def initialize(row, value = nil, options = {})
          @row = row
          type = options.delete(:type) || cell_type_from_value(value)
          self.type = type unless type == :string

          val = options.delete(:style)
          self.style = val unless val.nil? || val.zero?

          val = options.delete(:formula_value)
          self.formula_value = val unless val.nil?

          val = options.delete(:escape_formulas)
          self.escape_formulas = val unless val.nil?

          val = options.delete(:secure_formulas)
          self.secure_formulas = val unless val.nil?

          parse_options(options) unless options.empty?

          self.value = value
          @value.cell = self if contains_rich_text? && @value.respond_to?(:cell=)
        end

        #: () -> Integer
        def style
          defined?(@style) && @style ? @style : 0
        end

        #: (untyped v) -> void
        def style=(v)
          Caxlsx.validate_unsigned_int(v)
          count = styles.cellXfs.size
          raise ArgumentError, "Invalid cellXfs id" unless v < count

          @style = v.to_i
        end

        #: () -> Integer
        def effective_style_index
          if needs_quote_prefix?
            styles.quote_prefix_style_for(style)
          else
            style
          end
        end

        #: () -> String
        def style_str
          effective_style_index.to_s
        end

        # Merges an additional style definition into this cell
        # @param style [Hash]
        #: (Hash[Symbol, untyped] style) -> void
        def add_style(style)
          self.raw_style ||= {}
          new_style = Caxlsx.hash_deep_merge(@raw_style || {}, style)

          all_edges = %i[top right bottom left]
          if @raw_style && @raw_style[:border] && style[:border]
            border_at = ((@raw_style[:border][:edges] || all_edges) + (style[:border][:edges] || all_edges)).uniq.sort
            new_style[:border][:edges] = border_at
          elsif style[:border]
            new_style[:border] = style[:border]
          end

          self.raw_style = new_style
          wb = @row.worksheet.workbook
          wb.styled_cells << self
        end

        #: () -> Symbol
        def type
          defined?(@type) && @type ? @type : :string
        end

        #: (untyped v) -> void
        def type=(v)
          RestrictionValidator.validate :cell_type, CELL_TYPES, v
          @type = v.to_sym
          self.value = @value if defined?(@value) && !@value.nil?
        end

        #: () -> bool
        def escape_formulas
          defined?(@escape_formulas) && !@escape_formulas.nil? ? @escape_formulas : @row.worksheet.escape_formulas
        end

        #: (untyped v) -> void
        def escape_formulas=(v)
          Caxlsx.validate_boolean(v)
          @escape_formulas = Caxlsx.booleanize(v)
        end

        #: () -> bool
        def secure_formulas
          defined?(@secure_formulas) && !@secure_formulas.nil? ? @secure_formulas : @row.worksheet.secure_formulas
        end

        #: (untyped v) -> void
        def secure_formulas=(v)
          Caxlsx.validate_boolean(v)
          @secure_formulas = Caxlsx.booleanize(v)
        end

        #: () -> bool
        def needs_quote_prefix?
          return false unless secure_formulas
          return false unless %i[string text].include?(type)
          return false if @value.nil? || @value.empty?

          stripped = @value.lstrip
          return false if safe_negative?(stripped)

          stripped.start_with?(FORMULA_PREFIX, *SECONDARY_FORMULA_PREFIXES)
        end

        #: (untyped v) -> void
        def value=(v)
          @value = cast_value(v)
        end

        #: () -> bool
        def is_text_run?
          defined?(@is_text_run) && @is_text_run && !contains_rich_text?
        end

        #: () -> bool
        def contains_rich_text?
          type == :richtext
        end

        #: () -> bool
        def plain_string?
          %i[string text].include?(type) &&
            !value.nil? &&
            !value.to_s.empty? &&
            !is_text_run? &&
            !is_formula? &&
            !is_array_formula?
        end

        #: () -> bool
        def is_formula?
          return false if escape_formulas || secure_formulas

          type == :string && @value.to_s.start_with?(FORMULA_PREFIX)
        end

        #: () -> bool
        def is_array_formula?
          return false if escape_formulas || secure_formulas

          type == :string && @value.to_s.start_with?(ARRAY_FORMULA_PREFIX) && @value.to_s.end_with?(ARRAY_FORMULA_SUFFIX)
        end

        #: () -> Integer
        def index
          @row.cells.index(self) || 0
        end

        #: () -> String
        def r
          Caxlsx.cell_r(index, @row.row_index)
        end

        #: () -> String
        def r_abs
          m = CELL_REFERENCE_REGEX.match(r)
          m ? "$#{m[1]}$#{m[2]}" : "$A$1"
        end

        #: () -> Array[Integer]
        def pos
          [index, @row.row_index]
        end

        #: (String label) -> void
        def name=(label)
          @row.worksheet.workbook.add_defined_name "#{@row.worksheet.name}!#{r_abs}", name: label
          @name = label
        end

        #: (untyped target) -> void
        def merge(target)
          @row.worksheet.merge_cells [self, target]
        end

        #: () -> Float?
        def autowidth
          return nil if @value.nil? || is_formula?

          if contains_rich_text?
            string_width("", font_size) + @value.autowidth
          elsif @style && !@style.zero? && styles.cellXfs[@style]&.alignment&.wrapText
            max_w = 0.0
            @value.to_s.split(/\r?\n/).each do |line|
              w = string_width(line, font_size)
              max_w = w if w > max_w
            end
            max_w
          else
            string_width(@value, font_size)
          end
        end

        #: () -> String
        def clean_value
          if %i[string text].include?(type) && !Caxlsx.trust_input
            Caxlsx.sanitize(::CGI.escapeHTML(@value.to_s))
          else
            @value.to_s
          end
        end

        # Inline style setters
        #: (untyped v) -> void
        def font_name=(v)
          set_run_style :validate_string, :font_name, v
        end

        #: (untyped v) -> void
        def charset=(v)
          set_run_style :validate_unsigned_int, :charset, v
        end

        #: (untyped v) -> void
        def family=(v)
          set_run_style :validate_family, :family, v.to_i
        end

        #: (untyped v) -> void
        def b=(v)
          set_run_style :validate_boolean, :b, v
        end

        #: (untyped v) -> void
        def i=(v)
          set_run_style :validate_boolean, :i, v
        end

        #: (untyped v) -> void
        def strike=(v)
          set_run_style :validate_boolean, :strike, v
        end

        #: (untyped v) -> void
        def outline=(v)
          set_run_style :validate_boolean, :outline, v
        end

        #: (untyped v) -> void
        def shadow=(v)
          set_run_style :validate_boolean, :shadow, v
        end

        #: (untyped v) -> void
        def condense=(v)
          set_run_style :validate_boolean, :condense, v
        end

        #: (untyped v) -> void
        def extend=(v)
          set_run_style :validate_boolean, :extend, v
        end

        #: (untyped v) -> void
        def u=(v)
          set_run_style :validate_cell_u, :u, v
        end

        #: (untyped v) -> void
        def color=(v)
          @color = v.is_a?(Color) ? v : Color.new(rgb: v)
          @is_text_run = true
        end

        #: (untyped v) -> void
        def sz=(v)
          set_run_style :validate_unsigned_int, :sz, v
        end

        #: (untyped v) -> void
        def vertAlign=(v)
          RestrictionValidator.validate :cell_vertAlign, %i[baseline subscript superscript], v
          set_run_style nil, :vertAlign, v
        end

        #: (untyped v) -> void
        def scheme=(v)
          RestrictionValidator.validate :cell_scheme, %i[none major minor], v
          set_run_style nil, :scheme, v
        end

        #: (untyped v) -> void
        def ssti=(v)
          Caxlsx.validate_unsigned_int(v)
          @ssti = v.to_i
        end

        #: (Integer _r_index, Integer _c_index, ?String str) -> String
        def to_xml_string(_r_index, _c_index, str = +"")
          s_attr = " s=\"#{style_str}\""
          ref_attr = " r=\"#{r}\""

          if is_formula?
            f_expr = @value.to_s.delete_prefix("=")
            str << "<c#{ref_attr}#{s_attr}><f>#{::CGI.escapeHTML(f_expr)}</f>"
            str << "<v>#{@formula_value}</v>" unless @formula_value.nil?
            str << "</c>"
          elsif contains_rich_text?
            str << "<c#{ref_attr}#{s_attr} t=\"inlineStr\"><is>#{@value.to_xml_string}</is></c>"
          elsif is_text_run?
            str << "<c#{ref_attr}#{s_attr} t=\"inlineStr\"><is><r><rPr>"
            str << %(<rFont val="#{@font_name}"/>) if @font_name
            str << "<b/>" if @b
            str << "<i/>" if @i
            str << "<strike/>" if @strike
            str << %(<u val="#{@u}"/>) if @u
            str << %(<vertAlign val="#{@vertAlign}"/>) if @vertAlign
            str << %(<sz val="#{@sz}"/>) if @sz
            str << @color.to_xml_string if @color
            str << "</rPr><t>#{clean_value}</t></r></is></c>"
          else
            case type
            when :boolean
              str << "<c#{ref_attr}#{s_attr} t=\"b\"><v>#{@value}</v></c>"
            when :integer, :float
              str << "<c#{ref_attr}#{s_attr}><v>#{@value}</v></c>"
            when :date
              serial = Xlsxrb::Ooxml::Utils.date_to_serial(@value, date1904: @row.worksheet.workbook.date1904)
              str << "<c#{ref_attr}#{s_attr}><v>#{serial}</v></c>"
            when :time
              serial = Xlsxrb::Ooxml::Utils.datetime_to_serial(@value, date1904: @row.worksheet.workbook.date1904)
              str << "<c#{ref_attr}#{s_attr}><v>#{serial}</v></c>"
            else
              str << if plain_string? && @row.worksheet.workbook.use_shared_strings
                       "<c#{ref_attr}#{s_attr} t=\"s\"><v>#{@ssti || 0}</v></c>"
                     else
                       "<c#{ref_attr}#{s_attr} t=\"inlineStr\"><is><t>#{clean_value}</t></is></c>"
                     end
            end
          end
          str
        end

        # Converts adapter cell to native Xlsxrb::Elements::Cell
        # @param r_idx [Integer, nil]
        # @param c_idx [Integer, nil]
        # @return [Xlsxrb::Elements::Cell]
        #: (?Integer? r_idx, ?Integer? c_idx) -> Xlsxrb::Elements::Cell
        def to_xlsxrb(r_idx = nil, c_idx = nil)
          r_idx ||= @row.row_index
          c_idx ||= index

          cell_val = if contains_rich_text?
                       @value.to_xlsxrb
                     elsif is_text_run?
                       font = {}
                       font[:bold] = @b unless @b.nil?
                       font[:italic] = @i unless @i.nil?
                       font[:strike] = @strike unless @strike.nil?
                       font[:underline] = @u unless @u.nil?
                       font[:vert_align] = @vertAlign unless @vertAlign.nil?
                       font[:sz] = @sz unless @sz.nil?
                       font[:color] = @color&.rgb unless @color.nil?
                       font[:name] = @font_name unless @font_name.nil?
                       font[:family] = @family unless @family.nil?
                       font[:scheme] = @scheme unless @scheme.nil?
                       font_h = font.compact
                       run = Xlsxrb::Elements::RichTextRun.new(text: @value.to_s, font: font_h.empty? ? nil : font_h)
                       Xlsxrb::Elements::RichText.new(runs: [run])
                     elsif type == :boolean
                       @value == 1
                     else
                       @value
                     end

          formula_obj = if is_formula?
                          expr = @value.to_s.delete_prefix("=")
                          Xlsxrb::Elements::Formula.new(expression: expr, cached_value: @formula_value)
                        end

          Xlsxrb::Elements::Cell.new(
            row_index: r_idx,
            column_index: c_idx,
            value: cell_val,
            formula: formula_obj,
            style_index: effective_style_index
          )
        end

        private

        def styles
          @row.worksheet.styles
        end

        def string_width(string, font_size)
          font_scale = font_size / @row.worksheet.workbook.font_scale_divisor
          (string.to_s.size + 3) * font_scale
        end

        def font_size
          return sz if sz

          s_idx = style || 0
          if s_idx.zero?
            font = styles.fonts.first
            return font.sz if font && !font.b && !(defined?(@b) && @b)
          end

          font = styles.fonts[styles.cellXfs[s_idx]&.fontId || 0] || styles.fonts.first
          font.b || (defined?(@b) && @b) ? (font.sz * @row.worksheet.workbook.bold_font_multiplier) : font.sz
        end

        def set_run_style(validator, attr, value)
          return unless INLINE_STYLES.include?(attr.to_sym)

          Caxlsx.send(validator, value) unless validator.nil?
          instance_variable_set(:"@#{attr}", value)
          @is_text_run = true
        end

        def safe_negative?(string)
          return false unless string.start_with?("-")
          return true if string.start_with?("- ")

          Float(string, exception: false) != nil
        end

        def cell_type_from_value(v)
          if v.is_a?(Date)
            :date
          elsif v.is_a?(Time)
            :time
          elsif v.is_a?(TrueClass) || v.is_a?(FalseClass)
            :boolean
          elsif v.respond_to?(:to_i) && Caxlsx::NUMERIC_REGEX.match?(v.to_s)
            :integer
          elsif v.respond_to?(:to_f) && (Caxlsx::SAFE_FLOAT_REGEX.match?(v.to_s) || ((matchdata = Caxlsx::MAYBE_FLOAT_REGEX.match(v.to_s)) && matchdata[:exp].to_i.between?(Float::MIN_10_EXP, Float::MAX_10_EXP)))
            :float
          elsif Caxlsx::ISO_8601_REGEX.match?(v.to_s)
            :iso_8601
          elsif v.is_a?(RichText) || (defined?(Xlsxrb::Elements::RichText) && v.is_a?(Xlsxrb::Elements::RichText))
            :richtext
          else
            :string
          end
        end

        def cast_value(v)
          return v if v.nil? || v.is_a?(RichText)
          return RichText.from_xlsxrb(v) if defined?(Xlsxrb::Elements::RichText) && v.is_a?(Xlsxrb::Elements::RichText)

          case type
          when :date
            self.style = STYLE_DATE if style.zero?
            if !v.is_a?(Date) && v.respond_to?(:to_date)
              v.to_date
            else
              v
            end
          when :time
            self.style = STYLE_DATE if style.zero?
            if !v.is_a?(Time) && v.respond_to?(:to_time)
              v.to_time
            else
              v
            end
          when :float
            v.to_f
          when :integer
            v.to_i
          when :boolean
            v ? 1 : 0
          when :iso_8601
            v
          else
            v.to_s
          end
        end
      end
    end
  end
end
