# frozen_string_literal: true

# rbs_inline: enabled

require_relative "constants"

module Xlsxrb
  module Adapters
    module FastExcel
      # Helper module providing hash-based attribute setters and introspection
      module AttributeHelper
        # Sets attributes from a hash of options.
        #
        # @param values [Hash{Symbol => untyped}]
        # @return [self]
        #: (Hash[Symbol, untyped] values) -> self
        def set(values)
          values.each do |key, value|
            if respond_to?(:"#{key}=")
              send(:"#{key}=", value)
            elsif respond_to?(:"set_#{key}=")
              send(:"set_#{key}=", value)
            else
              self[key] = value
            end
          end

          self
        end

        # Returns a hash of all member field values.
        #
        # @return [Hash{Symbol => untyped}]
        #: () -> Hash[Symbol, untyped]
        def fields_hash
          res = {}
          members.each do |key|
            res[key] = respond_to?(key) ? send(key) : self[key]
          end
          res
        end

        # @param pp [PP]
        # @return [void]
        def pretty_print(pp)
          pp.pp fields_hash
        end
      end

      # Format class managing cell font, border, fill, alignment, and number formatting
      class Format
        include AttributeHelper

        MEMBERS = %i[
          file xf_format_indices dxf_format_indices num_xf_formats num_dxf_formats
          xf_index dxf_index xf_id num_format font_name font_scheme num_format_index
          font_index has_font has_dxf_font font_size bold italic font_color underline
          font_strikeout font_outline font_shadow font_script font_family font_charset
          font_condense font_extend theme hyperlink hidden locked text_h_align text_wrap
          text_v_align text_justlast rotation fg_color bg_color dxf_fg_color dxf_bg_color
          pattern has_fill has_dxf_fill fill_index fill_count border_index has_border
          has_dxf_border border_count bottom diag_border diag_type left right top
          bottom_color diag_color left_color right_color top_color indent shrink
          merge_range reading_order just_distrib color_indexed font_only quote_prefix list_pointers
        ].freeze

        # @return [Workbook, nil]
        attr_accessor :workbook

        # @param workbook [Workbook, nil]
        # @param options [Hash{Symbol => untyped}, nil]
        #: (?Workbook? workbook, ?Hash[Symbol, untyped]? options) -> void
        def initialize(workbook = nil, options = nil)
          @workbook = workbook
          @fields = {}
          @fields[:text_h_align] = 0
          @fields[:text_v_align] = 0
          @fields[:font_size] = 0
          @fields[:font_name] = "Calibri"
          @fields[:bottom] = 0
          @fields[:top] = 0
          @fields[:left] = 0
          @fields[:right] = 0
          @fields[:diag_border] = 0
          @fields[:pattern] = 0
          @fields[:rotation] = 0
          @fields[:indent] = 0
          @fields[:underline] = 0
          @fields[:font_script] = 0

          set(options) if options
        end

        # @return [Array<Symbol>]
        #: () -> Array[Symbol]
        def members
          MEMBERS
        end

        # Access member field value by symbol key.
        #
        # @param key [Symbol, String]
        # @return [untyped]
        #: (Symbol | String key) -> untyped
        def [](key)
          @fields[key.to_sym]
        end

        # Set member field value by symbol key.
        #
        # @param key [Symbol, String]
        # @param value [untyped]
        # @return [untyped]
        #: (Symbol | String key, untyped value) -> untyped
        def []=(key, value)
          @fields[key.to_sym] = value
        end

        # Numeric font properties
        %i[font_size underline font_script rotation indent pattern].each do |prop|
          define_method(prop) do
            self[prop]
          end

          define_method(:"#{prop}=") do |value|
            send(:"set_#{prop}", value)
          end
        end

        # Boolean font and layout properties
        %i[bold italic font_outline font_shadow hidden text_wrap font_strikeout shrink text_justlast].each do |prop|
          define_method(prop) do
            self[prop]
          end

          define_method(:"#{prop}=") do |value|
            value ? send(:"set_#{prop}") : self[prop] = 0
          end
        end

        # String properties
        %i[num_format font_name].each do |prop|
          define_method(prop) do
            val = self[prop]
            val&.to_s
          end

          define_method(:"#{prop}=") do |value|
            send(:"set_#{prop}", value)
          end
        end

        # Sets font size in points (0 for default).
        #
        # @param value [Numeric]
        # @return [Numeric]
        #: (Numeric value) -> Numeric
        def set_font_size(value)
          raise ArgumentError, "font size should be >= 0 (use 0 for user default font size)" if value.negative?

          self[:font_size] = value
        end

        # @return [String]
        #: () -> String
        def font_family
          font_name
        end

        # @param value [String]
        # @return [String]
        #: (String value) -> String
        def font_family=(value)
          self.font_name = value
        end

        # Alignments getter returning hash of :horizontal and :vertical symbols
        #
        # @return [Hash{Symbol => Symbol}]
        #: () -> Hash[Symbol, Symbol]
        def align
          {
            horizontal: ALIGN_ENUM.find(self[:text_h_align]) || :align_none,
            vertical: ALIGN_ENUM.find(self[:text_v_align]) || :align_none
          }
        end

        # Sets alignment using symbol, string, or hash with :horizontal, :vertical, :h, :v keys.
        #
        # @param value [Symbol, String, Hash{Symbol => untyped}]
        # @return [void]
        #: (Symbol | String | Hash[Symbol, untyped] value) -> void
        def align=(value)
          value = value.to_sym if value.is_a?(String)

          if value.is_a?(Symbol)
            prefixed = :"align_#{value}"
            if ALIGN_ENUM.find(value)
              set_align(value)
            elsif ALIGN_ENUM.find(prefixed)
              set_align(prefixed)
            else
              raise ArgumentError, "Can not set align = #{value.inspect}, possible values are: #{ALIGN_ENUM.symbols}"
            end
          elsif value.is_a?(Hash)
            self.align = :"align_#{value[:horizontal].to_s.sub(/^align_/, "")}" if value[:horizontal]
            self.align = :"align_#{value[:h].to_s.sub(/^align_/, "")}" if value[:h]
            self.align = :"align_vertical_#{value[:vertical].to_s.sub(/^align_vertical_/, "")}" if value[:vertical]
            self.align = :"align_vertical_#{value[:v].to_s.sub(/^align_vertical_/, "")}" if value[:v]

            possible = %i[horizontal h vertical v]
            extras = value.keys - possible
            raise ArgumentError, "Not allowed keys for align: #{extras.inspect}, possible keys: #{possible.inspect}" if extras.size.positive?
          else
            raise ArgumentError, "value must be a symbol or a hash"
          end
        end

        # @param value [Symbol, Integer]
        # @return [void]
        #: (Symbol | Integer value) -> void
        def set_align(value)
          idx = value.is_a?(Numeric) ? value.to_i : ALIGN_ENUM.find(value)
          return unless idx

          sym = ALIGN_ENUM.find(idx)
          if sym.to_s.include?("vertical")
            self[:text_v_align] = idx
          else
            self[:text_h_align] = idx
          end
        end

        # Color properties
        %i[font_color bg_color fg_color bottom_color diag_color left_color right_color top_color].each do |prop|
          define_method(:"#{prop}=") do |value|
            send(:"set_#{prop}", FastExcel.color_to_hex(value))
          end

          define_method(prop) do
            self[prop]
          end

          define_method(:"set_#{prop}") do |val|
            hex = FastExcel.color_to_hex(val)
            self[prop] = hex
          end
        end

        # @param val [Integer, Symbol, String]
        # @return [void]
        #: (untyped val) -> void
        def set_color(val)
          send(:set_font_color, val)
        end

        %i[bottom_color left_color right_color top_color].each do |prop|
          alias_method :"border_#{prop}=", :"#{prop}="
          alias_method :"border_#{prop}", prop
        end

        # Border properties
        %i[bottom diag_border left right top].each do |prop|
          define_method(:"#{prop}=") do |value|
            send(:"set_#{prop}", border_value(value))
          end

          define_method(prop) do
            BORDER_ENUM.find(self[prop])
          end

          define_method(:"set_#{prop}") do |val|
            self[prop] = border_value(val)
          end

          unless prop == :diag_border
            alias_method :"border_#{prop}=", :"#{prop}="
            alias_method :"border_#{prop}", prop
          end
        end

        # @return [Symbol, Integer, nil]
        def border
          bottom
        end

        # Sets all four border edges at once.
        #
        # @param value [Symbol, Integer, String]
        # @return [void]
        #: (Symbol | Integer | String value) -> void
        def border=(value)
          set_border(value)
        end

        # Sets all four border edges.
        #
        # @param value [Symbol, Integer, String]
        # @return [void]
        #: (Symbol | Integer | String value) -> void
        def set_border(value)
          val = border_value(value)
          %i[bottom top left right].each do |edge|
            send(:"set_#{edge}", val)
          end
        end

        # Converts border symbol, string, or integer into valid integer border value.
        #
        # @param value [Symbol, String, Integer]
        # @return [Integer]
        #: (untyped value) -> Integer
        def border_value(value)
          return value.to_i if value.is_a?(Numeric) && BORDER_ENUM.find(value.to_i)

          orig_value = value
          val = value.is_a?(String) ? value.to_sym : value

          found = BORDER_ENUM.find(val)
          return found if found

          found_prefixed = BORDER_ENUM.find(:"border_#{val}")
          return found_prefixed if found_prefixed

          short_symbols = BORDER_ENUM.symbols.map { |s| s.to_s.sub(/^border_/, "").to_sym }
          raise ArgumentError, "Unknown value #{orig_value.inspect} for border. Possible values: #{short_symbols}"
        end

        # Method setters for Libxlsxwriter compatibility
        # @return [void]
        def set_bold
          self[:bold] = 1
        end

        # @return [void]
        def set_italic
          self[:italic] = 1
        end

        # @param val [String]
        # @return [void]
        def set_font_name(val)
          self[:font_name] = val.to_s
        end

        # @param val [String]
        # @return [void]
        def set_num_format(val)
          self[:num_format] = val.to_s
        end

        # @return [void]
        def set_text_wrap
          self[:text_wrap] = 1
        end

        # @return [void]
        def set_font_strikeout
          self[:font_strikeout] = 1
        end

        # @return [void]
        def set_font_outline
          self[:font_outline] = 1
        end

        # @return [void]
        def set_font_shadow
          self[:font_shadow] = 1
        end

        # @return [void]
        def set_shrink
          self[:shrink] = 1
        end

        # @return [void]
        def set_text_justlast
          self[:text_justlast] = 1
        end

        # @param val [Integer, Symbol]
        # @return [void]
        def set_underline(val = 1)
          self[:underline] = val
        end

        # @param val [Integer]
        # @return [void]
        def set_rotation(val)
          self[:rotation] = val
        end

        # @param val [Integer]
        # @return [void]
        def set_indent(val)
          self[:indent] = val
        end

        # @param val [Integer]
        # @return [void]
        def set_pattern(val)
          self[:pattern] = val
        end

        # @param val [Integer, Symbol]
        # @return [void]
        def set_font_script(val)
          self[:font_script] = val
        end

        # Converts format font properties to a hash suitable for Xlsxrb rich text and style builder.
        #
        # @return [Hash{Symbol => untyped}]
        #: () -> Hash[Symbol, untyped]
        def to_font_hash
          color_hex = if self[:font_color]
                        val = self[:font_color]
                        val.is_a?(Integer) ? format("%06X", val & 0xFFFFFF) : Xlsxrb.to_hex_color(val)
                      end
          {
            name: self[:font_name] || "Calibri",
            sz: self[:font_size]&.positive? ? self[:font_size] : nil,
            bold: [true, 1].include?(self[:bold]) || nil,
            italic: [true, 1].include?(self[:italic]) || nil,
            underline: [true, 1].include?(self[:underline]) || nil,
            strike: [true, 1].include?(self[:font_strikeout]) || nil,
            color: color_hex
          }.compact
        end
      end
    end
  end
end
