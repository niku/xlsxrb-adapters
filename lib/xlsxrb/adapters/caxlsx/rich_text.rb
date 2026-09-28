# frozen_string_literal: true

# rbs_inline: enabled

require "cgi/escape"
require_relative "util"
require_relative "color"

module Xlsxrb
  module Adapters
    module Caxlsx
      # A single rich text run with formatting options
      class RichTextRun
        include OptionsParser

        attr_accessor :value, :cell #: untyped
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
          font_name charset family b i strike outline
          shadow condense extend u vertAlign sz color scheme
        ].freeze

        # @param value [Object]
        # @param options [Hash]
        #: (untyped value, ?Hash[Symbol, untyped] options) -> void
        def initialize(value, options = {})
          @value = value
          parse_options(options)
        end

        #: (untyped v) -> void

        alias text value
        alias text= value=

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
          v = :single if [true, 1, true, "true"].include?(v)
          set_run_style :validate_cell_u, :u, v
        end

        #: (untyped v) -> void
        def color=(v)
          @color = v.is_a?(Color) ? v : Color.new(rgb: v)
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

        #: (untyped validator, Symbol attr, untyped val) -> void
        def set_run_style(validator, attr, val)
          return unless INLINE_STYLES.include?(attr)

          Caxlsx.send(validator, val) unless validator.nil?
          instance_variable_set(:"@#{attr}", val)
        end

        #: (Array[Float] widtharray) -> Array[Float]
        def autowidth(widtharray)
          return widtharray if @value.nil?

          font_scale = (@sz || 11) / 10.0
          w = @value.to_s.size * font_scale
          widtharray[-1] = (widtharray[-1] || 0) + w
          widtharray
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<r><rPr>"
          str << %(<rFont val="#{@font_name}"/>) if @font_name
          str << "<b/>" if @b
          str << "<i/>" if @i
          str << "<strike/>" if @strike
          str << %(<u val="#{@u}"/>) if @u
          str << %(<vertAlign val="#{@vertAlign}"/>) if @vertAlign
          str << %(<sz val="#{@sz}"/>) if @sz
          str << @color.to_xml_string if @color
          str << "</rPr>"
          clean = Caxlsx.trust_input ? @value.to_s : ::CGI.escapeHTML(Caxlsx.sanitize(@value.to_s))
          str << "<t>#{clean}</t></r>"
          str
        end

        # Converts to native Xlsxrb::Elements::RichTextRun.
        #
        # @return [Xlsxrb::Elements::RichTextRun]
        #: () -> Xlsxrb::Elements::RichTextRun
        def to_xlsxrb
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
          Xlsxrb::Elements::RichTextRun.new(text: @value.to_s, font: font_h.empty? ? nil : font_h)
        end

        # Builds a RichTextRun from native Xlsxrb::Elements::RichTextRun or Hash.
        #
        # @param xlsxrb_run [Xlsxrb::Elements::RichTextRun, Hash{Symbol => untyped}]
        # @return [RichTextRun]
        #: (untyped xlsxrb_run) -> RichTextRun
        def self.from_xlsxrb(xlsxrb_run)
          text = xlsxrb_run.respond_to?(:text) ? xlsxrb_run.text : xlsxrb_run[:text]
          font = xlsxrb_run.respond_to?(:font) ? xlsxrb_run.font : xlsxrb_run[:font]
          opts = {}
          if font
            opts[:b] = font[:bold] if font.key?(:bold)
            opts[:i] = font[:italic] if font.key?(:italic)
            opts[:strike] = font[:strike] if font.key?(:strike)
            opts[:u] = font[:underline] if font.key?(:underline)
            opts[:vertAlign] = font[:vert_align] if font.key?(:vert_align)
            opts[:sz] = font[:sz] if font.key?(:sz)
            opts[:color] = Color.new(rgb: font[:color]) if font.key?(:color)
            opts[:font_name] = font[:name] if font.key?(:name)
            opts[:family] = font[:family] if font.key?(:family)
            opts[:scheme] = font[:scheme] if font.key?(:scheme)
          end
          new(text.to_s, opts)
        end
      end

      # Collection of rich text runs
      class RichText < SimpleTypedList
        attr_reader :cell #: untyped

        # @param text [String, nil]
        # @param options [Hash]
        #: (?String? text, ?Hash[Symbol, untyped] options) ?{ (self) -> void } -> void
        def initialize(text = nil, options = {})
          super(RichTextRun)
          add_run(text, options) unless text.nil?
          yield self if block_given?
        end

        #: (untyped c) -> void
        def cell=(c)
          @cell = c
          each { |run| run.cell = c }
        end

        #: (String text, ?Hash[Symbol, untyped] options) -> RichTextRun
        def add_run(text, options = {})
          run = RichTextRun.new(text, options)
          self << run
          run
        end

        #: () -> self
        def runs
          self
        end

        #: () -> Float
        def autowidth
          widtharray = [0.0]
          each { |run| run.autowidth(widtharray) }
          widtharray.max || 0.0
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          each { |run| run.to_xml_string(str) }
          str
        end

        # Converts to native Xlsxrb::Elements::RichText.
        #
        # @return [Xlsxrb::Elements::RichText]
        #: () -> Xlsxrb::Elements::RichText
        def to_xlsxrb
          Xlsxrb::Elements::RichText.new(runs: map(&:to_xlsxrb))
        end

        # Builds a RichText from native Xlsxrb::Elements::RichText.
        #
        # @param xlsxrb_rt [Xlsxrb::Elements::RichText]
        # @return [RichText]
        #: (untyped xlsxrb_rt) -> RichText
        def self.from_xlsxrb(xlsxrb_rt)
          rt = new
          (xlsxrb_rt.runs || []).each do |x_run|
            rt << RichTextRun.from_xlsxrb(x_run)
          end
          rt
        end
      end
    end
  end
end
