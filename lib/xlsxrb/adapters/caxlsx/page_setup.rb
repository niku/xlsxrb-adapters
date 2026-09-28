# frozen_string_literal: true

# rbs_inline: enabled

require_relative "util"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Page margins settings
      class PageMargins
        include OptionsParser
        include SerializedAttributes

        DEFAULT_LEFT_RIGHT_TOP_BOTTOM = 1.0
        DEFAULT_HEADER_FOOTER = 0.5

        attr_reader :left #: Float
        attr_reader :right #: Float
        attr_reader :top #: Float
        attr_reader :bottom #: Float
        attr_reader :header #: Float
        attr_reader :footer #: Float

        serializable_attributes :left, :right, :top, :bottom, :header, :footer

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          set(
            left: DEFAULT_LEFT_RIGHT_TOP_BOTTOM,
            right: DEFAULT_LEFT_RIGHT_TOP_BOTTOM,
            top: DEFAULT_LEFT_RIGHT_TOP_BOTTOM,
            bottom: DEFAULT_LEFT_RIGHT_TOP_BOTTOM,
            header: DEFAULT_HEADER_FOOTER,
            footer: DEFAULT_HEADER_FOOTER
          )
          set(options) unless options.empty?
        end

        #: (untyped v) -> void
        def left=(v)
          Caxlsx.validate_float(v)
          @left = v.to_f
        end

        #: (untyped v) -> void
        def right=(v)
          Caxlsx.validate_float(v)
          @right = v.to_f
        end

        #: (untyped v) -> void
        def top=(v)
          Caxlsx.validate_float(v)
          @top = v.to_f
        end

        #: (untyped v) -> void
        def bottom=(v)
          Caxlsx.validate_float(v)
          @bottom = v.to_f
        end

        #: (untyped v) -> void
        def header=(v)
          Caxlsx.validate_float(v)
          @header = v.to_f
        end

        #: (untyped v) -> void
        def footer=(v)
          Caxlsx.validate_float(v)
          @footer = v.to_f
        end

        # @param options [Hash]
        #: (Hash[Symbol, untyped] options) -> void
        def set(options)
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("pageMargins", str)
        end

        #: () -> Hash[Symbol, Float]
        def to_xlsxrb_hash
          {
            left: @left,
            right: @right,
            top: @top,
            bottom: @bottom,
            header: @header,
            footer: @footer
          }
        end
      end

      # Page setup properties
      class PageSetUpPr
        include OptionsParser
        include SerializedAttributes

        attr_reader :fit_to_page #: bool?

        serializable_attributes :fit_to_page

        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          parse_options options
        end

        #: (untyped v) -> void
        def fit_to_page=(v)
          Caxlsx.validate_boolean(v)
          @fit_to_page = Caxlsx.booleanize(v)
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("pageSetUpPr", str)
        end
      end

      # Page setup settings
      class PageSetup
        include OptionsParser
        include SerializedAttributes

        attr_reader :paper_size #: Integer?
        attr_reader :scale #: Integer?
        attr_reader :first_page_number #: Integer?
        attr_reader :fit_to_width #: Integer?
        attr_reader :fit_to_height #: Integer?
        attr_reader :page_order #: Symbol?
        attr_reader :orientation #: Symbol?
        attr_reader :use_printer_defaults #: bool?
        attr_reader :black_and_white #: bool?
        attr_reader :draft #: bool?
        attr_reader :cell_comments #: Symbol?
        attr_reader :use_first_page_number #: bool?
        attr_reader :horizontal_dpi #: Integer?
        attr_reader :vertical_dpi #: Integer?
        attr_reader :copies #: Integer?
        attr_reader :paper_width #: String?
        attr_reader :paper_height #: String?

        serializable_attributes :paper_size, :scale, :first_page_number, :fit_to_width, :fit_to_height,
                                :page_order, :orientation, :use_printer_defaults, :black_and_white,
                                :draft, :cell_comments, :use_first_page_number, :horizontal_dpi,
                                :vertical_dpi, :copies, :paper_width, :paper_height

        alias fitToWidth fit_to_width
        alias fitToHeight fit_to_height

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @orientation = nil
          @paper_size = nil
          @scale = nil
          @fit_to_width = nil
          @fit_to_height = nil
          parse_options options
        end

        # @param options [Hash]
        #: (Hash[Symbol, untyped] options) -> void
        def set(options)
          parse_options options
        end

        #: (untyped v) -> void
        def orientation=(v)
          Caxlsx.validate_page_orientation(v)
          @orientation = v.to_sym
        end

        #: (untyped v) -> void
        def paper_size=(v)
          Caxlsx.validate_unsigned_int(v)
          @paper_size = v.to_i
        end

        #: (untyped v) -> void
        def scale=(v)
          Caxlsx.validate_scale_10_400(v)
          @scale = v.to_i
        end

        #: (untyped v) -> void
        def fit_to_width=(v)
          Caxlsx.validate_unsigned_int(v)
          @fit_to_width = v.to_i
        end
        alias fitToWidth= fit_to_width=

        #: (untyped v) -> void
        def fit_to_height=(v)
          Caxlsx.validate_unsigned_int(v)
          @fit_to_height = v.to_i
        end
        alias fitToHeight= fit_to_height=

        #: (untyped v) -> void
        def paper_width=(v)
          Caxlsx.validate_number_with_unit(v)
          @paper_width = v.to_s
        end

        #: (untyped v) -> void
        def paper_height=(v)
          Caxlsx.validate_number_with_unit(v)
          @paper_height = v.to_s
        end

        #: () -> bool
        def fit_to_page?
          !(@fit_to_width.nil? && @fit_to_height.nil?)
        end
        alias fit_to_page fit_to_page?

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = +""
          attrs << %( fitToWidth="#{@fit_to_width}") if @fit_to_width
          attrs << %( fitToHeight="#{@fit_to_height}") if @fit_to_height
          attrs << %( orientation="#{@orientation}") if @orientation
          attrs << %( paperSize="#{@paper_size}") if @paper_size
          attrs << %( scale="#{@scale}") if @scale
          attrs << %( paperWidth="#{@paper_width}") if @paper_width
          attrs << %( paperHeight="#{@paper_height}") if @paper_height
          str << "<pageSetup#{attrs}/>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          {
            orientation: @orientation,
            paper_size: @paper_size,
            scale: @scale,
            fit_to_width: @fit_to_width,
            fit_to_height: @fit_to_height
          }.compact
        end
      end

      # Print options settings
      class PrintOptions
        include OptionsParser

        attr_reader :grid_lines #: bool
        attr_reader :headings #: bool
        attr_reader :horizontal_centered #: bool
        attr_reader :vertical_centered #: bool

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @grid_lines = false
          @headings = false
          @horizontal_centered = false
          @vertical_centered = false
          parse_options options
        end

        # @param options [Hash]
        #: (Hash[Symbol, untyped] options) -> void
        def set(options)
          parse_options options
        end

        #: (untyped v) -> void
        def grid_lines=(v)
          Caxlsx.validate_boolean(v)
          @grid_lines = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def headings=(v)
          Caxlsx.validate_boolean(v)
          @headings = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def horizontal_centered=(v)
          Caxlsx.validate_boolean(v)
          @horizontal_centered = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def vertical_centered=(v)
          Caxlsx.validate_boolean(v)
          @vertical_centered = Caxlsx.booleanize(v)
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = %(gridLines="#{@grid_lines ? 1 : 0}" headings="#{@headings ? 1 : 0}" horizontalCentered="#{@horizontal_centered ? 1 : 0}" verticalCentered="#{@vertical_centered ? 1 : 0}")
          str << "<printOptions #{attrs}/>"
          str
        end

        #: () -> Hash[Symbol, bool]
        def to_xlsxrb_hash
          {
            grid_lines: @grid_lines,
            headings: @headings,
            horizontal_centered: @horizontal_centered,
            vertical_centered: @vertical_centered
          }
        end
      end

      # Header and footer settings
      class HeaderFooter
        include OptionsParser

        attr_accessor :different_first #: bool?
        attr_accessor :different_odd_even #: bool?
        attr_accessor :odd_header #: String?
        attr_accessor :odd_footer #: String?
        attr_accessor :even_header #: String?
        attr_accessor :even_footer #: String?
        attr_accessor :first_header #: String?
        attr_accessor :first_footer #: String?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          parse_options options
        end

        # @param options [Hash]
        #: (Hash[Symbol, untyped] options) -> void
        def set(options)
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = +""
          attrs << ' differentFirst="1"' if @different_first
          attrs << ' differentOddEven="1"' if @different_odd_even
          str << "<headerFooter#{attrs}>"
          str << "<oddHeader>#{Caxlsx.coder.encode(@odd_header)}</oddHeader>" if @odd_header
          str << "<oddFooter>#{Caxlsx.coder.encode(@odd_footer)}</oddFooter>" if @odd_footer
          str << "<evenHeader>#{Caxlsx.coder.encode(@even_header)}</evenHeader>" if @even_header
          str << "<evenFooter>#{Caxlsx.coder.encode(@even_footer)}</evenFooter>" if @even_footer
          str << "<firstHeader>#{Caxlsx.coder.encode(@first_header)}</firstHeader>" if @first_header
          str << "<firstFooter>#{Caxlsx.coder.encode(@first_footer)}</firstFooter>" if @first_footer
          str << "</headerFooter>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          {
            different_first: @different_first,
            different_odd_even: @different_odd_even,
            odd_header: @odd_header,
            odd_footer: @odd_footer,
            even_header: @even_header,
            even_footer: @even_footer,
            first_header: @first_header,
            first_footer: @first_footer
          }.compact
        end
      end

      # Break represents a page break
      class Break
        include OptionsParser
        include SerializedAttributes

        attr_accessor :id #: Integer?
        attr_accessor :min #: Integer?
        attr_accessor :max #: Integer?
        attr_accessor :man #: bool?
        attr_accessor :pt #: bool?

        serializable_attributes :id, :min, :max, :man, :pt

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @id = 0
          @min = 0
          @max = 0
          @man = true
          @pt = nil
          parse_options options
          yield self if block_given?
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("brk", str)
        end
      end

      # RowBreaks collection
      class RowBreaks < SimpleTypedList
        def initialize
          super(Break, "rowBreaks")
        end
      end

      # ColBreaks collection
      class ColBreaks < SimpleTypedList
        def initialize
          super(Break, "colBreaks")
        end
      end
    end
  end
end
