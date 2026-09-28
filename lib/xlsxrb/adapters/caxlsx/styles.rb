# frozen_string_literal: true

# rbs_inline: enabled

require_relative "constants"
require_relative "util"
require_relative "color"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Font definition
      class Font
        include OptionsParser
        include SerializedAttributes

        attr_reader :name #: String
        attr_reader :sz #: Integer | Float
        attr_reader :family #: Integer?
        attr_reader :b #: bool?
        attr_reader :i #: bool?
        attr_reader :u #: Symbol?
        attr_reader :strike #: bool?
        attr_reader :outline #: bool?
        attr_reader :shadow #: bool?
        attr_reader :condense #: bool?
        attr_reader :extend #: bool?
        attr_reader :color #: Color?
        attr_reader :charset #: Integer?
        attr_reader :scheme #: Symbol?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @name = "Arial"
          @sz = 11
          @family = 1
          parse_options options
        end

        #: (untyped v) -> void
        def name=(v)
          return if v.nil?

          Caxlsx.validate_string(v)
          @name = v.to_s
        end

        #: (untyped v) -> void
        def sz=(v)
          return if v.nil?

          Caxlsx.validate_unsigned_numeric(v)
          @sz = v
        end

        #: (untyped v) -> void
        def family=(v)
          return if v.nil?

          Caxlsx.validate_family(v)
          @family = v.to_i
        end

        #: (untyped v) -> void
        def b=(v)
          return if v.nil?

          Caxlsx.validate_boolean(v)
          @b = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def i=(v)
          return if v.nil?

          Caxlsx.validate_boolean(v)
          @i = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def u=(v)
          return if v.nil?

          v = :single if [true, 1, true, "true"].include?(v)
          Caxlsx.validate_cell_u(v)
          @u = v.to_sym
        end

        #: (untyped v) -> void
        def strike=(v)
          return if v.nil?

          Caxlsx.validate_boolean(v)
          @strike = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def outline=(v)
          return if v.nil?

          Caxlsx.validate_boolean(v)
          @outline = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def shadow=(v)
          return if v.nil?

          Caxlsx.validate_boolean(v)
          @shadow = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def condense=(v)
          return if v.nil?

          Caxlsx.validate_boolean(v)
          @condense = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def extend=(v)
          return if v.nil?

          Caxlsx.validate_boolean(v)
          @extend = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def color=(v)
          @color = Color.from(v)
        end

        #: (untyped v) -> void
        def charset=(v)
          return if v.nil?

          Caxlsx.validate_unsigned_int(v)
          @charset = v.to_i
        end

        #: (untyped v) -> void
        def scheme=(v)
          return if v.nil?

          RestrictionValidator.validate :font_scheme, %i[none major minor], v
          @scheme = v.to_sym
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<font>"
          str << %(<sz val="#{@sz}"/>)
          str << %(<name val="#{Caxlsx.coder.encode(@name)}"/>)
          str << %(<family val="#{@family}"/>) if @family
          str << "<b/>" if @b
          str << "<i/>" if @i
          str << "<strike/>" if @strike
          str << %(<u val="#{@u}"/>) if @u
          str << @color.to_xml_string if @color
          str << %(<charset val="#{@charset}"/>) if @charset
          str << %(<scheme val="#{@scheme}"/>) if @scheme
          str << "</font>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          {
            name: @name,
            sz: @sz,
            bold: @b,
            italic: @i,
            underline: @u,
            strike: @strike,
            color: @color&.rgb,
            family: @family,
            charset: @charset,
            scheme: @scheme
          }.compact
        end
      end

      # PatternFill definition
      class PatternFill
        include OptionsParser

        attr_reader :patternType #: Symbol
        attr_reader :fgColor #: Color?
        attr_reader :bgColor #: Color?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @patternType = :none
          parse_options options
        end

        #: (untyped v) -> void
        def patternType=(v)
          Caxlsx.validate_pattern_type(v)
          @patternType = v.to_sym
        end

        #: (untyped v) -> void
        def fgColor=(v)
          @fgColor = Color.from(v)
        end

        #: (untyped v) -> void
        def bgColor=(v)
          @bgColor = Color.from(v)
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << %(<patternFill patternType="#{@patternType}">)
          str << @fgColor.to_xml_string(+"", "fgColor") if @fgColor
          str << @bgColor.to_xml_string(+"", "bgColor") if @bgColor
          str << "</patternFill>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          res = { pattern: @patternType.to_s }
          res[:fg_color] = @fgColor.rgb if @fgColor&.rgb
          res[:bg_color] = @bgColor.rgb if @bgColor&.rgb
          res
        end
      end

      # GradientStop definition
      class GradientStop
        include OptionsParser

        attr_accessor :position #: Float
        attr_accessor :color #: Color?

        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @position = 0.0
          parse_options options
        end
      end

      # GradientFill definition
      class GradientFill
        include OptionsParser

        attr_accessor :type #: Symbol
        attr_accessor :degree #: Float?
        attr_accessor :stop #: SimpleTypedList

        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @type = :linear
          @stop = SimpleTypedList.new(GradientStop)
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << %(<gradientFill type="#{@type}">)
          str << "</gradientFill>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          { pattern: "solid" }
        end
      end

      # Fill container
      class Fill
        attr_reader :fill_type #: PatternFill | GradientFill

        # @param fill_type [PatternFill, GradientFill]
        #: (PatternFill | GradientFill fill_type) -> void
        def initialize(fill_type)
          @fill_type = fill_type
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<fill>"
          @fill_type.to_xml_string(str)
          str << "</fill>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          @fill_type.to_xlsxrb_hash
        end
      end

      # Border property for an edge
      class BorderPr
        include OptionsParser

        attr_reader :name #: Symbol
        attr_reader :style #: Symbol?
        attr_reader :color #: Color?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @name = :left
          @style = nil
          @color = nil
          parse_options options
        end

        #: (untyped v) -> void
        def name=(v)
          @name = v.to_sym
        end

        #: (untyped v) -> void
        def style=(v)
          @style = v&.to_sym
        end

        #: (untyped v) -> void
        def color=(v)
          @color = Color.from(v)
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          tag = @name.to_s
          if @style.nil?
            str << "<#{tag}/>"
          else
            str << %(<#{tag} style="#{@style}">)
            str << @color.to_xml_string if @color
            str << "</#{tag}>"
          end
          str
        end
      end

      # Border definition
      class Border
        EDGES = %i[left right top bottom diagonal].freeze

        attr_reader :prs #: SimpleTypedList
        attr_accessor :diagonalUp #: bool?
        attr_accessor :diagonalDown #: bool?
        attr_accessor :outline #: bool?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @prs = SimpleTypedList.new(BorderPr)
          @diagonalUp = nil
          @diagonalDown = nil
          @outline = nil
          EDGES.each do |edge|
            @prs << BorderPr.new(name: edge)
          end
          options.each do |k, v|
            setter = :"#{k}="
            send(setter, v) if respond_to?(setter)
          end
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = +""
          attrs << ' diagonalUp="1"' if @diagonalUp
          attrs << ' diagonalDown="1"' if @diagonalDown
          attrs << ' outline="1"' if @outline
          str << "<border#{attrs}>"
          @prs.each { |pr| pr.to_xml_string(str) }
          str << "</border>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          res = {}
          @prs.each do |pr|
            next unless pr.style

            res[pr.name] = {
              style: pr.style.to_s,
              color: pr.color&.rgb
            }.compact
          end
          res[:diagonal_up] = true if @diagonalUp
          res[:diagonal_down] = true if @diagonalDown
          res
        end
      end

      # Number format definition
      class NumFmt
        include OptionsParser

        attr_accessor :numFmtId #: Integer
        attr_accessor :formatCode #: String

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @numFmtId = 0
          @formatCode = ""
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << %(<numFmt numFmtId="#{@numFmtId}" formatCode="#{Caxlsx.coder.encode(@formatCode)}"/>)
        end
      end

      # Cell alignment definition
      class CellAlignment
        include OptionsParser

        attr_reader :horizontal #: Symbol?
        attr_reader :vertical #: Symbol?
        attr_accessor :textRotation #: Integer?
        attr_accessor :wrapText #: bool?
        attr_accessor :indent #: Integer?
        attr_accessor :relativeIndent #: Integer?
        attr_accessor :justifyLastLine #: bool?
        attr_accessor :shrinkToFit #: bool?
        attr_accessor :readingOrder #: Integer?

        alias wrap_text wrapText
        alias wrap_text= wrapText=
        alias shrink_to_fit shrinkToFit
        alias shrink_to_fit= shrinkToFit=
        alias text_rotation textRotation
        alias text_rotation= textRotation=

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          parse_options options
        end

        #: (untyped v) -> void
        def horizontal=(v)
          Caxlsx.validate_horizontal_alignment(v)
          @horizontal = v.to_sym
        end

        #: (untyped v) -> void
        def vertical=(v)
          Caxlsx.validate_vertical_alignment(v)
          @vertical = v.to_sym
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = +""
          attrs << %( horizontal="#{@horizontal}") if @horizontal
          attrs << %( vertical="#{@vertical}") if @vertical
          attrs << %( textRotation="#{@textRotation}") if @textRotation
          attrs << ' wrapText="1"' if @wrapText
          attrs << %( indent="#{@indent}") if @indent
          attrs << ' shrinkToFit="1"' if @shrinkToFit
          str << "<alignment#{attrs}/>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          {
            horizontal: @horizontal,
            vertical: @vertical,
            text_rotation: @textRotation,
            wrap_text: @wrapText,
            indent: @indent,
            shrink_to_fit: @shrinkToFit
          }.compact
        end
      end

      # Cell protection definition
      class CellProtection
        include OptionsParser

        attr_accessor :hidden #: bool?
        attr_accessor :locked #: bool?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @hidden = nil
          @locked = nil
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = +""
          attrs << %( hidden="#{@hidden ? 1 : 0}") unless @hidden.nil?
          attrs << %( locked="#{@locked ? 1 : 0}") unless @locked.nil?
          str << "<protection#{attrs}/>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          {
            hidden: @hidden,
            locked: @locked
          }.compact
        end
      end

      # Formatting record (Xf)
      class Xf
        include OptionsParser

        attr_accessor :numFmtId #: Integer
        attr_accessor :fontId #: Integer
        attr_accessor :fillId #: Integer
        attr_accessor :borderId #: Integer
        attr_accessor :xfId #: Integer
        attr_accessor :quotePrefix #: bool?
        attr_accessor :pivotButton #: bool?
        attr_accessor :applyNumberFormat #: bool?
        attr_accessor :applyFont #: bool?
        attr_accessor :applyFill #: bool?
        attr_accessor :applyBorder #: bool?
        attr_accessor :applyAlignment #: bool?
        attr_accessor :applyProtection #: bool?
        attr_accessor :alignment #: CellAlignment?
        attr_accessor :protection #: CellProtection?

        alias quote_prefix quotePrefix
        alias quote_prefix= quotePrefix=

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @numFmtId = 0
          @fontId = 0
          @fillId = 0
          @borderId = 0
          @xfId = 0
          @quotePrefix = nil
          @pivotButton = nil
          @applyNumberFormat = nil
          @applyFont = nil
          @applyFill = nil
          @applyBorder = nil
          @applyAlignment = nil
          @applyProtection = nil
          @alignment = nil
          @protection = nil
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = %(numFmtId="#{@numFmtId}" fontId="#{@fontId}" fillId="#{@fillId}" borderId="#{@borderId}")
          attrs << %( xfId="#{@xfId}") if @xfId
          attrs << ' quotePrefix="1"' if @quotePrefix
          attrs << ' applyNumberFormat="1"' if @applyNumberFormat || @numFmtId.positive?
          attrs << ' applyFont="1"' if @applyFont || @fontId.positive?
          attrs << ' applyFill="1"' if @applyFill || @fillId.positive?
          attrs << ' applyBorder="1"' if @applyBorder || @borderId.positive?
          attrs << ' applyAlignment="1"' if @applyAlignment || @alignment
          attrs << ' applyProtection="1"' if @applyProtection || @protection
          if @alignment || @protection
            str << "<xf #{attrs}>"
            @alignment&.to_xml_string(str)
            @protection&.to_xml_string(str)
            str << "</xf>"
          else
            str << "<xf #{attrs}/>"
          end
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          res = {
            num_fmt_id: @numFmtId,
            font_id: @fontId,
            fill_id: @fillId,
            border_id: @borderId,
            xf_id: @xfId
          }
          res[:alignment] = @alignment.to_xlsxrb_hash if @alignment
          res[:protection] = @protection.to_xlsxrb_hash if @protection
          res[:quote_prefix] = true if @quotePrefix
          res
        end
      end

      # Named cell style definition
      class CellStyle
        include OptionsParser

        attr_accessor :name #: String
        attr_accessor :xfId #: Integer
        attr_accessor :builtinId #: Integer?
        attr_accessor :iLevel #: Integer?
        attr_accessor :hidden #: bool?
        attr_accessor :customBuiltin #: bool?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @name = "Normal"
          @xfId = 0
          @builtinId = nil
          @iLevel = nil
          @hidden = nil
          @customBuiltin = nil
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = %(name="#{Caxlsx.coder.encode(@name)}" xfId="#{@xfId}")
          attrs += %( builtinId="#{@builtinId}") if @builtinId
          str << "<cellStyle #{attrs}/>"
          str
        end
      end

      # Differential formatting record (dxf)
      class Dxf
        include OptionsParser

        attr_accessor :font #: Font?
        attr_accessor :fill #: Fill?
        attr_accessor :border #: Border?
        attr_accessor :numFmt #: NumFmt?
        attr_accessor :alignment #: CellAlignment?
        attr_accessor :protection #: CellProtection?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @font = nil
          @fill = nil
          @border = nil
          @numFmt = nil
          @alignment = nil
          @protection = nil
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<dxf>"
          @font&.to_xml_string(str)
          @numFmt&.to_xml_string(str)
          @fill&.to_xml_string(str)
          @border&.to_xml_string(str)
          @alignment&.to_xml_string(str)
          @protection&.to_xml_string(str)
          str << "</dxf>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          res = {}
          res[:font] = @font.to_xlsxrb_hash if @font
          res[:fill] = @fill.to_xlsxrb_hash if @fill
          res[:border] = @border.to_xlsxrb_hash if @border
          res[:num_fmt] = { num_fmt_id: @numFmt.numFmtId, format_code: @numFmt.formatCode } if @numFmt
          res
        end
      end

      # Table style element
      class TableStyleElement
        include OptionsParser

        attr_accessor :type #: Symbol
        attr_accessor :size #: Integer
        attr_accessor :dxfId #: Integer

        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @type = :wholeTable
          @size = 1
          @dxfId = 0
          parse_options options
        end
      end

      # Table style
      class TableStyle
        include OptionsParser

        attr_accessor :name #: String
        attr_accessor :pivot #: bool?
        attr_accessor :table #: bool?

        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @name = ""
          parse_options options
        end
      end

      # Table styles container
      class TableStyles < SimpleTypedList
        include OptionsParser

        attr_accessor :defaultTableStyle #: String?
        attr_accessor :defaultPivotStyle #: String?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          super(TableStyle)
          @defaultTableStyle = "TableStyleMedium9"
          @defaultPivotStyle = "PivotStyleLight16"
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          attrs = +""
          attrs << %( defaultTableStyle="#{@defaultTableStyle}") if @defaultTableStyle
          attrs << %( defaultPivotStyle="#{@defaultPivotStyle}") if @defaultPivotStyle
          if empty?
            str << "<tableStyles#{attrs}/>"
          else
            str << "<tableStyles count=\"#{size}\"#{attrs}>"
            each { |ts| ts.to_xml_string(str) }
            str << "</tableStyles>"
          end
          str
        end
      end

      # Theme definition
      class Theme
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          str << %(<a:theme xmlns:a="#{XML_NS_A}" name="Office Theme">)
          str << "<a:themeElements><a:clrScheme name=\"Office\"><a:dk1><a:sysClr val=\"windowText\" lastClr=\"000000\"/></a:dk1><a:lt1><a:sysClr val=\"window\" lastClr=\"FFFFFF\"/></a:lt1><a:dk2><a:srgbClr val=\"1F497D\"/></a:dk2><a:lt2><a:srgbClr val=\"EEECE1\"/></a:lt2><a:accent1><a:srgbClr val=\"4F81BD\"/></a:accent1><a:accent2><a:srgbClr val=\"C0504D\"/></a:accent2><a:accent3><a:srgbClr val=\"9BBB59\"/></a:accent3><a:accent4><a:srgbClr val=\"8064A2\"/></a:accent4><a:accent5><a:srgbClr val=\"4BACC6\"/></a:accent5><a:accent6><a:srgbClr val=\"F79646\"/></a:accent6><a:hlink><a:srgbClr val=\"0000FF\"/></a:hlink><a:folHlink><a:srgbClr val=\"800080\"/></a:folHlink></a:clrScheme></a:themeElements>"
          str << "</a:theme>"
          str
        end
      end

      # Styles manager
      class Styles
        attr_reader :numFmts #: SimpleTypedList
        attr_reader :fonts #: SimpleTypedList
        attr_reader :fills #: SimpleTypedList
        attr_reader :borders #: SimpleTypedList
        attr_reader :cellStyleXfs #: SimpleTypedList
        attr_reader :cellStyles #: SimpleTypedList
        attr_reader :cellXfs #: SimpleTypedList
        attr_reader :dxfs #: SimpleTypedList
        attr_reader :tableStyles #: TableStyles

        # @return [Hash]
        #: () -> Hash[Integer, Hash[Symbol, untyped]]
        def style_index
          @style_index ||= {}
        end

        def initialize
          load_default_styles
        end

        # Adds a new style definition and returns its index
        # @param options [Hash]
        # @return [Integer, Dxf]
        #: (?Hash[Symbol, untyped] options) -> (Integer | Dxf)
        def add_style(options = {})
          options = options.dup
          options[:type] ||= :xf
          raise ArgumentError, "Type must be one of [:xf, :dxf]" unless %i[xf dxf].include?(options[:type])

          if options[:border].is_a?(Hash)
            if options[:border][:edges] == :all
              options[:border][:edges] = Border::EDGES
            elsif options[:border][:edges]
              options[:border][:edges] = options[:border][:edges].map(&:to_sym)
            end
          end

          if options[:type] == :xf
            font_defaults = { name: @fonts.first.name, sz: @fonts.first.sz, family: @fonts.first.family }
            raw_style = { type: :xf }.merge(font_defaults, options)
            raw_style.delete(:num_fmt) if raw_style[:format_code]

            xf_idx = style_index.key(raw_style)
            return xf_idx if xf_idx
          end

          fill = parse_fill_options(options)
          font = parse_font_options(options)
          num_fmt = parse_num_fmt_options(options)
          border = parse_border_options(options)
          alignment = parse_alignment_options(options)
          protection = parse_protection_options(options)

          if options[:type] == :dxf
            dxf = Dxf.new(fill: fill, font: font, numFmt: num_fmt, border: border, alignment: alignment, protection: protection)
            dxfs << dxf
          else
            xf = Xf.new(
              fillId: fill || 0,
              fontId: font || 0,
              numFmtId: num_fmt || 0,
              borderId: border || 0,
              alignment: alignment,
              protection: protection,
              applyFill: !fill.nil?,
              applyFont: !font.nil?,
              applyNumberFormat: !num_fmt.nil?,
              applyBorder: !border.nil?,
              applyAlignment: !alignment.nil?,
              applyProtection: !protection.nil?
            )
            idx = (cellXfs << xf)
            style_index[idx] = raw_style
            idx
          end
        end

        # Returns quotePrefix enabled style index for base_style_index
        # @param base_style_index [Integer]
        # @return [Integer]
        #: (Integer base_style_index) -> Integer
        def quote_prefix_style_for(base_style_index)
          @quote_prefix_cache ||= {}
          @quote_prefix_cache[base_style_index] ||= begin
            old_xf = cellXfs[base_style_index] || cellXfs.first
            new_xf = Xf.new(
              fillId: old_xf.fillId,
              fontId: old_xf.fontId,
              numFmtId: old_xf.numFmtId,
              borderId: old_xf.borderId,
              alignment: old_xf.alignment,
              protection: old_xf.protection,
              applyFill: old_xf.applyFill,
              applyFont: old_xf.applyFont,
              applyNumberFormat: old_xf.applyNumberFormat,
              applyBorder: old_xf.applyBorder,
              applyAlignment: old_xf.applyAlignment,
              applyProtection: old_xf.applyProtection,
              quotePrefix: true
            )
            cellXfs << new_xf
          end
        end

        # Option parsing helpers

        #: (?Hash[Symbol, untyped] options) -> CellProtection?
        def parse_protection_options(options = {})
          return unless options.keys.intersect?(%i[hidden locked])

          CellProtection.new(options)
        end

        #: (?Hash[Symbol, untyped] options) -> CellAlignment?
        def parse_alignment_options(options = {})
          return unless options[:alignment]

          CellAlignment.new(options[:alignment])
        end

        #: (?Hash[Symbol, untyped] options) -> (Font | Integer | nil)
        def parse_font_options(options = {})
          return unless options.keys.intersect?(%i[fg_color sz b i u strike outline shadow charset family font_name])

          base = Caxlsx.instance_values_for(fonts.first).transform_keys(&:to_sym)
          font = Font.new(base.merge(options))
          font.color = Color.new(rgb: options[:fg_color]) if options[:fg_color]
          font.name = options[:font_name] if options[:font_name]
          options[:type] == :dxf ? font : (fonts << font)
        end

        #: (?Hash[Symbol, untyped] options) -> (Fill | Integer | nil)
        def parse_fill_options(options = {})
          return unless options[:bg_color] || options[:pattern_type] || options[:pattern_bg_color] || options[:pattern_fg_color]

          pattern_type = options[:pattern_type] || :solid
          dxf = options[:type] == :dxf

          pattern_options = { patternType: pattern_type }
          bg_color = options[:pattern_bg_color] || options[:bg_color]
          fg_color = options[:pattern_fg_color]

          if bg_color
            pattern_options[:bgColor] = Color.new(rgb: bg_color)
          elsif pattern_type == :solid && fg_color
            pattern_options[:bgColor] = Color.new(rgb: fg_color)
          end

          if fg_color
            pattern_options[:fgColor] = Color.new(rgb: fg_color)
          elsif pattern_type == :solid && bg_color
            pattern_options[:fgColor] = Color.new(rgb: bg_color)
          end

          pattern = PatternFill.new(pattern_options)
          fill = Fill.new(pattern)
          dxf ? fill : (fills << fill)
        end

        #: (?Hash[Symbol, untyped] options) -> (Border | Integer | nil)
        def parse_border_options(options = {})
          return nil if options[:border].nil? && Border::EDGES.all? { |x| options[:"border_#{x}"].nil? }

          if options[:border].is_a?(Integer)
            raise ArgumentError, format(ERR_INVALID_BORDER_ID, options[:border]) if options[:border] >= borders.size

            return options[:type] == :dxf ? borders[options[:border]].clone : options[:border]
          end

          borders_array = []
          base_border_opts = {}
          if options[:border].is_a?(Array)
            borders_array += options[:border]
            options[:border].each do |b_opts|
              base_border_opts = base_border_opts.merge(b_opts) if b_opts[:edges].nil?
            end
          elsif options[:border].is_a?(Hash)
            borders_array << options[:border]
            base_border_opts = options[:border]
          end

          border = Border.new
          Border::EDGES.each do |edge|
            edge_opts = options[:"border_#{edge}"] || {}
            edge_override = borders_array.find do |b_opts|
              b_opts[:edges]&.include?(edge)
            end || {}

            combined = base_border_opts.merge(edge_override).merge(edge_opts)
            next unless combined[:style] || combined[:color]

            pr = border.prs.find { |p| p.name == edge }
            if pr
              pr.style = combined[:style] if combined[:style]
              pr.color = Color.new(rgb: combined[:color]) if combined[:color]
            end
          end

          options[:type] == :dxf ? border : (borders << border)
        end

        #: (?Hash[Symbol, untyped] options) -> (NumFmt | Integer | nil)
        def parse_num_fmt_options(options = {})
          return unless options.keys.intersect?(%i[format_code num_fmt])

          if options[:format_code] || options[:type] == :dxf
            options[:num_fmt] ||= (@numFmts.map(&:numFmtId).max || 163) + 1 if options[:type] != :dxf
            num_fmt = NumFmt.new(numFmtId: options[:num_fmt] || 0, formatCode: options[:format_code].to_s)
            if options[:type] == :dxf
              num_fmt
            else
              numFmts << num_fmt
              num_fmt.numFmtId
            end
          else
            options[:num_fmt]
          end
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << %(<styleSheet xmlns="#{XML_NS}">)
          @numFmts.to_xml_string(str) unless @numFmts.empty?
          @fonts.to_xml_string(str)
          @fills.to_xml_string(str)
          @borders.to_xml_string(str)
          @cellStyleXfs.to_xml_string(str)
          @cellXfs.to_xml_string(str)
          @cellStyles.to_xml_string(str)
          @dxfs.to_xml_string(str) unless @dxfs.empty?
          @tableStyles.to_xml_string(str)
          str << "</styleSheet>"
          str
        end

        # Converts to xlsxrb styles hash
        # @return [Hash]
        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          {
            fonts: @fonts.map(&:to_xlsxrb_hash),
            fills: @fills.map(&:to_xlsxrb_hash),
            borders: @borders.map(&:to_xlsxrb_hash),
            num_fmts: @numFmts.map { |nf| { num_fmt_id: nf.numFmtId, format_code: nf.formatCode } },
            cell_xfs: @cellXfs.map(&:to_xlsxrb_hash),
            xf_entries: @cellXfs.map(&:to_xlsxrb_hash),
            dxfs: @dxfs.map(&:to_xlsxrb_hash)
          }
        end

        private

        def load_default_styles
          @numFmts = SimpleTypedList.new(NumFmt, "numFmts")
          @numFmts << NumFmt.new(numFmtId: NUM_FMT_YYYYMMDD, formatCode: "yyyy/mm/dd")
          @numFmts << NumFmt.new(numFmtId: NUM_FMT_YYYYMMDDHHMMSS, formatCode: "yyyy/mm/dd hh:mm:ss")
          @numFmts.lock

          @fonts = SimpleTypedList.new(Font, "fonts")
          @fonts << Font.new(name: "Arial", sz: 11, family: 1)
          @fonts.lock

          @fills = SimpleTypedList.new(Fill, "fills")
          @fills << Fill.new(PatternFill.new(patternType: :none))
          @fills << Fill.new(PatternFill.new(patternType: :gray125))
          @fills.lock

          @borders = SimpleTypedList.new(Border, "borders")
          @borders << Border.new
          black_border = Border.new
          %i[left right top bottom].each do |edge|
            pr = black_border.prs.find { |p| p.name == edge }
            if pr
              pr.style = :thin
              pr.color = Color.new(rgb: "FF000000")
            end
          end
          @borders << black_border
          @borders.lock

          @cellStyleXfs = SimpleTypedList.new(Xf, "cellStyleXfs")
          @cellStyleXfs << Xf.new(borderId: 0, numFmtId: 0, fontId: 0, fillId: 0)
          @cellStyleXfs.lock

          @cellStyles = SimpleTypedList.new(CellStyle, "cellStyles")
          @cellStyles << CellStyle.new(name: "Normal", builtinId: 0, xfId: 0)
          @cellStyles.lock

          @cellXfs = SimpleTypedList.new(Xf, "cellXfs")
          @cellXfs << Xf.new(borderId: 0, xfId: 0, numFmtId: 0, fontId: 0, fillId: 0)
          @cellXfs << Xf.new(borderId: 1, xfId: 0, numFmtId: 0, fontId: 0, fillId: 0)
          @cellXfs << Xf.new(borderId: 0, xfId: 0, numFmtId: 14, fontId: 0, fillId: 0, applyNumberFormat: 1)
          @cellXfs.lock

          @dxfs = SimpleTypedList.new(Dxf, "dxfs")
          @dxfs.lock

          @tableStyles = TableStyles.new(defaultTableStyle: "TableStyleMedium9", defaultPivotStyle: "PivotStyleLight16")
          @tableStyles.lock
        end
      end
    end
  end
end
