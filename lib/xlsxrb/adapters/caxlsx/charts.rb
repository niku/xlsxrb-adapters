# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # Scaling properties for an axis.
      class Scaling
        include OptionsParser

        # @return [Integer, nil]
        attr_reader :logBase

        # @return [Symbol, nil]
        attr_reader :orientation

        # @return [Float, nil]
        attr_reader :max

        # @return [Float, nil]
        attr_reader :min

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @orientation = :minMax
          @logBase = @min = @max = nil
          parse_options(options)
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def logBase=(v)
          DataTypeValidator.validate("Scaling.logBase", [Integer], v, ->(arg) { arg.between?(2, 1000) })
          @logBase = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def orientation=(v)
          RestrictionValidator.validate("Scaling.orientation", %i[minMax maxMin], v)
          @orientation = v
        end

        # @param v [Float]
        # @return [Float]
        #: (Float v) -> Float
        def max=(v)
          DataTypeValidator.validate("Scaling.max", Float, v)
          @max = v
        end

        # @param v [Float]
        # @return [Float]
        #: (Float v) -> Float
        def min=(v)
          DataTypeValidator.validate("Scaling.min", Float, v)
          @min = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:scaling>"
          str << '<c:logBase val="' << @logBase.to_s << '"/>' unless @logBase.nil?
          str << '<c:orientation val="' << @orientation.to_s << '"/>' unless @orientation.nil?
          str << '<c:min val="' << @min.to_s << '"/>' unless @min.nil?
          str << '<c:max val="' << @max.to_s << '"/>' unless @max.nil?
          str << "</c:scaling>"
        end
      end

      # Chart title representation.
      class Title
        # @return [String]
        attr_reader :text

        # @return [String]
        attr_reader :text_size

        # @return [Cell, nil]
        attr_reader :cell

        # @param title [String, Cell]
        # @param title_size [String]
        #: (?String | Cell title, ?String title_size) -> void
        def initialize(title = "", title_size = "")
          if title.is_a?(Cell)
            self.cell = title
          else
            self.text = title.to_s
          end
          self.text_size = title_size.to_s.empty? ? "1600" : title_size.to_s
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def text=(v)
          DataTypeValidator.validate("Title.text", String, v)
          @text = v
          @cell = nil
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def text_size=(v)
          DataTypeValidator.validate("Title.text_size", String, v)
          @text_size = v
          @cell = nil
        end

        # @param v [Cell]
        # @return [Cell]
        #: (Cell v) -> Cell
        def cell=(v)
          DataTypeValidator.validate("Title.cell", Cell, v)
          @cell = v
          @text = v.value.to_s
        end

        # @return [Boolean]
        #: () -> bool
        def empty?
          @text.empty? && @cell.nil?
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:title>"
          unless empty?
            clean_value = Caxlsx.trust_input ? @text.to_s : ::CGI.escapeHTML(Caxlsx.sanitize(@text.to_s))
            str << "<c:tx>"
            if @cell.is_a?(Cell)
              str << "<c:strRef>"
              str << "<c:f>" << Caxlsx.cell_range([@cell]) << "</c:f>"
              str << "<c:strCache>"
              str << '<c:ptCount val="1"/>'
              str << '<c:pt idx="0">'
              str << "<c:v>" << clean_value << "</c:v>"
              str << "</c:pt>"
              str << "</c:strCache>"
              str << "</c:strRef>"
            else
              str << "<c:rich>"
              str << "<a:bodyPr/>"
              str << "<a:lstStyle/>"
              str << "<a:p>"
              str << "<a:r>"
              str << '<a:rPr sz="' << @text_size.to_s << '"/>'
              str << "<a:t>" << clean_value << "</a:t>"
              str << "</a:r>"
              str << "</a:p>"
              str << "</c:rich>"
            end
            str << "</c:tx>"
          end
          str << "<c:layout/>"
          str << '<c:overlay val="0"/>'
          str << "</c:title>"
        end
      end

      # Series title representation.
      class SeriesTitle < Title
        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          clean_value = Caxlsx.trust_input ? @text.to_s : ::CGI.escapeHTML(Caxlsx.sanitize(@text.to_s))

          str << "<c:tx>"
          if @cell.is_a?(Cell)
            str << "<c:strRef>"
            str << "<c:f>" << Caxlsx.cell_range([@cell]) << "</c:f>"
            str << "<c:strCache>"
            str << '<c:ptCount val="1"/>'
            str << '<c:pt idx="0">'
            str << "<c:v>" << clean_value << "</c:v>"
            str << "</c:pt>"
            str << "</c:strCache>"
            str << "</c:strRef>"
          else
            str << "<c:v>" << clean_value << "</c:v>"
          end
          str << "</c:tx>"
        end
      end

      # Base axis class.
      class Axis
        include OptionsParser

        VALID_TICK_MARK_VALUES = %i[cross in none out].freeze

        # @return [String, nil]
        attr_accessor :color

        # @return [Integer]
        attr_reader :id
        alias axID id

        # @return [Axis, nil]
        attr_reader :cross_axis
        alias crossAx cross_axis

        # @return [Scaling]
        attr_reader :scaling

        # @return [Symbol]
        attr_reader :ax_pos
        alias axPos ax_pos

        # @return [Symbol]
        attr_reader :tick_lbl_pos
        alias tickLblPos tick_lbl_pos

        # @return [Symbol]
        attr_reader :major_tick_mark
        alias majorTickMark major_tick_mark

        # @return [Symbol]
        attr_reader :minor_tick_mark
        alias minorTickMark minor_tick_mark

        # @return [String]
        attr_reader :format_code

        # @return [Symbol]
        attr_reader :crosses

        # @return [Integer]
        attr_reader :label_rotation

        # @return [Boolean]
        attr_reader :gridlines

        # @return [Integer, Boolean]
        attr_accessor :delete

        # @return [Title, nil]
        attr_reader :title

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @id = rand(8**8)
          @format_code = "General"
          @delete = 0
          @label_rotation = 0
          @scaling = Scaling.new(orientation: :minMax)
          @title = @color = nil
          self.ax_pos = :b
          self.tick_lbl_pos = :nextTo
          self.major_tick_mark = :cross
          self.minor_tick_mark = :none
          self.format_code = "General"
          self.crosses = :autoZero
          self.gridlines = true
          parse_options(options)
        end

        # @param color_rgb [String]
        # @return [String]
        #: (String color_rgb) -> String

        # @param axis [Axis]
        # @return [void]
        #: (Axis axis) -> void
        def cross_axis=(axis)
          DataTypeValidator.validate("#{self.class}.cross_axis", [Axis], axis)
          @cross_axis = axis
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def ax_pos=(v)
          RestrictionValidator.validate("#{self.class}.ax_pos", %i[l r b t], v)
          @ax_pos = v
        end
        alias axPos= ax_pos=

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def tick_lbl_pos=(v)
          RestrictionValidator.validate("#{self.class}.tick_lbl_pos", %i[nextTo high low none], v)
          @tick_lbl_pos = v
        end
        alias tickLblPos= tick_lbl_pos=

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def major_tick_mark=(v)
          RestrictionValidator.validate("#{self.class}.major_tick_mark", VALID_TICK_MARK_VALUES, v)
          @major_tick_mark = v
        end
        alias majorTickMark= major_tick_mark=

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def minor_tick_mark=(v)
          RestrictionValidator.validate("#{self.class}.minor_tick_mark", VALID_TICK_MARK_VALUES, v)
          @minor_tick_mark = v
        end
        alias minorTickMark= minor_tick_mark=

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def format_code=(v)
          Caxlsx.validate_string(v)
          @format_code = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def gridlines=(v)
          Caxlsx.validate_boolean(v)
          @gridlines = v
        end

        # @param v [Boolean, Integer]
        # @return [untyped]
        #: (bool | Integer v) -> untyped

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def crosses=(v)
          RestrictionValidator.validate("#{self.class}.crosses", %i[autoZero min max], v)
          @crosses = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def label_rotation=(v)
          Caxlsx.validate_int(v)
          adjusted = v.to_i * 60_000
          Caxlsx.validate_angle(adjusted)
          @label_rotation = adjusted
        end

        # @param v [String, Cell]
        # @return [void]
        #: (String | Cell v) -> void
        def title=(v)
          DataTypeValidator.validate("#{self.class}.title", [String, Cell], v)
          @title ||= Title.new
          if v.is_a?(String)
            @title.text = v
          elsif v.is_a?(Cell)
            @title.cell = v
          end
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<c:axId val="' << @id.to_s << '"/>'
          @scaling.to_xml_string(str)
          str << '<c:delete val="' << @delete.to_s << '"/>'
          str << '<c:axPos val="' << @ax_pos.to_s << '"/>'
          str << "<c:majorGridlines>"
          str << "<c:spPr><a:ln><a:noFill/></a:ln></c:spPr>" if gridlines == false
          str << "</c:majorGridlines>"
          @title&.to_xml_string(str)
          str << '<c:numFmt formatCode="' << @format_code << '" sourceLinked="' << (@format_code == "General" ? "1" : "0") << '"/>'
          str << '<c:majorTickMark val="' << @major_tick_mark.to_s << '"/>'
          str << '<c:minorTickMark val="' << @minor_tick_mark.to_s << '"/>'
          str << '<c:tickLblPos val="' << @tick_lbl_pos.to_s << '"/>'
          str << '<c:spPr><a:ln><a:solidFill><a:srgbClr val="' << @color << '"/></a:solidFill></a:ln></c:spPr>' if @color
          str << '<c:txPr><a:bodyPr rot="' << @label_rotation.to_s << '"/><a:lstStyle/><a:p><a:pPr><a:defRPr/></a:pPr><a:endParaRPr/></a:p></c:txPr>'
          str << '<c:crossAx val="' << (@cross_axis&.id || 0).to_s << '"/>'
          str << '<c:crosses val="' << @crosses.to_s << '"/>'
        end
      end

      # Category axis.
      class CatAxis < Axis
        LBL_OFFSET_REGEX = /0*(([0-9])|([1-9][0-9])|([1-9][0-9][0-9])|1000)/

        # @return [Integer, Boolean]
        attr_accessor :auto

        # @return [Symbol]
        attr_reader :lbl_algn
        alias lblAlgn lbl_algn

        # @return [Integer, String]
        attr_reader :lbl_offset
        alias lblOffset lbl_offset

        # @return [Integer, nil]
        attr_reader :tick_lbl_skip
        alias tickLblSkip tick_lbl_skip

        # @return [Integer, nil]
        attr_reader :tick_mark_skip
        alias tickMarkSkip tick_mark_skip

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          self.auto = 1
          self.lbl_algn = :ctr
          self.lbl_offset = "100"
          super
        end

        # @param v [Integer, nil]
        # @return [Integer, nil]
        #: (Integer? v) -> Integer?
        def tick_lbl_skip=(v)
          Caxlsx.validate_unsigned_int(v) unless v.nil?
          @tick_lbl_skip = v
        end
        alias tickLblSkip= tick_lbl_skip=

        # @param v [Integer, nil]
        # @return [Integer, nil]
        #: (Integer? v) -> Integer?
        def tick_mark_skip=(v)
          Caxlsx.validate_unsigned_int(v) unless v.nil?
          @tick_mark_skip = v
        end
        alias tickMarkSkip= tick_mark_skip=

        # @param v [Boolean, Integer]
        # @return [untyped]
        #: (bool | Integer v) -> untyped

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def lbl_algn=(v)
          RestrictionValidator.validate("#{self.class}.lbl_algn", %i[ctr l r], v)
          @lbl_algn = v
        end
        alias lblAlgn= lbl_algn=

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def lbl_offset=(v)
          RegexValidator.validate("#{self.class}.lbl_offset", LBL_OFFSET_REGEX, v)
          @lbl_offset = v
        end
        alias lblOffset= lbl_offset=

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:catAx>"
          super
          str << '<c:auto val="' << @auto.to_s << '"/>'
          str << '<c:lblAlgn val="' << @lbl_algn.to_s << '"/>'
          str << '<c:lblOffset val="' << @lbl_offset.to_i.to_s << '"/>'
          str << '<c:tickLblSkip val="' << @tick_lbl_skip.to_s << '"/>' unless @tick_lbl_skip.nil?
          str << '<c:tickMarkSkip val="' << @tick_mark_skip.to_s << '"/>' unless @tick_mark_skip.nil?
          str << "</c:catAx>"
        end
      end

      # Value axis.
      class ValAxis < Axis
        # @return [Symbol]
        attr_reader :cross_between
        alias crossBetween cross_between

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          self.cross_between = :between
          super
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def cross_between=(v)
          RestrictionValidator.validate("ValAxis.cross_between", %i[between midCat], v)
          @cross_between = v
        end
        alias crossBetween= cross_between=

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:valAx>"
          super
          str << '<c:crossBetween val="' << @cross_between.to_s << '"/>'
          str << "</c:valAx>"
        end
      end

      # Series axis.
      class SerAxis < Axis
        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:serAx>"
          super
          str << "</c:serAx>"
        end
      end

      # Axes collection for a chart.
      class Axes
        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          options.each do |name, axis_class|
            add_axis(name, axis_class)
          end
        end

        # @param name [Symbol]
        # @return [Axis, nil]
        #: (Symbol name) -> Axis?
        def [](name)
          axis = axes.assoc(name)
          axis ? axis[1] : nil
        end

        # @param str [String]
        # @param options [Hash{Symbol => untyped}]
        # @return [String]
        #: (?String str, ?Hash[Symbol, untyped] options) -> String
        def to_xml_string(str = +"", options = {})
          if options[:ids]
            axes.each { |axis| str << '<c:axId val="' << axis[1].id.to_s << '"/>' }
          else
            axes.each { |axis| axis[1].to_xml_string(str) }
          end
          str
        end

        # @param name [Symbol]
        # @param axis_class [Class]
        # @return [void]
        #: (Symbol name, Class axis_class) -> void
        def add_axis(name, axis_class)
          axis = axis_class.new
          set_cross_axis(axis)
          axes << [name, axis]
        end

        private

        # @return [Array[Array[untyped]]]
        #: () -> Array[Array[untyped]]
        def axes
          @axes ||= []
        end

        # @param axis [Axis]
        # @return [void]
        #: (Axis axis) -> void
        def set_cross_axis(axis)
          axes.first[1].cross_axis = axis if axes.size == 1
          axis.cross_axis = axes.first[1] unless axes.empty?
        end
      end

      # String value data point.
      class StrVal
        include OptionsParser

        # @return [String]
        attr_reader :v

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @v = ""
          @idx = 0
          parse_options(options)
        end

        # @param v [untyped]
        # @return [String]
        #: (untyped v) -> String
        def v=(v)
          @v = v.to_s
        end

        # @param idx [Integer]
        # @param str [String]
        # @return [void]
        #: (Integer idx, ?String str) -> void
        def to_xml_string(idx, str = +"")
          Caxlsx.validate_unsigned_int(idx)
          return if v.to_s.empty?

          str << '<c:pt idx="' << idx.to_s << '"><c:v>' << ::CGI.escapeHTML(v.to_s) << "</c:v></c:pt>"
        end
      end

      # Numeric value data point.
      class NumVal < StrVal
        # @return [String]
        attr_reader :format_code

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @format_code = "General"
          super
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def format_code=(v)
          Caxlsx.validate_string(v)
          @format_code = v
        end

        # @param idx [Integer]
        # @param str [String]
        # @return [void]
        #: (Integer idx, ?String str) -> void
        def to_xml_string(idx, str = +"")
          Caxlsx.validate_unsigned_int(idx)
          return if v.to_s.empty?

          str << '<c:pt idx="' << idx.to_s << '" formatCode="' << format_code << '"><c:v>' << v.to_s << "</c:v></c:pt>"
        end
      end

      # String data collection.
      class StrData
        include OptionsParser

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @tag_prefix = :str
          @type = StrVal
          @pt = SimpleTypedList.new(@type)
          parse_options(options)
        end

        # @param values [Array[untyped]]
        # @return [void]
        #: (?Array[untyped] values) -> void
        def data=(values = [])
          @tag_name = values.first.is_a?(Cell) ? :strCache : :strLit
          values.each do |val|
            v = val.is_a?(Cell) ? val.value : val
            @pt << @type.new(v: v)
          end
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:" << @tag_name.to_s << ">"
          str << '<c:ptCount val="' << @pt.size.to_s << '"/>'
          @pt.each_with_index do |value, idx|
            value.to_xml_string(idx, str)
          end
          str << "</c:" << @tag_name.to_s << ">"
        end
      end

      # Numeric data collection.
      class NumData
        include OptionsParser

        # @return [String]
        attr_reader :format_code

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @format_code = "General"
          @pt = SimpleTypedList.new(NumVal)
          parse_options(options)
        end

        # @param values [Array[untyped]]
        # @return [void]
        #: (?Array[untyped] values) -> void
        def data=(values = [])
          @tag_name = values.first.is_a?(Cell) ? :numCache : :numLit
          values.each do |val|
            v = if val.is_a?(Cell)
                  val.is_formula? ? 0 : val.value
                else
                  val
                end
            @pt << NumVal.new(v: v)
          end
        end

        # @param v [String]
        # @return [String]
        #: (?String v) -> String
        def format_code=(v = "General")
          Caxlsx.validate_string(v)
          @format_code = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:" << @tag_name.to_s << ">"
          str << "<c:formatCode>" << format_code.to_s << "</c:formatCode>"
          str << '<c:ptCount val="' << @pt.size.to_s << '"/>'
          @pt.each_with_index do |num_val, idx|
            num_val.to_xml_string(idx, str)
          end
          str << "</c:" << @tag_name.to_s << ">"
        end
      end

      # Numeric data source for charts.
      class NumDataSource
        include OptionsParser

        # @return [Symbol]
        attr_reader :tag_name

        # @return [untyped]
        attr_reader :data

        # @return [Array[Symbol]]
        def self.allowed_tag_names
          %i[yVal val bubbleSize]
        end

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @data_type = NumData
          @tag_name = :val
          @ref_tag_name = :numRef
          @f = nil
          @data = @data_type.new(options)
          @f = Caxlsx.cell_range(options[:data]) if options[:data] && options[:data].first.is_a?(Cell)
          parse_options(options)
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def tag_name=(v)
          RestrictionValidator.validate("#{self.class.name}.tag_name", self.class.allowed_tag_names, v)
          @tag_name = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:" << tag_name.to_s << ">"
          if @f
            str << "<c:" << @ref_tag_name.to_s << ">"
            str << "<c:f>" << @f.to_s << "</c:f>"
          end
          @data.to_xml_string(str)
          str << "</c:" << @ref_tag_name.to_s << ">" if @f
          str << "</c:" << tag_name.to_s << ">"
        end
      end

      # Axis data source.
      class AxDataSource < NumDataSource
        # @return [Array[Symbol]]
        def self.allowed_tag_names
          %i[xVal cat]
        end

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @tag_name = :cat
          @data_type = StrData
          @ref_tag_name = :strRef
          super
        end
      end

      # 3D view properties for charts.
      class View3D
        include OptionsParser

        H_PERCENT_REGEX = /0*(([5-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)/
        DEPTH_PERCENT_REGEX = /0*(([2-9][0-9])|([1-9][0-9][0-9])|(1[0-9][0-9][0-9])|2000)/

        # @return [Integer, nil]
        attr_reader :rot_x
        alias rotX rot_x

        # @return [String, nil]
        attr_reader :h_percent
        alias hPercent h_percent

        # @return [Integer, nil]
        attr_reader :rot_y
        alias rotY rot_y

        # @return [String, nil]
        attr_reader :depth_percent
        alias depthPercent depth_percent

        # @return [Boolean, Integer, nil]
        attr_accessor :r_ang_ax
        alias rAngAx r_ang_ax

        # @return [Integer, nil]
        attr_reader :perspective

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @rot_x = @h_percent = @rot_y = @depth_percent = @r_ang_ax = @perspective = nil
          parse_options(options)
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def rot_x=(v)
          RangeValidator.validate("View3D.rot_x", -90, 90, v)
          @rot_x = v
        end
        alias rotX= rot_x=

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def h_percent=(v)
          RegexValidator.validate("View3D.h_percent", H_PERCENT_REGEX, v)
          @h_percent = v
        end
        alias hPercent= h_percent=

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def rot_y=(v)
          RangeValidator.validate("View3D.rot_y", 0, 360, v)
          @rot_y = v
        end
        alias rotY= rot_y=

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def depth_percent=(v)
          RegexValidator.validate("View3D.depth_percent", DEPTH_PERCENT_REGEX, v)
          @depth_percent = v
        end
        alias depthPercent= depth_percent=

        # @param v [Boolean, Integer]
        # @return [untyped]
        #: (bool | Integer v) -> untyped
        alias rAngAx= r_ang_ax=

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def perspective=(v)
          RangeValidator.validate("View3D.perspective", 0, 240, v)
          @perspective = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:view3D>"
          str << '<c:rotX val="' << @rot_x.to_s << '"/>' unless @rot_x.nil?
          str << '<c:hPercent val="' << @h_percent.to_s << '"/>' unless @h_percent.nil?
          str << '<c:rotY val="' << @rot_y.to_s << '"/>' unless @rot_y.nil?
          str << '<c:depthPercent val="' << @depth_percent.to_s << '"/>' unless @depth_percent.nil?
          str << '<c:rAngAx val="' << @r_ang_ax.to_s << '"/>' unless @r_ang_ax.nil?
          str << '<c:perspective val="' << @perspective.to_s << '"/>' unless @perspective.nil?
          str << "</c:view3D>"
        end
      end

      # Data label properties.
      class DLbls
        include Accessors
        include OptionsParser

        boolean_attr_accessor :show_legend_key,
                              :show_val,
                              :show_cat_name,
                              :show_ser_name,
                              :show_percent,
                              :show_bubble_size,
                              :show_leader_lines

        # @return [Class]
        attr_reader :chart_type

        # @param chart_type [Class]
        # @param options [Hash{Symbol => untyped}]
        #: (Class chart_type, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart_type, options = {})
          @chart_type = chart_type
          initialize_defaults
          parse_options(options)
        end

        # @return [void]
        #: () -> void
        def initialize_defaults
          %i[show_legend_key show_val show_cat_name show_ser_name show_percent show_bubble_size show_leader_lines].each do |attr|
            send(:"#{attr}=", false)
          end
        end

        # @return [Symbol, nil]
        #: () -> Symbol?
        def d_lbl_pos
          @d_lbl_pos ||= :bestFit
        end

        # @param label_position [Symbol]
        # @return [Symbol]
        #: (Symbol label_position) -> Symbol
        def d_lbl_pos=(label_position)
          RestrictionValidator.validate("DLbls#d_lbl_pos", %i[bestFit b ctr inBase inEnd l outEnd r t], label_position)
          @d_lbl_pos = label_position
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<c:dLbls>"
          instance_vals = Caxlsx.instance_values_for(self)
          %w[d_lbl_pos show_legend_key show_val show_cat_name show_ser_name show_percent show_bubble_size show_leader_lines].each do |key|
            next unless instance_vals.key?(key) && !instance_vals[key].nil?

            str << "<c:#{Caxlsx.camel(key, false)} val='#{instance_vals[key]}' />"
          end
          str << "</c:dLbls>"
        end
      end

      # Base Series class.
      class Series
        include OptionsParser

        # @return [Chart]
        attr_reader :chart

        # @return [SeriesTitle, nil]
        attr_reader :title

        # @param chart [Chart]
        # @param options [Hash{Symbol => untyped}]
        #: (Chart chart, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart, options = {})
          @order = nil
          self.chart = chart
          @chart.series << self
          parse_options(options)
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @chart.series.index(self) || 0
        end

        # @return [Integer]
        #: () -> Integer
        def order
          @order || index
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def order=(v)
          Caxlsx.validate_unsigned_int(v)
          @order = v
        end

        # @param v [String, Cell, SeriesTitle]
        # @return [void]
        #: (String | Cell | SeriesTitle v) -> void
        def title=(v)
          v = SeriesTitle.new(v) if v.is_a?(String) || v.is_a?(Cell)
          DataTypeValidator.validate("#{self.class}.title", SeriesTitle, v)
          @title = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) ?{ () -> void } -> String
        def to_xml_string(str = +"")
          str << "<c:ser>"
          str << '<c:idx val="' << index.to_s << '"/>'
          str << '<c:order val="' << (order || index).to_s << '"/>'
          title&.to_xml_string(str)
          yield if block_given?
          str << "</c:ser>"
        end

        private

        # @param v [Chart]
        # @return [void]
        #: (Chart v) -> void
        def chart=(v)
          DataTypeValidator.validate("Series.chart", Chart, v)
          @chart = v
        end
      end

      # Bar series.
      class BarSeries < Series
        # @return [NumDataSource, nil]
        attr_reader :data

        # @return [AxDataSource, nil]
        attr_reader :labels

        # @return [Symbol]
        attr_reader :shape

        # @return [Array[String]]
        attr_reader :colors

        # @return [String, nil]
        attr_accessor :series_color

        # @param chart [Chart]
        # @param options [Hash{Symbol => untyped}]
        #: (Chart chart, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart, options = {})
          @shape = :box
          @colors = []
          super
          self.labels = AxDataSource.new({ data: options[:labels] }) unless options[:labels].nil?
          self.data = NumDataSource.new(options) unless options[:data].nil?
        end

        # @param v [Array[String]]
        # @return [void]
        #: (Array[String] v) -> void
        def colors=(v)
          DataTypeValidator.validate("BarSeries.colors", [Array], v)
          @colors = v
        end

        # @param v [String, nil]
        # @return [void]
        #: (String? v) -> void

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def shape=(v)
          RestrictionValidator.validate("BarSeries.shape", %i[cone coneToMax box cylinder pyramid pyramidToMax], v)
          @shape = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            colors.each_with_index do |c, idx|
              str << "<c:dPt>"
              str << '<c:idx val="' << idx.to_s << '"/>'
              str << '<c:spPr><a:solidFill><a:srgbClr val="' << c << '"/></a:solidFill></c:spPr></c:dPt>'
            end

            str << '<c:spPr><a:solidFill><a:srgbClr val="' << series_color << '"/></a:solidFill></c:spPr>' if series_color

            @labels&.to_xml_string(str)
            @data&.to_xml_string(str)
            str << '<c:shape val="' << shape.to_s << '"></c:shape>'
          end
        end

        private

        # @param v [NumDataSource]
        # @return [void]
        #: (NumDataSource v) -> void
        def data=(v)
          DataTypeValidator.validate("Series.data", [NumDataSource], v)
          @data = v
        end

        # @param v [AxDataSource]
        # @return [void]
        #: (AxDataSource v) -> void
        def labels=(v)
          DataTypeValidator.validate("Series.labels", [AxDataSource], v)
          @labels = v
        end
      end

      # Line series.
      class LineSeries < Series
        # @return [NumDataSource, nil]
        attr_reader :data

        # @return [AxDataSource, nil]
        attr_reader :labels

        # @return [String, nil]
        attr_accessor :color

        # @return [Boolean]
        attr_accessor :show_marker

        # @return [Symbol, String]
        attr_accessor :marker_symbol

        # @return [Boolean]
        attr_accessor :smooth

        # @param chart [Chart]
        # @param options [Hash{Symbol => untyped}]
        #: (Chart chart, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart, options = {})
          @show_marker = false
          @marker_symbol = options[:marker_symbol] || :default
          @smooth = false
          @color = nil
          super
          self.labels = AxDataSource.new({ data: options[:labels] }) unless options[:labels].nil?
          self.data = NumDataSource.new(options) unless options[:data].nil?
        end

        # @param v [String]
        # @return [void]
        #: (String v) -> void

        # @param v [Boolean]
        # @return [void]
        #: (bool v) -> void

        # @param v [Symbol, String]
        # @return [void]
        #: (Symbol | String v) -> void

        # @param v [Boolean]
        # @return [void]
        #: (bool v) -> void

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << '<c:spPr><a:ln><a:solidFill><a:srgbClr val="' << color << '"/></a:solidFill></a:ln></c:spPr>' if color
            str << '<c:marker><c:symbol val="' << marker_symbol.to_s << '"/></c:marker>' if show_marker
            @labels&.to_xml_string(str)
            @data&.to_xml_string(str)
            str << '<c:smooth val="' << (@smooth ? "1" : "0") << '"/>'
          end
        end

        private

        # @param v [NumDataSource]
        # @return [void]
        #: (NumDataSource v) -> void
        def data=(v)
          DataTypeValidator.validate("Series.data", [NumDataSource], v)
          @data = v
        end

        # @param v [AxDataSource]
        # @return [void]
        #: (AxDataSource v) -> void
        def labels=(v)
          DataTypeValidator.validate("Series.labels", [AxDataSource], v)
          @labels = v
        end
      end

      # Pie series.
      class PieSeries < Series
        # @return [NumDataSource, nil]
        attr_reader :data

        # @return [AxDataSource, nil]
        attr_reader :labels

        # @return [Integer, nil]
        attr_reader :explosion

        # @return [Array[String]]
        attr_reader :colors

        # @param chart [Chart]
        # @param options [Hash{Symbol => untyped}]
        #: (Chart chart, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart, options = {})
          @explosion = nil
          @colors = []
          super
          self.labels = AxDataSource.new(data: options[:labels]) unless options[:labels].nil?
          self.data = NumDataSource.new(options) unless options[:data].nil?
        end

        # @param v [Array[String]]
        # @return [void]
        #: (Array[String] v) -> void
        def colors=(v)
          DataTypeValidator.validate("PieSeries.colors", [Array], v)
          @colors = v
        end

        # @param v [Integer]
        # @return [void]
        #: (Integer v) -> void
        def explosion=(v)
          Caxlsx.validate_unsigned_int(v)
          @explosion = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << '<c:explosion val="' << @explosion.to_s << '"/>' unless @explosion.nil?
            colors.each_with_index do |c, idx|
              str << "<c:dPt>"
              str << '<c:idx val="' << idx.to_s << '"/>'
              str << '<c:spPr><a:solidFill><a:srgbClr val="' << c << '"/></a:solidFill></c:spPr></c:dPt>'
            end
            @labels&.to_xml_string(str)
            @data&.to_xml_string(str)
          end
        end

        private

        # @param v [NumDataSource]
        # @return [void]
        #: (NumDataSource v) -> void
        def data=(v)
          DataTypeValidator.validate("Series.data", [NumDataSource], v)
          @data = v
        end

        # @param v [AxDataSource]
        # @return [void]
        #: (AxDataSource v) -> void
        def labels=(v)
          DataTypeValidator.validate("Series.labels", [AxDataSource], v)
          @labels = v
        end
      end

      # Scatter series.
      class ScatterSeries < Series
        # @return [AxDataSource, nil]
        attr_accessor :xData

        # @return [NumDataSource, nil]
        attr_accessor :yData

        # @return [String, nil]
        attr_accessor :color

        # @return [String, nil]
        attr_reader :ln_width

        # @return [Boolean]
        attr_accessor :smooth

        # @return [Boolean]
        attr_reader :show_marker

        # @return [Symbol, String]
        attr_reader :marker_symbol

        # @param chart [Chart]
        # @param options [Hash{Symbol => untyped}]
        #: (Chart chart, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart, options = {})
          @xData = @yData = nil
          @color = @ln_width = nil
          @smooth = false
          @show_marker = false
          @marker_symbol = :default
          super
          self.xData = AxDataSource.new(data: options[:xData]) unless options[:xData].nil?
          self.yData = NumDataSource.new(data: options[:yData]) unless options[:yData].nil?
        end

        # @param v [AxDataSource]
        # @return [void]
        #: (AxDataSource v) -> void

        # @param v [NumDataSource]
        # @return [void]
        #: (NumDataSource v) -> void

        # @param v [String]
        # @return [void]
        #: (String v) -> void

        # @param v [Boolean]
        # @return [void]
        #: (bool v) -> void

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << '<c:spPr><a:ln><a:solidFill><a:srgbClr val="' << color << '"/></a:solidFill></a:ln></c:spPr>' if color
            @xData&.to_xml_string(str)
            @yData&.to_xml_string(str)
            str << '<c:smooth val="' << (@smooth ? "1" : "0") << '"/>'
          end
        end
      end

      # Area series.
      class AreaSeries < Series
        # @return [NumDataSource, nil]
        attr_reader :data

        # @return [AxDataSource, nil]
        attr_reader :labels

        # @return [String, nil]
        attr_accessor :color

        # @param chart [Chart]
        # @param options [Hash{Symbol => untyped}]
        #: (Chart chart, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart, options = {})
          @color = nil
          super
          self.labels = AxDataSource.new(data: options[:labels]) unless options[:labels].nil?
          self.data = NumDataSource.new(options) unless options[:data].nil?
        end

        # @param v [String]
        # @return [void]
        #: (String v) -> void

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << '<c:spPr><a:solidFill><a:srgbClr val="' << color << '"/></a:solidFill></c:spPr>' if color
            @labels&.to_xml_string(str)
            @data&.to_xml_string(str)
          end
        end

        private

        # @param v [NumDataSource]
        # @return [void]
        #: (NumDataSource v) -> void
        def data=(v)
          DataTypeValidator.validate("Series.data", [NumDataSource], v)
          @data = v
        end

        # @param v [AxDataSource]
        # @return [void]
        #: (AxDataSource v) -> void
        def labels=(v)
          DataTypeValidator.validate("Series.labels", [AxDataSource], v)
          @labels = v
        end
      end

      # Bubble series.
      class BubbleSeries < Series
        # @return [AxDataSource, nil]
        attr_accessor :xData

        # @return [NumDataSource, nil]
        attr_accessor :yData

        # @return [NumDataSource, nil]
        attr_accessor :bubbleSize

        # @param chart [Chart]
        # @param options [Hash{Symbol => untyped}]
        #: (Chart chart, ?Hash[Symbol, untyped] options) -> void
        def initialize(chart, options = {})
          @xData = @yData = @bubbleSize = nil
          super
          self.xData = AxDataSource.new(data: options[:xData]) unless options[:xData].nil?
          self.yData = NumDataSource.new(data: options[:yData]) unless options[:yData].nil?
          self.bubbleSize = NumDataSource.new(data: options[:bubbleSize], tag_name: :bubbleSize) unless options[:bubbleSize].nil?
        end

        # @param v [AxDataSource]
        # @return [void]
        #: (AxDataSource v) -> void

        # @param v [NumDataSource]
        # @return [void]
        #: (NumDataSource v) -> void

        # @param v [NumDataSource]
        # @return [void]
        #: (NumDataSource v) -> void

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            @xData&.to_xml_string(str)
            @yData&.to_xml_string(str)
            @bubbleSize&.to_xml_string(str)
          end
        end
      end

      # Base Chart class.
      class Chart
        include OptionsParser

        # @return [GraphicFrame]
        attr_reader :graphic_frame

        # @return [SimpleTypedList]
        attr_reader :series

        # @return [Class]
        attr_reader :series_type

        # @return [Boolean, Integer]
        attr_accessor :vary_colors

        # @return [Title]
        attr_reader :title

        # @return [Integer, nil]
        attr_reader :style

        # @return [Boolean]
        attr_reader :show_legend

        # @return [Symbol]
        attr_reader :legend_position

        # @return [Symbol]
        attr_reader :display_blanks_as

        # @return [String, nil]
        attr_reader :bg_color

        # @return [Boolean]
        attr_reader :plot_visible_only

        # @return [Boolean]
        attr_reader :rounded_corners

        # @return [View3D, nil]
        attr_reader :view_3D
        alias view3D view_3D

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @graphic_frame = frame
          @series = SimpleTypedList.new(Series)
          @series_type = Series
          @title = Title.new
          @show_legend = true
          @legend_position = :r
          @display_blanks_as = :gap
          @plot_visible_only = true
          @rounded_corners = true
          @style = nil
          @bg_color = nil
          @view_3D = nil
          @d_lbls = nil
          parse_options(options)
        end

        # @return [DLbls]
        #: () -> DLbls
        def d_lbls
          @d_lbls ||= DLbls.new(self.class)
        end

        # @param v [Boolean, Integer]
        # @return [void]
        #: (bool | Integer v) -> void

        # @return [Relationship]
        #: () -> Relationship
        def relationship
          Relationship.new(self, CHART_R, "../#{pn}")
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @graphic_frame.anchor.drawing.worksheet.workbook.charts.index(self) || 0
        end

        # @return [String]
        #: () -> String
        def pn
          format(CHART_PN, index + 1)
        end

        # @param v [String, Cell]
        # @return [void]
        #: (String | Cell v) -> void
        def title=(v)
          DataTypeValidator.validate("#{self.class}.title", [String, Cell], v)
          if v.is_a?(String)
            @title.text = v
          elsif v.is_a?(Cell)
            @title.cell = v
          end
        end

        # @param v [String]
        # @return [void]
        #: (String v) -> void
        def title_size=(v)
          @title.text_size = v unless v.to_s.empty?
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def show_legend=(v)
          Caxlsx.validate_boolean(v)
          @show_legend = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def display_blanks_as=(v)
          Caxlsx.validate_display_blanks_as(v)
          @display_blanks_as = v
        end

        # @param v [Integer]
        # @return [void]
        #: (Integer v) -> void
        def style=(v)
          DataTypeValidator.validate("Chart.style", Integer, v, ->(arg) { arg.between?(1, 48) })
          @style = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def legend_position=(v)
          RestrictionValidator.validate("Chart.legend_position", %i[b l r t tr], v)
          @legend_position = v
        end

        # @return [Marker]
        #: () -> Marker
        def to
          @graphic_frame.anchor.to
        end

        # @return [Marker]
        #: () -> Marker
        def from
          @graphic_frame.anchor.from
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Series]
        #: (?Hash[Symbol, untyped] options) -> Series
        def add_series(options = {})
          @series_type.new(self, options)
          @series.last
        end

        # @param v [String]
        # @return [void]
        #: (String v) -> void
        def bg_color=(v)
          DataTypeValidator.validate(:color, Color, Color.new(rgb: v))
          @bg_color = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def plot_visible_only=(v)
          Caxlsx.validate_boolean(v)
          @plot_visible_only = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def rounded_corners=(v)
          Caxlsx.validate_boolean(v)
          @rounded_corners = v
        end

        # @param x [Integer, String, Cell, Array[Integer]]
        # @param y [Integer]
        # @return [Marker]
        #: (?Integer | String | Cell | Array[Integer] x, ?Integer y) -> Marker
        def start_at(x = 0, y = 0)
          @graphic_frame.anchor.start_at(x, y)
        end

        # @param x [Integer]
        # @param y [Integer]
        # @return [Marker]
        #: (?Integer x, ?Integer y) -> Marker
        def end_at(x = 10, y = 10)
          @graphic_frame.anchor.end_at(x, y)
        end

        # @param v [View3D]
        # @return [void]
        #: (View3D v) -> void
        def view_3D=(v)
          DataTypeValidator.validate("#{self.class}.view_3D", View3D, v)
          @view_3D = v
        end
        alias view3D= view_3D=

        # @param str [String]
        # @return [String]
        #: (?String str) ?{ () -> void } -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << '<c:chartSpace xmlns:c="' << XML_NS_C << '" xmlns:a="' << XML_NS_A << '" xmlns:r="' << XML_NS_R << '">'
          str << '<c:date1904 val="' << Caxlsx.date1904.to_s << '"/>'
          str << '<c:roundedCorners val="' << rounded_corners.to_s << '"/>'
          str << '<c:style val="' << style.to_s << '"/>'
          str << "<c:chart>"
          @title.to_xml_string(str) unless @title.empty?
          str << '<c:autoTitleDeleted val="' << @title.empty?.to_s << '"/>'
          @view_3D&.to_xml_string(str)
          str << '<c:floor><c:thickness val="0"/></c:floor>'
          str << '<c:sideWall><c:thickness val="0"/></c:sideWall>'
          str << '<c:backWall><c:thickness val="0"/></c:backWall>'
          str << "<c:plotArea>"
          str << "<c:layout/>"
          yield if block_given?
          str << "</c:plotArea>"
          if @show_legend
            str << "<c:legend>"
            str << '<c:legendPos val="' << @legend_position.to_s << '"/>'
            str << "<c:layout/>"
            str << '<c:overlay val="0"/>'
            str << "</c:legend>"
          end
          str << '<c:plotVisOnly val="' << @plot_visible_only.to_s << '"/>'
          str << '<c:dispBlanksAs val="' << display_blanks_as.to_s << '"/>'
          str << '<c:showDLblsOverMax val="1"/>'
          str << "</c:chart>"
          if bg_color
            str << "<c:spPr><a:solidFill>"
            str << '<a:srgbClr val="' << bg_color << '"/>'
            str << "</a:solidFill><a:ln><a:noFill/></a:ln></c:spPr>"
          end
          str << "<c:printSettings><c:headerFooter/>"
          str << '<c:pageMargins b="1.0" l="0.75" r="0.75" t="1.0" header="0.5" footer="0.5"/>'
          str << "<c:pageSetup/></c:printSettings>"
          str << "</c:chartSpace>"
        end

        # Converts chart settings to an options hash compatible with xlsxrb ChartBuilder / Writer.
        #
        # @return [Hash{Symbol => untyped}]
        #: () -> Hash[Symbol, untyped]
        def to_chart_options
          chart_type = case self
                       when Bar3DChart then :bar3d
                       when Line3DChart then :line3d
                       when LineChart then :line
                       when Pie3DChart then :pie3d
                       when PieChart then :pie
                       when ScatterChart then :scatter
                       when AreaChart then :area
                       when BubbleChart then :bubble
                       else :bar
                       end

          opts = {
            type: chart_type,
            title: @title&.text,
            style: @style,
            rounded_corners: @rounded_corners
          }

          anchor = @graphic_frame&.anchor
          if anchor
            opts[:from_col] = anchor.from.col if anchor.respond_to?(:from) && anchor.from
            opts[:from_row] = anchor.from.row if anchor.respond_to?(:from) && anchor.from
            opts[:to_col] = anchor.to.col if anchor.respond_to?(:to) && anchor.to
            opts[:to_row] = anchor.to.row if anchor.respond_to?(:to) && anchor.to
          end

          if @series && !@series.empty?
            opts[:series] = @series.map do |s|
              s_hash = {}
              s_title = s.respond_to?(:title) ? s.title : nil
              s_hash[:title] = s_title.is_a?(SeriesTitle) ? s_title.text : s_title.to_s if s_title
              if s.respond_to?(:xData) && s.respond_to?(:yData)
                s_hash[:cat_ref] = s.xData.is_a?(String) ? s.xData : nil
                s_hash[:cat_data] = s.xData if s.xData.is_a?(Array)
                s_hash[:val_ref] = s.yData.is_a?(String) ? s.yData : nil
                s_hash[:val_data] = s.yData if s.yData.is_a?(Array)
              else
                s_labels = s.respond_to?(:labels) ? s.labels : nil
                s_data = s.respond_to?(:data) ? s.data : nil
                s_hash[:cat_ref] = s_labels.is_a?(String) ? s_labels : nil
                s_hash[:cat_data] = s_labels if s_labels.is_a?(Array)
                s_hash[:val_ref] = s_data.is_a?(String) ? s_data : nil
                s_hash[:val_data] = s_data if s_data.is_a?(Array)
              end
              s_hash
            end
          end

          opts
        end
        alias to_hash to_chart_options
      end

      # Bar chart.
      class BarChart < Chart
        # @return [CatAxis]
        #: () -> CatAxis
        def cat_axis
          axes[:cat_axis] #: untyped
        end
        alias catAxis cat_axis

        # @return [ValAxis]
        #: () -> ValAxis
        def val_axis
          axes[:val_axis] #: untyped
        end
        alias valAxis val_axis

        # @return [Symbol]
        #: () -> Symbol
        def bar_dir
          @bar_dir ||= :bar
        end
        alias barDir bar_dir

        # @return [Integer]
        #: () -> Integer
        def gap_width
          @gap_width ||= 150
        end
        alias gapWidth gap_width

        # @return [Symbol]
        #: () -> Symbol
        def grouping
          @grouping ||= :clustered
        end

        # @return [Integer]
        #: () -> Integer
        def overlap
          @overlap ||= 0
        end

        # @return [Symbol]
        #: () -> Symbol
        def shape
          @shape ||= :box
        end

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = true
          @gap_width = @overlap = @shape = nil
          super
          @series_type = BarSeries
          @d_lbls = nil
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def bar_dir=(v)
          RestrictionValidator.validate("BarChart.bar_dir", %i[bar col], v)
          @bar_dir = v
        end
        alias barDir= bar_dir=

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def grouping=(v)
          RestrictionValidator.validate("BarChart.grouping", %i[percentStacked clustered standard stacked], v)
          @grouping = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def gap_width=(v)
          RangeValidator.validate("BarChart.gap_width", 0, 500, v)
          @gap_width = v
        end
        alias gapWidth= gap_width=

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def overlap=(v)
          RangeValidator.validate("BarChart.overlap", -100, 100, v)
          @overlap = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def shape=(v)
          RestrictionValidator.validate("BarChart.shape", %i[cone coneToMax box cylinder pyramid pyramidToMax], v)
          @shape = v
        end

        # @return [Axes]
        #: () -> Axes
        def axes
          @axes ||= begin
            a = Axes.new(cat_axis: CatAxis, val_axis: ValAxis)
            if bar_dir == :col
              a[:val_axis].ax_pos = :l
            else
              a[:cat_axis].ax_pos = :l
            end
            a
          end
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:barChart>"
            str << '<c:barDir val="' << bar_dir.to_s << '"/>'
            str << '<c:grouping val="' << grouping.to_s << '"/>'
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            @d_lbls&.to_xml_string(str)
            str << '<c:overlap val="' << @overlap.to_s << '"/>' unless @overlap.nil?
            str << '<c:gapWidth val="' << @gap_width.to_s << '"/>' unless @gap_width.nil?
            str << '<c:shape val="' << @shape.to_s << '"/>' unless @shape.nil?
            axes.to_xml_string(str, ids: true)
            str << "</c:barChart>"
            axes.to_xml_string(str)
          end
        end
      end

      # 3D Bar chart.
      class Bar3DChart < Chart
        # @return [CatAxis]
        #: () -> CatAxis
        def cat_axis
          axes[:cat_axis] #: untyped
        end
        alias catAxis cat_axis

        # @return [ValAxis]
        #: () -> ValAxis
        def val_axis
          axes[:val_axis] #: untyped
        end
        alias valAxis val_axis

        # @return [Symbol]
        #: () -> Symbol
        def bar_dir
          @bar_dir ||= :bar
        end
        alias barDir bar_dir

        # @return [Integer, nil]
        attr_reader :gap_depth
        alias gapDepth gap_depth

        # @return [Integer]
        #: () -> Integer
        def gap_width
          @gap_width ||= 150
        end
        alias gapWidth gap_width

        # @return [Symbol]
        #: () -> Symbol
        def grouping
          @grouping ||= :clustered
        end

        # @return [Symbol]
        #: () -> Symbol
        def shape
          @shape ||= :box
        end

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = true
          @gap_width = @gap_depth = @shape = nil
          super
          @series_type = BarSeries
          @view_3D = View3D.new({ r_ang_ax: 1 }.merge(options))
          @d_lbls = nil
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def bar_dir=(v)
          RestrictionValidator.validate("Bar3DChart.bar_dir", %i[bar col], v)
          @bar_dir = v
        end
        alias barDir= bar_dir=

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def grouping=(v)
          RestrictionValidator.validate("Bar3DChart.grouping", %i[percentStacked clustered standard stacked], v)
          @grouping = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def gap_width=(v)
          RangeValidator.validate("Bar3DChart.gap_width", 0, 500, v)
          @gap_width = v
        end
        alias gapWidth= gap_width=

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def gap_depth=(v)
          RangeValidator.validate("Bar3DChart.gap_depth", 0, 500, v)
          @gap_depth = v
        end
        alias gapDepth= gap_depth=

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def shape=(v)
          RestrictionValidator.validate("Bar3DChart.shape", %i[cone coneToMax box cylinder pyramid pyramidToMax], v)
          @shape = v
        end

        # @return [Axes]
        #: () -> Axes
        def axes
          @axes ||= begin
            a = Axes.new(cat_axis: CatAxis, val_axis: ValAxis)
            if bar_dir == :col
              a[:val_axis].ax_pos = :l
            else
              a[:cat_axis].ax_pos = :l
            end
            a
          end
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:bar3DChart>"
            str << '<c:barDir val="' << bar_dir.to_s << '"/>'
            str << '<c:grouping val="' << grouping.to_s << '"/>'
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            @d_lbls&.to_xml_string(str)
            str << '<c:gapWidth val="' << @gap_width.to_s << '"/>' unless @gap_width.nil?
            str << '<c:gapDepth val="' << @gap_depth.to_s << '"/>' unless @gap_depth.nil?
            str << '<c:shape val="' << @shape.to_s << '"/>' unless @shape.nil?
            axes.to_xml_string(str, ids: true)
            str << "</c:bar3DChart>"
            axes.to_xml_string(str)
          end
        end
      end

      # Line chart.
      class LineChart < Chart
        # @return [CatAxis]
        #: () -> CatAxis
        def cat_axis
          axes[:cat_axis] #: untyped
        end
        alias catAxis cat_axis

        # @return [ValAxis]
        #: () -> ValAxis
        def val_axis
          axes[:val_axis] #: untyped
        end
        alias valAxis val_axis

        # @return [Symbol]
        attr_reader :grouping

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = false
          @grouping = :standard
          super
          @series_type = LineSeries
          @d_lbls = nil
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def grouping=(v)
          RestrictionValidator.validate("LineChart.grouping", %i[percentStacked clustered standard stacked], v)
          @grouping = v
        end

        # @return [Axes]
        #: () -> Axes
        def axes
          @axes ||= Axes.new(cat_axis: CatAxis, val_axis: ValAxis)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:lineChart>"
            str << '<c:grouping val="' << grouping.to_s << '"/>'
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            @d_lbls&.to_xml_string(str)
            axes.to_xml_string(str, ids: true)
            str << "</c:lineChart>"
            axes.to_xml_string(str)
          end
        end
      end

      # 3D Line chart.
      class Line3DChart < LineChart
        # @return [String, nil]
        attr_reader :gap_depth
        alias gapDepth gap_depth

        GAP_AMOUNT_PERCENT = /0*(([0-9])|([1-9][0-9])|([1-4][0-9][0-9])|500)%/

        # @return [Axis, nil]
        #: () -> Axis?
        def ser_axis
          axes[:ser_axis]
        end
        alias serAxis ser_axis

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @gap_depth = nil
          @view_3D = View3D.new({ r_ang_ax: 1 }.merge(options))
          super
          axes.add_axis(:ser_axis, SerAxis)
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def gap_depth=(v)
          RegexValidator.validate("Line3DChart.gapWidth", GAP_AMOUNT_PERCENT, v)
          @gap_depth = v
        end
        alias gapDepth= gap_depth=

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << '<c:gapDepth val="' << @gap_depth.to_s << '"/>' unless @gap_depth.nil?
          end
        end
      end

      # Pie chart.
      class PieChart < Chart
        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = true
          super
          @series_type = PieSeries
          @d_lbls = nil
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:pieChart>"
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            d_lbls.to_xml_string(str) if @d_lbls
            str << "</c:pieChart>"
          end
        end
      end

      # 3D Pie chart.
      class Pie3DChart < Chart
        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = true
          super
          @series_type = PieSeries
          @view_3D = View3D.new({ rot_x: 30, perspective: 30 }.merge(options))
          @d_lbls = nil
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:pie3DChart>"
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            d_lbls.to_xml_string(str) if @d_lbls
            str << "</c:pie3DChart>"
          end
        end
      end

      # Scatter chart.
      class ScatterChart < Chart
        include OptionsParser

        # @return [Symbol]
        attr_reader :scatter_style
        alias scatterStyle scatter_style

        # @return [ValAxis]
        #: () -> ValAxis
        def x_val_axis
          axes[:x_val_axis] #: untyped
        end
        alias xValAxis x_val_axis

        # @return [ValAxis]
        #: () -> ValAxis
        def y_val_axis
          axes[:y_val_axis] #: untyped
        end
        alias yValAxis y_val_axis

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = 0
          @scatter_style = :lineMarker
          super
          @series_type = ScatterSeries
          @d_lbls = nil
          parse_options(options)
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def scatter_style=(v)
          Caxlsx.validate_scatter_style(v)
          @scatter_style = v
        end
        alias scatterStyle= scatter_style=

        # @return [Axes]
        #: () -> Axes
        def axes
          @axes ||= Axes.new(x_val_axis: ValAxis, y_val_axis: ValAxis)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:scatterChart>"
            str << '<c:scatterStyle val="' << scatter_style.to_s << '"/>'
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            @d_lbls&.to_xml_string(str)
            axes.to_xml_string(str, ids: true)
            str << "</c:scatterChart>"
            axes.to_xml_string(str)
          end
        end
      end

      # Area chart.
      class AreaChart < Chart
        # @return [CatAxis]
        #: () -> CatAxis
        def cat_axis
          axes[:cat_axis] #: untyped
        end
        alias catAxis cat_axis

        # @return [ValAxis]
        #: () -> ValAxis
        def val_axis
          axes[:val_axis] #: untyped
        end
        alias valAxis val_axis

        # @return [Symbol]
        attr_reader :grouping

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = false
          @grouping = :standard
          super
          @series_type = AreaSeries
          @d_lbls = nil
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def grouping=(v)
          RestrictionValidator.validate("AreaChart.grouping", %i[percentStacked clustered standard stacked], v)
          @grouping = v
        end

        # @return [Axes]
        #: () -> Axes
        def axes
          @axes ||= Axes.new(cat_axis: CatAxis, val_axis: ValAxis)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:areaChart>"
            str << '<c:grouping val="' << grouping.to_s << '"/>'
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            @d_lbls&.to_xml_string(str)
            axes.to_xml_string(str, ids: true)
            str << "</c:areaChart>"
            axes.to_xml_string(str)
          end
        end
      end

      # Bubble chart.
      class BubbleChart < Chart
        # @return [ValAxis]
        #: () -> ValAxis
        def x_val_axis
          axes[:x_val_axis] #: untyped
        end
        alias xValAxis x_val_axis

        # @return [ValAxis]
        #: () -> ValAxis
        def y_val_axis
          axes[:y_val_axis] #: untyped
        end
        alias yValAxis y_val_axis

        # @param frame [GraphicFrame]
        # @param options [Hash{Symbol => untyped}]
        #: (GraphicFrame frame, ?Hash[Symbol, untyped] options) -> void
        def initialize(frame, options = {})
          @vary_colors = 0
          super
          @series_type = BubbleSeries
          @d_lbls = nil
          parse_options(options)
        end

        # @return [Axes]
        #: () -> Axes
        def axes
          @axes ||= Axes.new(x_val_axis: ValAxis, y_val_axis: ValAxis)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          super do
            str << "<c:bubbleChart>"
            str << '<c:varyColors val="' << vary_colors.to_s << '"/>'
            @series.each { |ser| ser.to_xml_string(str) }
            @d_lbls&.to_xml_string(str)
            axes.to_xml_string(str, ids: true)
            str << "</c:bubbleChart>"
            axes.to_xml_string(str)
          end
        end
      end
    end
  end
end
