# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # Conditional Format Value Object.
      class Cfvo
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :type, :gte, :val

        # @return [Symbol, nil]
        attr_reader :type

        # @return [Boolean, nil]
        attr_reader :gte

        # @return [String, nil]
        attr_reader :val

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @gte = true
          parse_options(options)
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def type=(v)
          Caxlsx.validate_conditional_formatting_value_object_type(v)
          @type = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def gte=(v)
          Caxlsx.validate_boolean(v)
          @gte = v
        end

        # @param v [untyped]
        # @return [String]
        #: (untyped v) -> String
        def val=(v)
          raise ArgumentError, "#{v.inspect} must respond to to_s" unless v.respond_to?(:to_s)

          @val = v.to_s
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("cfvo", str)
        end
      end

      # Collection of Cfvo objects.
      class Cfvos < SimpleTypedList
        #: () -> void
        def initialize
          super(Cfvo)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          each { |cfvo| cfvo.to_xml_string(str) }
          str
        end
      end

      # ColorScale conditional formatting.
      class ColorScale
        class << self
          # @return [Array[Hash{Symbol => untyped}]]
          #: () -> Array[Hash[Symbol, untyped]]
          def default_cfvos
            [{ type: :min, val: 0, color: "FFFF7128" },
             { type: :max, val: 0, color: "FFFFEF9C" }]
          end

          # @return [ColorScale]
          #: () -> ColorScale
          def two_tone
            new
          end

          # @return [ColorScale]
          #: () -> ColorScale
          def three_tone
            new({ type: :min, val: 0, color: "FFF8696B" },
                { type: :percent, val: "50", color: "FFFFEB84" },
                { type: :max, val: 0, color: "FF63BE7B" })
          end
        end

        # @return [Cfvos]
        #: () -> Cfvos
        def value_objects
          @value_objects ||= Cfvos.new
        end

        # @return [SimpleTypedList]
        #: () -> SimpleTypedList
        def colors
          @colors ||= SimpleTypedList.new(Color)
        end

        # @param cfvos [Array[Hash{Symbol => untyped}]]
        #: (*Hash[Symbol, untyped] cfvos) ?{ (ColorScale) -> void } -> void
        def initialize(*cfvos)
          initialize_default_cfvos(cfvos)
          yield self if block_given?
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> Hash[Symbol, untyped]
        def add(options = {})
          value_objects << Cfvo.new(type: options[:type] || :min, val: options[:val] || 0)
          colors << Color.new(rgb: options[:color] || "FF000000")
          { cfvo: value_objects.last, color: colors.last }
        end

        # @param index [Integer]
        # @return [void]
        #: (?Integer index) -> void
        def delete_at(index = 2)
          value_objects.delete_at(index)
          colors.delete_at(index)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<colorScale>"
          value_objects.to_xml_string(str)
          colors.each { |color| color.to_xml_string(str) }
          str << "</colorScale>"
        end

        private

        # @param user_cfvos [Array[Hash{Symbol => untyped}]]
        # @return [void]
        #: (Array[Hash[Symbol, untyped]] user_cfvos) -> void
        def initialize_default_cfvos(user_cfvos)
          defaults = self.class.default_cfvos
          user_cfvos.each_with_index do |cfvo, idx|
            cfvo = defaults[idx].merge(cfvo) if idx < defaults.size
            add(cfvo)
          end
          add(defaults[colors.size - 1]) while colors.size < defaults.size
        end
      end

      # DataBar conditional formatting.
      class DataBar
        include OptionsParser
        include SerializedAttributes

        class << self
          # @return [Array[Hash{Symbol => untyped}]]
          #: () -> Array[Hash[Symbol, untyped]]
          def default_cfvos
            [{ type: :min, val: "0" },
             { type: :max, val: "0" }]
          end
        end

        serializable_attributes :min_length, :max_length, :show_value

        CHILD_ELEMENTS = %i[value_objects color].freeze

        # @return [Integer]
        attr_reader :min_length
        alias minLength min_length

        # @return [Integer]
        attr_reader :max_length
        alias maxLength max_length

        # @return [Boolean]
        attr_reader :show_value
        alias showValue show_value

        # @param options [Hash{Symbol => untyped}]
        # @param cfvos [Array[Hash{Symbol => untyped}]]
        #: (?Hash[Symbol, untyped] options, *Hash[Symbol, untyped] cfvos) ?{ (DataBar) -> void } -> void
        def initialize(options = {}, *cfvos)
          @min_length = 10
          @max_length = 90
          @show_value = true
          parse_options(options)
          initialize_cfvos(cfvos)
          yield self if block_given?
        end

        # @return [Cfvos]
        #: () -> Cfvos
        def value_objects
          @value_objects ||= Cfvos.new
        end

        # @return [Color]
        #: () -> Color
        def color
          @color ||= Color.new(rgb: "FF0000FF")
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def min_length=(v)
          Caxlsx.validate_unsigned_int(v)
          @min_length = v
        end
        alias minLength= min_length=

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def max_length=(v)
          Caxlsx.validate_unsigned_int(v)
          @max_length = v
        end
        alias maxLength= max_length=

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def show_value=(v)
          Caxlsx.validate_boolean(v)
          @show_value = v
        end
        alias showValue= show_value=

        # @param v [Color, String]
        # @return [void]
        #: (Color | String v) -> void
        def color=(v)
          @color = v if v.is_a?(Color)
          color.rgb = v if v.is_a?(String)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("dataBar", str) do
            value_objects.to_xml_string(str)
            color.to_xml_string(str)
          end
        end

        private

        # @param cfvos [Array[Hash{Symbol => untyped}]]
        # @return [void]
        #: (Array[Hash[Symbol, untyped]] cfvos) -> void
        def initialize_cfvos(cfvos)
          self.class.default_cfvos.each_with_index do |default, idx|
            value_objects << if idx < cfvos.size
                               Cfvo.new(default.merge(cfvos[idx]))
                             else
                               Cfvo.new(default)
                             end
          end
        end
      end

      # IconSet conditional formatting.
      class IconSet
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :iconSet, :percent, :reverse, :showValue

        # @return [String]
        attr_reader :iconSet

        # @return [Boolean]
        attr_reader :percent

        # @return [Boolean]
        attr_reader :reverse

        # @return [Boolean]
        attr_reader :showValue

        # @return [Array[Integer]]
        attr_reader :interpolationPoints

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) ?{ (IconSet) -> void } -> void
        def initialize(options = {})
          @percent = @showValue = true
          @reverse = false
          @iconSet = "3TrafficLights1"
          @interpolationPoints = [0, 33, 67]
          parse_options(options)
          yield self if block_given?
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def iconSet=(v)
          Caxlsx.validate_icon_set(v)
          @iconSet = v
        end

        # @param v [Array[Integer]]
        # @return [Array[Integer]]
        #: (Array[Integer] v) -> Array[Integer]
        def interpolationPoints=(v)
          v.each { |point| Caxlsx.validate_int(point) }
          @value_objects = nil
          @interpolationPoints = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def showValue=(v)
          Caxlsx.validate_boolean(v)
          @showValue = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def percent=(v)
          Caxlsx.validate_boolean(v)
          @percent = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def reverse=(v)
          Caxlsx.validate_boolean(v)
          @reverse = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          initialize_value_objects if @value_objects.nil?

          serialized_tag("iconSet", str) do
            @value_objects&.each { |cfvo| cfvo.to_xml_string(str) }
          end
        end

        private

        # @return [void]
        #: () -> void
        def initialize_value_objects
          @value_objects = SimpleTypedList.new(Cfvo)
          @interpolationPoints.each { |point| @value_objects << Cfvo.new(type: :percent, val: point) }
          @value_objects.lock
        end
      end

      # A rule for conditional formatting.
      class ConditionalFormattingRule
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :type, :aboveAverage, :bottom, :dxfId, :equalAverage,
                                :priority, :operator, :text, :percent, :rank, :stdDev,
                                :stopIfTrue, :timePeriod

        # @return [Array[String], nil]
        attr_reader :formula

        # @return [Symbol, nil]
        attr_reader :type

        # @return [Boolean, nil]
        attr_reader :aboveAverage

        # @return [Boolean, nil]
        attr_reader :bottom

        # @return [Integer, nil]
        attr_reader :dxfId

        # @return [Boolean, nil]
        attr_reader :equalAverage

        # @return [Symbol, nil]
        attr_reader :operator

        # @return [Integer, nil]
        attr_reader :priority

        # @return [String, nil]
        attr_reader :text

        # @return [Boolean, nil]
        attr_reader :percent

        # @return [Integer, nil]
        attr_reader :rank

        # @return [Integer, nil]
        attr_reader :stdDev

        # @return [Boolean, nil]
        attr_reader :stopIfTrue

        # @return [Symbol, nil]
        attr_reader :timePeriod

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @color_scale = @data_bar = @icon_set = @formula = nil
          parse_options(options)
        end

        # @return [ColorScale]
        #: () -> ColorScale
        def color_scale
          @color_scale ||= ColorScale.new
        end

        # @return [DataBar]
        #: () -> DataBar
        def data_bar
          @data_bar ||= DataBar.new
        end

        # @return [IconSet]
        #: () -> IconSet
        def icon_set
          @icon_set ||= IconSet.new
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def type=(v)
          Caxlsx.validate_conditional_formatting_type(v)
          @type = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def aboveAverage=(v)
          Caxlsx.validate_boolean(v)
          @aboveAverage = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def bottom=(v)
          Caxlsx.validate_boolean(v)
          @bottom = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def dxfId=(v)
          Caxlsx.validate_unsigned_numeric(v)
          @dxfId = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def equalAverage=(v)
          Caxlsx.validate_boolean(v)
          @equalAverage = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def priority=(v)
          Caxlsx.validate_unsigned_numeric(v)
          @priority = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def operator=(v)
          Caxlsx.validate_conditional_formatting_operator(v)
          @operator = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def text=(v)
          Caxlsx.validate_string(v)
          @text = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def percent=(v)
          Caxlsx.validate_boolean(v)
          @percent = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def rank=(v)
          Caxlsx.validate_unsigned_numeric(v)
          @rank = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def stdDev=(v)
          Caxlsx.validate_unsigned_numeric(v)
          @stdDev = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def stopIfTrue=(v)
          Caxlsx.validate_boolean(v)
          @stopIfTrue = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def timePeriod=(v)
          Caxlsx.validate_time_period_type(v)
          @timePeriod = v
        end

        # @param v [String, Array[String]]
        # @return [Array[String]]
        #: (String | Array[String] v) -> Array[String]
        def formula=(v)
          [*v].each { |x| Caxlsx.validate_string(x) }
          @formula = [*v].map { |form| ::CGI.escapeHTML(form) }
        end

        # @param v [ColorScale]
        # @return [void]
        #: (ColorScale v) -> void
        def color_scale=(v)
          DataTypeValidator.validate("conditional_formatting_rule.color_scale", ColorScale, v)
          @color_scale = v
        end

        # @param v [DataBar]
        # @return [void]
        #: (DataBar v) -> void
        def data_bar=(v)
          DataTypeValidator.validate("conditional_formatting_rule.data_bar", DataBar, v)
          @data_bar = v
        end

        # @param v [IconSet]
        # @return [void]
        #: (IconSet v) -> void
        def icon_set=(v)
          DataTypeValidator.validate("conditional_formatting_rule.icon_set", IconSet, v)
          @icon_set = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<cfRule "
          serialized_attributes(str)
          str << ">"
          str << "<formula>" << [*formula].join("</formula><formula>") << "</formula>" if @formula
          @color_scale&.to_xml_string(str) if @color_scale && @type == :colorScale
          @data_bar&.to_xml_string(str) if @data_bar && @type == :dataBar
          @icon_set&.to_xml_string(str) if @icon_set && @type == :iconSet
          str << "</cfRule>"
        end
      end

      # Conditional formatting element applying rules to a range.
      class ConditionalFormatting
        include OptionsParser

        # @return [String, nil]
        attr_reader :sqref

        # @return [Array[ConditionalFormattingRule]]
        attr_accessor :rules

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @rules = []
          parse_options(options)
        end

        # @param rules [Array[untyped], Hash{Symbol => untyped}]
        # @return [void]
        #: (Array[untyped] | Hash[Symbol, untyped] rules) -> void
        def add_rules(rules)
          rules = [rules] if rules.is_a?(Hash)
          rules.each do |rule|
            add_rule(rule)
          end
        end

        # @param rule [ConditionalFormattingRule, Hash{Symbol => untyped}]
        # @return [void]
        #: (ConditionalFormattingRule | Hash[Symbol, untyped] rule) -> void
        def add_rule(rule)
          if rule.is_a?(ConditionalFormattingRule)
            @rules << rule
          elsif rule.is_a?(Hash)
            @rules << ConditionalFormattingRule.new(rule)
          end
        end

        # @param v [Array[ConditionalFormattingRule]]
        # @return [Array[ConditionalFormattingRule]]
        #: (Array[ConditionalFormattingRule] v) -> Array[ConditionalFormattingRule]

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def sqref=(v)
          Caxlsx.validate_string(v)
          @sqref = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<conditionalFormatting sqref="' << sqref.to_s << '">'
          rules.each_with_index do |rule, idx|
            str << " " unless idx.zero?
            rule.to_xml_string(str)
          end
          str << "</conditionalFormatting>"
        end
      end

      # Collection of ConditionalFormatting items for a worksheet.
      class ConditionalFormattings < SimpleTypedList
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          super(ConditionalFormatting)
          @worksheet = worksheet
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          each { |item| item.to_xml_string(str) }
          str
        end
      end
    end
  end
end
