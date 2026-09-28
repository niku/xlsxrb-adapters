# frozen_string_literal: true

# rbs_inline: enabled

require_relative "util"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Defines column attributes
      class Col
        MAX_WIDTH = 255

        include OptionsParser
        include SerializedAttributes

        attr_reader :min #: Integer
        attr_reader :max #: Integer
        attr_reader :best_fit #: bool?
        attr_reader :collapsed #: bool?
        attr_reader :hidden #: bool?
        attr_reader :outline_level #: Integer?
        attr_reader :phonetic #: bool?
        attr_reader :style #: Integer?
        attr_reader :width #: Numeric?
        attr_reader :custom_width #: bool?

        alias bestFit best_fit
        alias customWidth custom_width
        alias outlineLevel outline_level

        serializable_attributes :collapsed, :hidden, :outline_level, :phonetic, :style, :width, :min, :max, :best_fit, :custom_width

        # @param min [Integer]
        # @param max [Integer]
        # @param options [Hash]
        #: (Integer min, Integer max, ?Hash[Symbol, untyped] options) -> void
        def initialize(min, max, options = {})
          Caxlsx.validate_unsigned_int(max)
          Caxlsx.validate_unsigned_int(min)
          @min = min
          @max = max
          @best_fit = nil
          @collapsed = nil
          @hidden = nil
          @outline_level = nil
          @phonetic = nil
          @style = nil
          @width = nil
          @custom_width = nil
          parse_options options
        end

        #: (untyped v) -> void
        def collapsed=(v)
          Caxlsx.validate_boolean(v)
          @collapsed = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def hidden=(v)
          Caxlsx.validate_boolean(v)
          @hidden = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def outline_level=(v)
          Caxlsx.validate_unsigned_numeric(v)
          raise ArgumentError, "outlineLevel must be between 0 and 7" unless v.between?(0, 7)

          @outline_level = v.to_i
        end
        alias outlineLevel= outline_level=

        #: (untyped v) -> void
        def phonetic=(v)
          Caxlsx.validate_boolean(v)
          @phonetic = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def style=(v)
          Caxlsx.validate_unsigned_int(v)
          @style = v.to_i
        end

        #: (untyped v) -> void
        def width=(v)
          @custom_width = @best_fit = !v.nil?
          @width = v.nil? ? nil : [v.to_f, MAX_WIDTH].min
        end

        # Updates column width based on cell autowidth and fixed width settings
        # @param cell [Cell]
        # @param fixed_width [Numeric, Symbol, nil]
        # @param use_autowidth [Boolean]
        #: (untyped cell, ?untyped fixed_width, ?bool use_autowidth) -> void
        def update_width(cell, fixed_width = nil, use_autowidth = true)
          return if @ignore_cell_width

          cell_width = case fixed_width
                       when Numeric
                         fixed_width
                       when nil, :auto
                         cell.autowidth if use_autowidth
                       when :ignore
                         nil
                       else
                         raise ArgumentError, "fixed_with must be a Numeric, :auto, :ignore or nil, but is '#{fixed_width.inspect}'"
                       end

          return unless cell_width && cell_width > (@width || 0)

          @custom_width = @best_fit = true
          @width = [cell_width.to_f, MAX_WIDTH].min
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("col", str)
        end
      end

      # Collection of column records
      class Cols < SimpleTypedList
        # @param worksheet [Worksheet]
        #: (untyped worksheet) -> void
        def initialize(worksheet)
          super(Col)
          @worksheet = worksheet
        end

        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<cols>"
          each { |item| item.to_xml_string(str) }
          str << "</cols>"
          str
        end
      end
    end
  end
end
