# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a column definition / range compatible with RubyXL::ColumnRange.
      class ColumnRange
        DEFAULT_WIDTH = 8
        MAX_DIGIT_WIDTH = 7

        attr_accessor :min #: Integer
        attr_accessor :max #: Integer
        attr_accessor :width #: Float?
        attr_accessor :style_index #: Integer
        attr_accessor :hidden #: bool
        attr_accessor :custom_width #: bool
        attr_accessor :outline_level #: Integer
        attr_accessor :collapsed #: bool

        # Converts character width to raw OOXML column width.
        #
        # @param width_in_chars [Numeric]
        # @return [Float]
        #: (Numeric width_in_chars) -> Float
        def self.chars2raw(width_in_chars)
          ((width_in_chars + (5.0 / MAX_DIGIT_WIDTH)) * 256).to_i / 256.0
        end

        # @param min [Integer]
        # @param max [Integer]
        # @param width [Float, nil]
        # @param style_index [Integer]
        # @param hidden [Boolean]
        # @param custom_width [Boolean]
        # @param outline_level [Integer]
        # @param collapsed [Boolean]
        #: (?min: Integer, ?max: Integer, ?width: Float?, ?style_index: Integer, ?hidden: bool, ?custom_width: bool, ?outline_level: Integer, ?collapsed: bool) -> void
        def initialize(min: 1, max: 1, width: nil, style_index: 0, hidden: false, custom_width: false, outline_level: 0, collapsed: false)
          @min = min
          @max = max
          @width = width
          @style_index = style_index
          @hidden = hidden
          @custom_width = custom_width
          @outline_level = outline_level
          @collapsed = collapsed
        end

        # @return [ColumnRange]
        #: () -> ColumnRange
        def dup
          self.class.new(
            min: @min,
            max: @max,
            width: @width,
            style_index: @style_index,
            hidden: @hidden,
            custom_width: @custom_width,
            outline_level: @outline_level,
            collapsed: @collapsed
          )
        end

        # @param col_index [Integer]
        #: (Integer col_index) -> void
        def delete_column(col_index)
          col = col_index + 1
          self.min -= 1 if min >= col
          self.max -= 1 if max >= col
        end

        # @param col_index [Integer]
        #: (Integer col_index) -> void
        def insert_column(col_index)
          col = col_index + 1
          self.min += 1 if min >= col
          self.max += 1 if max >= col - 1
        end

        # @param col_index [Integer]
        # @return [bool]
        #: (Integer col_index) -> bool
        def include?(col_index)
          ((min - 1)..(max - 1)).include?(col_index)
        end
      end

      # Collection of column ranges compatible with RubyXL::ColumnRanges.
      # @rbs inherits Array[ColumnRange]
      class ColumnRanges < Array
        # @param other_array [Array<ColumnRange>]
        #: (?Array[ColumnRange] other_array) -> void
        def initialize(other_array = [])
          super()
          other_array.each { |cr| self << cr }
        end

        # @param col_index [Integer]
        # @return [ColumnRange]
        #: (Integer col_index) -> ColumnRange
        def get_range(col_index)
          col_num = col_index + 1
          old_range = locate_range(col_index)

          if old_range.nil?
            new_range = ColumnRange.new(width: ColumnRange.chars2raw(ColumnRange::DEFAULT_WIDTH))
          elsif old_range.min == col_num && old_range.max == col_num
            return old_range
          elsif old_range.min == col_num
            new_range = old_range.dup
            old_range.min += 1
          elsif old_range.max == col_num
            new_range = old_range.dup
            old_range.max -= 1
          else
            prior_range = old_range.dup
            prior_range.max = col_index
            self << prior_range

            old_range.min = col_num + 1
            new_range = ColumnRange.new
          end

          new_range.min = new_range.max = col_num
          self << new_range
          new_range
        end

        # @param col_index [Integer]
        # @return [ColumnRange, nil]
        #: (Integer col_index) -> ColumnRange?
        def locate_range(col_index)
          find { |range| range.include?(col_index) }
        end

        # @param col_index [Integer]
        #: (Integer col_index) -> void
        def insert_column(col_index)
          each { |range| range.insert_column(col_index) }
        end
      end
    end
  end
end
