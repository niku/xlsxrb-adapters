# frozen_string_literal: true

# rbs_inline: enabled

require_relative "reference"

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a merged cell range.
      class MergedCell
        attr_accessor :ref #: Reference

        # @param ref [Reference, String]
        #: (?ref: (Reference | String)) -> void
        def initialize(ref: Reference.new(0, 0))
          @ref = ref.is_a?(Reference) ? ref : Reference.new(ref.to_s)
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(MergedCell)

          @ref == other.ref
        end

        # @return [String]
        #: () -> String
        def to_s
          @ref.to_s
        end
      end

      # Collection of merged cells in a worksheet.
      # @rbs inherits Array[MergedCell]
      class MergedCells < Array
        # @param other_array [Array<MergedCell>]
        #: (?Array[MergedCell] other_array) -> void
        def initialize(other_array = [])
          super()
          other_array.each { |mc| self << mc }
        end
      end
    end
  end
end
