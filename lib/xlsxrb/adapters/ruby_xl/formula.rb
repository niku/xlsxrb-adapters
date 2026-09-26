# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a formula in a cell, compatible with RubyXL::Formula interface.
      class Formula
        attr_accessor :expression #: String?

        # @param expression [String, Hash, nil] The formula expression (e.g. "SUM(A1:B1)").
        #: (?String | Hash[Symbol | String, untyped] | nil expression) -> void
        def initialize(expression = nil)
          @expression = expression.is_a?(Hash) ? expression[:expression] || expression["expression"] : expression
        end

        # Compares formula with another Formula or String.
        #
        # @param other [Object]
        # @return [Boolean]
        #: (untyped other) -> bool
        def ==(other)
          if other.is_a?(Formula)
            expression == other.expression
          elsif other.is_a?(String)
            expression == other
          else
            false
          end
        end
        alias eql? ==

        # Returns hash of the formula expression.
        #
        # @return [Integer]
        #: () -> Integer
        def hash
          expression.hash
        end

        # Returns the formula expression string.
        #
        # @return [String]
        #: () -> String
        def to_s
          expression.to_s
        end
        alias to_str to_s

        # Inspect representation.
        #
        # @return [String]
        #: () -> String
        def inspect
          "#<#{self.class.name} @expression=#{expression.inspect}>"
        end
      end
    end
  end
end
