# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a formula in a cell, compatible with RubyXL::Formula interface.
      class Formula
        attr_accessor :expression #: String?
        attr_accessor :t #: (String | Symbol)?
        attr_accessor :ref #: String?
        attr_accessor :si #: Integer?

        alias type t
        alias type= t=
        alias shared_index si
        alias shared_index= si=

        # @param expression [String, Hash, nil] The formula expression (e.g. "SUM(A1:B1)").
        # @param t [String, Symbol, nil] Formula type (:shared, :array, etc.).
        # @param ref [String, nil] Formula range reference.
        # @param si [Integer, nil] Shared formula index.
        #: (?String | Hash[Symbol | String, untyped] | nil expression, ?t: (String | Symbol)?, ?ref: String?, ?si: Integer?) -> void
        # rubocop:disable-next Naming/MethodParameterName
        def initialize(expression = nil, t: nil, ref: nil, si: nil)
          if expression.is_a?(Hash)
            @expression = expression[:expression] || expression["expression"]
            @t = expression[:t] || expression["t"] || expression[:type] || expression["type"]
            @ref = expression[:ref] || expression["ref"]
            raw_si = expression[:si] || expression["si"] || expression[:shared_index] || expression["shared_index"]
            @si = raw_si&.to_i
          else
            @expression = expression
            @t = t
            @ref = ref
            @si = si&.to_i
          end
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
