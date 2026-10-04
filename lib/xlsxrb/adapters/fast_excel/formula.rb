# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module FastExcel
      # Represents an Excel formula string expression compatible with FastExcel::Formula
      class Formula
        # @return [String]
        attr_accessor :fml

        # @param fml [String] Formula expression (e.g. "SUM(A1:B2)")
        #: (String fml) -> void
        def initialize(fml)
          @fml = fml
        end

        # @return [String]
        #: () -> String
        def to_s
          @fml.to_s
        end

        # @param other [Object]
        # @return [Boolean]
        #: (untyped other) -> bool
        def ==(other)
          other.is_a?(Formula) && other.fml == @fml
        end
        alias eql? ==

        # @return [Integer]
        #: () -> Integer
        def hash
          [self.class, @fml].hash
        end
      end
    end
  end
end
