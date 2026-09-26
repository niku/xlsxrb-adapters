# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a text run.
      class Text
        attr_accessor :value #: String?

        # @param params [Hash]
        #: (?Hash[Symbol, untyped] params) -> void
        def initialize(params = {})
          @value = params[:value]&.to_s
        end

        # @return [String]
        #: () -> String
        def to_s
          @value.to_s
        end

        # Writes XML representation of text node.
        #
        # @return [String]
        #: () -> String
        def write_xml
          "<t>#{@value}</t>"
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(Text)

          @value == other.value
        end
      end

      # Represents rich formatted text.
      class RichText
        attr_accessor :t #: Text?

        # @param params [Hash]
        #: (?Hash[Symbol, untyped] params) -> void
        def initialize(params = {})
          t_val = params[:t]
          @t = t_val.is_a?(Text) ? t_val : Text.new(value: t_val.to_s)
        end

        # @return [String]
        #: () -> String
        def to_s
          @t.to_s
        end

        # Writes XML representation of rich text node.
        #
        # @return [String]
        #: () -> String
        def write_xml
          "<si>#{@t&.write_xml}</si>"
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(RichText)

          @t == other.t
        end
      end
    end
  end
end
