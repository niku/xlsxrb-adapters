# frozen_string_literal: true

# rbs_inline: enabled

require_relative "text"

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a shared strings table.
      class SharedStringsTable
        attr_accessor :strings #: Array[RichText]

        #: () -> void
        def initialize
          @strings = []
          @index_by_content = {}
        end

        # @param index [Integer]
        # @return [RichText, nil]
        #: (Integer index) -> RichText?
        def [](index)
          @strings[index]
        end

        # @return [bool]
        #: () -> bool
        def empty?
          @strings.empty?
        end

        # @return [Integer]
        #: () -> Integer
        def size
          @strings.size
        end

        # @param val [Object]
        # @param index [Integer, nil]
        # @return [Integer]
        #: (untyped val, ?Integer? index) -> Integer
        def add(val, index = nil)
          index ||= @strings.size
          rich = case val
                 when RichText then val
                 when Text then RichText.new(t: val)
                 else RichText.new(t: Text.new(value: val.to_s))
                 end
          @strings[index] = rich
          @index_by_content[val.to_s] = index
          index
        end

        # @param str [String]
        # @param add_if_missing [bool]
        # @return [Integer, nil]
        #: (String str, ?bool add_if_missing) -> Integer?
        def get_index(str, add_if_missing = false)
          idx = @index_by_content[str]
          idx = add(str) if idx.nil? && add_if_missing
          idx
        end
      end
    end
  end
end
