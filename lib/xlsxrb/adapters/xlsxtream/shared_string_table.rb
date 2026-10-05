# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Xlsxtream
      # Hash-based shared string table keeping track of unique strings and reference counts.
      # @rbs inherits Hash[String, Integer]
      class SharedStringTable < Hash
        # @return [Integer]
        attr_reader :references

        def initialize
          @references = 0
          super { |hash, string| hash[string] = hash.size }
        end

        # Fetches or records a string index and increments the reference counter.
        #
        # @param string [String]
        # @return [Integer]
        def [](string)
          @references += 1
          super
        end
      end
    end
  end
end
