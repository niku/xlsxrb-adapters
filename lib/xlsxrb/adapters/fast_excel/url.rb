# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module FastExcel
      # Represents a hyperlink URL compatible with FastExcel::URL
      class URL
        # @return [String]
        attr_accessor :url

        # @param url [String] Hyperlink destination URL (e.g. "https://github.com")
        #: (String url) -> void
        def initialize(url)
          @url = url
        end

        # @return [String]
        #: () -> String
        def to_s
          @url.to_s
        end

        # @param other [Object]
        # @return [Boolean]
        #: (untyped other) -> bool
        def ==(other)
          other.is_a?(URL) && other.url == @url
        end
        alias eql? ==

        # @return [Integer]
        #: () -> Integer
        def hash
          [self.class, @url].hash
        end
      end
    end
  end
end
