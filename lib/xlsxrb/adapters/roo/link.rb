# frozen_string_literal: true

# rbs_inline: enabled

require "uri"

module Xlsxrb
  module Adapters
    module Roo
      # String subclass representing a hyperlinked cell value matching Roo::Link.
      class Link < String
        attr_reader :href

        alias url href

        # @param href [String] Target URL.
        # @param text [String] Display text.
        #: (?String href, ?String text) -> void
        def initialize(href = "", text = href)
          super(text)
          @href = href
        end

        # Parses href as a URI.
        #: () -> URI::Generic
        def to_uri
          URI.parse(@href)
        end
      end
    end
  end
end
