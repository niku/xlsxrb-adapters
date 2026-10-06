# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module SimpleXlsxReader
      # Represents a cell hyperlink value as a String subclass with #url and #friendly_name.
      # Matches SimpleXlsxReader::Hyperlink.
      class Hyperlink < String
        # Target URL of the hyperlink.
        # @return [String, nil]
        #: String?
        attr_reader :url

        # Visible label or friendly name of the hyperlink.
        # @return [String, nil]
        #: String?
        attr_reader :friendly_name

        # @param url [String] Target URL.
        # @param friendly_name [Object, nil] Optional display name.
        #: (String url, ?untyped friendly_name) -> void
        def initialize(url, friendly_name = nil)
          @url = url
          @friendly_name = friendly_name&.to_s
          super(@friendly_name || @url)
        end
      end
    end
  end
end
