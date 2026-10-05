# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Xlsxtream
      # XML utility helpers for character escaping and formatting.
      module XML
        XML_ESCAPES = {
          "&" => "&amp;",
          '"' => "&quot;",
          "<" => "&lt;",
          ">" => "&gt;"
        }.freeze

        HEX_ESCAPE_REGEXP = /_(x[0-9A-Fa-f]{4}_)/
        XML_ESCAPE_UNDERSCORE = '_x005f_\1'

        XML_DECLARATION = "<?xml version=\"1.0\" encoding=\"UTF-8\" standalone=\"yes\"?>\r\n"

        WS_AROUND_TAGS = /(?<=>)\s+|\s+(?=<)/

        UNSAFE_ATTR_CHARS = /[&"<>]/
        UNSAFE_VALUE_CHARS = /[&<>]/

        INVALID_XML10_CHARS = /[^\x09\x0A\x0D\x20-\uD7FF\uE000-\uFFFD\u{10000}-\u{10FFFF}]/

        ESCAPE_CHAR = ->(c) { format("_x%04X_", c.ord).freeze }

        class << self
          # Returns the XML declaration header.
          #
          # @return [String]
          #: () -> String
          def header
            XML_DECLARATION
          end

          # Strips whitespace around XML tags.
          #
          # @param xml [String]
          # @return [String]
          #: (String xml) -> String
          def strip(xml)
            xml.gsub(WS_AROUND_TAGS, "")
          end

          # Escapes characters for XML attribute values.
          #
          # @param string [String]
          # @return [String]
          #: (String string) -> String
          def escape_attr(string)
            string.gsub(UNSAFE_ATTR_CHARS, XML_ESCAPES)
          end

          # Escapes characters for XML element values.
          #
          # @param string [String]
          # @return [String]
          #: (String string) -> String
          def escape_value(string)
            string
              .gsub(UNSAFE_VALUE_CHARS, XML_ESCAPES)
              .gsub(HEX_ESCAPE_REGEXP, XML_ESCAPE_UNDERSCORE)
              .gsub(INVALID_XML10_CHARS) { |c| ESCAPE_CHAR.call(c) }
          end
        end
      end
    end
  end
end
