# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Xlsxtream
      # Generates OpenXML <cols> element from column definitions.
      class Columns
        # Pass an Array of column options Hashes. Symbol Hash keys:
        # +width_chars+:: Approximate column width in characters.
        # +width_pixels+:: Exact width of column in pixels (overrides +width_chars+).
        #
        # @param column_options_array [Array<Hash{Symbol => untyped}>]
        #: (Array[Hash[Symbol, untyped]] column_options_array) -> void
        def initialize(column_options_array)
          @columns = column_options_array
        end

        # Renders the <cols>...</cols> XML chunk.
        #
        # @return [String]
        #: () -> String
        def to_xml
          xml = String.new("<cols>")

          @columns.each_with_index do |column, index|
            width_chars  = column[:width_chars]
            width_pixels = column[:width_pixels]

            if width_chars.nil? && width_pixels.nil?
              xml << %(<col min="#{index + 1}" max="#{index + 1}"/>)
            else
              width_pixels ||= ((((width_chars * 7.0) + 5) / 7) * 256).truncate / 256.0
              xml << %(<col min="#{index + 1}" max="#{index + 1}" width="#{width_pixels}" customWidth="1"/>)
            end
          end

          xml << "</cols>"
        end
      end
    end
  end
end
