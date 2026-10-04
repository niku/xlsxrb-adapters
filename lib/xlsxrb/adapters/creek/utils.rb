# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Creek
      # Utility helpers matching Creek::Utils for relationship path expansion,
      # file existence, XML parsing, and coordinate math.
      module Utils
        # Expands a target XML path to its corresponding relationships path.
        #
        # @param filepath [String]
        # @return [String]
        #: (String filepath) -> String
        def expand_to_rels_path(filepath)
          filepath.sub(%r{(/[^/]+$)}, '/_rels\1.rels')
        end

        # Checks whether an entry exists in the underlying ZIP package.
        #
        # @param path [String]
        # @return [Boolean]
        #: (String path) -> bool
        def file_exist?(path)
          @book.files.file.exist?(path)
        end

        # Opens and parses an XML entry from the underlying ZIP package.
        #
        # @param xml_path [String]
        # @return [untyped]
        #: (String xml_path) -> untyped
        def parse_xml(xml_path)
          doc = @book.files.file.open(xml_path)
          if defined?(::Nokogiri::XML::Document)
            ::Nokogiri::XML::Document.parse(doc)
          else
            doc.read
          end
        end

        # Converts an Excel column string (e.g. "A", "Z", "AA") into a 0-based column index.
        #
        # @param col [String]
        # @return [Integer]
        #: (String col) -> Integer
        def self.column_to_index(col)
          Xlsxrb::Utils.col_name_to_index(col)
        end

        # Converts a 0-based column index into an Excel column letter string (e.g. 0 => "A", 26 => "AA").
        #
        # @param idx [Integer]
        # @return [String]
        #: (Integer idx) -> String
        def self.index_to_column(idx)
          Xlsxrb::Utils.col_index_to_name(idx)
        end
      end
    end
  end
end
