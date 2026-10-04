# frozen_string_literal: true

# rbs_inline: enabled

require_relative "styles"

module Xlsxrb
  module Adapters
    module Creek
      # Shared strings table parser matching Creek::SharedStrings.
      class SharedStrings
        SPREADSHEETML_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

        # @return [untyped]
        attr_reader :book

        # @return [Hash<Integer, String>]
        attr_reader :dictionary

        # @param book [untyped]
        #: (untyped book) -> void
        def initialize(book)
          @book = book
          @dictionary = {}
          parse_shared_shared_strings
        end

        # Parses shared strings from xl/sharedStrings.xml in the book's archive.
        #
        # @return [Hash<Integer, String>, nil]
        #: () -> Hash[Integer, String]?
        def parse_shared_shared_strings
          path = "xl/sharedStrings.xml"
          return unless @book.files.file.exist?(path)

          doc = @book.files.file.open(path)
          xml = if defined?(::Nokogiri::XML::Document)
                  ::Nokogiri::XML::Document.parse(doc)
                else
                  doc
                end
          parse_shared_string_from_document(xml)
        end

        # Parses shared strings from the provided XML document and updates dictionary.
        #
        # @param xml [untyped]
        # @return [Hash<Integer, String>]
        #: (untyped xml) -> Hash[Integer, String]
        def parse_shared_string_from_document(xml)
          @dictionary = self.class.parse_shared_string_from_document(xml)
        end

        # Parses shared strings dictionary from a Nokogiri document or XML string/IO.
        #
        # @param xml [untyped]
        # @return [Hash<Integer, String>]
        #: (untyped xml) -> Hash[Integer, String]
        def self.parse_shared_string_from_document(xml)
          if xml.respond_to?(:css)
            parse_with_nokogiri(xml)
          else
            raw_str = if xml.respond_to?(:read)
                        xml.rewind if xml.respond_to?(:rewind)
                        xml.read
                      else
                        xml.to_s
                      end
            parse_with_fast_scan(raw_str)
          end
        end

        # @param xml [untyped]
        # @return [Hash<Integer, String>]
        #: (untyped xml) -> Hash[Integer, String]
        def self.parse_with_nokogiri(xml)
          dictionary = {}
          namespace = xml.namespaces.detect { |_key, uri| uri == SPREADSHEETML_URI }
          prefix = if namespace && namespace[0].start_with?("xmlns:")
                     "#{namespace[0].delete_prefix("xmlns:")}|"
                   else
                     ""
                   end
          node_selector = "#{prefix}si"
          text_selector = ">#{prefix}t, #{prefix}r #{prefix}t"

          xml.css(node_selector).each_with_index do |si, idx|
            text_nodes = si.css(text_selector)
            dictionary[idx] = if text_nodes.one?
                                Styles::Converter.unescape_string(text_nodes.first.content)
                              else
                                text_nodes.map { |n| Styles::Converter.unescape_string(n.content) }.join
                              end
          end

          dictionary
        end

        # @param xml_str [String]
        # @return [Hash<Integer, String>]
        #: (String xml_str) -> Hash[Integer, String]
        def self.parse_with_fast_scan(xml_str)
          dict = {}
          idx = 0
          xml_str.scan(%r{<(?:[A-Za-z0-9_]+:)?si\b[^>]*>(.*?)</(?:[A-Za-z0-9_]+:)?si>}m) do |m|
            si_body = m[0]
            # Strip phonetic runs <rPh>...</rPh>
            cleaned = si_body.gsub(%r{<(?:[A-Za-z0-9_]+:)?rPh\b[^>]*>.*?</(?:[A-Za-z0-9_]+:)?rPh>}m, "")
            t_nodes = []
            cleaned.scan(%r{<(?:[A-Za-z0-9_]+:)?t\b[^>]*>(.*?)</(?:[A-Za-z0-9_]+:)?t>}m) do |t_m|
              t_nodes << Styles::Converter.unescape_string(decode_xml_entities(t_m[0]))
            end
            dict[idx] = t_nodes.join
            idx += 1
          end
          dict
        end

        # @param str [String]
        # @return [String]
        #: (String str) -> String
        def self.decode_xml_entities(str)
          return str unless str.include?("&")

          str.gsub("&amp;", "&").gsub("&lt;", "<").gsub("&gt;", ">").gsub("&quot;", '"').gsub("&apos;", "'")
        end
      end
    end
  end
end
