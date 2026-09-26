# frozen_string_literal: true

# rbs_inline: enabled

require "time"
require "stringio"
require "xlsxrb"
require_relative "workbook"
require_relative "merged_cells"

module Xlsxrb
  module Adapters
    module RubyXL
      # Parser providing RubyXL::Parser compatible reading interface backed by xlsxrb.
      class Parser
        # Parses an XLSX file from disk.
        #
        # @param file_path [String]
        # @return [Workbook]
        #: (String file_path) -> Workbook
        def self.parse(file_path)
          unless File.exist?(file_path)
            err_class = defined?(::Zip::Error) ? ::Zip::Error : StandardError
            raise err_class, "File #{file_path} not found"
          end

          xlsxrb_wb = Xlsxrb.read(file_path).load
          wb = Workbook.from_xlsxrb(xlsxrb_wb)
          enrich_workbook(file_path, wb)
          wb
        rescue StandardError => e
          raise ::Zip::Error, e.message if defined?(::Zip::Error) && !e.is_a?(::Zip::Error) && (e.is_a?(Errno::ENOENT) || e.message =~ /zip|entry|central directory/i)

          raise
        end

        # Parses an XLSX file from an in-memory buffer (StringIO, IO, or binary String).
        #
        # @param buffer [StringIO, IO, String]
        # @return [Workbook]
        #: (StringIO | IO | String buffer) -> Workbook
        def self.parse_buffer(buffer)
          buf = buffer.is_a?(String) ? StringIO.new(buffer.b) : buffer
          xlsxrb_wb = Xlsxrb.read(buf).load
          wb = Workbook.from_xlsxrb(xlsxrb_wb)
          enrich_workbook(buf, wb)
          wb
        end

        # Enriches workbook with metadata from ZIP entries not directly in xlsxrb DOM.
        #
        # @param source [Object]
        # @param wb [Workbook]
        # @return [void]
        #: (untyped source, Workbook wb) -> void
        def self.enrich_workbook(source, wb)
          io = if source.is_a?(String) && File.exist?(source)
                 File.open(source, "rb")
               elsif source.is_a?(String)
                 StringIO.new(source.b)
               elsif source.respond_to?(:read)
                 source.rewind if source.respond_to?(:rewind)
                 source
               end
          return unless io

          zip = Xlsxrb::Ooxml::ZipReader.open(io)
          begin
            # Custom number formats from styles.xml
            styles_xml = zip.read_entry("xl/styles.xml")
            if styles_xml
              num_fmt_matches = styles_xml.scan(/<numFmt\s+numFmtId="(\d+)"\s+formatCode="([^"]+)"/)
              num_fmt_matches.each do |id_str, code|
                id = id_str.to_i
                code = unescape_xml(code) || code
                wb.stylesheet.number_formats << NumberFormat.new(num_fmt_id: id, format_code: code) unless wb.stylesheet.number_formats.any? { |nf| nf.num_fmt_id == id }
              end
            end

            # Core properties
            core_xml = zip.read_entry("docProps/core.xml")
            if core_xml
              creator = core_xml[%r{<dc:creator[^>]*>([^<]+)</dc:creator>}, 1]
              modifier = core_xml[%r{<cp:lastModifiedBy[^>]*>([^<]+)</cp:lastModifiedBy>}, 1]
              title = core_xml[%r{<dc:title[^>]*>([^<]+)</dc:title>}, 1]
              created = core_xml[%r{<dcterms:created[^>]*>([^<]+)</dcterms:created>}, 1]
              modified = core_xml[%r{<dcterms:modified[^>]*>([^<]+)</dcterms:modified>}, 1]

              wb.creator = unescape_xml(creator) if creator
              wb.modifier = unescape_xml(modifier) if modifier
              wb.title = unescape_xml(title) if title
              wb.created_at = Time.parse(created) if created
              wb.modified_at = Time.parse(modified) if modified
            end

            # Defined names from workbook.xml
            workbook_xml = zip.read_entry("xl/workbook.xml")
            if workbook_xml
              dn_matches = workbook_xml.scan(%r{<definedName\s+name="([^"]+)"[^>]*>([^<]+)</definedName>})
              unless dn_matches.empty?
                wb.defined_names ||= DefinedNames.new
                dn_matches.each do |name, ref|
                  next if wb.defined_names.any? { |d| d.name == name }

                  wb.defined_names << DefinedName.new(name: unescape_xml(name) || name, reference: unescape_xml(ref) || ref)
                end
              end
            end

            # Merged cells per sheet
            wb.worksheets.each_with_index do |ws, idx|
              sheet_entry = "xl/worksheets/sheet#{idx + 1}.xml"
              sheet_xml = zip.read_entry(sheet_entry)
              next unless sheet_xml

              merge_matches = sheet_xml.scan(/<mergeCell\s+ref="([^"]+)"/)
              next if merge_matches.empty?

              ws.merged_cells ||= MergedCells.new
              merge_matches.each do |match|
                ref = match.first
                next unless ref

                ws.merged_cells << MergedCell.new(ref: ref) unless ws.merged_cells.any? { |m| m.ref.to_s == ref }
              end
            end
          ensure
            zip.close
            io.close if io.is_a?(File)
          end
        rescue StandardError
          # Best effort metadata recovery
        end

        # @param str [String, nil]
        # @return [String, nil]
        #: (String? str) -> String?
        def self.unescape_xml(str)
          return nil unless str

          str.gsub("&amp;", "&").gsub("&lt;", "<").gsub("&gt;", ">").gsub("&quot;", '"').gsub("&apos;", "'")
        end
      end
    end
  end
end
