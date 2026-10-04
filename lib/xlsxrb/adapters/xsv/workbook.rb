# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require_relative "helpers"
require_relative "sheet"

module Xlsxrb
  module Adapters
    module Xsv
      # An OOXML Spreadsheet document matching Xsv::Workbook.
      # Consists of multiple Sheets accessible via {#sheets} or index/name access.
      class Workbook
        [Enumerable].each { |m| include m }

        # Access Sheet objects in the workbook.
        # @return [Array<Sheet>, nil]
        #: Array[Sheet]?
        attr_reader :sheets

        # Shared strings table array.
        # @return [Array<String>, nil]
        #: Array[String]?
        attr_reader :shared_strings

        # Styles cell formatting records array.
        # @return [Array<Hash[Symbol, Integer]>, nil]
        #: Array[Hash[Symbol, Integer]]?
        attr_reader :xfs

        # Number format definitions mapping id to format string.
        # @return [Hash<Integer, String>, nil]
        #: Hash[Integer, String]?
        attr_reader :num_fmts

        # Parsed Elements::Styles definition.
        # @return [Elements::Styles, nil]
        #: Elements::Styles?
        attr_reader :styles

        # Whether the 1904 date system is active.
        # @return [Boolean]
        #: bool
        attr_reader :date1904

        # Whether empty trailing rows are trimmed.
        # @return [Boolean]
        #: bool
        attr_reader :trim_empty_rows

        # Returns whether the workbook uses the 1904 date system.
        # @return [Boolean]
        #: () -> bool
        def date1904?
          @date1904 ? true : false
        end

        # Returns whether empty trailing rows are trimmed.
        # @return [Boolean]
        #: () -> bool
        def trim_empty_rows?
          @trim_empty_rows ? true : false
        end

        # Deprecated: use Xsv.open instead.
        # @param data [String, IO]
        # @param kws [Hash]
        #: (untyped data, **untyped kws) ?{ (Workbook) -> untyped } -> untyped
        def self.open(data, **kws, &)
          Xsv.open(data, **kws, &)
        end

        # Converts an Xlsxrb::Elements::Workbook into an Xsv adapter Workbook.
        #
        # @param xlsxrb_wb [Xlsxrb::Elements::Workbook]
        # @param trim_empty_rows [Boolean]
        # @param parse_headers [Boolean]
        # @return [Workbook]
        #: (Xlsxrb::Elements::Workbook xlsxrb_wb, ?trim_empty_rows: bool, ?parse_headers: bool) -> Workbook
        def self.from_xlsxrb(xlsxrb_wb, trim_empty_rows: false, parse_headers: false)
          binary = Xlsxrb.write(xlsxrb_wb)
          Xsv.open(binary, trim_empty_rows: trim_empty_rows, parse_headers: parse_headers)
        end

        # Open a workbook from an open ZipReader.
        #
        # @param zip_reader [Ooxml::ZipReader]
        # @param trim_empty_rows [Boolean]
        # @param parse_headers [Boolean]
        #: (Ooxml::ZipReader zip_reader, ?trim_empty_rows: bool, ?parse_headers: bool) -> void
        def initialize(zip_reader, trim_empty_rows: false, parse_headers: false)
          @zip = zip_reader
          @trim_empty_rows = trim_empty_rows ? true : false

          @sheets = []
          @xfs, @num_fmts, @styles = fetch_styles
          @sheet_ids, @date1904 = fetch_sheet_ids
          @relationships = fetch_relationships
          @shared_strings = fetch_shared_strings
          @sheets = fetch_sheets(parse_headers)
        end

        # Human-readable representation.
        # @return [String]
        #: () -> String
        def inspect
          "#<#{self.class.name}:#{object_id} sheets=#{sheets ? sheets.count : 0} trim_empty_rows=#{@trim_empty_rows}>"
        end

        # Close the workbook handle and release resources.
        # @return [true]
        #: () -> bool
        def close
          @zip&.close
          @zip = nil
          @sheets = nil
          @xfs = nil
          @num_fmts = nil
          @relationships = nil
          @shared_strings = nil
          @sheet_ids = nil

          true
        end

        # Returns an array of sheets matching the given sheet name.
        #
        # @param name [String]
        # @return [Array<Sheet>]
        #: (String name) -> Array[Sheet]
        def sheets_by_name(name)
          return [] unless @sheets

          @sheets.select { |s| s.name == name }
        end

        # Get number format string for given style index.
        #
        # @param style [Integer]
        # @return [String, nil]
        #: (Integer style) -> String?
        def get_num_fmt(style)
          return nil unless style
          return @styles.number_format(style) if @styles

          return nil unless @xfs && @xfs[style]

          num_fmt_id = @xfs[style][:numFmtId]
          @num_fmts[num_fmt_id] if num_fmt_id
        end

        # Iterate over worksheet instances.
        #
        # @yield [sheet]
        # @yieldparam sheet [Sheet]
        # @return [Enumerator, self]
        #: () { (Sheet) -> void } -> untyped
        #: () -> Enumerator[Sheet, void]
        def each(&block)
          return enum_for(:each) unless block

          @sheets&.each(&block)
          self
        end

        # Access a worksheet by 0-based index or name.
        #
        # @param index_or_name [Integer, String]
        # @return [Sheet, nil]
        #: (Integer | String index_or_name) -> Sheet?
        def [](index_or_name)
          case index_or_name
          when Integer
            @sheets ? @sheets[index_or_name] : nil
          when String
            sheets_by_name(index_or_name).first
          else
            raise ArgumentError, "Sheets can be accessed by Integer of String only"
          end
        end

        # Converts this Xsv Workbook instance into an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          worksheets = (@sheets || []).map(&:to_xlsxrb)
          Elements::Workbook.new(sheets: worksheets, date1904: date1904?)
        end

        private

        #: () -> Array[String]
        def fetch_shared_strings
          sst_xml = @zip&.read_entry("xl/sharedStrings.xml")
          return [] unless sst_xml

          Ooxml::SharedStringsParser.parse(sst_xml)
        end

        #: () -> [Array[Hash[Symbol, Integer]], Hash[Integer, String], Elements::Styles?]
        def fetch_styles
          styles_xml = @zip&.read_entry("xl/styles.xml")
          return [[], Helpers::BUILT_IN_NUMBER_FORMATS.dup, nil] unless styles_xml

          parsed = Ooxml::StylesParser.parse(styles_xml)
          cell_xfs = (parsed[:cell_xfs] || []).map do |xf|
            { numFmtId: xf[:num_fmt_id] || 0 }
          end
          num_fmts = Helpers::BUILT_IN_NUMBER_FORMATS.dup.merge(parsed[:num_fmts] || {})

          [cell_xfs, num_fmts, parsed]
        end

        #: () -> [Array[Hash[Symbol, untyped]], bool]
        def fetch_sheet_ids
          wb_xml = @zip&.read_entry("xl/workbook.xml")
          return [[], false] unless wb_xml

          parsed = Ooxml::WorkbookParser.parse_with_properties(wb_xml)
          sheets = (parsed[:sheets] || []).map do |s|
            {
              name: s[:name].to_s.force_encoding(Encoding::UTF_8),
              sheetId: s[:sheet_id],
              id: s[:r_id],
              state: s[:state]
            }
          end
          [sheets, parsed[:date1904] ? true : false]
        end

        #: () -> Hash[String, String]
        def fetch_relationships
          rels_xml = @zip&.read_entry("xl/_rels/workbook.xml.rels")
          return {} unless rels_xml

          Ooxml::RelationshipsParser.parse(rels_xml)
        end

        #: (bool parse_headers) -> Array[Sheet]
        def fetch_sheets(parse_headers)
          return [] unless @zip

          @sheet_ids.map do |sheet_id_info|
            target = @relationships[sheet_id_info[:id]]
            next nil unless target

            sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
            next nil unless @zip.entry?(sheet_path)

            xml = @zip.read_entry(sheet_path) || ""
            stream_sheet = StreamSheet.new(
              sheet_id_info[:name],
              xml,
              @shared_strings || [],
              @styles,
              zip_reader: @zip,
              entry_name: sheet_path,
              state: sheet_id_info[:state] || :visible,
              date1904: @date1904,
              trim_empty_rows: @trim_empty_rows
            )
            sheet = Sheet.new(self, xml, sheet_id_info, stream_sheet: stream_sheet)
            sheet.parse_headers! if parse_headers
            sheet
          end.compact
        end
      end
    end
  end
end
