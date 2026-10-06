# frozen_string_literal: true

# rbs_inline: enabled

require "forwardable"
require "xlsxrb"

module Xlsxrb
  module Adapters
    module SimpleXlsxReader
      # Main document class matching SimpleXlsxReader::Document.
      class Document
        # Underlying file path, IO, or binary string.
        # @return [String, IO, StringIO, File]
        #: untyped
        attr_reader :string_or_io

        # @param legacy_file_path [String, nil] Positional file path.
        # @param file_path [String, nil] Keyword file path.
        # @param string_or_io [String, IO, StringIO, File, nil] Buffer or IO object.
        #: (?String? legacy_file_path, ?file_path: String?, ?string_or_io: untyped) -> void
        def initialize(legacy_file_path = nil, file_path: nil, string_or_io: nil)
          raise ArgumentError, "either file_path or string_or_io must be provided" if legacy_file_path.nil? && file_path.nil? && string_or_io.nil?

          @string_or_io = string_or_io || File.new(legacy_file_path || file_path)
          @sheets = nil
        end

        # Access sheets in this workbook.
        #
        # @return [Array<Sheet>]
        #: () -> Array[Sheet]
        def sheets
          @sheets ||= Loader.new(string_or_io).init_sheets
        end

        # Slurps all sheets and returns a Hash mapping sheet name to an array of rows.
        #
        # @return [Hash<String, Array<Array<Object>>>]
        #: () -> Hash[String, Array[Array[untyped]]]
        def to_hash
          sheets.to_h { |sheet| [sheet.name, sheet.rows.to_a] }
        end

        # Converts this SimpleXlsxReader Document into an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          Elements::Workbook.new(sheets: sheets.map(&:to_xlsxrb))
        end

        # Constructs a SimpleXlsxReader Document from an Xlsxrb::Elements::Workbook.
        #
        # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
        # @return [Document]
        #: (Xlsxrb::Elements::Workbook xlsxrb_workbook) -> Document
        def self.from_xlsxrb(xlsxrb_workbook)
          bin = Xlsxrb.write(xlsxrb_workbook)
          new(string_or_io: bin).tap(&:sheets)
        end

        # Worksheet representation matching SimpleXlsxReader::Document::Sheet.
        class Sheet
          extend Forwardable

          # Sheet name.
          # @return [String]
          #: String
          attr_reader :name

          # Rows proxy responding to #each and #slurp.
          # @return [RowsProxy]
          #: RowsProxy
          attr_reader :rows

          def_delegators :rows, :load_errors, :slurp

          # @param name [String] Sheet name.
          # @param sheet_parser [Loader::SheetParser] Parser instance.
          #: (name: String, sheet_parser: Loader::SheetParser) -> void
          def initialize(name:, sheet_parser:)
            @name = name
            @rows = RowsProxy.new(sheet_parser: sheet_parser)
          end

          # Returns the first row as headers from slurped rows.
          #
          # @return [Array<Object>]
          #: () -> Array[untyped]
          def headers
            rows.slurped![0]
          end

          # Returns rows excluding the header row from slurped rows.
          #
          # @return [Array<Array<Object>>]
          #: () -> Array[Array[untyped]]
          def data
            rows.slurped![1..] || []
          end

          # Converts this Sheet into an Xlsxrb::Elements::Worksheet.
          #
          # @return [Xlsxrb::Elements::Worksheet]
          #: () -> Xlsxrb::Elements::Worksheet
          def to_xlsxrb
            row_arrays = rows.to_a
            worksheet_rows = row_arrays.map.with_index do |row_vals, r_idx|
              cells = (row_vals || []).map.with_index do |val, c_idx|
                if val.is_a?(Hyperlink)
                  Elements::Cell.new(row_index: r_idx, column_index: c_idx, value: val.to_s, hyperlink: val.url)
                else
                  Elements::Cell.new(row_index: r_idx, column_index: c_idx, value: val)
                end
              end
              Elements::Row.new(index: r_idx, cells: cells)
            end
            Elements::Worksheet.new(name: name, rows: worksheet_rows)
          end
        end

        # Rows proxy supporting lazy streaming and slurping matching SimpleXlsxReader::Document::RowsProxy.
        class RowsProxy
          [Enumerable].each { |m| include m }

          # Cached array of slurped rows.
          # @return [Array<Array<Object>>, nil]
          #: Array[Array[untyped]]?
          attr_reader :slurped

          # Cell load errors recorded when catch_cell_load_errors is enabled.
          # @return [Hash<Array(Integer, Integer), String>]
          #: Hash[[Integer, Integer], String]
          attr_reader :load_errors

          # @param sheet_parser [Loader::SheetParser]
          #: (sheet_parser: Loader::SheetParser) -> void
          def initialize(sheet_parser:)
            @sheet_parser = sheet_parser
            @slurped = nil
            @load_errors = {}
          end

          # Iterates over rows, yielding arrays by default or hashes if headers are configured.
          #
          # @param headers [Boolean, Proc, Hash] Header parsing configuration.
          # @yield [row]
          # @yieldparam row [Array<Object>, Hash<Object, Object>]
          # @return [Enumerator, untyped]
          #: (?headers: untyped) ?{ (untyped) -> void } -> untyped
          def each(headers: false, &)
            if slurped?
              raise "#each does not support headers with slurped rows" if headers

              slurped.each(&)
            elsif block_given?
              sheet_parser = @sheet_parser
              @sheet_parser.parse(headers: headers, &).tap do
                @load_errors = sheet_parser.load_errors
              end
            else
              to_enum(:each, headers: headers)
            end
          end

          # Loads all rows into memory and releases parser resources.
          #
          # @return [Array<Array<Object>>]
          #: () -> Array[Array[untyped]]
          def slurp
            # rubocop:disable-next Naming/MemoizedInstanceVariableName
            @slurped ||= to_a.tap { @sheet_parser = nil }
          end

          # Returns whether rows have been slurped into memory.
          #
          # @return [Boolean]
          #: () -> bool
          def slurped?
            !@slurped.nil?
          end

          # Returns slurped rows, raising RuntimeError unless slurped (or auto_slurp is true).
          #
          # @return [Array<Array<Object>>]
          #: () -> Array[Array[untyped]]
          def slurped!
            check_slurped
            slurped || []
          end

          # Access slurped rows by index or range.
          #
          # @param args [Array<Object>] Index, slice, or range arguments.
          # @return [Object]
          #: (*untyped args) -> untyped
          def [](*)
            check_slurped
            (slurped || [])[*]
          end

          # Shifts the first row from slurped rows.
          #
          # @param args [Array<Object>] Optional count argument.
          # @return [Object]
          #: (*untyped args) -> untyped
          def shift(*)
            check_slurped
            (slurped || []).shift(*)
          end

          private

          #: () -> void
          def check_slurped
            slurp if SimpleXlsxReader.configuration.auto_slurp
            return if slurped?

            raise "Called a slurp-y method without explicitly slurping; use #each or call rows.slurp first"
          end
        end
      end
    end
  end
end
