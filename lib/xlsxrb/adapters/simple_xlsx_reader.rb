# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "xlsxrb"
require_relative "simple_xlsx_reader/version"
require_relative "simple_xlsx_reader/hyperlink"
require_relative "simple_xlsx_reader/document"
require_relative "simple_xlsx_reader/loader"

module Xlsxrb
  module Adapters
    # SimpleXlsxReader adapter module providing a drop-in compatible interface for
    # simple_xlsx_reader without hijacking the global ::SimpleXlsxReader namespace by default.
    module SimpleXlsxReader
      # Base date for the 1900 date system.
      DATE_SYSTEM_1900 = Date.new(1899, 12, 30).freeze

      # Base date for the 1904 date system.
      DATE_SYSTEM_1904 = Date.new(1904, 1, 1).freeze

      # Raised when a cell fails to cast into its designated type.
      class CellLoadError < StandardError; end

      # Alias for Document::Sheet
      Sheet = Document::Sheet

      # Configuration options container matching SimpleXlsxReader.configuration.
      Configuration = Struct.new(:catch_cell_load_errors, :auto_slurp) do
        # @param catch_cell_load_errors [Boolean]
        # @param auto_slurp [Boolean]
        #: (?catch_cell_load_errors: bool, ?auto_slurp: bool) -> void
        def initialize(catch_cell_load_errors: false, auto_slurp: false)
          super(catch_cell_load_errors, auto_slurp)
        end
      end

      class << self
        # Global configuration object.
        #
        # @return [Configuration]
        #: () -> Configuration
        def configuration
          @configuration ||= Configuration.new
        end

        # Opens and loads an XLSX spreadsheet from a file path.
        #
        # @param file_path [String]
        # @return [Document, Object]
        #: (String file_path) ?{ (Document) -> untyped } -> untyped
        def open(file_path)
          doc = Document.new(file_path: file_path).tap(&:sheets)
          block_given? ? yield(doc) : doc
        end

        # Parses an XLSX spreadsheet from a string or IO stream.
        #
        # @param string_or_io [String, IO, StringIO]
        # @return [Document, Object]
        #: (untyped string_or_io) ?{ (Document) -> untyped } -> untyped
        def parse(string_or_io)
          doc = Document.new(string_or_io: string_or_io).tap(&:sheets)
          block_given? ? yield(doc) : doc
        end

        # Converts an Xlsxrb::Elements::Workbook into a SimpleXlsxReader Document adapter instance.
        #
        # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
        # @return [Document]
        #: (Xlsxrb::Elements::Workbook xlsxrb_workbook) -> Document
        def from_xlsxrb(xlsxrb_workbook)
          Document.from_xlsxrb(xlsxrb_workbook)
        end
      end
    end
  end
end
