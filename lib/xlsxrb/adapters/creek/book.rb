# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "tempfile"
require "stringio"
require "open-uri"
require "xlsxrb"
require_relative "shared_strings"
require_relative "styles"
require_relative "sheet"

module Xlsxrb
  module Adapters
    module Creek
      # Workbook parser matching Creek::Book.
      # Opens and extracts worksheets, styles, and shared strings from an XLSX / XLSM file.
      class Book
        DATE_1900 = Date.new(1899, 12, 30).freeze
        DATE_1904 = Date.new(1904, 1, 1).freeze

        # Wrapper mimicking Creek's @files.file.open / exist? interface over ZipReader.
        class FilesWrapper
          # File proxy supporting .open and .exist?
          class FileProxy
            # @param zip_reader [Xlsxrb::Ooxml::ZipReader]
            #: (Xlsxrb::Ooxml::ZipReader zip_reader) -> void
            def initialize(zip_reader)
              @zip_reader = zip_reader
            end

            # @param path [String]
            # @return [Boolean]
            #: (String path) -> bool
            def exist?(path)
              norm = path.delete_prefix("/")
              @zip_reader.entry?(norm) || @zip_reader.entry?(path)
            end

            # @param path [String]
            # @yield [io]
            # @return [StringIO, untyped]
            #: (String path) ?{ (StringIO) -> untyped } -> untyped
            def open(path, &block)
              norm = path.delete_prefix("/")
              data = @zip_reader.read_entry(norm) || @zip_reader.read_entry(path)
              raise Errno::ENOENT, "No such entry in archive: #{path}" unless data

              io = StringIO.new(data.b)
              if block
                begin
                  block.call(io)
                ensure
                  io.close
                end
              else
                io
              end
            end
          end

          # @return [Xlsxrb::Ooxml::ZipReader, nil]
          attr_reader :zip_reader

          # @param zip_reader [Xlsxrb::Ooxml::ZipReader]
          #: (Xlsxrb::Ooxml::ZipReader zip_reader) -> void
          def initialize(zip_reader)
            @zip_reader = zip_reader
            @file = FileProxy.new(zip_reader)
          end

          # Returns the entry file accessor.
          #
          # @return [FileProxy]
          #: () -> FileProxy
          attr_reader :file

          # Closes the archive reader.
          #
          # @return [void]
          #: () -> void
          def close
            @zip_reader&.close
            @zip_reader = nil
          end
        end

        # @return [FilesWrapper]
        attr_reader :files

        # @return [SharedStrings]
        attr_reader :shared_strings

        # @return [Boolean]
        attr_reader :with_headers

        # Opens an XLSX file or buffer.
        #
        # @param path [String, IO, StringIO, Pathname]
        # @param options [Hash]
        #: (untyped path, ?Hash[Symbol, untyped] options) -> void
        def initialize(path, options = {})
          check_file_extension = options.fetch(:check_file_extension, true)
          if check_file_extension
            original_name = (options[:original_filename] || (path.is_a?(String) ? path : "")).to_s
            extension = File.extname(original_name).downcase
            raise "Not a valid file format." unless [".xlsx", ".xlsm"].include?(extension)
          end

          input_path = options[:remote] ? download_file(path.to_s) : path
          zip_reader = Xlsxrb::Ooxml::ZipReader.open(input_path)

          @files = FilesWrapper.new(zip_reader)
          @shared_strings = SharedStrings.new(self)
          @with_headers = options.fetch(:with_headers, false)
          @sheets = nil
        end

        # Returns the list of Sheet objects parsed from xl/workbook.xml.
        #
        # @return [Array<Sheet>]
        #: () -> Array[Sheet]
        def sheets
          return @sheets if @sheets

          wb_path = "xl/workbook.xml"
          rels_path = "xl/_rels/workbook.xml.rels"

          rels = {}
          if @files.file.exist?(rels_path)
            rels_doc = @files.file.open(rels_path).read
            rels_doc.scan(%r{<Relationship\b([^>]*?)(?:/>|>.*?</Relationship>)}m) do |attrs_str,|
              id = attrs_str[/Id="([^"]*)"/, 1]
              target = attrs_str[/Target="([^"]*)"/, 1]
              rels[id] = target if id && target
            end
          end

          @sheets = []
          if @files.file.exist?(wb_path)
            wb_doc = @files.file.open(wb_path).read
            wb_doc.scan(%r{<(?:[A-Za-z0-9_]+:)?sheet\b([^>]*?)(?:/>|>.*?</(?:[A-Za-z0-9_]+:)?sheet>)}m) do |attrs_str,|
              name = attrs_str[/name="([^"]*)"/, 1]
              sheetid = attrs_str[/sheetid="([^"]*)"/, 1]
              state = attrs_str[/state="([^"]*)"/, 1]
              visible = attrs_str[/visible="([^"]*)"/, 1]
              rid = attrs_str[/(?:r:id|id)="([^"]*)"/, 1]
              sheetfile = rels[rid] || ""

              sheet = Sheet.new(self, name || "", sheetid, state, visible, rid, sheetfile)
              sheet.with_headers = @with_headers
              @sheets << sheet
            end
          end
          @sheets
        end

        # Returns array of style types defined in the workbook.
        #
        # @return [Array<Symbol?>]
        #: () -> Array[Symbol?]
        def style_types
          @style_types ||= Styles.new(self).style_types
        end

        # Closes the workbook file handles.
        #
        # @return [void]
        #: () -> void
        def close
          @files&.close
        end

        # Returns base date for serial arithmetic (DATE_1900 or DATE_1904).
        #
        # @return [Date]
        #: () -> Date
        def base_date
          @base_date ||= begin
            result = DATE_1900
            wb_path = "xl/workbook.xml"
            if @files.file.exist?(wb_path)
              wb_doc = @files.file.open(wb_path).read
              result = DATE_1904 if wb_doc =~ /<workbookPr\b[^>]*\bdate1904=(?:"(true|1)"|'(true|1)')/i
            end
            result
          end
        end

        # Converts this Creek Book into an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          worksheets = sheets.map(&:to_xlsxrb)
          Xlsxrb::Elements::Workbook.new(sheets: worksheets, date1904: base_date == DATE_1904)
        end

        # Converts an Xlsxrb::Elements::Workbook into a Creek Book adapter instance.
        #
        # @param xlsxrb_wb [Xlsxrb::Elements::Workbook]
        # @param options [Hash]
        # @return [Book]
        #: (Xlsxrb::Elements::Workbook xlsxrb_wb, ?Hash[Symbol, untyped] options) -> Book
        def self.from_xlsxrb(xlsxrb_wb, options = {})
          binary = Xlsxrb.write(xlsxrb_wb)
          opts = options.merge(check_file_extension: false)
          new(StringIO.new(binary), opts)
        end

        private

        # Downloads a remote file to a temporary location.
        #
        # @param url [String]
        # @return [String]
        #: (String url) -> String
        def download_file(url)
          downloaded = URI.parse(url).open
          if downloaded.is_a?(StringIO)
            temp = Tempfile.new(["creek-file", ".xlsx"])
            temp.binmode
            temp.write(downloaded.read)
            temp.close
            temp.path
          else
            downloaded.path
          end
        end
      end
    end
  end
end
