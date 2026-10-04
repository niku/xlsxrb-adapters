# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require_relative "xsv/helpers"
require_relative "xsv/sheet"
require_relative "xsv/workbook"

module Xlsxrb
  module Adapters
    # Xsv adapter module providing a drop-in compatible interface for Xsv (ISO/IEC 29500 OOXML)
    # without hijacking the global ::Xsv namespace by default.
    module Xsv
      class Error < StandardError; end
      class DuplicateHeaders < StandardError; end
      class AssertionFailed < StandardError; end
      VERSION = "1.4.1"

      # Open the workbook of the given filename, string or buffer.
      #
      # @param filename_or_string [String, IO, StringIO]
      # @param trim_empty_rows [Boolean]
      # @param parse_headers [Boolean]
      # @yield [workbook]
      # @yieldparam workbook [Workbook]
      # @return [Workbook, Object]
      #: (untyped filename_or_string, ?trim_empty_rows: bool, ?parse_headers: bool) ?{ (Workbook) -> untyped } -> untyped
      def self.open(filename_or_string, trim_empty_rows: false, parse_headers: false, &block)
        err_class = defined?(::Zip::Error) ? ::Zip::Error : Error

        # Validate non-empty file / input
        input = if filename_or_string.is_a?(String) && (filename_or_string.start_with?("PK\x03\x04") || filename_or_string.include?("\x00"))
                  StringIO.new(filename_or_string.b)
                else
                  filename_or_string
                end

        if input.is_a?(String)
          raise err_class, "File #{input} not found" unless File.exist?(input)
          raise err_class, "File #{input} has zero size." if File.empty?(input)
        elsif input.respond_to?(:size) && input.size <= 0
          raise err_class, "File has zero size."
        end

        reader = Ooxml::ZipReader.open(input)
        workbook = Workbook.new(reader, trim_empty_rows: trim_empty_rows, parse_headers: parse_headers)

        if block_given?
          begin
            yield workbook
          ensure
            workbook.close
          end
        else
          workbook
        end
      end

      # Converts an Xlsxrb::Elements::Workbook into an Xsv Workbook adapter instance.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @param trim_empty_rows [Boolean]
      # @param parse_headers [Boolean]
      # @return [Workbook]
      #: (Xlsxrb::Elements::Workbook xlsxrb_workbook, ?trim_empty_rows: bool, ?parse_headers: bool) -> Workbook
      def self.from_xlsxrb(xlsxrb_workbook, trim_empty_rows: false, parse_headers: false)
        Workbook.from_xlsxrb(xlsxrb_workbook, trim_empty_rows: trim_empty_rows, parse_headers: parse_headers)
      end
    end
  end
end
