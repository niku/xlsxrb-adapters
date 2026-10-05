# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require_relative "write_xlsx/constants"
require_relative "write_xlsx/colors"
require_relative "write_xlsx/utility"
require_relative "write_xlsx/format"
require_relative "write_xlsx/chart"
require_relative "write_xlsx/worksheet"
require_relative "write_xlsx/workbook"

module Xlsxrb
  module Adapters
    # WriteXLSX adapter class providing drop-in compatibility with the write_xlsx gem
    class WriteXLSX < Writexlsx::Workbook
      ROW_MAX = Writexlsx::ROW_MAX
      COL_MAX = Writexlsx::COL_MAX
      STR_MAX = Writexlsx::STR_MAX
      SHEETNAME_MAX = Writexlsx::SHEETNAME_MAX
      MAX_URL_LENGTH = Writexlsx::MAX_URL_LENGTH
      URL_MAX = MAX_URL_LENGTH
      OFFICE_URL = Writexlsx::OFFICE_URL

      Workbook = Writexlsx::Workbook
      Worksheet = Writexlsx::Worksheet
      Format = Writexlsx::Format
      Chart = Writexlsx::Chart
      Colors = Writexlsx::Colors
      Utility = Writexlsx::Utility

      WriteXLSXInsufficientArgumentError = Writexlsx::WriteXLSXInsufficientArgumentError
      WriteXLSXDimensionError = Writexlsx::WriteXLSXDimensionError
      WriteXLSXOptionParameterError = Writexlsx::WriteXLSXOptionParameterError

      InsufficientArgumentError = Writexlsx::InsufficientArgumentError
      DimensionError = Writexlsx::DimensionError
      OptionParameterError = Writexlsx::OptionParameterError

      # Converts an adapter Workbook to an immutable Xlsxrb::Elements::Workbook.
      #
      # @param obj [Writexlsx::Workbook]
      # @return [Xlsxrb::Elements::Workbook]
      #: (Writexlsx::Workbook obj) -> Xlsxrb::Elements::Workbook
      def self.to_xlsxrb(obj)
        obj.to_xlsxrb
      end

      # Converts an Xlsxrb::Elements::Workbook into a WriteXLSX adapter instance.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @param file [Object, nil]
      # @param options [Hash{Symbol => untyped}]
      # @return [Workbook]
      def self.from_xlsxrb(xlsxrb_workbook, file = nil, options = {})
        Writexlsx::Workbook.from_xlsxrb(xlsxrb_workbook, file, options)
      end
    end

    # Module namespace matching upstream Writexlsx
    module Writexlsx
      WriteXLSX = Xlsxrb::Adapters::WriteXLSX

      # Compatibility namespaces matching write_xlsx structure
      # rubocop:disable Lint/EmptyClass
      module Package
        class Table; end
        class ConditionalFormat; end
        class XMLWriterSimple; end
      end

      class Shape; end
      class Sparkline; end
      # rubocop:enable Lint/EmptyClass

      # @param obj [Workbook]
      # @return [Xlsxrb::Elements::Workbook]
      def self.to_xlsxrb(obj)
        WriteXLSX.to_xlsxrb(obj)
      end

      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @param file [Object, nil]
      # @param options [Hash{Symbol => untyped}]
      # @return [Workbook]
      def self.from_xlsxrb(xlsxrb_workbook, file = nil, options = {})
        WriteXLSX.from_xlsxrb(xlsxrb_workbook, file, options)
      end
    end
  end
end
