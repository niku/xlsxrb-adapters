# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Writexlsx
      # Maximum number of rows supported in Excel 2007+ (OOXML)
      ROW_MAX = 1_048_576

      # Maximum number of columns supported in Excel 2007+ (OOXML)
      COL_MAX = 16_384

      # Maximum string length in a single cell
      STR_MAX = 32_767

      # Maximum length of a worksheet name
      SHEETNAME_MAX = 31

      # Maximum length of a hyperlink URL in Excel
      MAX_URL_LENGTH = 2_079
      URL_MAX = MAX_URL_LENGTH

      # Microsoft Office schema URL namespace prefix
      OFFICE_URL = "http://schemas.microsoft.com/office/"

      # Raised when insufficient arguments are passed to a WriteXLSX method
      class WriteXLSXInsufficientArgumentError < StandardError; end

      # Raised when cell row/column coordinates exceed Excel dimensions
      class WriteXLSXDimensionError < StandardError; end

      # Raised when an invalid option parameter is provided
      class WriteXLSXOptionParameterError < StandardError; end

      InsufficientArgumentError = WriteXLSXInsufficientArgumentError
      DimensionError = WriteXLSXDimensionError
      OptionParameterError = WriteXLSXOptionParameterError
    end
  end
end

# Top-level error compatibility matching upstream write_xlsx
WriteXLSXInsufficientArgumentError = Xlsxrb::Adapters::Writexlsx::WriteXLSXInsufficientArgumentError unless defined?(WriteXLSXInsufficientArgumentError)
WriteXLSXDimensionError = Xlsxrb::Adapters::Writexlsx::WriteXLSXDimensionError unless defined?(WriteXLSXDimensionError)
WriteXLSXOptionParameterError = Xlsxrb::Adapters::Writexlsx::WriteXLSXOptionParameterError unless defined?(WriteXLSXOptionParameterError)
