# frozen_string_literal: true

# rbs_inline: enabled

require_relative "roo/constants"
require_relative "roo/link"
require_relative "roo/font"
require_relative "roo/formatters/base"
require_relative "roo/formatters/csv"
require_relative "roo/formatters/matrix"
require_relative "roo/formatters/xml"
require_relative "roo/formatters/yaml"
require_relative "roo/base"
require_relative "roo/coordinate"
require_relative "roo/utils"
require_relative "roo/format"
require_relative "roo/cell"
require_relative "roo/sheet"
require_relative "roo/excelx"
require_relative "roo/spreadsheet"

module Xlsxrb
  module Adapters
    # Roo adapter module providing a drop-in compatible interface for Roo (.xlsx & .xlsm)
    # without hijacking the global ::Roo namespace by default.
    module Roo
      # Converts an Xlsxrb::Elements::Workbook to an adapter Excelx workbook.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @return [Excelx]
      #: (Xlsxrb::Elements::Workbook xlsxrb_workbook) -> Excelx
      def self.from_xlsxrb(xlsxrb_workbook)
        Excelx.from_xlsxrb(xlsxrb_workbook)
      end
    end
  end
end
