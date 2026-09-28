# frozen_string_literal: true

# rbs_inline: enabled

require_relative "caxlsx/constants"
require_relative "caxlsx/util"
require_relative "caxlsx/relationship"
require_relative "caxlsx/color"
require_relative "caxlsx/doc_props"
require_relative "caxlsx/rich_text"
require_relative "caxlsx/styles"
require_relative "caxlsx/col"
require_relative "caxlsx/cell"
require_relative "caxlsx/row"
require_relative "caxlsx/page_setup"
require_relative "caxlsx/sheet_view"
require_relative "caxlsx/sheet_protection"
require_relative "caxlsx/auto_filter"
require_relative "caxlsx/table"
require_relative "caxlsx/pivot_table"
require_relative "caxlsx/data_validation"
require_relative "caxlsx/conditional_formatting"
require_relative "caxlsx/comments"
require_relative "caxlsx/hyperlink"
require_relative "caxlsx/drawing"
require_relative "caxlsx/charts"
require_relative "caxlsx/worksheet"
require_relative "caxlsx/workbook"
require_relative "caxlsx/package"
require_relative "caxlsx/streaming_package"

module Xlsxrb
  module Adapters
    # Caxlsx (Axlsx) adapter module providing a drop-in compatible interface for caxlsx
    # without hijacking the global ::Axlsx / ::Caxlsx namespace by default.
    module Caxlsx
      # Converts an adapter Package or Workbook to an immutable Xlsxrb::Elements::Workbook.
      #
      # @param obj [Package, Workbook]
      # @return [Xlsxrb::Elements::Workbook]
      #: (Package | Workbook obj) -> Xlsxrb::Elements::Workbook
      def self.to_xlsxrb(obj)
        obj.to_xlsxrb
      end
    end

    Axlsx = Caxlsx
  end
end
