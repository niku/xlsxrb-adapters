# frozen_string_literal: true

# rbs_inline: enabled

require_relative "ruby_xl/formula"
require_relative "ruby_xl/color"
require_relative "ruby_xl/styles"
require_relative "ruby_xl/reference"
require_relative "ruby_xl/data_type"
require_relative "ruby_xl/text"
require_relative "ruby_xl/column_range"
require_relative "ruby_xl/merged_cells"
require_relative "ruby_xl/defined_names"
require_relative "ruby_xl/cell"
require_relative "ruby_xl/row"
require_relative "ruby_xl/worksheet"
require_relative "ruby_xl/workbook"
require_relative "ruby_xl/parser"

module Xlsxrb
  module Adapters
    # RubyXL adapter module providing a drop-in compatible interface for RubyXL
    # without hijacking the global ::RubyXL namespace.
    module RubyXL
      # Converts an Xlsxrb::Elements::Workbook to an adapter Workbook.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @return [Workbook]
      #: (Xlsxrb::Elements::Workbook xlsxrb_workbook) -> Workbook
      def self.from_xlsxrb(xlsxrb_workbook)
        Workbook.from_xlsxrb(xlsxrb_workbook)
      end
    end
  end
end
