# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require_relative "creek/utils"
require_relative "creek/styles"
require_relative "creek/shared_strings"
require_relative "creek/drawing"
require_relative "creek/sheet"
require_relative "creek/book"

module Xlsxrb
  module Adapters
    # Creek adapter module providing a drop-in compatible interface for Creek
    # stream parser without hijacking the global ::Creek namespace by default.
    module Creek
      VERSION = "2.6.3"

      # Converts an Xlsxrb::Elements::Workbook into a Creek Book adapter instance.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @param options [Hash]
      # @return [Book]
      #: (Xlsxrb::Elements::Workbook xlsxrb_workbook, ?Hash[Symbol, untyped] options) -> Book
      def self.from_xlsxrb(xlsxrb_workbook, options = {})
        Book.from_xlsxrb(xlsxrb_workbook, options)
      end
    end
  end
end
