# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require_relative "xlsxtream/version"
require_relative "xlsxtream/errors"
require_relative "xlsxtream/xml"
require_relative "xlsxtream/shared_string_table"
require_relative "xlsxtream/columns"
require_relative "xlsxtream/row"
require_relative "xlsxtream/zip_kit_writer"
require_relative "xlsxtream/worksheet"
require_relative "xlsxtream/workbook"

module Xlsxrb
  module Adapters
    # Xlsxtream adapter module providing a drop-in compatible interface for xlsxtream
    # without hijacking the global ::Xlsxtream namespace by default.
    module Xlsxtream
      # Opens or creates a new workbook compatible with Xlsxtream::Workbook.open.
      #
      # @param output [String, IO, Object] Destination target.
      # @param options [Hash{Symbol => untyped}] Options.
      # @yield [workbook]
      # @yieldparam workbook [Workbook]
      # @return [Workbook, untyped]
      def self.open(output, options = {}, &)
        Workbook.open(output, options, &)
      end

      # Creates a new workbook instance.
      #
      # @param output [String, IO, Object] Destination target.
      # @param options [Hash{Symbol => untyped}] Options.
      # @return [Workbook]
      def self.new(output, options = {})
        Workbook.new(output, options)
      end

      # Converts an adapter Workbook to an immutable Xlsxrb::Elements::Workbook.
      #
      # @param obj [Workbook]
      # @return [Xlsxrb::Elements::Workbook]
      #: (Workbook obj) -> Xlsxrb::Elements::Workbook
      def self.to_xlsxrb(obj)
        obj.to_xlsxrb
      end

      # Converts an Xlsxrb::Elements::Workbook into an Xlsxtream Workbook adapter instance.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @param output [Object, nil]
      # @param options [Hash{Symbol => untyped}]
      # @return [Workbook]
      def self.from_xlsxrb(xlsxrb_workbook, output = nil, options = {})
        Workbook.from_xlsxrb(xlsxrb_workbook, output, options)
      end
    end
  end
end
