# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "stringio"
require "time"
require "xlsxrb"
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
      SUBPATH_FEATURES = %w[
        rubyXL
        rubyXL/cell
        rubyXL/worksheet
        rubyXL/parser
        rubyXL/convenience_methods
        rubyXL/convenience_methods/cell
        rubyXL/convenience_methods/color
        rubyXL/convenience_methods/font
        rubyXL/convenience_methods/workbook
        rubyXL/convenience_methods/worksheet
      ].freeze

      # Converts an Xlsxrb::Elements::Workbook to an adapter Workbook.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @return [Workbook]
      #: (Xlsxrb::Elements::Workbook xlsxrb_workbook) -> Workbook
      def self.from_xlsxrb(xlsxrb_workbook)
        Workbook.from_xlsxrb(xlsxrb_workbook)
      end

      # Activates RubyXL drop-in mode by binding ::RubyXL and registering sub-path features.
      #
      # @return [void]
      #: () -> void
      def self.activate!
        Object.send(:remove_const, :RubyXL) if Object.const_defined?(:RubyXL, false)
        Object.const_set(:RubyXL, self)
        register_subpath_features!
      end

      # Registers rubyXL sub-paths (such as convenience_methods/*) in $LOADED_FEATURES
      # so explicit sub-path requires succeed without loading conflicting gem files.
      #
      # @return [void]
      #: () -> void
      def self.register_subpath_features!
        SUBPATH_FEATURES.each do |sub|
          feature = "#{sub}.rb"
          $LOADED_FEATURES << feature unless $LOADED_FEATURES.include?(feature)
          $LOAD_PATH.each do |lp|
            full = File.join(lp, feature)
            $LOADED_FEATURES << full if File.exist?(full) && !$LOADED_FEATURES.include?(full)
          end
        end
      end
    end
  end
end
