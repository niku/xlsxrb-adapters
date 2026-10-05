# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "xlsxrb"
require_relative "xml"
require_relative "row"
require_relative "columns"

module Xlsxrb
  module Adapters
    module Xlsxtream
      # Represents a worksheet stream writing rows directly to the ZIP package.
      class Worksheet
        # @return [Integer, nil]
        attr_reader :id

        # @return [String, nil]
        attr_reader :name

        # @return [Array<Array<Object>>]
        attr_reader :rows

        # @return [Hash{Symbol => untyped}]
        attr_reader :options

        # @param io [Object] Target IO / ZIP writer.
        # @param options [Hash{Symbol => untyped}] Worksheet options.
        #: (untyped io, ?Hash[Symbol, untyped] options) -> void
        def initialize(io, options = {})
          @io = io
          @rownum = 1
          @closed = false
          @options = options
          @id = options[:id]
          @name = options[:name]
          @retain_rows = options.fetch(:retain_rows, true)
          @rows = []

          write_header
        end

        # Appends a row of values to the worksheet.
        #
        # @param row [Enumerable<Object>] Row cell values.
        # @return [void]
        #: (untyped row) -> void
        def <<(row)
          row_array = if row.respond_to?(:to_a)
                        row.to_a
                      else
                        Array(row)
                      end
          @rows << row_array if @retain_rows

          @io << Row.new(row, @rownum, @options).to_xml
          @rownum += 1
        end
        alias add_row <<

        # Finalizes the worksheet and writes the XML closing tags.
        #
        # @return [void]
        #: () -> void
        def close
          return if @closed

          write_footer
          @closed = true
        end

        # Returns whether the worksheet is closed.
        #
        # @return [Boolean]
        #: () -> bool
        def closed?
          @closed
        end

        # Converts worksheet data to an immutable Xlsxrb::Elements::Worksheet.
        #
        # @param default_style_index [Integer] Default style index.
        # @return [Xlsxrb::Elements::Worksheet]
        def to_xlsxrb(default_style_index = 0)
          elements_columns = nil
          if (cols = @options[:columns]) && !cols.empty?
            elements_columns = cols.map.with_index do |col, idx|
              width_chars = col[:width_chars]
              width_pixels = col[:width_pixels]
              width_pixels ||= ((((width_chars * 7.0) + 5) / 7) * 256).truncate / 256.0 if width_chars
              Xlsxrb::Elements::Column.new(
                index: idx,
                width: width_pixels,
                custom_width: !width_pixels.nil?
              )
            end
          end

          elements_rows = @rows.map.with_index do |row_values, r_idx|
            elements_cells = row_values.map.with_index do |val, c_idx|
              next nil if val.nil?

              val = auto_format_value(val) if @options[:auto_format] && val.is_a?(String)
              next nil if val == ""

              style_idx = case val
                          when Time, DateTime then 2
                          when Date then 1
                          else default_style_index
                          end

              Xlsxrb::Elements::Cell.fast_create(r_idx, c_idx, val, style_idx)
            end

            Xlsxrb::Elements::Row.new(
              index: r_idx,
              cells: elements_cells
            )
          end

          Xlsxrb::Elements::Worksheet.new(
            name: @name,
            rows: elements_rows,
            columns: elements_columns
          )
        end

        private

        def auto_format_value(value)
          case value
          when Row::TRUE_STRING then true
          when Row::FALSE_STRING then false
          when Row::NUMBER_PATTERN
            value.include?(".") ? value.to_f : value.to_i
          when Row::DATE_PATTERN
            begin
              Date.parse(value)
            rescue StandardError
              value
            end
          when Row::TIME_PATTERN
            begin
              DateTime.parse(value)
            rescue StandardError
              value
            end
          else
            value
          end
        end

        def write_header
          @io << XML.header
          @io << XML.strip(<<-XML)
            <worksheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
          XML

          columns = Array(@options[:columns])
          @io << Columns.new(columns).to_xml unless columns.empty?

          @io << XML.strip(<<-XML)
              <sheetData>
          XML
        end

        def write_footer
          @io << XML.strip(<<-XML)
              </sheetData>
            </worksheet>
          XML
        end
      end
    end
  end
end
