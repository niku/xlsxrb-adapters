# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require_relative "helpers"

module Xlsxrb
  module Adapters
    module Xsv
      # Represents a single worksheet from a workbook matching Xsv::Sheet.
      # Designed for worksheets with a single table of data, supporting both
      # array mode (default) and hash mode (via {#parse_headers!}).
      class Sheet
        [Enumerable].each { |m| include m }
        include Helpers

        ROW_RE = %r{<row\b([^>]*?)>(.*?)</row>|<row\b([^>]*?)/>}m
        CELL_RE = %r{<c\b([^>]*?)(?:>\s*(?:<f\b[^>]*?(?:/>|>([^<]*)</f>))?\s*(?:<v>([^<]*)</v>)?\s*(?:<is>(.*?)</is>)?.*?</c>|/>)}m

        # Returns current mode (:array or :hash).
        # @return [Symbol]
        #: Symbol
        attr_reader :mode

        # Returns sheet ID.
        # @return [Integer]
        #: Integer
        attr_reader :id

        # Returns sheet name.
        # @return [String]
        #: String
        attr_reader :name

        # Number of rows to skip at the top of the worksheet.
        # @return [Integer]
        #: Integer
        attr_accessor :row_skip

        # Underlying StreamSheet instance if available.
        # @return [StreamSheet, nil]
        #: StreamSheet?
        attr_reader :stream_sheet

        # Returns sheet dimension reference string (e.g. "A1:Z50"), or nil.
        # @return [String, nil]
        #: () -> String?
        def dimension
          @stream_sheet&.dimension || raw_sheet_xml[/<dimension\s+ref="([^"]+)"/, 1]
        end

        # Returns the 1-based index of the first row.
        # @return [Integer]
        #: () -> Integer
        def first_row
          @stream_sheet&.first_row || 1
        end

        # Returns the 1-based index of the first column.
        # @return [Integer]
        #: () -> Integer
        def first_column
          @stream_sheet&.first_column || 1
        end
        alias first_col first_column

        # Returns the 1-based index of the last row.
        # @return [Integer]
        #: Integer
        attr_reader :last_row

        # Returns the 1-based index of the last column.
        # @return [Integer]
        #: () -> Integer
        def last_column
          @column_count
        end
        alias last_col last_column

        # Returns whether the workbook uses the 1904 date system.
        # @return [Boolean]
        #: () -> bool
        def date1904?
          @workbook ? @workbook.date1904? : false
        end

        # Returns whether empty trailing rows are trimmed.
        # @return [Boolean]
        #: () -> bool
        def trim_empty_rows?
          @workbook ? @workbook.trim_empty_rows? : false
        end
        alias trim_empty_rows trim_empty_rows?

        # @param workbook [Workbook]
        # @param sheet_source [String, Proc, IO]
        # @param ids [Hash[Symbol, untyped]]
        # @param last_row [Integer, nil]
        # @param column_count [Integer, nil]
        # @param stream_sheet [StreamSheet, nil]
        #: (Workbook workbook, untyped sheet_source, Hash[Symbol, untyped] ids, ?Integer? last_row, ?Integer? column_count, ?stream_sheet: StreamSheet?) -> void
        def initialize(workbook, sheet_source, ids, last_row = nil, column_count = nil, stream_sheet: nil)
          @workbook = workbook
          @sheet_source = sheet_source
          @stream_sheet = stream_sheet
          @id = (ids[:sheet_id] || ids[:sheetId] || 1).to_i
          @name = ids[:name].to_s
          @headers = []
          @mode = :array
          @row_skip = 0
          @hidden = ["hidden", :hidden].include?(ids[:state])

          if last_row && column_count
            @last_row = last_row
            @column_count = column_count
          else
            @last_row, @column_count = compute_bounds
          end
        end

        # Returns true if the worksheet is hidden.
        # @return [Boolean]
        #: () -> bool
        def hidden?
          @hidden
        end

        # Human-readable representation.
        # @return [String]
        #: () -> String
        def inspect
          "#<#{self.class.name}:#{object_id} mode=#{@mode}>"
        end

        # Switch to hash mode using the first data row (after row_skip) as header names.
        #
        # @return [self]
        # @raise [Xsv::DuplicateHeaders] If duplicate non-nil headers are detected.
        #: () -> self
        def parse_headers!
          @headers = parse_headers

          # Check for duplicate headers, but don't care about nil columns
          duplicate_header = @headers.compact.detect { |h| @headers.count(h) > 1 }
          raise Xsv::DuplicateHeaders, "Duplicate header '#{duplicate_header}' found, consider parsing this sheet in array mode." if duplicate_header

          @mode = :hash
          self
        end

        # Returns the header row array for the worksheet.
        # @return [Array<Object>]
        #: () -> Array[untyped]
        def headers
          if @headers.any?
            @headers
          else
            parse_headers
          end
        end

        # Iterate over worksheet rows, yielding arrays in :array mode or hashes in :hash mode.
        #
        # @yield [row]
        # @yieldparam row [Array<Object>, Hash<Object, Object>]
        # @return [true, Enumerator]
        #: () { (untyped) -> void } -> (bool | Enumerator[untyped, void])
        #: () -> Enumerator[untyped, void]
        def each_row(&block)
          return to_enum(__method__) unless block

          row_index = 0
          curr_row_num = 0
          adjusted_last_row = @last_row - @row_skip

          xml = raw_sheet_xml
          sd_start = xml.index("<sheetData")
          return true unless sd_start

          sd_open_end = xml.index(">", sd_start)
          return true unless sd_open_end

          sd_end = xml.index("</sheetData>", sd_open_end)
          sd_xml = xml.byteslice(sd_open_end + 1, (sd_end || xml.bytesize) - sd_open_end - 1)
          return true if sd_xml.nil? || sd_xml.empty?

          sd_xml.scan(ROW_RE) do |attrs1, body, attrs2|
            attrs = attrs1 || attrs2
            r_attr = attrs[/r="(\d+)"/, 1]
            curr_row_num = r_attr ? r_attr.to_i : (curr_row_num + 1)
            next if curr_row_num <= @row_skip

            adjusted_row_num = curr_row_num - @row_skip
            row_index += 1

            # Skip header row in hash mode
            next if adjusted_row_num == 1 && @mode == :hash

            # Pad empty rows
            while row_index < adjusted_row_num
              block.call(empty_row)
              row_index += 1
            end

            # Do not return empty trailing rows
            next if row_index > adjusted_last_row

            # Build and yield current row
            if @mode == :array
              row_array = Array.new(@column_count)
              if body
                col_idx = 0
                body.scan(CELL_RE) do |c_attrs, _f, v, is|
                  r = c_attrs[/r="([A-Za-z0-9]+)"/, 1]
                  s = c_attrs[/s="(\d+)"/, 1]
                  t = c_attrs[/t="([a-zA-Z]+)"/, 1]
                  c_idx = r ? column_index(r) : col_idx
                  val = format_cell_value(t, s, v, is)
                  row_array.concat(Array.new(c_idx - row_array.size + 1)) if c_idx >= row_array.size
                  row_array[c_idx] = val
                  col_idx = c_idx + 1
                end
              end
              block.call(row_array)
            else
              row_hash = empty_row.dup
              if body
                col_idx = 0
                body.scan(CELL_RE) do |c_attrs, _f, v, is|
                  r = c_attrs[/r="([A-Za-z0-9]+)"/, 1]
                  s = c_attrs[/s="(\d+)"/, 1]
                  t = c_attrs[/t="([a-zA-Z]+)"/, 1]
                  c_idx = r ? column_index(r) : col_idx
                  hdr = @headers[c_idx]
                  row_hash[hdr] = format_cell_value(t, s, v, is) if hdr
                  col_idx = c_idx + 1
                end
              end
              block.call(row_hash)
            end
          end

          true
        end
        alias each each_row

        # Converts this Xsv Sheet instance into an Xlsxrb::Elements::Worksheet.
        #
        # @return [Xlsxrb::Elements::Worksheet]
        #: () -> Xlsxrb::Elements::Worksheet
        def to_xlsxrb
          rows = []
          orig_mode = @mode
          orig_skip = @row_skip
          @mode = :array
          @row_skip = 0
          begin
            each_with_index do |row_data, r_idx|
              cells = []
              row_data.each_with_index do |val, c_idx|
                next if val.nil?

                cells << Elements::Cell.new(row_index: r_idx, column_index: c_idx, value: val, date1904: date1904?)
              end
              rows << Elements::Row.new(index: r_idx, cells: cells)
            end
          ensure
            @mode = orig_mode
            @row_skip = orig_skip
          end
          state = hidden? ? :hidden : :visible
          Elements::Worksheet.new(name: @name, rows: rows, state: state, date1904: date1904?, dimension: dimension)
        end

        # Iterates over row values directly as Arrays without constructing hash records.
        #
        # @param type_cast [Boolean] Whether to cast date/time serial values.
        # @param trim_empty_rows [Boolean, nil]
        # @yield [values]
        # @yieldparam values [Array<Object>]
        # @return [void, Enumerator]
        #: (?type_cast: bool, ?trim_empty_rows: bool?) { (Array[untyped]) -> void } -> void
        #: (?type_cast: bool, ?trim_empty_rows: bool?) -> Enumerator[Array[untyped], void]
        def each_row_values(type_cast: true, trim_empty_rows: nil, &block)
          return enum_for(:each_row_values, type_cast: type_cast, trim_empty_rows: trim_empty_rows) unless block

          if @stream_sheet && @row_skip.zero?
            @stream_sheet.each_row_values(type_cast: type_cast, trim_empty_rows: trim_empty_rows, &block)
          else
            each_row do |row|
              block.call(row.is_a?(Hash) ? row.values : row)
            end
          end
        end

        # Get row by 0-based index or range.
        #
        # @param number_or_range [Integer, Range]
        # @return [Array, Hash]
        #: (untyped number_or_range) -> untyped
        def [](number_or_range)
          case number_or_range
          when Range
            rows = []
            each_with_index do |row, i|
              rows << row if number_or_range.cover?(i)
            end
            rows
          when Integer
            each_with_index do |row, i|
              return row if i == number_or_range
            end
            empty_row
          else
            raise ArgumentError, "Expected Integer or Range, got #{number_or_range.class}"
          end
        end

        private

        #: () -> (Array[untyped] | Hash[untyped, untyped])
        def empty_row
          if @mode == :hash
            @headers.compact.zip([]).to_h
          else
            Array.new(@column_count.to_i)
          end
        end

        #: () -> Array[untyped]
        def parse_headers
          case @mode
          when :array
            first
          when :hash
            @mode = :array
            h = headers
            @mode = :hash
            h
          end || []
        end

        #: () -> String
        def raw_sheet_xml
          @raw_sheet_xml ||= if @sheet_source.is_a?(String)
                               @sheet_source
                             elsif @sheet_source.respond_to?(:call)
                               buf = +""
                               @sheet_source.call { |chunk| buf << chunk }
                               buf
                             elsif @sheet_source.respond_to?(:read)
                               @sheet_source.rewind if @sheet_source.respond_to?(:rewind)
                               @sheet_source.read || ""
                             else
                               ""
                             end
        end

        #: () -> [Integer, Integer]
        def compute_bounds
          if @stream_sheet && !@workbook&.trim_empty_rows
            dim_last_row = @stream_sheet.last_row
            dim_last_col = @stream_sheet.last_column
            return [dim_last_row, dim_last_col] if dim_last_row && dim_last_col
          end

          xml = raw_sheet_xml
          max_column = -1
          dim_match = xml[/<dimension\s+ref="([^"]+)"/, 1]
          if dim_match
            _first_cell, last_cell = dim_match.split(":")
            if last_cell
              max_column = column_index(last_cell)
              unless @workbook&.trim_empty_rows
                max_row = last_cell[/\d+$/].to_i
                return [max_row, max_column.negative? ? 0 : max_column + 1]
              end
            end
          end

          # Scan sheetData looking for cell bounds
          max_row = 0
          curr_row = 0
          xml.scan(ROW_RE) do |attrs1, body, attrs2|
            attrs = attrs1 || attrs2
            r_attr = attrs[/r="(\d+)"/, 1]
            curr_row = r_attr ? r_attr.to_i : (curr_row + 1)
            next unless body

            col_idx = 0
            body.scan(CELL_RE) do |c_attrs, _f, v, is|
              r = c_attrs[/r="([A-Za-z0-9]+)"/, 1]
              col = r ? column_index(r) : col_idx
              if (v && !v.empty?) || (is && !is.empty?)
                max_column = col if col > max_column
                max_row = curr_row if curr_row > max_row
              end
              col_idx = col + 1
            end
          end

          [max_row, max_column.negative? ? 0 : max_column + 1]
        end

        #: (String? cell_type, String? style_id, String? val_str, String? is_xml) -> untyped
        def format_cell_value(cell_type, style_id, val_str, is_xml)
          val = if is_xml
                  text = +""
                  is_xml.scan(%r{<t\b[^>]*>(.*?)</t>}m) { |m| text << m[0] }
                  text
                else
                  val_str&.strip
                end
          return nil if val.nil? || val.empty?

          case cell_type
          when "s"
            @workbook.shared_strings[val.to_i]
          when "str", "inlineStr"
            val.strip
          when "e"
            nil
          when "b"
            val == "1"
          when "d"
            DateTime.parse(val)
          when nil, "n"
            if style_id
              fmt = @workbook.get_num_fmt(style_id.to_i)
              parse_number_format(val, fmt, style_id: style_id.to_i)
            else
              parse_number(val)
            end
          else
            raise Xsv::Error, "Encountered unknown column type #{cell_type}"
          end
        end
      end
    end
  end
end
