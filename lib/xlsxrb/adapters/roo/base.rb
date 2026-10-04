# frozen_string_literal: true

# rbs_inline: enabled

require_relative "formatters/base"
require_relative "formatters/csv"
require_relative "formatters/matrix"
require_relative "formatters/xml"
require_relative "formatters/yaml"

module Xlsxrb
  module Adapters
    module Roo
      # Base spreadsheet abstraction matching Roo::Base.
      class Base
        [
          Enumerable,
          Formatters::Base,
          Formatters::CSV,
          Formatters::Matrix,
          Formatters::XML,
          Formatters::YAML
        ].each { |m| include m }

        attr_reader :headers
        attr_accessor :header_line

        def initialize
          @header_line = 1
          @headers = nil
          @default_sheet = nil
          @cleaned = {}
        end

        # Returns array of sheet names.
        #
        # @return [Array<String>]
        #: () -> Array[String]
        def sheets
          raise NotImplementedError
        end

        # Returns current default sheet name.
        #
        # @return [String]
        #: () -> String
        def default_sheet
          @default_sheet ||= sheets.first || "Sheet1"
        end

        # Sets current default sheet.
        #
        # @param sheet [String, Integer]
        # @return [String]
        #: (untyped sheet) -> untyped
        def default_sheet=(sheet)
          validate_sheet!(sheet)
          @default_sheet = sheet.is_a?(String) ? sheet : sheets[sheet]
        end

        # Switches default sheet or returns tuple [sheet_name, self].
        #
        # @param index [String, Integer]
        # @param name [Boolean]
        # @return [Base, Array(String, Base)]
        #: (untyped index, ?bool name) -> untyped
        def sheet(index, name = false)
          self.default_sheet = index
          name ? [default_sheet, self] : self
        end

        # Iterates through all worksheets.
        #
        # @yield [name, sheet]
        # @return [Enumerator, void]
        #: () { ([String, untyped]) -> void } -> void
        #: () -> Enumerator[[String, untyped], void]
        def each_with_pagename
          return to_enum(:each_with_pagename) { sheets.size } unless block_given?

          sheets.each do |s|
            yield sheet(s, true)
          end
        end

        # First non-empty row.
        #
        # @param sheet [String, Integer, nil]
        # @return [Integer, nil]
        #: (?untyped sheet) -> Integer?
        def first_row(sheet = default_sheet)
          sheet_for(sheet)&.first_row
        end

        # Last non-empty row.
        #
        # @param sheet [String, Integer, nil]
        # @return [Integer, nil]
        #: (?untyped sheet) -> Integer?
        def last_row(sheet = default_sheet)
          sheet_for(sheet)&.last_row
        end

        # First non-empty column.
        #
        # @param sheet [String, Integer, nil]
        # @return [Integer, nil]
        #: (?untyped sheet) -> Integer?
        def first_column(sheet = default_sheet)
          sheet_for(sheet)&.first_column
        end

        # Last non-empty column.
        #
        # @param sheet [String, Integer, nil]
        # @return [Integer, nil]
        #: (?untyped sheet) -> Integer?
        def last_column(sheet = default_sheet)
          sheet_for(sheet)&.last_column
        end
        alias first_col first_column
        alias last_col last_column

        # Sheet dimension reference string (e.g. "A1:Z50"), or nil.
        #
        # @param sheet [String, Integer, nil]
        # @return [String, nil]
        #: (?untyped sheet) -> String?
        def dimension(sheet = default_sheet)
          target_sheet = sheet_for(sheet)
          return nil unless target_sheet

          if target_sheet.respond_to?(:dimension) && target_sheet.dimension
            target_sheet.dimension
          elsif target_sheet.respond_to?(:dimensions)
            target_sheet.dimensions
          end
        end
        alias dimensions dimension

        # Returns whether workbook uses 1904 date system.
        #
        # @return [Boolean]
        #: () -> bool
        def date1904?
          false
        end

        # First non-empty column as a letter string (e.g. "A").
        #
        # @param sheet [String, Integer, nil]
        # @return [String]
        #: (?untyped sheet) -> String
        def first_column_as_letter(sheet = default_sheet)
          Utils.number_to_letter(first_column(sheet))
        end

        # Last non-empty column as a letter string (e.g. "Z").
        #
        # @param sheet [String, Integer, nil]
        # @return [String]
        #: (?untyped sheet) -> String
        def last_column_as_letter(sheet = default_sheet)
          Utils.number_to_letter(last_column(sheet))
        end

        # Returns row values as Array.
        #
        # @param rownumber [Integer]
        # @param sheet [String, Integer, nil]
        # @return [Array<Object>]
        #: (Integer rownumber, ?untyped sheet) -> Array[untyped]
        def row(rownumber, sheet = default_sheet)
          sheet_for(sheet)&.row(rownumber) || []
        end

        # Returns column values as Array.
        #
        # @param column_number [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Array<Object>]
        #: (untyped column_number, ?untyped sheet) -> Array[untyped]
        def column(column_number, sheet = default_sheet)
          c_num = column_number.is_a?(String) || column_number.is_a?(Symbol) ? Utils.letter_to_number(column_number) : column_number.to_i
          sheet_for(sheet)&.column(c_num) || []
        end

        # Returns cell value.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Object, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> untyped
        def cell(row, col, sheet = default_sheet)
          r, c = normalize(row, col)
          sheet_for(sheet)&.cell_at([r, c])&.value
        end

        # Returns whether cell is empty.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Boolean]
        #: (untyped row, untyped col, ?untyped sheet) -> bool
        def empty?(row, col, sheet = default_sheet)
          target_sheet = sheet_for(sheet)
          return true unless target_sheet

          r, c = normalize(row, col)
          c_obj = target_sheet.cells[[r, c]]
          return true if c_obj.nil? || c_obj.empty?
          return true if c_obj.value.nil? || (c_obj.value.is_a?(String) && c_obj.value.empty?)

          f_row = target_sheet.first_row
          l_row = target_sheet.last_row
          f_col = target_sheet.first_column
          l_col = target_sheet.last_column
          return true unless f_row && l_row && f_col && l_col

          r < f_row || r > l_row || c < f_col || c > l_col
        end

        # Sets cell value in-memory.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param value [Object]
        # @param sheet [String, Integer, nil]
        # @return [void]
        #: (untyped row, untyped col, untyped value, ?untyped sheet) -> void
        def set(row, col, value, sheet = default_sheet)
          r, c = normalize(row, col)
          coord = Excelx::Coordinate.new(r, c)
          target_sheet = sheet_for(sheet)
          return unless target_sheet

          type = case value
                 when Numeric then :number
                 when Date then :date
                 when Time, DateTime then :datetime
                 when TrueClass, FalseClass then :boolean
                 else :string
                 end

          cell_obj = case type
                     when :string then Excelx::Cell::String.new(value, nil, nil, nil, coord)
                     when :boolean then Excelx::Cell::Boolean.new(value, nil, nil, nil, coord)
                     when :number then Excelx::Cell::Number.new(value, nil, "General", nil, nil, coord)
                     when :date then Excelx::Cell::Date.new(value, nil, "yyyy-mm-dd", nil, nil, nil, coord)
                     when :datetime then Excelx::Cell::DateTime.new(value, nil, "yyyy-mm-dd hh:mm:ss", nil, nil, nil, coord)
                     else Excelx::Cell::String.new(value.to_s, nil, nil, nil, coord)
                     end

          target_sheet.cells[coord] = cell_obj
          target_sheet.reset_bounds!
        end

        # Returns document and sheet information summary string.
        #
        # @return [String]
        #: () -> String
        def info
          without_changing_default_sheet do
            fn = @filename ? File.basename(@filename.to_s) : "spreadsheet"
            result = "File: #{fn}\nNumber of sheets: #{sheets.size}\nSheets: #{sheets.join(", ")}\n"
            sheets.each_with_index do |s_name, idx|
              self.default_sheet = s_name
              result << "Sheet #{idx + 1}:\n"
              f_row = first_row(s_name)
              if f_row
                result << "  First row: #{f_row}\n"
                result << "  Last row: #{last_row(s_name)}\n"
                result << "  First column: #{first_column_as_letter(s_name)}\n"
                result << "  Last column: #{last_column_as_letter(s_name)}"
              else
                result << "  - empty -"
              end
              result << "\n" if s_name != sheets.last
            end
            result
          end
        end

        # Iterates through rows or maps to columns hash.
        #
        # @param options [Hash]
        # @yield [row]
        # @return [Enumerator, void]
        #: (?Hash[untyped, untyped] options) { (untyped) -> void } -> void
        #: (?Hash[untyped, untyped] options) -> Enumerator[untyped, void]
        def each(options = {}, &block)
          return to_enum(:each, options) unless block

          l_row = last_row
          return unless l_row

          if options.empty?
            1.upto(l_row) do |line|
              block.call(row(line))
            end
          else
            clean_sheet_if_needed(options)
            search_or_set_header(options)

            f_col = first_column || 1
            l_col = last_column || 1
            hdrs = @headers || (f_col..l_col).to_h do |col|
              [cell(@header_line, col), col]
            end

            header_map = hdrs.is_a?(Array) ? hdrs.to_h : hdrs

            @header_line.upto(l_row) do |line|
              row_hash = {}
              header_map.each do |k, col_idx|
                row_hash[k] = cell(line, col_idx)
              end
              block.call(row_hash)
            end
          end
        end

        # Iterates over row values directly as Arrays.
        #
        # @param sheet [String, Integer, nil]
        # @yield [values]
        # @yieldparam values [Array<Object>]
        # @return [Enumerator, void]
        #: (?untyped sheet) { (Array[untyped]) -> void } -> void
        #: (?untyped sheet) -> Enumerator[Array[untyped], void]
        def each_row_values(sheet = default_sheet, &block)
          return enum_for(:each_row_values, sheet) unless block

          l_row = last_row(sheet)
          return unless l_row

          1.upto(l_row) do |line|
            block.call(row(line, sheet))
          end
        end

        # Parses spreadsheet content into an Array of rows or hashes.
        #
        # @param options [Hash]
        # @yield [row]
        # @return [Array<Object>]
        #: (?Hash[untyped, untyped] options) ?{ (untyped) -> untyped } -> Array[untyped]
        def parse(options = {}, &block)
          results = each(options).map do |r|
            block ? block.call(r) : r
          end
          options[:headers] == true ? results : results.drop(1)
        end

        # Finds the header row matching query expressions.
        #
        # @param query [Array<Object>]
        # @param return_headers [Boolean]
        # @return [Integer, Array<Object>]
        #: (Array[untyped] query, ?bool return_headers) -> untyped
        def row_with(query, return_headers = false)
          line_no = 0
          closest_mismatched = []

          each do |r_data|
            line_no += 1
            matched = query.map do |q|
              r_data.grep(q).first
            end.compact

            if matched.length == query.length
              @header_line = line_no
              return return_headers ? matched : line_no
            else
              closest_mismatched = matched if matched.length > closest_mismatched.length
              break if line_no > 100
            end
          end

          missing = query.select { |q| closest_mismatched.grep(q).empty? }
          raise HeaderRowNotFoundError, missing
        end

        # Finds rows by condition or index.
        #
        # @param args [Array<Object>]
        # @return [Object]
        #: (*untyped args) -> untyped
        def find(*args)
          opts = args.last.is_a?(Hash) ? args.pop : {}
          case args[0]
          when Integer
            find_by_row(args[0])
          when :all
            find_by_conditions(opts)
          else
            raise ArgumentError, "unexpected arg #{args[0].inspect}, pass a row index or :all"
          end
        end

        # Handles method missing for coordinate access (e.g. `xlsx.a1`, `xlsx.b2`) and defined names.
        #: (Symbol method_name, *untyped args) -> untyped
        def method_missing(method_name, *args)
          str = method_name.to_s
          if str =~ /^([a-z]+)(\d+)$/
            c = Utils.letter_to_number(::Regexp.last_match(1))
            r = ::Regexp.last_match(2).to_i
            s = args.first
            cell(r, c, s)
          else
            super
          end
        end

        #: (Symbol method_name, ?bool include_private) -> bool
        def respond_to_missing?(method_name, include_private = false)
          method_name.to_s.match?(/^([a-z]+)(\d+)$/) || super
        end

        # Closes reader and clears in-memory state.
        #: () -> nil
        def close
          nil
        end

        # Validates that sheet name or index exists.
        #
        # @param sheet [String, Integer, nil]
        # @return [void]
        #: (untyped sheet) -> void
        def validate_sheet!(sheet)
          case sheet
          when nil
            raise ArgumentError, "Error: sheet 'nil' not valid"
          when Integer
            sheets.fetch(sheet) do
              raise RangeError, "sheet index #{sheet} not found"
            end
          when String
            raise RangeError, "sheet '#{sheet}' not found" unless sheets.include?(sheet)
          else
            raise TypeError, "not a valid sheet type: #{sheet.inspect}"
          end
        end

        # Normalizes row and column arguments to [row_integer, col_integer].
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @return [Array(Integer, Integer)]
        #: (untyped row, untyped col) -> [Integer, Integer]
        def normalize(row, col)
          r = row
          c = col

          r, c = c, r if r.is_a?(String) && (c.is_a?(Integer) || c.is_a?(Numeric))

          c_num = if c.is_a?(String) || c.is_a?(Symbol)
                    Utils.letter_to_number(c)
                  else
                    c.to_i
                  end

          [r.to_i, c_num]
        end

        # Returns internal sheet instance for given sheet name or index.
        #
        # @param sheet [String, Integer, nil]
        # @return [Excelx::Sheet, nil]
        #: (?untyped sheet) -> Excelx::Sheet?
        def sheet_for(sheet = nil)
          raise NotImplementedError
        end

        private

        #: (Hash[untyped, untyped] options) -> void
        def clean_sheet_if_needed(options)
          return unless options[:clean]

          options.delete(:clean)
          target_sheet = default_sheet
          return if @cleaned[target_sheet]

          s_obj = sheet_for(target_sheet)
          s_obj&.cells&.each_value do |c|
            c.value = sanitize_value(c.value) if c.value.is_a?(String)
          end
          @cleaned[target_sheet] = true
        end

        #: (String str) -> String
        def sanitize_value(str)
          str.gsub(/[[:cntrl:]]|^\p{Space}+|\p{Space}+$/, "")
        end

        #: (Hash[untyped, untyped] options) -> void
        def search_or_set_header(options)
          if options[:header_search]
            @headers = nil
            @header_line = row_with(options[:header_search])
          elsif [:first_row, true].include?(options[:headers])
            @headers = []
            f_row = first_row || 1
            row(f_row).each_with_index { |x, i| @headers << [x, i + 1] }
          else
            set_headers(options)
          end
        end

        #: (Hash[untyped, untyped] hash) -> void
        def set_headers(hash = {})
          header_row = row_with(hash.values, true)
          @headers = {}
          hash.each_with_index do |(k, _), idx|
            matching_col = header_index(header_row[idx])
            @headers[k] = matching_col if matching_col
          end
        end

        #: (untyped query) -> Integer?
        def header_index(query)
          r_data = row(@header_line)
          idx = r_data.index(query)
          idx ? (idx + (first_column || 1)) : nil
        end

        #: (Integer row_index) -> Array[untyped]
        def find_by_row(row_index)
          r_idx = row_index + (@header_line ? (@header_line - 1) : 0)
          row(r_idx)
        end

        #: (Hash[untyped, untyped] options) -> Array[untyped]
        def find_by_conditions(options)
          f_row = first_row || 1
          l_row = last_row || 1
          f_col = first_column || 1
          l_col = last_column || 1

          header_for = (f_col..l_col).to_h do |c|
            [c, cell(@header_line, c)]
          end

          rows_indices = (f_row..l_row).to_a
          conditions = options[:conditions]
          if conditions && !conditions.empty?
            col_with = header_for.invert
            rows_indices = rows_indices.select do |i|
              conditions.all? { |k, v| cell(i, col_with[k]) == v }
            end
          end

          if options[:array]
            rows_indices.map { |i| row(i) }
          else
            rows_indices.map do |i|
              (f_col..l_col).to_h do |j|
                [header_for[j], cell(i, j)]
              end
            end
          end
        end

        #: [T] () { () -> T } -> T
        def without_changing_default_sheet(&block)
          original = default_sheet
          block.call
        ensure
          self.default_sheet = original
        end
      end

      # Forward declaration for Excelx inheriting from Base
      class Excelx < Base
      end
    end
  end
end

require_relative "constants"
require_relative "utils"
