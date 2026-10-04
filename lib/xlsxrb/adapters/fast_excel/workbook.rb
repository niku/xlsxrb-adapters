# frozen_string_literal: true

# rbs_inline: enabled

require "tmpdir"
require "fileutils"
require "xlsxrb"
require_relative "constants"
require_relative "format"
require_relative "worksheet"

module Xlsxrb
  module Adapters
    module FastExcel
      # Represents a workbook compatible with FastExcel::WorkbookExt and Libxlsxwriter::Workbook
      class Workbook
        include AttributeHelper

        # Standard built-in OpenXML number format mappings
        BUILTIN_NUM_FMTS = {
          "0" => 1,
          "0.00" => 2,
          "#,##0" => 3,
          "#,##0.00" => 4,
          "0%" => 9,
          "0.00%" => 10,
          "0.00E+00" => 11,
          "mm-dd-yy" => 14,
          "d-mmm-yy" => 15,
          "d-mmm" => 16,
          "mmm-yy" => 17,
          "h:mm AM/PM" => 18,
          "h:mm:ss AM/PM" => 19,
          "h:mm" => 20,
          "h:mm:ss" => 21,
          "m/d/yy h:mm" => 22
        }.freeze

        # @return [Boolean]
        attr_accessor :tmp_file

        # @return [Boolean]
        attr_accessor :is_open

        # @return [String]
        attr_accessor :filename

        # @return [Format]
        attr_reader :default_format

        # @return [Array<Worksheet>]
        attr_reader :sheets
        alias worksheets sheets

        # @param filename [String, Pathname, nil]
        # @param constant_memory [Boolean]
        # @param default_format [Hash{Symbol => untyped}, nil]
        #: (?String? filename, ?constant_memory: bool, ?default_format: Hash[Symbol, untyped]?) -> void
        def initialize(filename = nil, constant_memory: false, default_format: nil)
          @tmp_file = false

          if filename
            fn_str = filename.to_s
            raise ArgumentError, "File '#{fn_str}' already exists. FastExcel can not open existing files, only create new files" if File.exist?(fn_str) && File.size(fn_str).positive?

            @filename = fn_str
          else
            @filename = File.join(Dir.mktmpdir, "fast_excel.xlsx")
            @tmp_file = true
          end

          @constant_memory = constant_memory ? true : false
          @is_open = true
          @sheet_names = Set.new
          @sheets = []
          @formats = []
          @defined_names = []
          @core_properties = {}
          @app_properties = {}
          @custom_properties = []

          @default_format = add_format
          @default_format.font_size = 11

          return unless default_format
          raise "default_format argument must be a hash" unless default_format.is_a?(Hash)

          @default_format.set(default_format)
        end

        # Returns whether constant_memory mode is active.
        #
        # @return [Boolean]
        #: () -> bool
        def constant_memory?
          @constant_memory
        end

        # Adds a new format to the workbook.
        #
        # @param options [Hash{Symbol => untyped}, nil]
        # @return [Format]
        #: (?Hash[Symbol, untyped]? options) -> Format
        def add_format(options = nil)
          fmt = Format.new(self, options)
          @formats << fmt
          fmt
        end

        # Creates a new format with bold text.
        #
        # @return [Format]
        #: () -> Format
        def bold_cell_format
          bold = add_format
          bold.set_bold
          bold
        end
        alias bold_format bold_cell_format

        # Creates a new format with a specified number formatting pattern.
        #
        # @param pattern [String]
        # @return [Format]
        #: (String pattern) -> Format
        def number_format(pattern)
          fmt = add_format
          fmt.set_num_format(pattern)
          fmt
        end

        # Adds a new worksheet to the workbook.
        #
        # @param sheetname [String, nil]
        # @yield [worksheet]
        # @yieldparam worksheet [Worksheet]
        # @return [Worksheet]
        #: (?String? sheetname) ?{ (Worksheet) -> void } -> Worksheet
        def add_worksheet(sheetname = nil)
          target_name = if sheetname.nil?
                          "Sheet#{@sheets.size + 1}"
                        else
                          sheetname.to_s
                        end

          if sheetname && !target_name.empty?
            error = validate_worksheet_name(target_name)
            if error != :no_error
              error_code = ERROR_ENUM.find(error)
              error_str = ERROR_STRINGS[error_code] || ""
              raise ArgumentError, "Invalid worksheet name '#{target_name}': (#{error_code} - #{error}) #{error_str}"
            end
          end

          sheet = Worksheet.new(self, target_name, @sheets.size)
          @sheets << sheet
          @sheet_names << target_name

          yield sheet if block_given?
          sheet
        end

        # Retrieves an existing worksheet by name.
        #
        # @param name [String]
        # @return [Worksheet, nil]
        #: (String name) -> Worksheet?
        def get_worksheet_by_name(name)
          @sheets.find { |s| s.name == name }
        end

        # Validates a candidate worksheet name against Excel specification limits.
        #
        # @param sheetname [String]
        # @return [Symbol]
        #: (String sheetname) -> Symbol
        def validate_worksheet_name(sheetname)
          return :no_error if sheetname.empty?
          return :error_sheetname_length_exceeded if sheetname.length > 31
          return :error_invalid_sheetname_character if sheetname =~ %r{[\[\]:*?/\\]}
          return :error_sheetname_start_end_apostrophe if sheetname.start_with?("'") || sheetname.end_with?("'")
          return :error_sheetname_already_used if @sheet_names.include?(sheetname)

          :no_error
        end

        # Defines a global workbook-level defined name / named range.
        #
        # @param name [String]
        # @param formula [String]
        # @return [Symbol]
        #: (String name, String formula) -> Symbol
        def define_name(name, formula)
          @defined_names << { name: name.to_s, value: formula.to_s }
          :no_error
        end

        # Sets core document properties.
        #
        # @param properties [Hash{Symbol => untyped}, untyped]
        # @return [Symbol]
        def set_properties(properties)
          props = properties.respond_to?(:to_h) ? properties.to_h : properties
          if props.is_a?(Hash)
            @core_properties[:title] = props[:title] if props[:title]
            @core_properties[:creator] = props[:author] if props[:author]
            @app_properties[:company] = props[:company] if props[:company]
          end
          :no_error
        end

        # Sets custom document property string.
        #
        # @param name [String]
        # @param value [String]
        # @return [Symbol]
        def set_custom_property_string(name, value)
          @custom_properties << { name: name.to_s, value: value.to_s, type: :string }
          :no_error
        end

        # Sets custom document property number.
        #
        # @param name [String]
        # @param value [Float]
        # @return [Symbol]
        def set_custom_property_number(name, value)
          @custom_properties << { name: name.to_s, value: value.to_f, type: :float }
          :no_error
        end

        # Sets custom document property integer.
        #
        # @param name [String]
        # @param value [Integer]
        # @return [Symbol]
        def set_custom_property_integer(name, value)
          @custom_properties << { name: name.to_s, value: value.to_i, type: :integer }
          :no_error
        end

        # Sets custom document property boolean.
        #
        # @param name [String]
        # @param value [Boolean, Integer]
        # @return [Symbol]
        def set_custom_property_boolean(name, value)
          @custom_properties << { name: name.to_s, value: [true, 1].include?(value), type: :boolean }
          :no_error
        end

        # Sets custom document property datetime.
        #
        # @param name [String]
        # @param datetime [Time, DateTime, Datetime]
        # @return [Symbol]
        def set_custom_property_datetime(name, datetime)
          @custom_properties << { name: name.to_s, value: datetime, type: :datetime }
          :no_error
        end

        # Closes the workbook and serializes contents to disk.
        #
        # @return [Symbol]
        #: () -> Symbol
        def close
          return :no_error unless @is_open

          @is_open = false
          @sheets.each(&:close)

          FileUtils.mkdir_p(File.dirname(@filename))
          elements_wb = to_xlsxrb
          Xlsxrb.write(@filename, elements_wb)
          :no_error
        end

        # Finalizes workbook and returns binary content string.
        # Automatically removes temporary directory if temporary file was used.
        #
        # @return [String]
        #: () -> String
        def read_string
          close if @is_open
          File.binread(@filename)
        ensure
          remove_tmp_folder
        end

        # Cleans up temporary directory if created.
        #
        # @return [void]
        def remove_tmp_folder
          return unless @tmp_file && @filename

          dir = File.dirname(@filename)
          FileUtils.rm_rf(dir)
        end

        # Converts mutable workbook into an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          styles_def, style_map = compile_styles

          elements_sheets = @sheets.map { |ws| ws.to_xlsxrb(style_map) }

          facade_meta = {}
          facade_meta[:core_properties] = @core_properties unless @core_properties.empty?
          facade_meta[:app_properties] = @app_properties unless @app_properties.empty?
          facade_meta[:custom_properties] = @custom_properties unless @custom_properties.empty?

          unmapped = facade_meta.empty? ? {} : { facade: facade_meta }

          Xlsxrb::Elements::Workbook.new(
            sheets: elements_sheets,
            styles: styles_def,
            defined_names: @defined_names.empty? ? nil : @defined_names,
            unmapped_data: unmapped
          )
        end

        # Builds a FastExcel::Workbook instance from an immutable Xlsxrb::Elements::Workbook.
        #
        # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
        # @param filename [String, nil]
        # @param constant_memory [Boolean]
        # @return [Workbook]
        #: (Xlsxrb::Elements::Workbook xlsxrb_workbook, ?String? filename, ?constant_memory: bool) -> Workbook
        def self.from_xlsxrb(xlsxrb_workbook, filename = nil, constant_memory: false)
          wb = xlsxrb_workbook.respond_to?(:load) ? xlsxrb_workbook.load : xlsxrb_workbook
          adapter = new(filename, constant_memory: constant_memory)

          wb.sheets.each do |sheet|
            ws = adapter.add_worksheet(sheet.name)
            sheet.rows.each do |row|
              row.cells.each do |cell|
                next if cell.nil? || (cell.value.nil? && cell.formula.nil?)

                if cell.formula
                  ws.write_formula(cell.row_index, cell.column_index, cell.formula.expression)
                else
                  ws.write_value(cell.row_index, cell.column_index, cell.value)
                end
              end
            end
          end

          adapter
        end

        private

        #: () -> [Hash[Symbol, untyped], Hash[Format, Integer]]
        def compile_styles
          default_sz = @default_format.font_size.positive? ? @default_format.font_size : 11
          default_font = { name: @default_format.font_name || "Calibri", sz: default_sz, family: 2 }

          fonts_list = [default_font]
          fills_list = [{ pattern: "none" }, { pattern: "gray125" }]
          borders_list = [{}]
          num_fmts_list = []
          xf_entries_list = [{ num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 }]
          style_map = {}

          custom_num_fmt_id = 165
          num_fmt_cache = {}

          @formats.each do |fmt|
            next if style_map.key?(fmt)

            font_key = compile_font_key(fmt, default_sz)
            font_id = fonts_list.index(font_key) || (fonts_list << font_key
                                                     fonts_list.size - 1)

            fill_key = compile_fill_key(fmt)
            fill_id = if fill_key
                        fills_list.index(fill_key) || (fills_list << fill_key
                                                       fills_list.size - 1)
                      else
                        0
                      end

            border_key = compile_border_key(fmt)
            border_id = if border_key
                          borders_list.index(border_key) || (borders_list << border_key
                                                             borders_list.size - 1)
                        else
                          0
                        end

            num_fmt_id = resolve_num_fmt(fmt.num_format, num_fmt_cache, num_fmts_list, custom_num_fmt_id)
            custom_num_fmt_id += 1 if num_fmt_id >= custom_num_fmt_id

            alignment_hash = compile_alignment(fmt)

            xf = {
              num_fmt_id: num_fmt_id,
              font_id: font_id,
              fill_id: fill_id,
              border_id: border_id,
              alignment: alignment_hash
            }.compact

            xf_id = xf_entries_list.size
            xf_entries_list << xf
            style_map[fmt] = xf_id
          end

          styles_hash = {
            fonts: fonts_list,
            fills: fills_list,
            borders: borders_list,
            num_fmts: num_fmts_list,
            cell_xfs: xf_entries_list,
            xf_entries: xf_entries_list
          }

          [styles_hash, style_map]
        end

        #: (Format fmt, Numeric default_sz) -> Hash[Symbol, untyped]
        def compile_font_key(fmt, default_sz)
          sz = fmt.font_size.positive? ? fmt.font_size : default_sz
          u_val = case fmt.underline
                  when 1, true, :single, :underline_single then true
                  when 2, :double, :underline_double then "double"
                  when :underline_single_accounting then "singleAccounting"
                  when :underline_double_accounting then "doubleAccounting"
                  end

          color_hex = fmt.font_color ? format("%06X", fmt.font_color) : nil

          {
            name: fmt.font_name || "Calibri",
            sz: sz,
            bold: [1, true].include?(fmt.bold) || nil,
            italic: [1, true].include?(fmt.italic) || nil,
            underline: u_val,
            strike: [1, true].include?(fmt.font_strikeout) || nil,
            outline: [1, true].include?(fmt.font_outline) || nil,
            shadow: [1, true].include?(fmt.font_shadow) || nil,
            color: color_hex
          }.compact
        end

        #: (Format fmt) -> Hash[Symbol, untyped]?
        def compile_fill_key(fmt)
          has_fill = fmt.bg_color || fmt.fg_color || (fmt.pattern && fmt.pattern != :border_none && fmt.pattern != 0)
          return nil unless has_fill

          pattern_str = if fmt.pattern && fmt.pattern != 0
                          fmt.pattern.to_s
                        else
                          "solid"
                        end

          fg = if fmt.fg_color
                 format("%06X", fmt.fg_color)
               else
                 (fmt.bg_color ? format("%06X", fmt.bg_color) : nil)
               end
          bg = fmt.bg_color ? format("%06X", fmt.bg_color) : nil

          {
            pattern: pattern_str,
            fg_color: fg,
            bg_color: bg
          }.compact
        end

        #: (Format fmt) -> Hash[Symbol, untyped]?
        def compile_border_key(fmt)
          b_val = fmt[:bottom].to_i
          t_val = fmt[:top].to_i
          l_val = fmt[:left].to_i
          r_val = fmt[:right].to_i

          return nil if b_val.zero? && t_val.zero? && l_val.zero? && r_val.zero?

          {
            bottom: border_side_hash(fmt.bottom, fmt.bottom_color),
            top: border_side_hash(fmt.top, fmt.top_color),
            left: border_side_hash(fmt.left, fmt.left_color),
            right: border_side_hash(fmt.right, fmt.right_color)
          }.compact
        end

        #: (Symbol? side, Integer? color) -> Hash[Symbol, String]?
        def border_side_hash(side, color)
          return nil if side.nil? || side == :border_none || side == :none

          style_name = side.to_s.sub(/^border_/, "")
          col_str = color ? format("%06X", color) : "000000"
          { style: style_name, color: col_str }
        end

        #: (String? pattern, Hash[String, Integer] cache, Array[Hash[Symbol, untyped]] list, Integer next_id) -> Integer
        def resolve_num_fmt(pattern, cache, list, next_id)
          return 0 if pattern.nil? || pattern.empty?
          return BUILTIN_NUM_FMTS[pattern] if BUILTIN_NUM_FMTS.key?(pattern)
          return cache[pattern] if cache.key?(pattern)

          id = next_id
          cache[pattern] = id
          list << { num_fmt_id: id, format_code: pattern }
          id
        end

        #: (Format fmt) -> Hash[Symbol, untyped]?
        def compile_alignment(fmt)
          res = {}
          h = fmt.align[:horizontal]
          v = fmt.align[:vertical]

          res[:horizontal] = h.to_s.sub(/^align_/, "") unless h == :align_none
          res[:vertical] = v.to_s.sub(/^align_vertical_/, "").sub(/^align_/, "") unless v == :align_none
          res[:wrap_text] = true if [1, true].include?(fmt.text_wrap)
          res[:text_rotation] = fmt.rotation if fmt.rotation&.positive?
          res[:indent] = fmt.indent if fmt.indent&.positive?
          res[:shrink_to_fit] = true if [1, true].include?(fmt.shrink)

          res.empty? ? nil : res
        end
      end
    end
  end
end
