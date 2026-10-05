# frozen_string_literal: true

# rbs_inline: enabled

require "fileutils"
require "stringio"
require "xlsxrb"
require_relative "constants"
require_relative "colors"
require_relative "utility"
require_relative "format"
require_relative "worksheet"
require_relative "chart"

module Xlsxrb
  module Adapters
    module Writexlsx
      # Represents a workbook in WriteXLSX
      class Workbook
        include Utility::Common
        include Utility::CellReference
        include Utility::DateTime

        # Built-in OpenXML number format mappings
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

        # @return [Object]
        attr_reader :file

        # @return [Array<Worksheet>]
        attr_reader :worksheets
        alias sheets worksheets

        # @return [Array<Format>]
        attr_reader :formats

        # @return [Array<Chart>]
        attr_reader :charts

        # @return [Colors]
        attr_reader :palette

        # @return [Format]
        attr_reader :default_url_format

        # @return [Boolean]
        attr_accessor :strings_to_urls

        # @return [Hash{Symbol => untyped}]
        attr_reader :doc_properties

        # @return [Array<Hash{Symbol => untyped}>]
        attr_reader :custom_properties

        # @return [Array<Hash{Symbol => untyped}>]
        attr_reader :defined_names

        # @param file [String, Pathname, IO, StringIO, nil]
        # @param options [Hash{Symbol => untyped}]
        #: (?untyped file, ?Hash[Symbol, untyped] options) -> void
        def initialize(file = nil, options = {})
          @file = file
          @options = options.dup
          @date_1904 = options[:date_1904] ? true : false
          @strings_to_urls = options[:strings_to_urls].nil? || options[:strings_to_urls]
          @tempdir = options[:tempdir]

          @worksheets = []
          @formats = []
          @charts = []
          @defined_names = []
          @doc_properties = {}
          @custom_properties = []
          @palette = Colors.new
          @fileclosed = false

          @default_url_format = add_format(color: "blue", underline: 1)
        end

        # Returns whether 1904 date system is active
        #
        # @return [Boolean]
        #: () -> bool
        def date_1904?
          @date_1904
        end

        # Sets whether 1904 date system is active
        #
        # @param val [Boolean]
        # @return [Boolean]
        #: (?bool val) -> bool
        def set_1904(val = true)
          @date_1904 = [true, 1].include?(val)
        end

        # Adds a new worksheet to the workbook
        #
        # @param name [String, nil]
        # @yield [worksheet]
        # @return [Worksheet]
        #: (?String? name) ?{ (Worksheet) -> void } -> Worksheet
        def add_worksheet(name = nil)
          target_name = if name.nil? || name.to_s.empty?
                          "Sheet#{@worksheets.size + 1}"
                        else
                          name.to_s
                        end

          ws = Worksheet.new(self, target_name, @worksheets.size)
          @worksheets << ws
          yield ws if block_given?
          ws
        end

        # Adds a new format to the workbook
        #
        # @param properties [Hash{Symbol => untyped}]
        # @return [Format]
        #: (?Hash[Symbol, untyped] properties) -> Format
        def add_format(properties = {})
          fmt = Format.new(self, properties)
          @formats << fmt
          fmt
        end

        # Adds a chart to the workbook
        #
        # @param options [Hash{Symbol => untyped}]
        # @return [Chart]
        #: (?Hash[Symbol, untyped] options) -> Chart
        def add_chart(options = {})
          ch = Chart.factory(options[:type] || :column, options[:subtype])
          ch.instance_variable_set(:@workbook, self)
          ch.palette = @palette
          @charts << ch
          ch
        end

        # Adds a chartsheet to the workbook
        #
        # @param name [String, nil]
        # @return [Worksheet]
        def add_chartsheet(name = nil)
          add_worksheet(name)
        end

        # Adds a shape to the workbook
        #
        # @param options [Hash{Symbol => untyped}]
        # @return [Hash{Symbol => untyped}]
        def add_shape(options = {})
          options
        end

        # Defines a global defined name / named range
        #
        # @param name [String]
        # @param formula [String]
        # @return [void]
        #: (String name, String formula) -> void
        def define_name(name, formula)
          @defined_names << { name: name.to_s, value: formula.to_s }
        end

        # Sets document properties
        #
        # @param properties [Hash{Symbol => untyped}]
        # @return [void]
        def set_properties(properties)
          return unless properties.is_a?(Hash)

          @doc_properties.merge!(properties)
        end

        # Sets a custom document property
        #
        # @param name [String]
        # @param value [Object]
        # @param type [Symbol, String, nil]
        # @return [void]
        def set_custom_property(name, value, type = nil)
          @custom_properties << { name: name.to_s, value: value, type: type }
        end

        # Sets window size
        def set_size(width, height)
          @window_width = width
          @window_height = height
        end

        # Sets tab ratio
        def set_tab_ratio(ratio)
          @tab_ratio = ratio
        end

        # Sets calculation mode
        def set_calc_mode(mode)
          @calc_mode = mode
        end

        # Looks up worksheet by name
        #
        # @param name [String]
        # @return [Worksheet, nil]
        #: (String name) -> Worksheet?
        def worksheet_by_name(name)
          @worksheets.find { |ws| ws.name == name }
        end
        alias get_worksheet_by_name worksheet_by_name

        # Closes the workbook and serializes contents
        #
        # @return [void]
        #: () -> void
        def close
          return if @fileclosed

          raise "'' must be valid filename String of IO object." if @file.nil?

          @fileclosed = true
          elements_wb = to_xlsxrb

          if @file.is_a?(String) || (defined?(Pathname) && @file.is_a?(Pathname))
            fn = @file.to_s
            FileUtils.mkdir_p(File.dirname(fn))
            Xlsxrb.write(fn, elements_wb)
          else
            Xlsxrb.write(@file, elements_wb)
          end
        end

        # Finalizes workbook and returns binary content string
        #
        # @return [String]
        #: () -> String
        def read_string
          if @file.nil?
            elements_wb = to_xlsxrb
            sio = StringIO.new
            Xlsxrb.write(sio, elements_wb)
            return sio.string
          end

          close unless @fileclosed

          if @file.respond_to?(:string)
            @file.string
          elsif @file.is_a?(String) || (defined?(Pathname) && @file.is_a?(Pathname))
            File.binread(@file.to_s)
          elsif @file.respond_to?(:rewind) && @file.respond_to?(:read)
            @file.rewind
            @file.read
          else
            ""
          end
        end

        # Converts mutable workbook into an immutable Xlsxrb::Elements::Workbook
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          styles_def, style_map = compile_styles

          elements_sheets = @worksheets.map { |ws| ws.to_xlsxrb(style_map) }

          facade_meta = {}
          core_props = {
            title: @doc_properties[:title],
            creator: @doc_properties[:author],
            subject: @doc_properties[:subject],
            category: @doc_properties[:category],
            keywords: @doc_properties[:keywords],
            comments: @doc_properties[:comments],
            status: @doc_properties[:status]
          }.compact
          facade_meta[:core_properties] = core_props unless core_props.empty?

          app_props = {
            company: @doc_properties[:company],
            manager: @doc_properties[:manager]
          }.compact
          facade_meta[:app_properties] = app_props unless app_props.empty?

          facade_meta[:custom_properties] = @custom_properties unless @custom_properties.empty?

          unmapped = facade_meta.empty? ? {} : { facade: facade_meta }
          unmapped[:workbook_properties] = { date1904: true } if @date_1904

          Xlsxrb::Elements::Workbook.new(
            sheets: elements_sheets,
            styles: styles_def,
            defined_names: @defined_names.empty? ? nil : @defined_names,
            unmapped_data: unmapped
          )
        end

        # Builds a WriteXLSX::Workbook instance from an Xlsxrb::Elements::Workbook
        #
        # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
        # @param file [Object, nil]
        # @param options [Hash{Symbol => untyped}]
        # @return [Workbook]
        def self.from_xlsxrb(xlsxrb_workbook, file = nil, options = {})
          wb = xlsxrb_workbook.respond_to?(:load) ? xlsxrb_workbook.load : xlsxrb_workbook
          target_file = file || StringIO.new
          adapter = new(target_file, options)

          wb.sheets.each do |sheet|
            ws = adapter.add_worksheet(sheet.name)
            sheet.rows.each do |row|
              row.cells.each do |cell|
                next if cell.nil? || (cell.value.nil? && cell.formula.nil?)

                if cell.formula
                  ws.write_formula(cell.row_index, cell.column_index, cell.formula.expression)
                else
                  ws.write(cell.row_index, cell.column_index, cell.value)
                end
              end
            end
          end

          adapter
        end

        private

        def compile_styles
          default_font = { name: "Calibri", sz: 11, family: 2 }
          fonts_list = [default_font]
          fills_list = [{ pattern: "none" }, { pattern: "gray125" }]
          borders_list = [{}]
          num_fmts_list = []
          xf_entries_list = [{ num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 }]
          style_map = {}

          custom_num_fmt_id = 165
          num_fmt_cache = {}

          all_formats = @formats.dup
          @worksheets.each do |ws|
            # Gather formats from worksheet if not already registered
            formats_used = ws.instance_variable_get(:@formats_used)
            formats_used = ws.instance_variable_get(:@table).values.map(&:format).compact if formats_used.nil? || formats_used.empty?
            row_formats = ws.instance_variable_get(:@set_rows).values.map { |r| r[:format] }.compact
            col_formats = ws.instance_variable_get(:@col_info).values.map { |c| c[:format] }.compact
            all_formats.concat(formats_used)
            all_formats.concat(row_formats)
            all_formats.concat(col_formats)
          end
          all_formats.uniq!

          all_formats.each do |fmt|
            next if style_map.key?(fmt)

            # Font
            font_key = compile_font_key(fmt)
            font_id = fonts_list.index(font_key) || (fonts_list << font_key
                                                     fonts_list.size - 1)

            # Fill
            fill_key = compile_fill_key(fmt)
            fill_id = if fill_key
                        fills_list.index(fill_key) || (fills_list << fill_key
                                                       fills_list.size - 1)
                      else
                        0
                      end

            # Border
            border_key = compile_border_key(fmt)
            border_id = if border_key
                          borders_list.index(border_key) || (borders_list << border_key
                                                             borders_list.size - 1)
                        else
                          0
                        end

            # Number Format
            num_fmt_id = resolve_num_fmt(fmt, num_fmt_cache, num_fmts_list, custom_num_fmt_id)
            custom_num_fmt_id += 1 if num_fmt_id >= custom_num_fmt_id

            # Alignment
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
            fmt.xf_index = xf_id
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

        def compile_font_key(fmt)
          sz = fmt.size&.positive? ? fmt.size : 11
          u_val = if fmt.underline?
                    fmt.underline == 2 ? "double" : true
                  end

          color_hex = Colors.to_hex(fmt.font_color)

          {
            name: fmt.font || "Calibri",
            sz: sz,
            bold: fmt.bold? || nil,
            italic: fmt.italic? || nil,
            underline: u_val,
            strike: fmt.strikeout? || nil,
            outline: fmt.outline? || nil,
            shadow: fmt.shadow? || nil,
            color: color_hex
          }.compact
        end

        def compile_fill_key(fmt)
          return nil unless fmt.has_fill?

          pattern = fmt.pattern ? fmt.pattern.to_s : "solid"
          pattern = "solid" if pattern == "1"
          fg = Colors.to_hex(fmt.fg_color)
          bg = Colors.to_hex(fmt.bg_color)

          {
            pattern: pattern,
            fg_color: fg || bg,
            bg_color: bg || fg
          }.compact
        end

        def compile_border_key(fmt)
          return nil unless fmt.has_border?

          res = {}
          res[:bottom] = { style: fmt.bottom, color: Colors.to_hex(fmt.bottom_color) }.compact if fmt.bottom
          res[:top] = { style: fmt.top, color: Colors.to_hex(fmt.top_color) }.compact if fmt.top
          res[:left] = { style: fmt.left, color: Colors.to_hex(fmt.left_color) }.compact if fmt.left
          res[:right] = { style: fmt.right, color: Colors.to_hex(fmt.right_color) }.compact if fmt.right
          res[:diagonal] = { style: fmt.diag_border, color: Colors.to_hex(fmt.diag_color) }.compact if fmt.diag_border
          res.empty? ? nil : res
        end

        def resolve_num_fmt(fmt, cache, num_fmts_list, current_custom_id)
          return fmt.num_format_index if fmt.num_format_index

          code = fmt.num_format
          return 0 if code.nil? || code.empty?

          return BUILTIN_NUM_FMTS[code] if BUILTIN_NUM_FMTS.key?(code)

          if cache.key?(code)
            cache[code]
          else
            new_id = current_custom_id
            num_fmts_list << { num_fmt_id: new_id, format_code: code }
            cache[code] = new_id
            new_id
          end
        end

        def compile_alignment(fmt)
          h_align = case fmt.text_h_align
                    when 1 then "left"
                    when 2 then "center"
                    when 3 then "right"
                    when 4 then "fill"
                    when 5 then "justify"
                    when 6 then "centerContinuous"
                    when 7 then "distributed"
                    end

          v_align = case fmt.text_v_align
                    when 1 then "top"
                    when 2 then "center"
                    when 3 then "bottom"
                    when 4 then "justify"
                    when 5 then "distributed"
                    end

          {
            horizontal: h_align,
            vertical: v_align,
            wrap_text: [1, true].include?(fmt.text_wrap) || nil,
            text_rotation: fmt.rotation,
            indent: fmt.indent,
            shrink_to_fit: [1, true].include?(fmt.shrink) || nil
          }.compact
        end
      end
    end
  end
end
