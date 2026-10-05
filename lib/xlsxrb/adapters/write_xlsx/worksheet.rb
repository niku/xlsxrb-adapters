# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "xlsxrb"
require_relative "constants"
require_relative "colors"
require_relative "utility"
require_relative "format"
require_relative "chart"

module Xlsxrb
  module Adapters
    module Writexlsx
      # Represents a worksheet in WriteXLSX
      class Worksheet
        include Utility::Common
        include Utility::CellReference
        include Utility::DateTime

        # Record representing a single cell entry in the worksheet
        CellRecord = Struct.new(
          :value,
          :format,
          :formula,
          :cached_value,
          :url,
          :comment,
          :rich_text,
          :array_formula,
          :boolean
        )

        # @return [Object]
        attr_accessor :workbook

        # @return [String]
        attr_accessor :name

        # @return [Integer]
        attr_accessor :index

        # @return [Numeric]
        attr_accessor :default_row_height

        # @return [Hash{Array<Integer> => CellRecord}]
        attr_reader :table

        # @param workbook [Object]
        # @param name [String]
        # @param index [Integer]
        #: (untyped workbook, String name, ?Integer index) -> void
        def initialize(workbook, name, index = 0)
          @workbook = workbook
          @name = name
          @index = index

          @table = {}
          @rows = {}
          @formats_used = []
          @set_rows = {}
          @col_info = {}
          @merge = []
          @autofilter_ref = nil
          @filter_columns = {}
          @filter_cells = {}
          @panes = nil
          @top_left_cell = nil
          @comments = []
          @hyperlinks = []
          @images = []
          @charts = []
          @tables = []
          @sparklines = []
          @data_validations = []
          @cond_formats = []

          @page_setup = {}
          @page_margins = { left: 0.7, right: 0.7, top: 0.75, bottom: 0.75, header: 0.3, footer: 0.3 }
          @header_footer = {}
          @print_options = {}
          @print_area = nil
          @clean_print_area = nil
          @repeat_rows = nil
          @repeat_cols = nil
          @h_breaks = []
          @v_breaks = []

          @hidden = false
          @active = false
          @selected = false
          @right_to_left = false
          @hide_gridlines = false
          @hide_row_col_headers = false
          @hide_zero = false
          @tab_color = nil
          @zoom = nil
          @default_row_height = 15.0
          @leading_zeros = false

          @dim_rowmin = nil
          @dim_rowmax = nil
          @dim_colmin = nil
          @dim_colmax = nil
        end

        attr_reader :dim_rowmin, :dim_rowmax, :dim_colmin, :dim_colmax, :tab_color, :sparklines, :clean_print_area

        # Returns dimensions of used area
        #
        # @return [Hash{Symbol => Integer}]
        #: () -> Hash[Symbol, Integer]
        def dimension
          {
            row_min: @dim_rowmin || 0,
            row_max: @dim_rowmax || 0,
            col_min: @dim_colmin || 0,
            col_max: @dim_colmax || 0
          }
        end

        # Primary cell write dispatcher supporting A1 and (row, col) notation
        #
        # @param row [Integer, String]
        # @param col [Integer, Object]
        # @param token [Object, nil]
        # @param format [Format, nil]
        # @param value1 [Object, nil]
        # @param value2 [Object, nil]
        # @return [Object]
        def write(row, col = nil, token = nil, format = nil, value1 = nil, value2 = nil)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            t = col
            fmt = token
            v1 = format
            v2 = value1
          else
            r = row
            c = col
            t = token
            fmt = format
            v1 = value1
            v2 = value2
          end

          return if r.nil? || c.nil?

          return write_string(r, c, t, fmt) if fmt.respond_to?(:force_text_format?) && fmt.force_text_format?

          return write_row(r, c, t, fmt, v1, v2) if t.respond_to?(:to_ary)

          return write_number(r, c, t, fmt) if t.is_a?(Numeric)

          return write_boolean(r, c, t, fmt) if t.is_a?(TrueClass) || t.is_a?(FalseClass)

          return write_date_time(r, c, t, fmt) if t.is_a?(Time) || t.is_a?(Date)

          if t.is_a?(String)
            return write_string(r, c, t, fmt) if @leading_zeros && t =~ /^0\d+$/

            return write_number(r, c, t, fmt) if t =~ /\A([+-]?)(?=\d|\.\d)\d*(\.\d*)?([Ee]([+-]?\d+))?\Z/ && !t.empty? && t != "."

            return write_formula(r, c, t, fmt, v1) if t.start_with?("=")

            return write_formula(r, c, t, fmt, v1) if t =~ /^\{=.*\}$/

            return write_blank(r, c, fmt) if t.empty?

            if @workbook.respond_to?(:strings_to_urls) && @workbook.strings_to_urls &&
               t =~ %r{\A(?:(?:https?|ftp)://|mailto:|(?:in|ex)ternal:)}
              return write_url(r, c, t, fmt, v1, v2)
            end

            return write_string(r, c, t, fmt)
          end

          return write_blank(r, c, fmt) if t.nil?

          write_string(r, c, t.to_s, fmt)
        end

        # Writes a row of values starting at (row, col)
        #
        # @param row [Integer, String]
        # @param col [Integer, Object]
        # @param tokens [Array<Object>, nil]
        # @param options [Array<Object>]
        # @return [void]
        def write_row(row, col = nil, tokens = nil, *options)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            t = col
            opts = [tokens] + options
          else
            r = row
            c = col
            t = tokens
            opts = options
          end

          raise WriteXLSXInsufficientArgumentError, "Tokens must be array in write_row" unless t.respond_to?(:to_ary)

          cur_col = c
          t.each do |item|
            if item.respond_to?(:to_ary)
              write_col(r, cur_col, item, *opts)
            else
              write(r, cur_col, item, *opts)
            end
            cur_col += 1
          end
        end

        # Writes a column of values starting at (row, col)
        #
        # @param row [Integer, String]
        # @param col [Integer, Object]
        # @param tokens [Array<Object>, nil]
        # @param options [Array<Object>]
        # @return [void]
        def write_col(row, col = nil, tokens = nil, *options)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            t = col
            opts = [tokens] + options
          else
            r = row
            c = col
            t = tokens
            opts = options
          end

          raise WriteXLSXInsufficientArgumentError, "Tokens must be array in write_col" unless t.respond_to?(:to_ary)

          cur_row = r
          t.each do |item|
            if item.respond_to?(:to_ary)
              write_row(cur_row, c, item, *opts)
            else
              write(cur_row, c, item, *opts)
            end
            cur_row += 1
          end
        end

        # Writes a string to a cell
        #
        # @param row [Integer, String]
        # @param col [Integer, Object]
        # @param str [Object, nil]
        # @param format [Format, nil]
        # @return [void]
        def write_string(row, col = nil, str = nil, format = nil)
          r, c, s, fmt = normalize_cell_args(row, col, str, format)
          update_dimensions(r, c)
          record = CellRecord.new(s.to_s, fmt, nil, nil, nil, nil, nil, nil, nil)
          store_cell_record(r, c, record)
        end

        # Writes a numeric value to a cell
        #
        # @param row [Integer, String]
        # @param col [Integer, Object]
        # @param num [Numeric, String, nil]
        # @param format [Format, nil]
        # @return [void]
        def write_number(row, col = nil, num = nil, format = nil)
          r, c, n, fmt = normalize_cell_args(row, col, num, format)
          update_dimensions(r, c)

          parsed_num = if n.is_a?(Numeric)
                         n
                       elsif n.to_s =~ /^[+-]?\d+$/
                         n.to_i
                       else
                         n.to_f
                       end

          record = CellRecord.new(parsed_num, fmt, nil, nil, nil, nil, nil, nil, nil)
          store_cell_record(r, c, record)
        end

        # Writes a blank formatted cell
        #
        # @param row [Integer, String]
        # @param col [Integer, Object]
        # @param format [Format, nil]
        # @return [void]
        def write_blank(row, col = nil, format = nil)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            fmt = col
          else
            r = row
            c = col
            fmt = format
          end
          update_dimensions(r, c)
          record = CellRecord.new(nil, fmt, nil, nil, nil, nil, nil, nil, nil)
          store_cell_record(r, c, record)
        end

        # Writes a formula to a cell
        #
        # @param row [Integer, String]
        # @param col [Integer, Object]
        # @param formula [String]
        # @param format [Format, nil]
        # @param value [Object, nil]
        # @return [void]
        def write_formula(row, col = nil, formula = nil, format = nil, value = nil)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            fml = col
            fmt = formula
            val = format
          else
            r = row
            c = col
            fml = formula
            fmt = format
            val = value
          end

          update_dimensions(r, c)
          clean_fml = fml.to_s
          clean_fml = clean_fml.sub(/^\{/, "").sub(/\}$/, "") if clean_fml =~ /^\{=.*\}$/

          record = CellRecord.new(val || 0, fmt, clean_fml, val || 0, nil, nil, nil, nil, nil)
          store_cell_record(r, c, record)
        end

        # Writes an array formula to a cell or cell range
        def write_array_formula(row1, col1 = nil, row2 = nil, col2 = nil, formula = nil, format = nil, value = nil)
          if (row_col_array = row_col_notation(row1))
            if row_col_array.size == 4
              r1, c1, r2, c2 = row_col_array
              fml = col1
              fmt = row2
              val = col2
            else
              r1, c1 = row_col_array
              r2 = col1
              c2 = row2
              fml = col2
              fmt = formula
              val = format
            end
          else
            r1 = row1
            c1 = col1
            r2 = row2
            c2 = col2
            fml = formula
            fmt = format
            val = value
          end

          r1, r2 = r2, r1 if r1 > r2
          c1, c2 = c2, c1 if c1 > c2

          update_dimensions(r1, c1)
          update_dimensions(r2, c2)

          range_ref = xl_range(r1, r2, c1, c2)
          clean_fml = fml.to_s.sub(/^\{/, "").sub(/\}$/, "")

          (r1..r2).each do |r|
            (c1..c2).each do |c|
              is_top_left = r == r1 && c == c1
              f_expr = is_top_left ? clean_fml : nil
              record = CellRecord.new(val || 0, fmt, f_expr, val || 0, nil, nil, nil, range_ref, nil)
              store_cell_record(r, c, record)
            end
          end
        end

        alias write_dynamic_array_formula write_array_formula

        # Writes a hyperlink to a cell
        def write_url(row, col = nil, url = nil, format = nil, str = nil, tip = nil, _ignore = false)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            u = col
            fmt = url
            s = format
            t = str
          else
            r = row
            c = col
            u = url
            fmt = format
            s = str
            t = tip
          end

          update_dimensions(r, c)
          display = s || u
          fmt ||= @workbook.respond_to?(:default_url_format) ? @workbook.default_url_format : nil

          cell_ref = xl_rowcol_to_cell(r, c)
          @hyperlinks << { ref: cell_ref, url: u.to_s, display: display.to_s, tooltip: t }

          record = CellRecord.new(display, fmt, nil, nil, u.to_s, nil, nil, nil, nil)
          store_cell_record(r, c, record)
        end

        # Writes a date or time to a cell
        def write_date_time(row, col = nil, str = nil, format = nil)
          r, c, val, fmt = normalize_cell_args(row, col, str, format)
          update_dimensions(r, c)

          is_1904 = @workbook.respond_to?(:date_1904?) && @workbook.date_1904?
          serial = convert_date_time(val, is_1904)

          record = if serial
                     CellRecord.new(serial, fmt, nil, nil, nil, nil, nil, nil, nil)
                   else
                     CellRecord.new(val.to_s, fmt, nil, nil, nil, nil, nil, nil, nil)
                   end
          store_cell_record(r, c, record)
        end

        # Writes a boolean to a cell
        def write_boolean(row, col = nil, val = true, format = nil)
          r, c, v, fmt = normalize_cell_args(row, col, val, format)
          update_dimensions(r, c)
          bool_val = [true, 1].include?(v)
          record = CellRecord.new(bool_val, fmt, nil, nil, nil, nil, nil, nil, true)
          store_cell_record(r, c, record)
        end

        # Writes a rich text formatted string to a cell
        def write_rich_string(row, col = nil, *fragments)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            frags = [col] + fragments
          else
            r = row
            c = col
            frags = fragments
          end

          update_dimensions(r, c)

          fmt = frags.last.is_a?(Format) ? frags.pop : nil
          rich_text_obj = Xlsxrb.rich_text(*frags)
          record = CellRecord.new(rich_text_obj, fmt, nil, nil, nil, nil, rich_text_obj, nil, nil)
          store_cell_record(r, c, record)
        end

        # Writes a comment to a cell
        def write_comment(row, col = nil, string = nil, options = nil)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            s = col
            opts = string
          else
            r = row
            c = col
            s = string
            opts = options
          end

          update_dimensions(r, c)
          opts_hash = opts.is_a?(Hash) ? opts : {}
          cell_ref = xl_rowcol_to_cell(r, c)

          cmt = {
            ref: cell_ref,
            text: s.to_s,
            author: opts_hash[:author] || @workbook.respond_to?(:comments_author) ? @workbook.comments_author : "Author",
            visible: opts_hash[:visible] ? true : false
          }.compact

          @comments << cmt

          existing = (@rows[r] && @rows[r][c]) || @table[[r, c]]
          if existing
            existing.comment = cmt
          else
            store_cell_record(r, c, CellRecord.new(nil, nil, nil, nil, nil, cmt, nil, nil, nil))
          end
        end

        # Sets row properties (height, format, hidden, level, collapsed)
        def set_row(row, height = nil, format = nil, hidden = 0, level = 0, collapsed = 0)
          return unless row

          update_dimensions(row, 0)
          h = height || @default_row_height
          is_hidden = hidden == 1 || height.zero?
          @formats_used << format if format
          @set_rows[row] = {
            height: h,
            format: format,
            hidden: is_hidden,
            level: [level.to_i, 7].min,
            collapsed: collapsed == 1
          }
        end

        def set_row_pixels(row, pixels, *)
          height = pixels ? pixels * 0.75 : nil
          set_row(row, height, *)
        end

        # Sets column properties (width, format, hidden, level, collapsed)
        def set_column(*args)
          if args[0].respond_to?(:=~) && args[0].to_s =~ /^\D/
            _r1, firstcol, _r2, lastcol, *data = substitute_cellref(*args)
          else
            firstcol, lastcol, *data = args
          end

          return unless firstcol && lastcol && !data.empty?

          lastcol = firstcol unless ptrue?(lastcol)
          firstcol, lastcol = lastcol, firstcol if firstcol > lastcol

          width, format, hidden, level, collapsed = data
          update_dimensions(0, firstcol)
          update_dimensions(0, lastcol)
          @formats_used << format if format

          (firstcol..lastcol).each do |c|
            @col_info[c] = {
              width: width ? [width.to_f, 255.0].min : nil,
              format: format,
              hidden: hidden == 1,
              level: [level.to_i, 7].min,
              collapsed: collapsed == 1
            }
          end
        end

        def set_column_pixels(firstcol, lastcol, pixels, *)
          width = pixels ? (pixels / 7.0) : nil
          set_column(firstcol, lastcol, width, *)
        end

        # Merges a range of cells with data in the first cell
        def merge_range(*args)
          if (row_col_array = row_col_notation(args.first))
            r1, c1, r2, c2 = row_col_array
            string, format = args[1..2]
          else
            r1, c1, r2, c2, string, format = args
          end

          raise WriteXLSXInsufficientArgumentError if [r1, c1, r2, c2, format].include?(nil)

          r1, r2 = r2, r1 if r1 > r2
          c1, c2 = c2, c1 if c1 > c2

          update_dimensions(r1, c1)
          update_dimensions(r2, c2)

          @merge << [r1, c1, r2, c2]

          write(r1, c1, string, format)

          (r1..r2).each do |r|
            (c1..c2).each do |c|
              next if r == r1 && c == c1

              write_blank(r, c, format)
            end
          end
        end

        def merge_range_type(_type, *)
          merge_range(*)
        end

        # Sets autofilter range on the sheet
        def autofilter(row1, col1 = nil, row2 = nil, col2 = nil)
          if (row_col_array = row_col_notation(row1))
            r1, c1, r2, c2 = row_col_array
          else
            r1 = row1
            c1 = col1
            r2 = row2
            c2 = col2
          end

          return if [r1, c1, r2, c2].include?(nil)

          r1, r2 = r2, r1 if r2 < r1
          c1, c2 = c2, c1 if c2 < c1

          @autofilter_ref = xl_range(r1, r2, c1, c2)
        end

        def filter_column(col, string)
          c = col.is_a?(String) ? xl_cell_to_rowcol("#{col}1")[1] : col
          @filter_columns[c] = string
        end

        def filter_column_list(col, list)
          c = col.is_a?(String) ? xl_cell_to_rowcol("#{col}1")[1] : col
          @filter_columns[c] = list
        end

        # Freezes rows/columns panes
        def freeze_panes(*args)
          return if args.empty?

          if (row_col_array = row_col_notation(args.first))
            row, col, top_row, left_col = row_col_array
            type = args[1] || 0
          else
            row, col, top_row, left_col, type = args
          end

          col ||= 0
          top_row ||= row
          left_col ||= col
          type ||= 0

          @panes = [row, col, top_row, left_col, type]
        end

        # Splits worksheet panes
        def split_panes(*args)
          freeze_panes(args[0], args[1], args[2], args[3], 2)
        end

        def set_top_left_cell(row, col = nil)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
          else
            r = row
            c = col
          end
          @top_left_cell = xl_rowcol_to_cell(r, c)
        end

        # Adds an Excel table to the worksheet
        def add_table(*args)
          if (row_col_array = row_col_notation(args.first))
            r1, c1, r2, c2 = row_col_array
            options = args[1] || {}
          else
            r1, c1, r2, c2, options = args
            options ||= {}
          end

          r1, r2 = r2, r1 if r1 > r2
          c1, c2 = c2, c1 if c1 > c2

          tbl = options.dup
          tbl[:range] = xl_range(r1, r2, c1, c2)
          tbl[:name] ||= "Table#{@tables.size + 1}"
          @tables << tbl
          tbl
        end

        # Adds sparklines to the worksheet
        #
        # @param options [Hash{Symbol => untyped}]
        # @return [void]
        #: (Hash[Symbol, untyped] options) -> void
        def add_sparkline(options)
          raise WriteXLSXInsufficientArgumentError, "Parameter ':location' is required in add_sparkline()" unless options[:location]
          raise WriteXLSXInsufficientArgumentError, "Parameter ':range' is required in add_sparkline()" unless options[:range]

          @sparklines << options
        end

        # Adds data validation rule
        def data_validation(*args)
          if (row_col_array = row_col_notation(args.first))
            if row_col_array.size == 4
              r1, c1, r2, c2 = row_col_array
              options = args[1]
            else
              r1, c1 = row_col_array
              r2, c2, options = args[1..3]
            end
          else
            r1, c1, r2, c2, options = args
          end

          opts = options || {}
          sqref = xl_range(r1, r2, c1, c2)
          dv = opts.dup
          dv[:sqref] = sqref
          @data_validations << dv
        end

        # Adds conditional formatting rule
        def conditional_formatting(*args)
          if (row_col_array = row_col_notation(args.first))
            if row_col_array.size == 4
              r1, c1, r2, c2 = row_col_array
              options = args[1]
            else
              r1, c1 = row_col_array
              options = args[1]
              r2 = r1
              c2 = c1
            end
          else
            r1, c1, r2, c2, options = args
          end

          opts = options || {}
          sqref = xl_range(r1, r2, c1, c2)
          cf = opts.dup
          cf[:sqref] = sqref
          @cond_formats << cf
        end

        # Inserts image into worksheet
        def insert_image(*args)
          if (row_col_array = row_col_notation(args.first))
            r, c = row_col_array
            img = args[1]
            x_off = args[2] || 0
            y_off = args[3] || 0
            s_x = args[4] || 1
            s_y = args[5] || 1
          else
            r = args[0]
            c = args[1]
            img = args[2]
            x_off = args[3] || 0
            y_off = args[4] || 0
            s_x = args[5] || 1
            s_y = args[6] || 1
          end

          @images << {
            row: r,
            col: c,
            image: img,
            x_offset: x_off,
            y_offset: y_off,
            scale_x: s_x,
            scale_y: s_y
          }
        end

        # Inserts chart into worksheet
        def insert_chart(*args)
          if (row_col_array = row_col_notation(args.first))
            r, c = row_col_array
            ch = args[1]
            x_off = args[2] || 0
            y_off = args[3] || 0
            s_x = args[4] || 1
            s_y = args[5] || 1
          else
            r = args[0]
            c = args[1]
            ch = args[2]
            x_off = args[3] || 0
            y_off = args[4] || 0
            s_x = args[5] || 1
            s_y = args[6] || 1
          end

          @charts << {
            chart: ch,
            row: r,
            col: c,
            x_offset: x_off,
            y_offset: y_off,
            scale_x: s_x,
            scale_y: s_y
          }
        end

        # Page setup & print option methods
        def set_landscape
          @page_setup[:orientation] = :landscape
        end

        def set_portrait
          @page_setup[:orientation] = :portrait
        end

        def set_paper(paper_size)
          @page_setup[:paper_size] = paper_size
        end

        def set_margins(left, right, top, bottom)
          @page_margins[:left] = left
          @page_margins[:right] = right
          @page_margins[:top] = top
          @page_margins[:bottom] = bottom
        end

        def set_margin_left(m)
          @page_margins[:left] = m
        end

        def set_margin_right(m)
          @page_margins[:right] = m
        end

        def set_margin_top(m)
          @page_margins[:top] = m
        end

        def set_margin_bottom(m)
          @page_margins[:bottom] = m
        end

        def set_margins_LR(m)
          @page_margins[:left] = m
          @page_margins[:right] = m
        end

        def set_margins_TB(m)
          @page_margins[:top] = m
          @page_margins[:bottom] = m
        end

        def set_header(string, margin = nil)
          @header_footer[:odd_header] = string
          @page_margins[:header] = margin if margin
        end

        def set_footer(string, margin = nil)
          @header_footer[:odd_footer] = string
          @page_margins[:footer] = margin if margin
        end

        def print_area(*args)
          if (row_col_array = row_col_notation(args.first))
            r1, c1, r2, c2 = row_col_array
          else
            r1, c1, r2, c2 = args
          end
          @print_area = xl_range_formula(@name, r1, r2, c1, c2)
          @clean_print_area = "#{xl_rowcol_to_cell(r1, c1)}:#{xl_rowcol_to_cell(r2, c2)}"
        end

        def fit_to_pages(width, height)
          @page_setup[:fit_to_width] = width
          @page_setup[:fit_to_height] = height
          @page_setup[:fit_to_page] = true
        end

        def print_scale=(scale)
          @page_setup[:scale] = scale
        end

        def set_print_scale(scale)
          self.print_scale = scale
        end

        def repeat_rows(r1, r2)
          @repeat_rows = [r1, r2]
        end

        def repeat_columns(c1, c2)
          @repeat_cols = [c1, c2]
        end

        def print_across(_val = true)
          @page_setup[:page_order] = :overThenDown
        end

        def print_row_col_headers(val = true)
          @print_options[:headings] = [true, 1].include?(val)
        end

        def print_black_and_white(val = true)
          @page_setup[:black_and_white] = [true, 1].include?(val)
        end

        def center_horizontally(val = true)
          @print_options[:horizontal_centered] = [true, 1].include?(val)
        end

        def center_vertically(val = true)
          @print_options[:vertical_centered] = [true, 1].include?(val)
        end

        def set_h_pagebreaks(*breaks)
          @h_breaks.concat(breaks.flatten)
        end

        def set_v_pagebreaks(*breaks)
          @v_breaks.concat(breaks.flatten)
        end

        def set_page_view(type = 1)
          @page_setup[:page_view] = type
        end

        # Worksheet state and display options
        def hide
          @hidden = true
        end

        def hidden?
          @hidden
        end

        def activate
          @active = true
        end

        def select
          @selected = true
        end

        def hide_gridlines(option = 1)
          @hide_gridlines = option != 0
        end

        def hide_row_col_headers(option = 1)
          @hide_row_col_headers = option != 0
        end

        def hide_zero(option = 1)
          @hide_zero = option != 0
        end

        def right_to_left(option = true)
          @right_to_left = [true, 1].include?(option)
        end

        def right_to_left?
          @right_to_left
        end

        def set_tab_color(color)
          @tab_color = color
        end

        def set_zoom(zoom)
          @zoom = zoom
        end

        def keep_leading_zeros(option = true)
          @leading_zeros = [true, 1].include?(option)
        end

        # Converts mutable worksheet to an immutable Xlsxrb::Elements::Worksheet
        #
        # @param style_map [Hash{Format => Integer}, nil]
        # @return [Xlsxrb::Elements::Worksheet]
        def to_xlsxrb(style_map = nil)
          style_map ||= if @workbook.respond_to?(:compile_styles, true)
                          @workbook.send(:compile_styles)[1]
                        else
                          {}
                        end
          if @rows.empty? && !@table.empty?
            @table.each do |(r, c), rec|
              (@rows[r] ||= {})[c] = rec
            end
          end

          all_row_indices = (@rows.keys + @set_rows.keys).uniq.sort
          elements_rows = []

          all_row_indices.each do |r_idx|
            cells_hash = @rows[r_idx] || {}
            max_c_idx = cells_hash.keys.max || -1

            row_cells = []
            0.upto(max_c_idx) do |c_idx|
              rec = cells_hash[c_idx]
              if rec
                val = rec.value
                val = format_datetime_val(val) if val.is_a?(Time) || val.is_a?(Date)

                fml = if rec.array_formula && rec.formula
                        Xlsxrb.array_formula(rec.formula, rec.array_formula, cached_value: rec.cached_value || 0)
                      elsif rec.formula
                        Xlsxrb.formula(rec.formula, cached_value: rec.cached_value || 0)
                      end

                fmt = rec.format || @col_info.dig(c_idx, :format) || @set_rows.dig(r_idx, :format)
                s_idx = fmt ? style_map[fmt] : nil

                cell_val = rec.formula ? (rec.cached_value || 0) : val

                row_cells << Xlsxrb::Elements::Cell.fast_create(
                  r_idx,
                  c_idx,
                  cell_val,
                  s_idx,
                  fml,
                  cell_val&.to_s
                )
              else
                col_fmt = @col_info.dig(c_idx, :format) || @set_rows.dig(r_idx, :format)
                s_idx = col_fmt ? style_map[col_fmt] : nil
                row_cells << Xlsxrb::Elements::Cell.fast_create(r_idx, c_idx, nil, s_idx)
              end
            end

            row_setting = @set_rows[r_idx] || {}
            elements_rows << Xlsxrb::Elements::Row.new(
              index: r_idx,
              cells: row_cells,
              height: row_setting[:height]
            )
          end

          cols_data = []
          all_col_indices = @col_info.keys.sort
          all_col_indices.each do |c_idx|
            info = @col_info[c_idx]
            fmt = info[:format]
            s_idx = fmt ? style_map[fmt] : nil
            col_unmapped = {}
            col_unmapped[:style_index] = s_idx if s_idx
            cols_data << Xlsxrb::Elements::Column.new(
              index: c_idx,
              width: info[:width] || 8.43,
              hidden: info[:hidden] || false,
              outline_level: info[:level],
              style_index: s_idx,
              unmapped_data: col_unmapped
            )
          end

          facade_meta = {}
          facade_meta[:merge_cells] = @merge.map { |m| "#{xl_rowcol_to_cell(m[0], m[1])}:#{xl_rowcol_to_cell(m[2], m[3])}" } unless @merge.empty?
          facade_meta[:auto_filter] = @autofilter_ref if @autofilter_ref
          facade_meta[:page_margins] = @page_margins.compact unless @page_margins.empty?
          facade_meta[:page_setup] = @page_setup.compact unless @page_setup.empty?
          facade_meta[:header_footer] = @header_footer.compact unless @header_footer.empty?
          facade_meta[:print_options] = @print_options.compact unless @print_options.empty?
          facade_meta[:tables] = @tables unless @tables.empty?

          hl_hash = {}
          @hyperlinks.each do |h|
            ref = h[:ref] || h[:cell]
            hl_hash[ref] = { cell: ref, ref: ref, url: h[:url], display: h[:display], tooltip: h[:tooltip], location: h[:location] }.compact
          end
          facade_meta[:hyperlinks] = @hyperlinks unless @hyperlinks.empty?

          pt = nil
          if @repeat_rows || @repeat_cols
            pt = {}
            pt[:rows] = "#{@repeat_rows[0] + 1}:#{@repeat_rows[1] + 1}" if @repeat_rows
            if @repeat_cols
              c1_name = @repeat_cols[0].is_a?(Integer) ? xl_col_to_name(@repeat_cols[0]) : @repeat_cols[0].to_s
              c2_name = @repeat_cols[1].is_a?(Integer) ? xl_col_to_name(@repeat_cols[1]) : @repeat_cols[1].to_s
              pt[:cols] = "#{c1_name}:#{c2_name}"
            end
          end

          sparkline_groups_data = @sparklines.map do |param|
            type = param[:type]&.to_s || "line"
            type = "stacked" if type == "win_loss"

            locations = Array(param[:location]).map { |l| l.to_s.delete("$") }
            ranges = Array(param[:range]).map do |rng|
              r_str = rng.to_s.delete("$")
              r_str.include?("!") ? r_str : "'#{@name}'!#{r_str}"
            end

            sparklines = locations.zip(ranges).map do |loc, rng|
              { location_ref: loc, data_ref: rng }
            end

            {
              type: type,
              sparklines: sparklines,
              high: [true, 1].include?(param[:high_point] || param[:high]) || nil,
              low: [true, 1].include?(param[:low_point] || param[:low]) || nil,
              negative: [true, 1].include?(param[:negative_points] || param[:negative]) || nil,
              first: [true, 1].include?(param[:first_point] || param[:first]) || nil,
              last: [true, 1].include?(param[:last_point] || param[:last]) || nil,
              markers: [true, 1].include?(param[:markers]) || nil,
              min: param[:min],
              max: param[:max],
              line_weight: param[:weight] || param[:line_weight],
              color_series: Colors.to_hex(param[:series_color]),
              color_negative: Colors.to_hex(param[:negative_color]),
              color_markers: Colors.to_hex(param[:markers_color]),
              color_first: Colors.to_hex(param[:first_color]),
              color_last: Colors.to_hex(param[:last_color]),
              color_high: Colors.to_hex(param[:high_color]),
              color_low: Colors.to_hex(param[:low_color])
            }.compact
          end

          view_props = {}
          view_props[:show_grid_lines] = false if @hide_gridlines
          view_props[:show_row_col_headers] = false if @hide_row_col_headers
          view_props[:show_zeros] = false if @hide_zero
          view_props[:right_to_left] = true if @right_to_left
          view_props[:zoom_scale] = @zoom if @zoom

          if @panes
            r, c, _tr, _lc, type = @panes
            view_props[:pane] = {
              state: type == 2 ? :split : :frozen,
              x_split: c,
              y_split: r,
              top_left_cell: @top_left_cell
            }.compact
          end

          facade_meta[:views] = [view_props] unless view_props.empty?

          if @tab_color
            hex_color = Colors.to_hex(@tab_color)
            facade_meta[:tab_color] = hex_color if hex_color
          end

          charts_data = @charts.map do |ch_entry|
            chart_obj = ch_entry[:chart]
            ch_opts = chart_obj.respond_to?(:to_chart_options) ? chart_obj.to_chart_options : {}
            ch_r = ch_entry[:row] || 0
            ch_c = ch_entry[:col] || 0
            sx = ch_entry[:scale_x] || 1
            sy = ch_entry[:scale_y] || 1
            w_cols = (15 * sx).round
            h_rows = (15 * sy).round
            ch_opts.merge(
              from_col: ch_c,
              from_row: ch_r,
              to_col: ch_c + w_cols,
              to_row: ch_r + h_rows
            )
          end

          Xlsxrb::Elements::Worksheet.new(
            name: @name,
            rows: elements_rows,
            columns: cols_data,
            charts: charts_data,
            conditional_formatting: @cond_formats.empty? ? nil : @cond_formats,
            data_validations: @data_validations,
            unmapped_data: facade_meta,
            state: @hidden ? :hidden : :visible,
            hyperlinks: hl_hash,
            print_area: @clean_print_area,
            print_titles: pt,
            sparkline_groups: sparkline_groups_data
          )
        end

        private

        def normalize_cell_args(row, col, value, format)
          if (row_col_array = row_col_notation(row))
            r, c = row_col_array
            val = col
            fmt = value
          else
            r = row
            c = col
            val = value
            fmt = format
          end
          [r, c, val, fmt]
        end

        def update_dimensions(row, col)
          @dim_rowmin = row if @dim_rowmin.nil? || row < @dim_rowmin
          @dim_rowmax = row if @dim_rowmax.nil? || row > @dim_rowmax
          @dim_colmin = col if @dim_colmin.nil? || col < @dim_colmin
          @dim_colmax = col if @dim_colmax.nil? || col > @dim_colmax
        end

        def format_datetime_val(val)
          is_1904 = @workbook.respond_to?(:date_1904?) && @workbook.date_1904?
          convert_date_time(val, is_1904)
        end

        def store_cell_record(row, col, record)
          @table[[row, col]] = record
          (@rows[row] ||= {})[col] = record
          @formats_used << record.format if record.format
        end
      end
    end
  end
end
