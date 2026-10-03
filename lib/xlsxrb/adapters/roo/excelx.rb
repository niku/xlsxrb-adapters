# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "stringio"
require "tempfile"
require "uri"
require "xlsxrb"
require_relative "constants"
require_relative "coordinate"
require_relative "link"
require_relative "font"
require_relative "utils"
require_relative "format"
require_relative "cell"
require_relative "sheet"
require_relative "base"

module Xlsxrb
  module Adapters
    module Roo
      # High-performance, low-memory Excel (.xlsx and .xlsm) reader adapter
      # matching Roo::Excelx backed by the xlsxrb engine.
      class Excelx < Base
        ERROR_VALUES = Roo::ERROR_VALUES
        class ExceedsMaxError < StandardError
        end

        # Defined name label matching Roo::Excelx::Workbook::Label.
        class Label
          attr_reader :name, :sheet, :row, :col

          # @param name [String]
          # @param sheet [String]
          # @param row [Integer] 1-based row index.
          # @param col [Integer] 1-based column index.
          #: (String name, String sheet, Integer row, Integer col) -> void
          def initialize(name, sheet, row, col)
            @name = name
            @sheet = sheet
            @row = row.to_i
            @col = col.to_i
          end

          # Coordinate key [row, col].
          #
          # @return [Array(Integer, Integer)]
          #: () -> [Integer, Integer]
          def key
            [@row, @col]
          end
        end

        # Workbook metadata proxy matching Roo::Excelx::Workbook.
        class WorkbookProxy
          attr_reader :defined_names, :base_date, :base_timestamp, :sheets_data

          # @param defined_names [Hash[String, Label]]
          # @param base_date [::Date]
          # @param sheets_data [Array[Hash[Symbol, untyped]]]
          #: (Hash[String, Label] defined_names, ::Date base_date, Array[Hash[Symbol, untyped]] sheets_data) -> void
          def initialize(defined_names, base_date, sheets_data)
            @defined_names = defined_names
            @base_date = base_date
            @base_timestamp = base_date.to_datetime.to_time.to_i
            @sheets_data = sheets_data
          end

          # Sheet metadata list.
          #: () -> Array[Hash[Symbol, untyped]]
          def sheets
            @sheets_data
          end
        end

        attr_reader :workbook

        # Initializes and parses an XLSX/XLSM file or stream.
        #
        # @param filename_or_stream [String, Pathname, File, IO, StringIO]
        # @param options [Hash]
        #: (untyped filename_or_stream, ?Hash[Symbol, untyped] options) -> void
        def initialize(filename_or_stream, options = {})
          super()
          @filename = filename_or_stream
          @options = options.dup

          packed = @options[:packed]
          file_warning = @options.fetch(:file_warning, :error)
          cell_max = @options.delete(:cell_max)
          only_visible = @options[:only_visible_sheets] || false

          io = prepare_input(filename_or_stream, packed, file_warning)
          parse_archive(io, only_visible)

          @default_sheet = @sheet_names.first

          return unless cell_max

          target_s = sheet_for(@options.delete(:sheet))
          dims = target_s&.dimensions
          return unless dims

          cell_count = Utils.num_cells_in_range(dims)
          return unless cell_count > cell_max

          raise ExceedsMaxError, "Excel file exceeds cell maximum: #{cell_count} > #{cell_max}"
        end

        # Returns list of sheet names.
        #
        # @return [Array<String>]
        #: () -> Array[String]
        def sheets
          @sheet_names
        end

        # Returns sheet instance for name or index.
        #
        # @param sheet [String, Integer, nil]
        # @return [Excelx::Sheet]
        #: (?untyped sheet) -> Excelx::Sheet?
        def sheet_for(sheet = nil)
          target = sheet || default_sheet
          validate_sheet!(target)
          target.is_a?(String) ? @sheets_by_name[target] : @sheets[target]
        end

        # Returns whether the specified sheet is hidden.
        #
        # @param sheet [String, Integer, nil]
        # @return [Boolean]
        #: (?untyped sheet) -> bool
        def sheet_hidden?(sheet = nil)
          sheet_for(sheet)&.hidden? || false
        end

        # Returns whether the specified sheet is visible.
        #
        # @param sheet [String, Integer, nil]
        # @return [Boolean]
        #: (?untyped sheet) -> bool
        def sheet_visible?(sheet = nil)
          sheet_for(sheet)&.visible? || false
        end

        # Returns internal type of cell: [:numeric_or_formula, format] or :string or :boolean.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Object, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> untyped
        def excelx_type(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.cell_at([r, c])&.cell_type
        end

        # Returns unformatted string value as stored in XML.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [String, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> untyped
        def excelx_value(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.cell_at([r, c])&.cell_value
        end

        # Returns formatted value string matching cell format code.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [String, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> untyped
        def formatted_value(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.cell_at([r, c])&.formatted_value
        end

        # Returns format code string for cell.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [String, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> String?
        def excelx_format(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.excelx_format([r, c])
        end

        # Returns formula expression string (without '=') or nil.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [String, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> String?
        def formula(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.cell_at([r, c])&.formula
        end

        # Returns whether the cell contains a formula.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Boolean]
        #: (*untyped args) -> bool
        def formula?(*)
          !!formula(*)
        end

        # Returns all formulas in sheet as [[row, col, formula], ...].
        #
        # @param sheet [String, Integer, nil]
        # @return [Array<Array(Integer, Integer, String)>]
        #: (?untyped sheet) -> Array[[Integer, Integer, String]]
        def formulas(sheet = nil)
          s_obj = sheet_for(sheet)
          return [] unless s_obj

          s_obj.cells.filter_map do |(r, c), cell_obj|
            [r, c, cell_obj.formula] if cell_obj.formula?
          end
        end

        # Returns Font object for cell.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Font, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> Font?
        def font(row, col, sheet = nil)
          r, c = normalize(row, col)
          cell_obj = sheet_for(sheet)&.cells&.[]([r, c])
          return nil unless cell_obj

          style_id = cell_obj.style || 0
          font_id = @cell_xfs[style_id]&.[](:font_id) || 0
          @fonts[font_id]
        end

        # Returns cell type Symbol.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Symbol, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> Symbol?
        def celltype(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.cell_at([r, c])&.type
        end

        # Returns comment string for cell or nil.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [String, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> String?
        def comment(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.comments&.[]([r, c])
        end

        # Returns whether the cell contains a comment.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Boolean]
        #: (untyped row, untyped col, ?untyped sheet) -> bool
        def comment?(row, col, sheet = nil)
          !comment(row, col, sheet).nil?
        end

        # Returns all comments in sheet as [[row, col, comment_text], ...].
        #
        # @param sheet [String, Integer, nil]
        # @return [Array<Array(Integer, Integer, String)>]
        #: (?untyped sheet) -> Array[[Integer, Integer, String]]
        def comments(sheet = nil)
          s_obj = sheet_for(sheet)
          return [] unless s_obj

          s_obj.comments.map do |(r, c), text|
            [r, c, text]
          end
        end

        # Returns hyperlink target URL string or nil.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [String, nil]
        #: (untyped row, untyped col, ?untyped sheet) -> String?
        def hyperlink(row, col, sheet = nil)
          r, c = normalize(row, col)
          sheet_for(sheet)&.hyperlinks&.[]([r, c])
        end

        # Returns whether the cell has a hyperlink.
        #
        # @param row [Integer, String]
        # @param col [Integer, String, Symbol]
        # @param sheet [String, Integer, nil]
        # @return [Boolean]
        #: (untyped row, untyped col, ?untyped sheet) -> bool
        def hyperlink?(row, col, sheet = nil)
          !hyperlink(row, col, sheet).nil?
        end

        # Returns coordinate [row, col, sheet] for defined name or [nil, nil, nil].
        #
        # @param name [String, Symbol]
        # @return [Array(Integer, Integer, String), Array(nil, nil, nil)]
        #: (untyped name) -> Array[untyped]
        def label(name)
          lbl = @workbook.defined_names[name.to_s]
          return [nil, nil, nil] unless lbl

          [lbl.row, lbl.col, lbl.sheet]
        end

        # Returns all defined names as [[name, [row, col, sheet]], ...].
        #
        # @return [Array<Array(String, Array(Integer, Integer, String))>]
        #: () -> Array[[String, [Integer, Integer, String]]]
        def labels
          @workbook.defined_names.map do |name, lbl|
            [name, [lbl.row, lbl.col, lbl.sheet]]
          end
        end

        # Returns list of image paths for sheet.
        #
        # @param sheet [String, Integer, nil]
        # @return [Array<String>]
        #: (?untyped sheet) -> Array[String]
        def images(sheet = nil)
          sheet_for(sheet)&.images || []
        end

        # Yields each row as an Array of Excelx::Cell objects.
        #
        # @param options [Hash]
        # @yield [row]
        # @return [Enumerator, void]
        #: (?Hash[Symbol, untyped] options) { (Array[untyped]) -> void } -> void
        #: (?Hash[Symbol, untyped] options) -> Enumerator[Array[untyped], void]
        def each_row_streaming(options = {}, &)
          sheet_target = options.delete(:sheet)
          s_obj = sheet_for(sheet_target)
          return [].to_enum unless s_obj

          s_obj.each_row(options, &)
        end

        alias each_row each_row_streaming

        # Resolves defined names or coordinates via method_missing.
        #: (Symbol method_name, *untyped args) -> untyped
        def method_missing(method_name, *args)
          lbl = @workbook.defined_names[method_name.to_s]
          if lbl
            cell(lbl.row, lbl.col, lbl.sheet)
          else
            super
          end
        end

        #: (Symbol method_name, ?bool include_private) -> bool
        def respond_to_missing?(method_name, include_private = false)
          @workbook&.defined_names&.key?(method_name.to_s) || super
        end

        # Converts this Excelx adapter instance to an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          elements_sheets = @sheets.map do |s|
            rows_hash = {}
            sheet_hyperlinks = {}
            sheet_comments = []

            s.cells.each do |(r, c), cell_obj|
              row_idx = r - 1
              col_idx = c - 1
              rows_hash[row_idx] ||= []
              formula_obj = cell_obj.formula ? Xlsxrb::Elements::Formula.new(expression: cell_obj.formula) : nil
              ref = "#{Xlsxrb::Elements::Cell.column_letter(col_idx)}#{r}"

              link_url = cell_obj.value.is_a?(Link) ? cell_obj.value.url : s.hyperlinks[[r, c]]
              hl = if link_url
                     h_entry = { ref: ref, url: link_url }
                     sheet_hyperlinks[ref] = h_entry
                     h_entry
                   end

              comment_str = s.comments[[r, c]]
              cm = if comment_str
                     c_entry = { ref: ref, text: comment_str }
                     sheet_comments << c_entry
                     c_entry
                   end

              fmt = cell_obj.format if cell_obj.respond_to?(:format)
              raw_v = cell_obj.cell_value.to_s if cell_obj.respond_to?(:cell_value)

              rows_hash[row_idx] << Xlsxrb::Elements::Cell.new(
                row_index: row_idx,
                column_index: col_idx,
                value: cell_obj.value,
                formula: formula_obj,
                style_index: cell_obj.style,
                hyperlink: hl,
                comment: cm,
                format_code: fmt,
                raw_value: raw_v
              )
            end

            elem_rows = rows_hash.sort_by { |idx, _| idx }.map do |idx, row_cells|
              sorted_cells = row_cells.sort_by(&:column_index)
              Xlsxrb::Elements::Row.new(index: idx, cells: sorted_cells)
            end

            sheet_state = s.state ? s.state.to_sym : :visible
            Xlsxrb::Elements::Worksheet.new(
              name: s.name,
              rows: elem_rows,
              state: sheet_state,
              hyperlinks: sheet_hyperlinks,
              comments: sheet_comments
            )
          end

          dns = @workbook.defined_names.map do |name, lbl|
            { name: name, value: "'#{lbl.sheet}'!$#{Xlsxrb::Elements::Cell.column_letter(lbl.col - 1)}$#{lbl.row}" }
          end

          Xlsxrb::Elements::Workbook.new(
            sheets: elements_sheets,
            defined_names: dns
          )
        end

        # Constructs an Excelx instance from an Xlsxrb::Elements::Workbook.
        #
        # @param xlsxrb_wb [Xlsxrb::Elements::Workbook]
        # @return [Excelx]
        #: (Xlsxrb::Elements::Workbook xlsxrb_wb) -> Excelx
        def self.from_xlsxrb(xlsxrb_wb)
          binary_data = Xlsxrb.write(xlsxrb_wb)
          new(StringIO.new(binary_data))
        end

        private

        #: (untyped source, untyped packed, untyped file_warning) -> untyped
        def prepare_input(source, _packed, file_warning)
          return source if is_stream?(source)

          if source.is_a?(String) || (defined?(Pathname) && source.is_a?(Pathname))
            str = source.to_s
            validate_file_extension(str, file_warning)
            if str.start_with?("http://", "https://")
              download_to_io(str)
            else
              File.open(str, "rb")
            end
          elsif source.respond_to?(:path)
            validate_file_extension(source.path, file_warning)
            File.open(source.path, "rb")
          else
            source
          end
        end

        #: (String path, Symbol file_warning) -> void
        def validate_file_extension(path, file_warning)
          return if file_warning == :ignore

          ext = File.extname(path.split("?").first.to_s).downcase
          return if %w[.xlsx .xlsm].include?(ext)

          msg = "#{path} is not an Excel 2007 file"
          raise TypeError, msg unless file_warning == :warning

          warn msg
        end

        # rubocop:disable Security/Open
        #: (String uri_str) -> StringIO
        def download_to_io(uri_str)
          require "open-uri"
          data = URI.open(uri_str, "rb", &:read)
          StringIO.new(data.b)
        end
        # rubocop:enable Security/Open

        #: (untyped obj) -> bool
        def is_stream?(obj)
          obj.respond_to?(:read) && obj.respond_to?(:seek)
        end

        #: (untyped io, bool only_visible) -> void
        def parse_archive(io, only_visible)
          io.rewind if io.respond_to?(:rewind)
          zip = Xlsxrb::Ooxml::ZipReader.open(io)

          # 1. Styles
          styles_xml = read_xml(zip, "xl/styles.xml")
          @parsed_styles = styles_xml ? Xlsxrb::Ooxml::StylesParser.parse(styles_xml) : {}
          num_fmts = @parsed_styles[:num_fmts] || {}
          @cell_xfs = @parsed_styles[:cell_xfs] || []
          raw_fonts = @parsed_styles[:fonts] || []
          @fonts = raw_fonts.map do |f_hash|
            Font.new(
              bold: f_hash[:bold] || false,
              italic: f_hash[:italic] || false,
              underline: !f_hash[:underline].nil?
            )
          end
          @fonts[0] ||= Font.new

          # 2. Shared Strings (both plain text and rich text HTML)
          sst_xml = read_xml(zip, "xl/sharedStrings.xml")
          shared_strings, html_strings = parse_shared_strings(sst_xml)

          # 3. Workbook
          wb_xml = read_xml(zip, "xl/workbook.xml")
          wb_rels_xml = read_xml(zip, "xl/_rels/workbook.xml.rels")
          wb_rels = wb_rels_xml ? Xlsxrb::Ooxml::RelationshipsParser.parse(wb_rels_xml) : {}

          sheets_info, defined_names, base_date = parse_workbook_xml(wb_xml, only_visible)
          @workbook = WorkbookProxy.new(defined_names, base_date, sheets_info)

          # 4. Images in archive
          media_images = zip.entry_names.select { |name| name.start_with?("xl/media/") }

          # 5. Parse Worksheets
          @sheet_names = []
          @sheets = []
          @sheets_by_name = {}

          sheets_info.each_with_index do |s_info, idx|
            s_name = s_info[:name]
            r_id = s_info[:r_id]
            target = wb_rels[r_id]
            next unless target

            sheet_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
            next unless zip.entry?(sheet_path)

            sheet_xml = read_xml(zip, sheet_path) || ""
            sheet_rels_path = sheet_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
            sheet_rels_xml = read_xml(zip, sheet_rels_path)

            sheet_obj = parse_sheet(
              s_name,
              sheet_xml,
              sheet_rels_xml,
              zip,
              shared_strings,
              html_strings,
              num_fmts,
              base_date,
              media_images,
              state: s_info[:state] || :visible
            )

            @sheet_names << s_name
            @sheets[idx] = sheet_obj
            @sheets_by_name[s_name] = sheet_obj
            @sheets_by_name[s_name.downcase] ||= sheet_obj
          end

          @default_sheet = @sheet_names.first
        end

        #: (String? xml) -> [Array[String], Hash[Integer, String]]
        def parse_shared_strings(xml)
          return [[], {}] unless xml && !xml.empty?

          plain_strings = []
          html_strings = {}

          idx = 0
          xml.scan(%r{<si>(.*?)</si>}m) do |match|
            si_content = match[0]
            if si_content.include?("<r>") || si_content.include?("<r ")
              plain_buf = String.new(encoding: Encoding::UTF_8)
              html_buf = String.new(encoding: Encoding::UTF_8)
              has_formatting = false

              si_content.scan(%r{<r>(.*?)</r>}m) do |r_match|
                r_content = r_match[0]
                t_match = r_content[%r{<t(?:\s+[^>]*)?>(.*?)</t>}m, 1]
                t_text = t_match ? unescape_xml(t_match) : ""
                t_text = t_text.gsub("_x000D_", "\n").gsub("\r\n", "\n").gsub("\r", "\n")
                plain_buf << t_text

                rpr = r_content[%r{<rPr>(.*?)</rPr>}m, 1]
                if rpr
                  open_tags = []
                  close_tags = []
                  if rpr =~ /<vertAlign[^>]*val="subscript"/
                    open_tags << "<sub>"
                    close_tags.unshift("</sub>")
                    has_formatting = true
                  elsif rpr =~ /<vertAlign[^>]*val="superscript"/
                    open_tags << "<sup>"
                    close_tags.unshift("</sup>")
                    has_formatting = true
                  end
                  if rpr.include?("<b/>") || rpr.include?("<b ") || rpr.include?("<b>")
                    open_tags << "<b>"
                    close_tags.unshift("</b>")
                    has_formatting = true
                  end
                  if rpr.include?("<i/>") || rpr.include?("<i ") || rpr.include?("<i>")
                    open_tags << "<i>"
                    close_tags.unshift("</i>")
                    has_formatting = true
                  end
                  if rpr.include?("<u/>") || rpr.include?("<u ") || rpr.include?("<u>")
                    open_tags << "<u>"
                    close_tags.unshift("</u>")
                    has_formatting = true
                  end
                  html_buf << open_tags.join << t_text << close_tags.join
                else
                  html_buf << t_text
                end
              end

              plain_strings << plain_buf
              html_strings[idx] = has_formatting ? "<html>#{html_buf}</html>" : plain_buf
            else
              t_match = si_content[%r{<t(?:\s+[^>]*)?>(.*?)</t>}m, 1]
              text = t_match ? unescape_xml(t_match) : ""
              text = text.gsub("_x000D_", "\n").gsub("\r\n", "\n").gsub("\r", "\n")
              plain_strings << text
              html_strings[idx] = text
            end
            idx += 1
          end

          [plain_strings, html_strings]
        end

        #: (String? xml, bool only_visible) -> [Array[Hash[Symbol, untyped]], Hash[String, Label], ::Date]
        def parse_workbook_xml(xml, only_visible)
          return [[], {}, ::Date.new(1899, 12, 30)] unless xml && !xml.empty?

          parsed = Xlsxrb::Ooxml::WorkbookParser.parse_with_properties(xml)

          sheets_info = []
          parsed[:sheets].each do |s|
            state_str = s[:state].to_s
            next if only_visible && %i[hidden very_hidden].include?(s[:state])

            s_id = s[:sheet_id]&.to_s
            r_id = s[:r_id]&.to_s
            sheets_info << {
              "name" => s[:name],
              :name => s[:name],
              "sheetId" => s_id,
              :sheet_id => s_id,
              "r:id" => r_id,
              :r_id => r_id,
              "state" => state_str,
              :state => s[:state]
            }
          end

          base_date = parsed[:date1904] ? ::Date.new(1904, 1, 1) : ::Date.new(1899, 12, 30)

          defined_names = {}
          parsed[:defined_names].each do |dn|
            formula_text = dn[:value]
            next unless formula_text
            next unless formula_text =~ /^'?(.*?)'?!(\$?[A-Za-z]+)(\$?\d+)$/

            s_name = ::Regexp.last_match(1)
            col_letter = ::Regexp.last_match(2).delete("$")
            row_num = ::Regexp.last_match(3).delete("$").to_i
            col_num = Utils.letter_to_number(col_letter)
            defined_names[dn[:name]] = Label.new(dn[:name], s_name, row_num, col_num)
          end

          [sheets_info, defined_names, base_date]
        end

        #: (String s_name, String sheet_xml, String? sheet_rels_xml, untyped zip, Array[String] shared_strings, Hash[Integer, String] html_strings, Hash[Integer, String] num_fmts, ::Date base_date, Array[String] media_images, ?state: (Symbol | String)) -> Excelx::Sheet
        def parse_sheet(s_name, sheet_xml, sheet_rels_xml, zip, shared_strings, html_strings, num_fmts, base_date, media_images, state: :visible)
          dim_ref = sheet_xml[/<dimension\s+ref="([^"]+)"/, 1]

          # Parse rels for hyperlinks, comments, drawings
          hyperlinks_map = {}
          comments_map = {}
          sheet_images = []

          if sheet_rels_xml && !sheet_rels_xml.empty?
            rels_by_id = {}
            comments_target = nil
            drawing_target = nil

            sheet_rels_xml.scan(%r{<Relationship\s+([^>]+)/?>}) do |m|
              attrs = m[0]
              r_id = attrs[/Id="([^"]+)"/, 1]
              r_type = attrs[/Type="([^"]+)"/, 1]
              target = attrs[/Target="([^"]+)"/, 1]

              if r_type&.include?("/hyperlink")
                rels_by_id[r_id] = target
              elsif r_type&.include?("/comments")
                comments_target = target
              elsif r_type&.include?("/drawing")
                drawing_target = target
              end
            end

            # Map hyperlinks
            unless @options[:no_hyperlinks]
              Xlsxrb::Ooxml::WorksheetParser.parse_hyperlinks(sheet_xml).each do |hl|
                ref = hl[:ref]
                r_id = hl[:rid]
                location = hl[:location]
                target_url = rels_by_id[r_id] || location
                next unless target_url

                full_url = rels_by_id[r_id] && location ? "#{rels_by_id[r_id]}##{location}" : target_url
                Utils.coordinates_in_range(ref) do |coord|
                  hyperlinks_map[coord] = full_url
                end
              end
            end

            # Parse comments
            if comments_target
              c_path = resolve_zip_path("xl/worksheets", comments_target)
              if zip.entry?(c_path)
                c_xml = read_xml(zip, c_path)
                comments_map = parse_comments_xml(c_xml)
              end
            end

            # Parse drawings for images
            if drawing_target
              d_path = resolve_zip_path("xl/worksheets", drawing_target)
              d_rels_path = d_path.sub(%r{([^/]+)$}, '_rels/\1.rels')
              if zip.entry?(d_rels_path)
                d_rels_xml = read_xml(zip, d_rels_path)
                d_rels_xml&.scan(/<Relationship[^>]+Target="([^"]+)"/) do |im|
                  img_target = im[0]
                  if img_target =~ %r{media/image\d+}
                    resolved = resolve_zip_path(File.dirname(d_path), img_target)
                    sheet_images << resolved if media_images.include?(resolved)
                  end
                end
              end
            end
          end

          # Parse merged ranges
          merged_ranges = []
          sheet_xml.scan(/<mergeCell\s+ref="([^"]+)"/) do |m|
            merged_ranges << m[0]
          end

          # Parse cells
          cells = {}
          formats = {}

          base_timestamp = base_date.to_datetime.to_time.to_i

          xml = sheet_xml.b
          pos = 0
          ycoord = 0
          row_start_pattern = "<row".b
          row_end_pattern = "</row>".b

          while (row_start = xml.index(row_start_pattern, pos))
            row_tag_end = xml.index(">".b, row_start + 4)
            break unless row_tag_end

            ycoord += 1
            row_tag = xml.byteslice(row_start, row_tag_end - row_start)
            r_row_idx = row_tag.index("r=\"".b)
            if r_row_idx
              val_start = r_row_idx + 3
              val_end = row_tag.index("\"".b, val_start)
              ycoord = row_tag.byteslice(val_start, val_end - val_start).to_i if val_end
            end

            if xml.getbyte(row_tag_end - 1) == 47 # self-closing <row ... />
              pos = row_tag_end + 1
              next
            end

            row_end = xml.index(row_end_pattern, row_tag_end + 1)
            break unless row_end

            c_pos = row_tag_end + 1
            xcoord = 0

            while (c_start = xml.index("<c".b, c_pos))
              break if c_start >= row_end

              c_tag_end = xml.index(">".b, c_start + 2)
              break unless c_tag_end && c_tag_end < row_end

              xcoord += 1
              c_tag = xml.byteslice(c_start, c_tag_end - c_start)

              r_idx = c_tag.index("r=\"".b)
              r_attr = if r_idx
                         v_st = r_idx + 3
                         v_en = c_tag.index("\"".b, v_st)
                         c_tag.byteslice(v_st, v_en - v_st).force_encoding(Encoding::UTF_8) if v_en
                       end

              coord = if r_attr
                        c_coord = Utils.extract_coordinate(r_attr)
                        xcoord = c_coord.column
                        c_coord
                      else
                        Coordinate.new(ycoord, xcoord)
                      end

              s_idx = c_tag.index("s=\"".b)
              style_id = if s_idx
                           v_st = s_idx + 3
                           v_en = c_tag.index("\"".b, v_st)
                           c_tag.byteslice(v_st, v_en - v_st).to_i if v_en
                         else
                           0
                         end

              t_idx = c_tag.index("t=\"".b)
              t_attr = if t_idx
                         v_st = t_idx + 3
                         v_en = c_tag.index("\"".b, v_st)
                         c_tag.byteslice(v_st, v_en - v_st).force_encoding(Encoding::UTF_8) if v_en
                       end

              link_url = hyperlinks_map[coord]

              xf = @cell_xfs[style_id] || {}
              num_fmt_id = xf[:num_fmt_id] || 0
              format_code = num_fmts[num_fmt_id] || Format::STANDARD_FORMATS[num_fmt_id] || "General"
              formats[coord] = format_code

              if xml.getbyte(c_tag_end - 1) == 47 # self-closing <c ... />
                cells[coord] = Cell::Empty.new(coord)
                c_pos = c_tag_end + 1
                next
              end

              c_end = xml.index("</c>".b, c_tag_end + 1)
              break unless c_end && c_end <= row_end

              inner = xml.byteslice(c_tag_end + 1, c_end - c_tag_end - 1)

              # Formula
              formula_str = nil
              f_idx = inner.index("<f".b)
              if f_idx
                f_tag_end = inner.index(">".b, f_idx + 2)
                if f_tag_end
                  if inner.getbyte(f_tag_end - 1) == 47
                    formula_str = ""
                  else
                    f_end = inner.index("</f>".b, f_tag_end + 1)
                    formula_str = unescape_xml(inner.byteslice(f_tag_end + 1, f_end - f_tag_end - 1).force_encoding(Encoding::UTF_8)) if f_end
                  end
                end
              end

              # Value <v>
              v_str = nil
              v_idx = inner.index("<v>".b)
              if v_idx
                v_end = inner.index("</v>".b, v_idx + 3)
                v_str = inner.byteslice(v_idx + 3, v_end - v_idx - 3).force_encoding(Encoding::UTF_8) if v_end
              end

              # InlineStr <is>
              is_str = nil
              is_idx = inner.index("<is>".b)
              if is_idx
                is_end = inner.index("</is>".b, is_idx + 4)
                is_str = inner.byteslice(is_idx + 4, is_end - is_idx - 4).force_encoding(Encoding::UTF_8) if is_end
              end

              cell_obj = if is_str
                           t_text = is_str.scan(%r{<t(?:\s+[^>]*)?>(.*?)</t>}m).map { |tm| unescape_xml(tm[0]) }.join
                           t_text = t_text.gsub("_x000D_", "\n").gsub("\r\n", "\n").gsub("\r", "\n")
                           if t_text.empty?
                             Cell::Empty.new(coord)
                           else
                             Cell::String.new(t_text, formula_str, style_id, link_url, coord)
                           end
                         elsif t_attr == "s" && v_str
                           sst_idx = v_str.to_i
                           raw_val = shared_strings[sst_idx] || ""
                           html_val = html_strings[sst_idx] || raw_val
                           effective_val = @options[:disable_html_wrapper] ? raw_val : html_val
                           Cell::String.new(effective_val, formula_str, style_id, link_url, coord)
                         elsif t_attr == "b" && v_str
                           Cell::Boolean.new(v_str, formula_str, style_id, link_url, coord)
                         elsif t_attr == "str" && v_str
                           v_clean = unescape_xml(v_str).gsub("_x000D_", "\n").gsub("\r\n", "\n").gsub("\r", "\n")
                           Cell::String.new(v_clean, formula_str, style_id, link_url, coord)
                         elsif v_str && !v_str.empty?
                           cell_val_type = Format.to_type(format_code)
                           numeric_raw = v_str.strip
                           excelx_type_arg = [:numeric_or_formula, format_code]

                           case cell_val_type
                           when :date
                             Cell::Date.new(numeric_raw, formula_str, excelx_type_arg, style_id, link_url, base_date, coord)
                           when :time, :datetime
                             num_f = numeric_raw.to_f
                             if num_f < 1.0
                               Cell::Time.new(numeric_raw, formula_str, excelx_type_arg, style_id, link_url, base_date, coord)
                             elsif (num_f - num_f.floor).abs > 0.000001
                               Cell::DateTime.new(numeric_raw, formula_str, excelx_type_arg, style_id, link_url, base_timestamp, coord)
                             else
                               Cell::Date.new(numeric_raw, formula_str, excelx_type_arg, style_id, link_url, base_date, coord)
                             end
                           else
                             Cell::Number.new(numeric_raw, formula_str, excelx_type_arg, style_id, link_url, coord)
                           end
                         elsif formula_str
                           Cell::Number.new("", formula_str, [:numeric_or_formula, format_code], style_id, link_url, coord)
                         else
                           Cell::Empty.new(coord)
                         end

              cells[coord] = cell_obj
              c_pos = c_end + 4
            end

            pos = row_end + 6
          end

          # Expand merged ranges if requested
          if @options[:expand_merged_ranges]
            merged_ranges.each do |range_str|
              next unless range_str.include?(":")

              tl_str, br_str = range_str.split(":", 2)
              tl = Utils.extract_coordinate(tl_str)
              br = Utils.extract_coordinate(br_str)
              master_cell = cells[tl]
              next unless master_cell

              (tl.row..br.row).each do |r|
                (tl.column..br.column).each do |c|
                  next if r == tl.row && c == tl.column

                  coord = Coordinate.new(r, c)
                  cells[coord] = master_cell
                end
              end
            end
          end

          Sheet.new(
            s_name,
            cells: cells,
            hyperlinks: hyperlinks_map,
            comments: comments_map,
            images: sheet_images,
            dimensions: dim_ref,
            formats: formats,
            styles: @parsed_styles,
            state: state
          )
        end

        #: (String? xml) -> Hash[Excelx::Coordinate, String]
        def parse_comments_xml(xml)
          return {} unless xml && !xml.empty?

          comments = {}
          xml.scan(%r{<comment\s+[^>]*ref="([^"]+)"[^>]*>(.*?)</comment>}m) do |ref, inner|
            coord = Utils.extract_coordinate(ref)
            text_parts = []
            inner.scan(%r{<t(?:\s+[^>]*)?>(.*?)</t>}m) do |t_match|
              text_parts << unescape_xml(t_match[0])
            end
            comments[coord] = text_parts.join
          end
          comments
        end

        #: (untyped zip, String path) -> String?
        def read_xml(zip, path)
          raw = zip.read_entry(path)
          return nil unless raw

          raw_b = raw.b
          xml = if raw_b.start_with?("\xFF\xFE".b)
                  raw_b.byteslice(2..-1).force_encoding(Encoding::UTF_16LE).encode(Encoding::UTF_8, invalid: :replace, undef: :replace)
                elsif raw_b.start_with?("\xFE\xFF".b)
                  raw_b.byteslice(2..-1).force_encoding(Encoding::UTF_16BE).encode(Encoding::UTF_8, invalid: :replace, undef: :replace)
                elsif raw_b.start_with?("\xEF\xBB\xBF".b)
                  raw_b.byteslice(3..-1).force_encoding(Encoding::UTF_8)
                else
                  raw.dup.force_encoding(Encoding::UTF_8)
                end
          xml = strip_tag_namespaces(xml) if xml =~ %r{<(/)?[a-zA-Z0-9_-]+:}
          xml
        end

        #: (String? xml) -> String
        def strip_tag_namespaces(xml)
          return "" unless xml

          xml.gsub(%r{<(/)?([a-zA-Z0-9_-]+:)?([a-zA-Z0-9_-]+)}) do
            slash = ::Regexp.last_match(1)
            tag = ::Regexp.last_match(3)
            "<#{slash}#{tag}"
          end
        end

        #: (String base_dir, String target) -> String
        def resolve_zip_path(base_dir, target)
          return target.delete_prefix("/") if target.start_with?("/")

          norm_base = base_dir.start_with?("/") ? base_dir : "/#{base_dir}"
          File.expand_path(target, norm_base).delete_prefix("/")
        end

        #: (String? str) -> String?
        def unescape_xml(str)
          return nil unless str

          str.gsub("&amp;", "&")
             .gsub("&lt;", "<")
             .gsub("&gt;", ">")
             .gsub("&quot;", '"')
             .gsub("&apos;", "'")
        end
      end
    end
  end
end
