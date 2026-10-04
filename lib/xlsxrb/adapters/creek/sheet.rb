# frozen_string_literal: true

# rbs_inline: enabled

require_relative "utils"
require_relative "styles"
require_relative "drawing"

module Xlsxrb
  module Adapters
    module Creek
      # Represents a single worksheet matching Creek::Sheet.
      class Sheet
        include Utils

        HEADERS_ROW_NUMBER = "1"
        SPREADSHEETML_URI = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"

        ROW_RE = %r{<(?:[A-Za-z0-9_]+:)?row\b([^>]*?)(/>|>(.*?)</(?:[A-Za-z0-9_]+:)?row>)}m
        CELL_RE = %r{<(?:[A-Za-z0-9_]+:)?c\b([^>]*?)(/>|>(.*?)</(?:[A-Za-z0-9_]+:)?c>)}m

        # @return [Boolean]
        attr_accessor :with_headers

        # @return [untyped]
        attr_reader :book

        # @return [String]
        attr_reader :name

        # @return [untyped]
        attr_reader :sheetid

        # @return [String, nil]
        attr_reader :state

        # @return [String, nil]
        attr_reader :visible

        # @return [String, nil]
        attr_reader :rid

        # @return [Integer, nil]
        attr_reader :index

        # @return [Hash<String, untyped>, nil]
        attr_reader :headers

        # @param book [untyped]
        # @param name [String]
        # @param sheetid [untyped]
        # @param state [String, nil]
        # @param visible [String, nil]
        # @param rid [String, nil]
        # @param sheetfile [String]
        #: (untyped book, String name, untyped sheetid, String? state, String? visible, String? rid, String sheetfile) -> void
        def initialize(book, name, sheetid, state, visible, rid, sheetfile)
          @book = book
          @name = name
          @sheetid = sheetid
          @visible = visible
          @rid = rid
          @state = state
          @sheetfile = sheetfile
          @images_present = false
          @with_headers = false
        end

        # Preloads images info from related drawing.xml and drawing rels.
        #
        # @return [self]
        #: () -> self
        def with_images
          @drawingfile = extract_drawing_filepath
          if @drawingfile
            @drawing = Drawing.new(@book, @drawingfile.sub("..", "xl"))
            @images_present = @drawing.has_images?
          end
          self
        end

        # Extracts images for a cell to a temporary folder.
        #
        # @param cell [String]
        # @return [Array<Pathname>, nil]
        #: (String cell) -> Array[Pathname]?
        def images_at(cell)
          @drawing.images_at(cell) if @images_present && @drawing
        end

        # Provides an Enumerator that returns a hash representing each row with column letters.
        #
        # @return [Enumerator]
        #: () -> Enumerator[Hash[String, untyped], void]
        def simple_rows
          rows_generator(false, true)
        end

        # Provides an Enumerator that returns a hash representing each row with cell coordinates.
        #
        # @return [Enumerator]
        #: () -> Enumerator[Hash[String, untyped], void]
        def rows
          rows_generator(false, false)
        end

        # Provides an Enumerator that returns a hash per row with metadata and a 'cells' sub-hash.
        #
        # @return [Enumerator]
        #: () -> Enumerator[Hash[String, untyped], void]
        def rows_with_meta_data
          rows_generator(true, false)
        end

        # Provides an Enumerator that returns a hash per row with metadata and simple column letters.
        #
        # @return [Enumerator]
        #: () -> Enumerator[Hash[String, untyped], void]
        def simple_rows_with_meta_data
          rows_generator(true, true)
        end

        # Converts this Creek Sheet into an immutable Xlsxrb::Elements::Worksheet.
        #
        # @return [Xlsxrb::Elements::Worksheet]
        #: () -> Xlsxrb::Elements::Worksheet
        def to_xlsxrb
          row_elements = []
          rows.each_with_index do |row_hash, r_idx|
            cells = []
            target_row_idx = nil
            row_hash.each do |coord, val|
              next if val.nil?

              row_col = Xlsxrb::Utils.ref_to_row_col(coord)
              actual_r_idx = row_col ? row_col[0] : r_idx
              c_idx = row_col ? row_col[1] : 0
              target_row_idx ||= actual_r_idx
              cells << Elements::Cell.new(row_index: actual_r_idx, column_index: c_idx, value: val)
            end
            row_elements << Elements::Row.new(index: target_row_idx || r_idx, cells: cells)
          end
          st = case @state
               when "hidden", :hidden then :hidden
               when "veryHidden", :veryHidden then :very_hidden
               else :visible
               end
          extracted_images = []
          if @images_present && @drawing
            @drawing.images_pathnames.each do |coord, pathnames|
              pathnames.each do |pn|
                filename = pn.basename.to_s
                data = pn.exist? ? File.binread(pn.to_s) : nil
                extracted_images << Elements::Image.new(
                  filename: filename,
                  from_row: coord[0],
                  from_col: coord[1],
                  data: data
                )
              end
            end
          end
          Elements::Worksheet.new(name: @name, rows: row_elements, state: st, images: extracted_images)
        end

        private

        # Generates an Enumerator for iterating over sheet rows.
        #: (bool include_meta_data, bool use_simple_rows_format) -> Enumerator[Hash[String, untyped], void]
        def rows_generator(include_meta_data = false, use_simple_rows_format = false)
          path = if @sheetfile.start_with?("/xl/") || @sheetfile.start_with?("xl/")
                   @sheetfile
                 else
                   "xl/#{@sheetfile}"
                 end

          Enumerator.new do |y|
            @headers = nil
            next unless @book.files.file.exist?(path)

            xml_str = @book.files.file.open(path).read

            xml_str.scan(ROW_RE) do |attrs_str, close_type, body|
              row = parse_attributes(attrs_str)
              row["cells"] = {}
              cells = {}
              last_cell = nil

              if close_type == "/>" || body.nil? || body.empty?
                y << (include_meta_data ? row : cells)
              else
                body.scan(CELL_RE) do |c_attrs_str, c_close, c_body|
                  c_attrs = parse_attributes(c_attrs_str)
                  cell_ref = c_attrs["r"]
                  last_cell = cell_ref if cell_ref

                  next if c_close == "/>" || c_body.nil? || c_body.empty?

                  cell_type = c_attrs["t"]
                  cell_style_idx = c_attrs["s"]

                  v_match = c_body[%r{<(?:[A-Za-z0-9_]+:)?v>([^<]*)</(?:[A-Za-z0-9_]+:)?v>}, 1]
                  t_match = c_body[%r{<(?:[A-Za-z0-9_]+:)?t\b[^>]*>([^<]*)</(?:[A-Za-z0-9_]+:)?t>}, 1]
                  val_raw = v_match || t_match

                  if val_raw && cell_ref
                    val_decoded = decode_entities(val_raw)
                    cells[cell_ref] = convert(val_decoded, cell_type, cell_style_idx)
                  end
                end

                processed_cells = fill_in_empty_cells(cells, row["r"], last_cell, use_simple_rows_format)
                @headers = processed_cells if with_headers && row["r"] == HEADERS_ROW_NUMBER

                if @images_present
                  processed_cells.each do |cell_name, cell_value|
                    next unless cell_value.nil?

                    processed_cells[cell_name] = images_at(cell_name)
                  end
                end

                row["cells"] = processed_cells
                y << (include_meta_data ? row : processed_cells)
              end
            end
          end
        end

        #: (String? value, String? type, String? style_idx) -> untyped
        def convert(value, type, style_idx)
          style = @book.style_types[style_idx.to_i]
          Styles::Converter.call(value, type, style, converter_options)
        end

        #: () -> Hash[Symbol, untyped]
        def converter_options
          @converter_options ||= {
            shared_strings: @book.shared_strings.dictionary,
            base_date: @book.base_date
          }
        end

        #: (Hash[String, untyped] cells, String? row_number, String? last_col, bool use_simple_rows_format) -> Hash[String, untyped]
        def fill_in_empty_cells(cells, row_number, last_col, use_simple_rows_format)
          new_cells = {}
          return new_cells if cells.empty? || last_col.nil?

          col_letter = last_col.gsub(row_number.to_s, "")
          return new_cells if col_letter.empty?

          ("A"..col_letter).to_a.each do |column|
            id = cell_id(column, use_simple_rows_format, row_number)
            new_cells[id] = cells["#{column}#{row_number}"]
          end

          new_cells
        end

        #: (String column, bool use_simple_rows_format, String? row_number) -> untyped
        def cell_id(column, use_simple_rows_format, row_number)
          return "#{column}#{row_number}" unless use_simple_rows_format

          with_headers && headers ? headers[column] : column
        end

        #: () -> String?
        def extract_drawing_filepath
          sheet_filepath = "xl/#{@sheetfile}"
          return nil unless file_exist?(sheet_filepath)

          sheet_xml = @book.files.file.open(sheet_filepath).read
          drawing_rid = sheet_xml[/<(?:[A-Za-z0-9_]+:)?drawing\b[^>]*?(?:r:id|id)="([^"]*)"/, 1]
          return nil unless drawing_rid

          sheet_rels_filepath = expand_to_rels_path(sheet_filepath)
          return nil unless file_exist?(sheet_rels_filepath)

          rels_xml = @book.files.file.open(sheet_rels_filepath).read
          rels_xml[/<Relationship\b[^>]*?Id="#{Regexp.escape(drawing_rid)}"[^>]*?Target="([^"]*)"/, 1]
        end

        #: (String attrs_str) -> Hash[String, String]
        def parse_attributes(attrs_str)
          attrs = {}
          attrs_str.scan(/([A-Za-z0-9_:]+)=(?:"([^"]*)"|'([^']*)')/) do |k, v1, v2|
            key = k.include?(":") ? k.split(":").last : k
            attrs[key] = (v1 || v2).to_s
          end
          attrs
        end

        #: (String str) -> String
        def decode_entities(str)
          return str unless str.include?("&")

          str.gsub("&amp;", "&").gsub("&lt;", "<").gsub("&gt;", ">").gsub("&quot;", '"').gsub("&apos;", "'")
        end
      end
    end
  end
end
