# frozen_string_literal: true

# rbs_inline: enabled

require "cgi"
require "date"
require_relative "helpers"

module Xlsxrb
  module Adapters
    module Xsv
      # Minimal streaming XML parser compatible with Xsv::SaxParser.
      class SaxParser
        CHUNK_SIZE = 65_536

        # @param io [String, IO]
        # @return [void]
        #: (untyped io) -> void
        def parse(io)
          responds_to_end_element = respond_to?(:end_element)
          responds_to_characters = respond_to?(:characters)

          if io.is_a?(String)
            buf = io.dup.force_encoding(Encoding::UTF_8)
            eof_reached = true
          else
            buf = String.new(capacity: CHUNK_SIZE * 2, encoding: Encoding::UTF_8)
            eof_reached = false
          end

          pos = 0
          gt = 0
          space = -1
          equals = -1
          colon = -1

          loop do
            lt = buf.byteindex("<", gt)
            gt = lt && buf.byteindex(">", lt)

            unless gt
              if eof_reached
                raise Xsv::Error, "Malformed XML document, looking for end of tag beyond EOF" if lt

                break
              end

              buf = buf.byteslice(pos, buf.bytesize - pos) if pos.positive?
              pos = 0
              gt = 0
              space = equals = colon = -1

              if (chunk = read_chunk(io))
                buf << chunk
              else
                eof_reached = true
              end

              next
            end

            if responds_to_characters && lt > pos
              chars = buf.byteslice(pos, lt - pos)
              characters(chars.include?("&") ? CGI.unescapeHTML(chars) : chars)
            end

            pos = gt + 1

            space = buf.byteindex(" ", lt) || buf.bytesize if space < lt
            colon = buf.byteindex(":", lt) || buf.bytesize if colon < lt

            if buf.getbyte(lt + 1) == 47
              name_start = colon < gt ? colon + 1 : lt + 2
              end_element(buf.byteslice(name_start, gt - name_start)) if responds_to_end_element
              next
            end

            name_end = [space, gt].min
            name_start = colon < name_end ? colon + 1 : lt + 1
            tag_name = buf.byteslice(name_start, name_end - name_start)

            if space > gt
              start_element(tag_name, nil)
              next
            end

            attributes = {}
            attr_start = space
            loop do
              equals = buf.byteindex('="', attr_start) || buf.bytesize if equals < attr_start
              break if equals > gt

              quote = begin
                buf.byteindex('"', equals + 2)
              rescue IndexError
                buf.b.byteindex('"', equals + 2)
              end
              break if quote.nil? || quote > gt

              colon = buf.byteindex(":", attr_start) || buf.bytesize if colon < attr_start

              if colon < equals
                key = buf.byteslice(colon + 1, equals - colon - 1)
              else
                key = buf.byteslice(attr_start + 1, equals - attr_start - 1)
                key.lstrip!
              end

              attributes[key.to_sym] = buf.byteslice(equals + 2, quote - equals - 2)
              attr_start = quote
            end

            start_element(tag_name, attributes)
          end
        end

        private

        #: (untyped io) -> String?
        def read_chunk(io)
          chunk = io.sysread(CHUNK_SIZE)
          chunk&.dup&.force_encoding(Encoding::UTF_8)
        rescue EOFError
          nil
        end
      end

      # SharedStringsParser compatible with Xsv::SharedStringsParser.
      class SharedStringsParser < SaxParser
        # @param io [String, IO]
        # @return [Array<String>]
        #: (untyped io) -> Array[String]
        def self.parse(io)
          strings = []
          new { |s| strings << -s }.parse(io)
          strings
        end

        #: () ?{ (String) -> void } -> void
        def initialize(&block)
          super()
          @block = block
          @state = nil
          @skip = false
          @current_string = ""
        end

        #: (String name, Hash[Symbol, String]? _attrs) -> void
        def start_element(name, _attrs)
          case name
          when "si"
            @current_string = ""
            @skip = false
          when "rPh"
            @skip = true
          when "t"
            @state = name
          end
        end

        #: (String value) -> void
        def characters(value)
          @current_string += value if @state == "t" && !@skip
        end

        #: (String name) -> void
        def end_element(name)
          case name
          when "si"
            @block&.call(@current_string)
          when "rPh"
            @skip = false
          when "t"
            @state = nil
          end
        end
      end

      # SheetBoundsHandler compatible with Xsv::SheetBoundsHandler.
      class SheetBoundsHandler < SaxParser
        include Helpers

        # @param sheet [IO, String]
        # @param workbook [Workbook]
        # @return [Array<Integer>]
        #: (untyped sheet, Workbook workbook) -> [Integer, Integer]
        def self.get_bounds(sheet, workbook)
          rows = 0
          cols = 0

          handler = new(workbook.trim_empty_rows) do |row, col|
            rows = row
            cols = col.zero? ? 0 : col + 1

            return [rows, cols]
          end

          sheet.rewind if sheet.respond_to?(:rewind)
          handler.parse(sheet)
          [rows, cols]
        end

        #: (bool trim_empty_rows) ?{ (Integer, Integer) -> void } -> void
        def initialize(trim_empty_rows, &block)
          super()
          @block = block
          @state = nil
          @cell = nil
          @row = nil
          @max_row = 0
          @max_column = 0
          @trim_empty_rows = trim_empty_rows
        end

        #: (String name, Hash[Symbol, String]? attrs) -> void
        def start_element(name, attrs)
          case name
          when "c"
            @state = name
            @cell = attrs ? attrs[:r] : nil
          when "v", "is"
            if @cell && @row
              col = column_index(@cell)
              @max_column = col if col > @max_column
              @max_row = @row if @row > @max_row
            end
          when "row"
            @state = name
            @row = attrs ? attrs[:r].to_i : 0
          when "dimension"
            @state = name
            ref = attrs ? attrs[:ref] : nil
            if ref
              _first_cell, last_cell = ref.split(":")
              if last_cell
                @max_column = column_index(last_cell)
                unless @trim_empty_rows
                  @max_row = last_cell[/\d+$/].to_i
                  @block&.call(@max_row, @max_column)
                end
              end
            end
          end
        end

        #: (String name) -> void
        def end_element(name)
          @block&.call(@max_row, @max_column) if name == "sheetData"
        end
      end

      # SheetRowsHandler compatible with Xsv::SheetRowsHandler.
      class SheetRowsHandler < SaxParser
        include Helpers

        SLASH = 47
        LOWER_C = 99
        LOWER_I = 105
        LOWER_O = 111
        LOWER_R = 114
        LOWER_S = 115
        LOWER_T = 116
        LOWER_V = 118
        LOWER_W = 119
        SPACE = 32
        QUOTE = 34
        COLON = 58
        LT = 60
        EQUALS = 61
        GT = 62
        ZERO = 48
        NINE = 57

        SINGLE_BYTE_CELL_TYPES = { "s".ord => :s, "n".ord => :n, "b".ord => :b, "e".ord => :e, "d".ord => :d }.freeze
        CELL_TYPES = { "str" => :str, "inlineStr" => :inlineStr }.freeze

        #: (Symbol mode, Array[untyped]? headers, untyped empty_row, Workbook workbook, Integer row_skip, Integer last_row) ?{ (untyped) -> void } -> void
        def initialize(mode, headers, empty_row, workbook, row_skip, last_row, &block)
          super()
          @mode = mode
          @headers = headers
          @empty_row = empty_row
          @workbook = workbook
          @row_skip = row_skip
          @last_row = last_row - @row_skip
          @block = block

          @row_index = 0
          @col_index = 0
          @current_row = {}
          @current_row_number = 0
          @number_formats = {}
        end

        # @param io [String, IO]
        # @return [void]
        #: (untyped io) -> void
        def parse(io)
          if io.is_a?(String)
            buf = io.dup.force_encoding(Encoding::UTF_8)
            eof_reached = true
          else
            buf = String.new(capacity: CHUNK_SIZE * 2, encoding: Encoding::UTF_8)
            eof_reached = false
          end

          pos = 0
          waiting_for_row_end = false

          loop do
            unless waiting_for_row_end && !eof_reached && !buf.byteindex("row>", pos.positive? ? pos - 1 : 0)
              if (next_pos = parse_next(buf, pos))
                pos = next_pos
                waiting_for_row_end = false
                next
              end

              waiting_for_row_end = true
            end

            if eof_reached
              last_lt = buf.byterindex("<")
              raise Xsv::Error, "Malformed XML document, looking for end of tag beyond EOF" if last_lt && !buf.byteindex(">", last_lt)

              break
            end

            buf = buf.byteslice(pos, buf.bytesize - pos) if pos.positive?
            pos = 0

            if (chunk = read_chunk(io))
              buf << chunk
            else
              eof_reached = true
            end
          end
        end

        private

        def parse_next(buf, pos)
          return unless (lt = buf.byteindex("<", pos.positive? ? pos - 1 : 0))
          return unless (gt = buf.byteindex(">", lt))
          return gt + 1 if buf.getbyte(lt + 1) == SLASH

          name_end = lt + 1
          name_end += 1 while (byte = buf.getbyte(name_end)) != SPACE && byte != GT
          name_start = name_end
          name_start -= 1 while name_start > lt + 1 && buf.getbyte(name_start - 1) != COLON

          unless name_end - name_start == 3 && buf.getbyte(name_start) == LOWER_R &&
                 buf.getbyte(name_start + 1) == LOWER_O && buf.getbyte(name_start + 2) == LOWER_W
            return gt + 1
          end

          number = row_number(buf, name_end, gt) || (@current_row_number + 1)

          if buf.getbyte(gt - 1) == SLASH
            @current_row_number = number
            return gt + 1
          end

          parse_row(buf, gt, number)
        end

        def row_number(buf, offset, gt_pos)
          while offset < gt_pos
            offset += 1 while (byte = buf.getbyte(offset)) <= SPACE
            return if [GT, SLASH].include?(byte)

            key_start = offset
            while (byte = buf.getbyte(offset)) != EQUALS
              return if byte == GT

              key_start = offset + 1 if byte == COLON
              offset += 1
            end
            return unless buf.getbyte(offset + 1) == QUOTE

            return integer_at(buf, offset + 2) if offset - key_start == 1 && buf.getbyte(key_start) == LOWER_R

            offset += 2
            offset += 1 while (byte = buf.getbyte(offset)) != QUOTE && byte
            offset += 1
          end
        end

        def parse_row(buf, row_gt, number)
          row = @mode == :array ? [] : @empty_row.dup
          col_index = @col_index

          store_characters = false
          value = nil
          cell_column = nil
          cell_type = nil
          cell_style = nil

          space = equals = colon = -1
          gt = row_gt

          loop do
            pos = gt + 1
            if buf.getbyte(pos) == LT
              lt = pos
            else
              return unless (lt = buf.byteindex("<", gt))

              if store_characters
                chars = buf.byteslice(pos, lt - pos)
                chars = CGI.unescapeHTML(chars) if chars.include?("&")
                value = value ? value + chars : chars
              end
            end

            first = buf.getbyte(lt + 1)

            if first == SLASH
              if buf.getbyte(lt + 3) == GT
                gt = lt + 3
                name_start = lt + 2
              else
                return unless (gt = buf.byteindex(">", lt))

                colon = buf.byteindex(":", lt) || buf.bytesize if colon < lt
                name_start = colon < gt ? colon + 1 : lt + 2
              end

              name_length = gt - name_start
              first = buf.getbyte(name_start)

              if name_length == 1
                if first == LOWER_C
                  column = cell_column || col_index
                  if @mode == :array
                    row[column] = format_cell(value, cell_type, cell_style)
                  else
                    header = @headers ? @headers[column] : nil
                    row[header] = format_cell(value, cell_type, cell_style) unless header.nil?
                  end
                  col_index += 1
                elsif [LOWER_V, LOWER_T].include?(first)
                  store_characters = false
                end
              elsif name_length == 2
                store_characters = false if first == LOWER_I && buf.getbyte(name_start + 1) == LOWER_S
              elsif name_length == 3 && first == LOWER_R && buf.getbyte(name_start + 1) == LOWER_O && buf.getbyte(name_start + 2) == LOWER_W
                @current_row_number = number
                @current_row = row
                @col_index = col_index
                end_row
                return gt + 1
              end

              next
            end

            second = buf.getbyte(lt + 2)
            if second == GT
              gt = lt + 2
              name_length = 1
              attr_start = gt
            elsif second == SPACE && first == LOWER_C
              return unless (gt = buf.byteindex(">", lt))

              name_length = 1
              attr_start = lt + 2
            else
              return unless (gt = buf.byteindex(">", lt))

              space = buf.byteindex(" ", lt) || buf.bytesize if space < lt
              colon = buf.byteindex(":", lt) || buf.bytesize if colon < lt

              attr_start = [space, gt].min
              name_start = colon < attr_start ? colon + 1 : lt + 1
              name_length = attr_start - name_start
              first = buf.getbyte(name_start)
              second = buf.getbyte(name_start + 1)
            end

            if name_length == 1
              if first == LOWER_C
                value = nil
                cell_column = cell_type = cell_style = nil

                while attr_start < gt
                  if buf.getbyte(attr_start + 2) == EQUALS && buf.getbyte(attr_start + 3) == QUOTE
                    key = buf.getbyte(attr_start + 1)
                    value_start = attr_start + 4
                  else
                    equals = buf.byteindex('="', attr_start) || buf.bytesize if equals < attr_start
                    break if equals > gt

                    before = buf.getbyte(equals - 2)
                    key = before <= SPACE || before == COLON ? buf.getbyte(equals - 1) : nil
                    value_start = equals + 2
                  end

                  quote = begin
                    buf.byteindex('"', value_start)
                  rescue IndexError
                    buf.b.byteindex('"', value_start)
                  end
                  break if quote.nil? || quote > gt

                  case key
                  when LOWER_R
                    column = 0
                    while (byte = buf.getbyte(value_start)) >= A_CODEPOINT
                      column = (column * 26) + byte - A_CODEPOINT + 1
                      value_start += 1
                    end
                    cell_column = column - 1
                  when LOWER_T
                    cell_type = SINGLE_BYTE_CELL_TYPES[buf.getbyte(value_start)] if quote - value_start == 1
                    if cell_type.nil?
                      name = buf.byteslice(value_start, quote - value_start)
                      cell_type = CELL_TYPES[name] || name
                    end
                  when LOWER_S
                    cell_style = integer_at(buf, value_start)
                  end

                  attr_start = quote + 1
                end
              elsif [LOWER_V, LOWER_T].include?(first)
                store_characters = true
              end
            elsif name_length == 2
              store_characters = true if first == LOWER_I && second == LOWER_S
            end
          end
        end

        def integer_at(buf, offset)
          number = 0
          while (byte = buf.getbyte(offset)) && byte >= ZERO && byte <= NINE
            number = (number * 10) + byte - ZERO
            offset += 1
          end
          number
        end

        def end_row
          return if @current_row_number <= @row_skip

          adjusted_row_number = @current_row_number - @row_skip
          @row_index += 1
          @col_index = 0

          return if adjusted_row_number == 1 && @mode == :hash

          while @row_index < adjusted_row_number
            @block&.call(@empty_row)
            @row_index += 1
          end

          return if @row_index > @last_row

          if @mode == :array && @current_row.length < @empty_row.length
            @block&.call(@current_row + @empty_row[@current_row.length..])
          else
            @block&.call(@current_row)
          end
        end

        def format_cell(value, type, style)
          return nil if value.nil? || value.empty?

          case type
          when :s
            @workbook.shared_strings[value.to_i]
          when :str, :inlineStr
            value.strip!
            -value
          when :e
            nil
          when nil, :n
            if style
              format_number(value, style)
            else
              parse_number(value)
            end
          when :b
            value == "1"
          when :d
            DateTime.parse(value)
          else
            raise Xsv::Error, "Encountered unknown column type #{type}"
          end
        end

        def format_number(value, style)
          number = parse_number(value)

          case @number_formats.fetch(style) { @number_formats[style] = number_format_kind(@workbook.get_num_fmt(style)) }
          when :date
            parse_date(number)
          when :time
            parse_time(number)
          when :datetime
            parse_datetime(number)
          else
            number
          end
        end
      end

      # RelationshipsHandler compatible with Xsv::RelationshipsHandler.
      class RelationshipsHandler < SaxParser
        #: (untyped io) -> Array[Hash[Symbol, String]]
        def self.get_relations(io)
          relations = []
          new { |rel| relations << rel }.parse(io)
          relations
        end

        #: () ?{ (Hash[Symbol, String]) -> void } -> void
        def initialize(&block)
          super()
          @block = block
        end

        #: (String name, Hash[Symbol, String]? attrs) -> void
        def start_element(name, attrs)
          @block&.call(attrs.slice(:Id, :Type, :Target)) if name == "Relationship" && attrs
        end
      end

      # SheetsIdsHandler compatible with Xsv::SheetsIdsHandler.
      class SheetsIdsHandler < SaxParser
        #: (untyped io) -> Array[Hash[Symbol, String]]
        def self.get_sheets_ids(io)
          sheets_ids = []
          new { |ids| sheets_ids << ids }.parse(io)
          sheets_ids.each { |s| s[:name]&.force_encoding(Encoding::UTF_8) }
        end

        #: () ?{ (Hash[Symbol, String]) -> void } -> void
        def initialize(&block)
          super()
          @block = block
        end

        #: (String name, Hash[Symbol, String]? attrs) -> void
        def start_element(name, attrs)
          @block&.call(attrs.slice(:name, :sheetId, :state, :id)) if name == "sheet" && attrs
        end
      end

      # StylesHandler compatible with Xsv::StylesHandler.
      class StylesHandler < SaxParser
        #: (untyped io) -> [Array[Hash[Symbol, Integer]], Hash[Integer, String]]
        def self.get_styles(io)
          xfs_res = []
          num_fmts_res = {}
          handler = new(Helpers::BUILT_IN_NUMBER_FORMATS.dup) do |xfs, num_fmts|
            xfs_res = xfs
            num_fmts_res = num_fmts
          end
          handler.parse(io)
          [xfs_res, num_fmts_res]
        end

        #: (Hash[Integer, String] num_fmts) ?{ (Array[Hash[Symbol, Integer]], Hash[Integer, String]) -> void } -> void
        def initialize(num_fmts, &block)
          super()
          @block = block
          @state = nil
          @xfs = []
          @num_fmts = num_fmts
        end

        #: (String name, Hash[Symbol, String]? attrs) -> void
        def start_element(name, attrs)
          case name
          when "cellXfs"
            @state = "cellXfs"
          when "xf"
            @xfs << attrs.transform_values(&:to_i) if @state == "cellXfs" && attrs
          when "numFmt"
            @num_fmts[attrs[:numFmtId].to_i] = attrs[:formatCode] if attrs
          end
        end

        #: (String name) -> void
        def end_element(name)
          case name
          when "styleSheet"
            @block&.call(@xfs, @num_fmts)
          when "cellXfs"
            @state = nil
          end
        end
      end
    end
  end
end
