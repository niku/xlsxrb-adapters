# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require "stringio"
require "xlsxrb"
require_relative "hyperlink"

module Xlsxrb
  module Adapters
    module SimpleXlsxReader
      # Main loading and parsing coordinator matching SimpleXlsxReader::Loader.
      class Loader < Struct.new(:string_or_io)
        # Shared strings table array.
        # @return [Array<String>, nil]
        #: Array[String]?
        attr_accessor :shared_strings

        # Worksheet parsers array.
        # @return [Array<SheetParser>, nil]
        #: Array[SheetParser]?
        attr_accessor :sheet_parsers

        # Table of contents mapping sheet name to 0-based sheet index.
        # @return [Hash<String, Integer>, nil]
        #: Hash[String, Integer]?
        attr_accessor :sheet_toc

        # Style types array matching cellXfs indices.
        # @return [Array<Symbol>, nil]
        #: Array[Symbol]?
        attr_accessor :style_types

        # Base date for date serial calculation.
        # @return [Date, nil]
        #: Date?
        attr_accessor :base_date

        # Sheet relationship IDs mapping sheet name to r:id.
        # @return [Hash<String, String>, nil]
        #: Hash[String, String]?
        attr_accessor :sheet_r_ids

        # Initializes and returns worksheet representations for all sheets in the document.
        #
        # @return [Array<Document::Sheet>]
        #: () -> Array[Document::Sheet]
        def init_sheets
          zip_reader = ZipReader.new(string_or_io: string_or_io, loader: self)
          zip_reader.read

          (sheet_toc || {}).each_with_index.map do |(sheet_name, _sheet_number), i|
            parsers = sheet_parsers || []
            parser = parsers[i] || parsers.find { |p| p&.name == sheet_name }
            Document::Sheet.new(
              name: sheet_name,
              sheet_parser: parser || SheetParser.new(file_io: StringIO.new(""), loader: self)
            )
          end
        end

        # Entry proxy wrapping raw entry data with an IO-like interface.
        class EntryProxy
          # @param data [String]
          #: (String data) -> void
          def initialize(data)
            @data = data
          end

          # Yields or returns a StringIO stream for this entry.
          #
          # @yield [io]
          # @return [StringIO, untyped]
          #: () ?{ (StringIO) -> untyped } -> untyped
          def get_input_stream(&block)
            io = StringIO.new(read.b)
            if block
              begin
                block.call(io)
              ensure
                io.close
              end
            else
              io
            end
          end

          # Returns decoded entry string (auto-decoding UTF-16LE/BE if BOM is detected).
          #
          # @return [String]
          #: () -> String
          def read
            data = @data
            if data.b.start_with?("\xFF\xFE".b)
              data.force_encoding("UTF-16LE").encode("UTF-8").delete_prefix("\uFEFF")
            elsif data.b.start_with?("\xFE\xFF".b)
              data.force_encoding("UTF-16BE").encode("UTF-8").delete_prefix("\uFEFF")
            else
              data
            end
          end
        end

        # Internal archive reader matching SimpleXlsxReader::Loader::ZipReader.
        ZipReader = Struct.new(:string_or_io, :loader) do
          # @return [Xlsxrb::Ooxml::ZipReader]
          #: Xlsxrb::Ooxml::ZipReader
          attr_reader :zip

          # @param string_or_io [untyped]
          # @param loader [Loader, nil]
          #: (?untyped string_or_io, ?Loader? loader, **untyped) -> void
          def initialize(string_or_io = nil, loader = nil, **kwargs)
            if kwargs.key?(:string_or_io) || kwargs.key?(:loader)
              super(kwargs[:string_or_io] || string_or_io, kwargs[:loader] || loader)
            else
              super(string_or_io, loader)
            end

            src = if self.string_or_io.is_a?(String) && (self.string_or_io.start_with?("PK\x03\x04") || self.string_or_io.include?("\x00"))
                    StringIO.new(self.string_or_io.b)
                  elsif self.string_or_io.is_a?(File)
                    self.string_or_io.path
                  else
                    self.string_or_io
                  end
            @zip = Xlsxrb::Ooxml::ZipReader.open(src)
          end

          # Reads workbook structure, styles, shared strings, and sheet references.
          #
          # @return [void]
          #: () -> void
          def read
            entry_at("xl/workbook.xml") do |file_io|
              loader.sheet_toc, loader.base_date, loader.sheet_r_ids = *WorkbookParser.parse(file_io)
            end

            entry_at("xl/styles.xml") do |file_io|
              loader.style_types = StyleTypesParser.parse(file_io)
            end

            if (ss_entry = entry_at("xl/sharedStrings.xml"))
              ss_entry.get_input_stream do |file|
                loader.shared_strings = SharedStringsParser.parse(file).map do |s|
                  s ? s.gsub("\r\n", "\n").tr("\r", "\n") : s
                end
              end
            else
              loader.shared_strings = []
            end

            loader.sheet_parsers = []

            rels = {}
            if (wb_rels_entry = entry_at("xl/_rels/workbook.xml.rels"))
              wb_rels_entry.get_input_stream do |file|
                rels = Xlsxrb::Ooxml::RelationshipsParser.parse(file.read)
              end
            end

            (loader.sheet_toc || {}).each_with_index do |(sheet_name, _sheet_idx), idx|
              parser = nil
              r_id = loader.sheet_r_ids ? loader.sheet_r_ids[sheet_name] : nil
              target = rels[r_id] if r_id
              if target
                target_path = target.start_with?("/") ? target.delete_prefix("/") : "xl/#{target}"
                if (entry = entry_at(target_path))
                  parser = create_sheet_parser(entry, target_path)
                end
              end

              unless parser
                fallback_names = [
                  "xl/worksheets/sheet#{idx + 1}.xml",
                  "xl/worksheets/sheet#{idx}.xml",
                  "xl/worksheets/sheet.xml"
                ]
                fallback_names.each do |fn|
                  if (entry = entry_at(fn))
                    parser = create_sheet_parser(entry, fn)
                    break
                  end
                end
              end

              parser ||= SheetParser.new(file_io: StringIO.new(""), loader: loader)
              parser.name = sheet_name
              loader.sheet_parsers << parser
            end
          end

          # Creates a SheetParser instance for an archive entry.
          #
          # @param entry [EntryProxy]
          # @param sheet_file_name [String]
          # @return [SheetParser]
          #: (EntryProxy entry, String sheet_file_name) -> SheetParser
          def create_sheet_parser(entry, sheet_file_name)
            parser = SheetParser.new(
              file_io: entry.get_input_stream,
              loader: loader
            )
            base_dir = File.dirname(sheet_file_name)
            base_name = File.basename(sheet_file_name)
            rel_file_name = "#{base_dir}/_rels/#{base_name}.rels"
            if (rel = entry_at(rel_file_name))
              parser.xrels_file = rel.get_input_stream
            end
            parser
          end

          # Finds an archive entry flexibly, supporting Windows backslashes and case differences.
          #
          # @param path [String]
          # @yield [io]
          # @return [EntryProxy, untyped]
          #: (String path) ?{ (StringIO) -> untyped } -> untyped
          def entry_at(path, &block)
            norm = path.delete_prefix("/")
            data = zip.read_entry(norm)
            return nil unless data

            proxy = EntryProxy.new(data)
            if block
              proxy.get_input_stream(&block)
            else
              proxy
            end
          end
        end

        # Core cell typecasting method matching SimpleXlsxReader::Loader.cast.
        #
        # @param value [String, nil] Raw string value from cell.
        # @param type [String, Symbol, nil] Cell type code or style symbol.
        # @param style [Symbol, String, nil] Resolved style symbol.
        # @param options [Hash] Context options (:shared_strings, :base_date, :url).
        # @return [Object, nil]
        # rubocop:disable Lint/DuplicateBranch
        #: (String? value, (String | Symbol | nil) type, (Symbol | String | nil) style, ?Hash[Symbol, untyped] options) -> untyped
        def self.cast(value, type, style, options = {})
          return nil if value.nil? || value.empty?

          type = style if type.nil? || (type == "n" && %i[date time date_time].include?(style))

          casted = case type
                   when "s"
                     options[:shared_strings] ? options[:shared_strings][value.to_i] : nil
                   when "n"
                     value.to_f
                   when "b"
                     value.to_i == 1
                   when "str", "inlineStr"
                     value
                   when nil, :string
                     retval = Integer(value, exception: false)
                     retval ||= Float(value, exception: false)
                     retval || value
                   when :unsupported
                     value
                   when :fixnum
                     value.to_i
                   when :float, :percentage
                     value.to_f
                   when :date, :time, :date_time
                     val = Float(value)
                     days_since_start = val.to_i
                     fraction_of_day = val - days_since_start

                     date = options.fetch(:base_date, DATE_SYSTEM_1900) + days_since_start
                     if fraction_of_day.positive?
                       seconds = (fraction_of_day * 86_400).round
                       Time.utc(date.year, date.month, date.day) + seconds
                     else
                       date
                     end
                   when :bignum
                     if defined?(BigDecimal)
                       BigDecimal(value)
                     else
                       value.to_f
                     end
                   else
                     value
                   end

          if options[:url]
            Hyperlink.new(options[:url], casted)
          else
            casted
          end
        end
        # rubocop:enable Lint/DuplicateBranch

        # Workbook XML parser matching SimpleXlsxReader::Loader::WorkbookParser.
        WorkbookParser = Struct.new(:file_io) do
          # Parses sheet names, IDs, base date, and relationship IDs from workbook.xml.
          #
          # @param file_io [IO, StringIO]
          # @return [Array(Hash<String, Integer>, Date, Hash<String, String>)]
          #: (untyped file_io) -> [Hash[String, Integer], Date, Hash[String, String]]
          def self.parse(file_io)
            parser = new(file_io).tap(&:parse)
            [parser.sheet_toc, parser.base_date, parser.sheet_r_ids]
          end

          # @return [void]
          #: () -> void
          def parse
            raw = file_io.read
            @raw_xml = if raw.b.start_with?("\xFF\xFE".b)
                         raw.force_encoding("UTF-16LE").encode("UTF-8").delete_prefix("\uFEFF")
                       elsif raw.b.start_with?("\xFE\xFF".b)
                         raw.force_encoding("UTF-16BE").encode("UTF-8").delete_prefix("\uFEFF")
                       else
                         raw
                       end
          end

          # Returns sheet table of contents mapping sheet name to 0-based sheet index.
          #
          # @return [Hash<String, Integer>]
          #: () -> Hash[String, Integer]
          def sheet_toc
            toc = {}
            @raw_xml.scan(%r{<(?:[a-zA-Z0-9_]+:)?sheet\b([^>]*?)(?:/>|>(?:.*?)</(?:[a-zA-Z0-9_]+:)?sheet>)}m) do |(attrs)|
              name = attrs[/name="([^"]+)"/, 1]
              sheet_id = attrs[/sheetId="(\d+)"/, 1]
              name = decode_entities(name) if name
              toc[name] = sheet_id.to_i - 1 if name && sheet_id
            end
            toc
          end

          # Returns map of sheet name to r:id relationship identifier.
          #
          # @return [Hash<String, String>]
          #: () -> Hash[String, String]
          def sheet_r_ids
            r_ids = {}
            @raw_xml.scan(%r{<(?:[a-zA-Z0-9_]+:)?sheet\b([^>]*?)(?:/>|>(?:.*?)</(?:[a-zA-Z0-9_]+:)?sheet>)}m) do |(attrs)|
              name = attrs[/name="([^"]+)"/, 1]
              rid = attrs[/(?:[a-zA-Z0-9_]+:)?id="([^"]+)"/, 1]
              name = decode_entities(name) if name
              r_ids[name] = rid if name && rid
            end
            r_ids
          end

          # Returns the base date (1900 or 1904) for serial number conversion.
          #
          # @return [Date]
          #: () -> Date
          def base_date
            return DATE_SYSTEM_1900 if @raw_xml.nil?

            if @raw_xml =~ /<(?:[a-zA-Z0-9_]+:)?workbookPr\b[^>]*?\bdate1904="(?:true|1)"/i
              DATE_SYSTEM_1904
            else
              DATE_SYSTEM_1900
            end
          end

          private

          #: (String str) -> String
          def decode_entities(str)
            str.gsub("&amp;", "&").gsub("&quot;", "\"").gsub("&lt;", "<").gsub("&gt;", ">").gsub("&apos;", "'")
          end
        end

        # Map of non-custom numFmtId to casting symbol
        NUM_FMT_MAP = {
          0 => :string,
          1 => :fixnum,
          2 => :float,
          3 => :fixnum,
          4 => :float,
          5 => :unsupported,
          6 => :unsupported,
          7 => :unsupported,
          8 => :unsupported,
          9 => :percentage,
          10 => :percentage,
          11 => :bignum,
          12 => :unsupported,
          13 => :unsupported,
          14 => :date,
          15 => :date,
          16 => :date,
          17 => :date,
          18 => :time,
          19 => :time,
          20 => :time,
          21 => :time,
          22 => :date_time,
          37 => :unsupported,
          38 => :unsupported,
          39 => :unsupported,
          40 => :unsupported,
          44 => :float,
          45 => :time,
          46 => :time,
          47 => :time,
          48 => :bignum,
          49 => :unsupported
        }.freeze
        NumFmtMap = NUM_FMT_MAP

        # Styles XML parser matching SimpleXlsxReader::Loader::StyleTypesParser.
        StyleTypesParser = Struct.new(:file_io) do
          # Parses style types array from styles.xml.
          #
          # @param file_io [IO, StringIO]
          # @return [Array<Symbol>]
          #: (untyped file_io) -> Array[Symbol]
          def self.parse(file_io)
            new(file_io).tap(&:parse).style_types
          end

          # @return [void]
          #: () -> void
          def parse
            raw = file_io.read
            @raw_xml = if raw.b.start_with?("\xFF\xFE".b)
                         raw.force_encoding("UTF-16LE").encode("UTF-8").delete_prefix("\uFEFF")
                       elsif raw.b.start_with?("\xFE\xFF".b)
                         raw.force_encoding("UTF-16BE").encode("UTF-8").delete_prefix("\uFEFF")
                       else
                         raw
                       end
          end

          # Returns array of style type symbols corresponding to cellXfs entries.
          #
          # @return [Array<Symbol>]
          #: () -> Array[Symbol]
          def style_types
            xfs = []
            if (cell_xfs_part = @raw_xml[%r{<(?:[a-zA-Z0-9_]+:)?cellXfs\b[^>]*>(.*?)</(?:[a-zA-Z0-9_]+:)?cellXfs>}m, 1])
              cell_xfs_part.scan(/<(?:[a-zA-Z0-9_]+:)?xf\b([^>]*?)>/) do |(attrs)|
                num_fmt_id = attrs[/numFmtId="(\d+)"/, 1]
                xfs << style_type_by_num_fmt_id(num_fmt_id)
              end
            end
            xfs
          end

          # Finds the style type symbol for a given numFmtId.
          #
          # @param id [String, Integer, nil]
          # @return [Symbol, nil]
          #: ((String | Integer | nil) id) -> Symbol?
          def style_type_by_num_fmt_id(id)
            return nil if id.nil?

            id = id.to_i
            NUM_FMT_MAP[id] || custom_style_types[id]
          end

          # Returns map of custom numFmtId (>= 164) to guessed style symbol.
          #
          # @return [Hash<Integer, Symbol>]
          #: () -> Hash[Integer, Symbol]
          def custom_style_types
            @custom_style_types ||= begin
              types = {}
              if (num_fmts_part = @raw_xml[%r{<(?:[a-zA-Z0-9_]+:)?numFmts\b[^>]*>(.*?)</(?:[a-zA-Z0-9_]+:)?numFmts>}m, 1])
                num_fmts_part.scan(/<(?:[a-zA-Z0-9_]+:)?numFmt\b([^>]*?)>/) do |(attrs)|
                  fid = attrs[/numFmtId="(\d+)"/, 1]
                  fcode = attrs[/formatCode="([^"]*)"/, 1]
                  types[fid.to_i] = determine_custom_style_type(fcode) if fid && fcode
                end
              end
              types
            end
          end

          # Determines custom style type from a formatCode string.
          #
          # @param string [String]
          # @return [Symbol]
          #: (String string) -> Symbol
          def determine_custom_style_type(string)
            return :float if string[0] == "_"
            return :float if string.start_with?(" 0")
            return :date_time if string =~ /(^|\])[^\[]*[ymdhis]/i

            :unsupported
          end
        end

        # Shared strings parser matching SimpleXlsxReader::Loader::SharedStringsParser.
        SharedStringsParser = Struct.new(:file) do
          # Parses shared strings XML into an array of Strings.
          #
          # @param file [IO, StringIO, String]
          # @return [Array<String>]
          #: (untyped file) -> Array[String]
          def self.parse(file)
            xml = file.respond_to?(:read) ? file.read : file.to_s
            Xlsxrb::Ooxml::SharedStringsParser.parse(xml)
          end
        end

        # Worksheet parser matching SimpleXlsxReader::Loader::SheetParser.
        class SheetParser
          # Stream IO for worksheet relationships file (_rels/sheetN.xml.rels).
          # @return [IO, StringIO, nil]
          #: untyped
          attr_accessor :xrels_file

          # Hyperlinks mapping cell coordinate (e.g. "A1") to URL.
          # @return [Hash<String, String>, nil]
          #: Hash[String, String]?
          attr_accessor :hyperlinks_by_cell

          # Sheet name.
          # @return [String, nil]
          #: String?
          attr_accessor :name

          # Cell load errors recorded during parsing.
          # @return [Hash<Array(Integer, Integer), String>]
          #: Hash[[Integer, Integer], String]
          attr_reader :load_errors

          # @param file_io [IO, StringIO]
          # @param loader [Loader]
          #: (file_io: untyped, loader: Loader) -> void
          def initialize(file_io:, loader:)
            @file_io = file_io
            @loader = loader
            @load_errors = {}
            @dimension = nil
            @column_length = nil
            @name = nil
            @headers = nil
            @each_callback = nil
            @current_row_num = nil
            @last_seen_row_idx = 0
          end

          # Parses worksheet rows, invoking the block for each row.
          #
          # @param headers [Boolean, Proc, Hash]
          # @yield [row]
          # @return [void]
          #: (?headers: untyped) { (untyped) -> void } -> void
          def parse(headers: false, &block)
            raise "parse called without a block; what should this do?" unless block_given?

            @headers = headers
            @each_callback = block
            @load_errors = {}
            @current_row_num = nil
            @last_seen_row_idx = 0
            @hyperlinks_by_cell = nil

            @file_io.rewind if @file_io.respond_to?(:rewind)
            raw = @file_io.read
            raw_xml = if raw.b.start_with?("\xFF\xFE".b)
                        raw.force_encoding("UTF-16LE").encode("UTF-8").delete_prefix("\uFEFF")
                      elsif raw.b.start_with?("\xFE\xFF".b)
                        raw.force_encoding("UTF-16BE").encode("UTF-8").delete_prefix("\uFEFF")
                      else
                        raw
                      end

            load_gui_hyperlinks(raw_xml) if xrels_file

            @dimension = raw_xml[/<(?:[a-zA-Z0-9_]+:)?dimension\b[^>]*?\bref="([^"]+)"/, 1]
            @column_length = nil

            row_counter = 0
            raw_xml.scan(%r{<(?:[a-zA-Z0-9_]+:)?row\b([^>]*?)(?:/>|>(.*?)</(?:[a-zA-Z0-9_]+:)?row>)}m) do |row_attrs, row_body|
              r_val = row_attrs[/r="(\d+)"/, 1]
              @current_row_num = r_val ? r_val.to_i : (row_counter + 1)
              row_counter = @current_row_num

              current_row = Array.new(column_length)
              col_idx = 0
              last_col_letter = nil

              row_body&.scan(%r{<(?:[a-zA-Z0-9_]+:)?c\b([^>]*?)(?:/>|>(.*?)</(?:[a-zA-Z0-9_]+:)?c>)}m) do |c_attrs, c_body|
                c_ref = c_attrs[/r="([A-Za-z0-9]+)"/, 1]
                type = c_attrs[/t="([a-zA-Z]+)"/, 1]
                s_attr = c_attrs[/s="(\d+)"/, 1]
                style_idx = s_attr&.to_i
                style = style_idx && @loader.style_types ? @loader.style_types[style_idx] : nil

                cell_col_letter = if c_ref
                                    c_ref.scan(/[A-Za-z]+/).first&.upcase
                                  else
                                    column_number_to_letter(col_idx)
                                  end
                c_idx = column_letter_to_number(cell_col_letter) - 1
                last_col_letter = cell_col_letter
                col_idx = c_idx + 1

                cell_name = c_ref || "#{cell_col_letter}#{@current_row_num}"

                url = nil
                raw_value = nil

                if c_body
                  f_match = c_body[%r{<(?:[a-zA-Z0-9_]+:)?f\b[^>]*?(?:/>|>([^<]*)</(?:[a-zA-Z0-9_]+:)?f>)}, 1]
                  url = f_match.slice(/HYPERLINK\("(.*?)"/, 1) if f_match

                  v_match = c_body[%r{<(?:[a-zA-Z0-9_]+:)?v>([^<]*)</(?:[a-zA-Z0-9_]+:)?v>}, 1]
                  is_match = c_body[%r{<(?:[a-zA-Z0-9_]+:)?is>(.*?)</(?:[a-zA-Z0-9_]+:)?is>}m, 1]

                  if is_match
                    t_parts = []
                    is_match.scan(%r{<(?:[a-zA-Z0-9_]+:)?t\b[^>]*>([^<]*)</(?:[a-zA-Z0-9_]+:)?t>}) do |(t_str)|
                      t_parts << decode_xml(t_str)
                    end
                    raw_value = t_parts.join
                  elsif v_match
                    raw_value = decode_xml(v_match)
                  end
                end

                url ||= @hyperlinks_by_cell[cell_name] if @hyperlinks_by_cell

                val = if raw_value
                        begin
                          Loader.cast(
                            raw_value, type, style,
                            url: url,
                            shared_strings: @loader.shared_strings,
                            base_date: @loader.base_date
                          )
                        rescue StandardError => e
                          row_0_idx = @current_row_num - 1
                          if SimpleXlsxReader.configuration.catch_cell_load_errors
                            @load_errors[[row_0_idx, c_idx]] = e.message
                            raw_value
                          else
                            error = CellLoadError.new("Row #{row_0_idx}, Col #{c_idx}: #{e.message}")
                            error.set_backtrace(e.backtrace)
                            raise error
                          end
                        end
                      end

                current_row[c_idx] = val
              end

              if @dimension.nil? && last_col_letter
                @dimension = "A1:#{last_col_letter}#{@current_row_num}"
                @column_length = column_letter_to_number(last_col_letter)
              end

              process_row(current_row)
              @last_seen_row_idx += 1
            end
          end

          # Returns the column length (number of columns) defined for the sheet.
          #
          # @return [Integer]
          #: () -> Integer
          def column_length
            return 0 unless @dimension

            @column_length ||= column_letter_to_number(last_cell_letter)
          end

          # Returns the column letter string of the last cell from sheet dimension.
          #
          # @return [String, nil]
          #: () -> String?
          def last_cell_letter
            return unless @dimension

            first_match = @dimension.scan(/:([A-Za-z]+)/).first
            return first_match.first.upcase if first_match

            prefix_match = @dimension.scan(/\A([A-Za-z]+)/).first
            return prefix_match.first.upcase if prefix_match

            "A"
          end

          # Converts a column letter string (e.g. "A", "Z", "AA") to a 1-based column number.
          #
          # @param column_letter [String]
          # @return [Integer]
          #: (String column_letter) -> Integer
          def column_letter_to_number(column_letter)
            pow = column_letter.length - 1
            result = 0
            column_letter.each_byte do |b|
              result += (26**pow) * (b - 64)
              pow -= 1
            end
            result
          end

          # Converts a 0-based column number to a column letter string (e.g. 0 -> "A").
          #
          # @param col_num [Integer]
          # @return [String]
          #: (Integer col_num) -> String
          def column_number_to_letter(col_num)
            result = []
            loop do
              result.unshift(((col_num % 26) + 65).chr)
              col_num = (col_num / 26) - 1
              break if col_num.negative?
            end
            result.join
          end

          private

          #: (Array[untyped] current_row) -> void
          def process_row(current_row)
            if @headers == true
              @headers = current_row
            elsif @headers.is_a?(Hash)
              test_headers_hash_against_current_row(current_row)
              @last_seen_row_idx = @current_row_num - 1
            elsif @headers.respond_to?(:call)
              @headers = current_row if @headers.call(current_row)
              @last_seen_row_idx = @current_row_num - 1
            elsif @headers
              possibly_yield_empty_rows(headers: true)
              yield_row(current_row, headers: true)
            else
              possibly_yield_empty_rows(headers: false)
              yield_row(current_row, headers: false)
            end
          end

          #: (Array[untyped] current_row) -> void
          def test_headers_hash_against_current_row(current_row)
            found = false
            current_row.each_with_index do |cell, cell_idx|
              @headers.each_pair do |key, search|
                matched = if search.is_a?(String)
                            cell == search
                          elsif cell
                            cell.to_s.match?(search)
                          else
                            false
                          end
                if matched
                  found = true
                  current_row[cell_idx] = key
                end
              end
            end
            @headers = current_row if found
          end

          #: (headers: bool) -> void
          def possibly_yield_empty_rows(headers:)
            while @current_row_num && @current_row_num > @last_seen_row_idx + 1
              @last_seen_row_idx += 1
              yield_row(Array.new(column_length), headers: headers)
            end
          end

          #: (Array[untyped] row, headers: bool) -> void
          def yield_row(row, headers:)
            if headers
              h_keys = @headers.is_a?(Array) ? @headers : []
              @each_callback&.call(h_keys.zip(row).to_h)
            else
              @each_callback&.call(row)
            end
          end

          #: (String sheet_xml) -> void
          def load_gui_hyperlinks(sheet_xml)
            @hyperlinks_by_cell = {}
            return unless xrels_file

            xrels_file.rewind if xrels_file.respond_to?(:rewind)
            xrels_content = xrels_file.read
            rels_map = Xlsxrb::Ooxml::RelationshipsParser.parse(xrels_content)

            sheet_xml.scan(%r{<(?:[a-zA-Z0-9_]+:)?hyperlink\b([^>]*?)/>}m) do |(attrs)|
              ref = attrs[/ref="([^"]+)"/, 1]
              id = attrs[/(?:[a-zA-Z0-9_]+:)?id="([^"]+)"/, 1]
              target = rels_map[id]
              @hyperlinks_by_cell[ref] = target if ref && target
            end
          end

          #: (String? str) -> String
          def decode_xml(str)
            return "" if str.nil? || str.empty?

            str = str.gsub("&amp;", "&") if str.include?("&amp;")
            str = str.gsub("&lt;", "<") if str.include?("&lt;")
            str = str.gsub("&gt;", ">") if str.include?("&gt;")
            str = str.gsub("&quot;", "\"") if str.include?("&quot;")
            str = str.gsub("&apos;", "'") if str.include?("&apos;")
            str.gsub("\r\n", "\n").force_encoding(Encoding::UTF_8)
          end
        end
      end
    end
  end
end
