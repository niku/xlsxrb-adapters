# frozen_string_literal: true

# rbs_inline: enabled

require "tmpdir"
require_relative "utils"

module Xlsxrb
  module Adapters
    module Creek
      # Extracts and maps embedded spreadsheet images matching Creek::Drawing.
      class Drawing
        include Utils

        COLUMNS = ("A".."AZ").to_a.freeze

        # @return [untyped]
        attr_reader :book

        # @return [String]
        attr_reader :drawing_filepath

        # @return [Array<untyped>]
        attr_reader :drawings

        # @return [untyped]
        attr_reader :drawings_rels

        # @return [Hash<Array<Integer>, Array<Pathname>>]
        attr_reader :images_pathnames

        # @param book [untyped]
        # @param drawing_filepath [String]
        #: (untyped book, String drawing_filepath) -> void
        def initialize(book, drawing_filepath)
          @book = book
          @drawing_filepath = drawing_filepath
          @drawings = []
          @drawings_rels = []
          @images_pathnames = Hash.new { |hash, key| hash[key] = [] }

          return unless file_exist?(@drawing_filepath)

          load_drawings_and_rels
          load_images_pathnames_by_cells if has_images?
        end

        # Returns whether there are images in the drawing file.
        #
        # @return [Boolean]
        #: () -> bool
        def has_images?
          @has_images ||= !@drawings.empty?
        end

        # Extracts images for a cell coordinate to a temporary folder and returns Pathname objects.
        #
        # @param cell_name [String]
        # @return [Array<Pathname>, nil]
        #: (String cell_name) -> Array[Pathname]?
        def images_at(cell_name)
          coordinate = calc_coordinate(cell_name)
          pathnames_at_coordinate = @images_pathnames[coordinate]
          return if pathnames_at_coordinate.empty?

          pathnames_at_coordinate.map do |image_pathname|
            unless image_pathname.exist?
              excel_image_path = "xl/media#{image_pathname.to_path.split(tmpdir).last}"
              IO.copy_stream(@book.files.file.open(excel_image_path), image_pathname.to_path)
            end
            image_pathname
          end
        end

        private

        # Transforms cell coordinate name (e.g. "A1", "B3") into 0-based [row, col].
        #
        # @param cell_name [String]
        # @return [Array<Integer>]
        #: (String cell_name) -> [Integer, Integer]
        def calc_coordinate(cell_name)
          coords = Xlsxrb::Utils.ref_to_row_col(cell_name)
          coords || [0, 0]
        end

        # Returns the temporary directory path for extracted images.
        #
        # @return [String]
        #: () -> String
        def tmpdir
          @tmpdir ||= ::Dir.mktmpdir("creek__drawing")
        end

        # Loads drawing and drawing rels.
        #: () -> void
        def load_drawings_and_rels
          if defined?(::Nokogiri::XML::Document)
            doc = parse_xml(@drawing_filepath)
            if doc.respond_to?(:css)
              @drawings = doc.css("xdr|twoCellAnchor", "xdr|oneCellAnchor")
              drawing_rels_filepath = expand_to_rels_path(@drawing_filepath)
              @drawings_rels = parse_xml(drawing_rels_filepath).css("Relationships") if file_exist?(drawing_rels_filepath)
              return
            end
          end

          # Fallback without Nokogiri
          xml_str = @book.files.file.open(@drawing_filepath).read
          drawing_rels_filepath = expand_to_rels_path(@drawing_filepath)
          rels_str = file_exist?(drawing_rels_filepath) ? @book.files.file.open(drawing_rels_filepath).read : ""

          @drawings = []
          xml_str.scan(%r{<(?:[A-Za-z0-9_]+:)?(twoCellAnchor|oneCellAnchor)\b([^>]*?)>(.*?)</(?:[A-Za-z0-9_]+:)?\1>}m) do |type, attrs, body|
            @drawings << { type: type, attrs: attrs, body: body }
          end

          @drawings_rels = {}
          rels_str.scan(%r{<Relationship\b([^>]*?)/>}m) do |m|
            attrs = m[0]
            id = attrs[/Id="([^"]*)"/, 1]
            target = attrs[/Target="([^"]*)"/, 1]
            @drawings_rels[id] = target if id && target
          end
        end

        # Populates @images_pathnames by cell coordinate.
        #: () -> void
        def load_images_pathnames_by_cells
          if @drawings.first.is_a?(Hash)
            load_images_pathnames_from_hashes
          else
            load_images_pathnames_from_nokogiri
          end
        end

        #: () -> void
        def load_images_pathnames_from_nokogiri
          image_selector = "xdr:pic/xdr:blipFill/a:blip"
          row_from_selector = "xdr:from/xdr:row"
          row_to_selector = "xdr:to/xdr:row"
          col_from_selector = "xdr:from/xdr:col"
          col_to_selector = "xdr:to/xdr:col"

          @drawings.each do |drawing|
            temp = drawing.xpath(image_selector).first
            embed = temp ? temp.attributes["embed"] : nil
            next if embed.nil?

            rid = embed.value
            path = Pathname.new("#{tmpdir}/#{extract_drawing_path(rid).slice(%r{[^/]*$})}")

            row_from = drawing.xpath(row_from_selector).text.to_i
            col_from = drawing.xpath(col_from_selector).text.to_i

            if drawing.name == "oneCellAnchor"
              @images_pathnames[[row_from, col_from]].push(path)
            else
              row_to = drawing.xpath(row_to_selector).text.to_i
              col_to = drawing.xpath(col_to_selector).text.to_i

              (col_from..col_to).each do |col|
                (row_from..row_to).each do |row|
                  @images_pathnames[[row, col]].push(path)
                end
              end
            end
          end
        end

        #: () -> void
        def load_images_pathnames_from_hashes
          @drawings.each do |entry|
            body = entry[:body]
            embed = body[/<(?:[A-Za-z0-9_]+:)?blip\b[^>]*?(?:r:embed|embed)="([^"]*)"/, 1]
            next unless embed

            target = @drawings_rels.is_a?(Hash) ? @drawings_rels[embed] : nil
            next unless target

            filename = target.slice(%r{[^/]*$})
            path = Pathname.new("#{tmpdir}/#{filename}")

            from_col = body[%r{<[^:]*:from>.*?<[^:]*:col>(\d+)</[^:]*:col>}m, 1].to_i
            from_row = body[%r{<[^:]*:from>.*?<[^:]*:row>(\d+)</[^:]*:row>}m, 1].to_i

            if entry[:type] == "oneCellAnchor"
              @images_pathnames[[from_row, from_col]].push(path)
            else
              to_col = body[%r{<[^:]*:to>.*?<[^:]*:col>(\d+)</[^:]*:col>}m, 1].to_i
              to_row = body[%r{<[^:]*:to>.*?<[^:]*:row>(\d+)</[^:]*:row>}m, 1].to_i
              (from_col..to_col).each do |col|
                (from_row..to_row).each do |row|
                  @images_pathnames[[row, col]].push(path)
                end
              end
            end
          end
        end

        # Resolves drawing target path from relationship ID.
        #
        # @param rid [String]
        # @return [String]
        #: (String rid) -> String
        def extract_drawing_path(rid)
          if @drawings_rels.respond_to?(:css)
            @drawings_rels.css("Relationship[@Id='#{rid}']").first.attributes["Target"].value
          elsif @drawings_rels.is_a?(Hash)
            @drawings_rels[rid] || ""
          else
            ""
          end
        end
      end
    end
  end
end
