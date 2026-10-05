# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require "stringio"
require_relative "errors"
require_relative "xml"
require_relative "shared_string_table"
require_relative "worksheet"
require_relative "zip_kit_writer"

module Xlsxrb
  module Adapters
    module Xlsxtream
      # Streaming XLSX workbook generator compatible with Xlsxtream::Workbook.
      class Workbook
        FONT_FAMILY_IDS = {
          "" => 0,
          "roman" => 1,
          "swiss" => 2,
          "modern" => 3,
          "script" => 4,
          "decorative" => 5
        }.freeze

        # @return [Array<Worksheet>]
        attr_reader :worksheets
        alias sheets worksheets

        # @return [Hash{Symbol => untyped}]
        attr_reader :options

        # Opens or creates a new workbook and yields it to the block, ensuring it closes.
        #
        # @param output [String, IO, ZipKit::Streamer] Destination target.
        # @param options [Hash{Symbol => untyped}] Workbook options.
        # @yield [workbook]
        # @yieldparam workbook [Workbook]
        # @return [Workbook, untyped]
        def self.open(output, options = {})
          workbook = new(output, options)
          if block_given?
            begin
              yield workbook
            ensure
              workbook.close
            end
          else
            workbook
          end
        end

        # @param output [String, IO, ZipKit::Streamer] Destination target.
        # @param options [Hash{Symbol => untyped}] Workbook options.
        def initialize(output, options = {})
          @writer = ZipKitWriter.with_output_to(output)
          @options = options
          @sst = SharedStringTable.new
          @worksheets = []
          @closed = false
        end

        # Adds a new worksheet sequentially.
        #
        # @param args [Array<Object>] Worksheet arguments (name and/or options).
        # @yield [worksheet]
        # @return [Worksheet]
        def add_worksheet(*, &)
          if block_given?
            warn "#{caller(1..1).first[/.*:\d+:(?=in `)/]} warning: Calling #{self.class}#add_worksheet with a block is deprecated, use #write_worksheet instead."
            return write_worksheet(*, &)
          end

          raise Error, "Close the current worksheet before adding a new one" unless @worksheets.all?(&:closed?)

          build_worksheet(*)
        end

        # Opens, yields, and closes a worksheet sequentially.
        #
        # @param args [Array<Object>] Worksheet arguments.
        # @yield [worksheet]
        # @return [nil]
        def write_worksheet(*)
          worksheet = build_worksheet(*)
          yield worksheet if block_given?
          worksheet.close
          nil
        end

        # Finalizes workbook metadata and writes package structures.
        #
        # @return [nil]
        def close
          return if @closed

          @closed = true
          write_workbook
          write_styles
          write_sst unless @sst.empty?
          write_workbook_rels
          write_root_rels
          write_content_types
          @writer.close
          nil
        end

        # Converts mutable workbook to an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        def to_xlsxrb
          font_options = @options.fetch(:font, {})
          font_size = font_options.fetch(:size, 12).to_f
          font_name = font_options.fetch(:name, "Calibri").to_s
          font_family = font_options.fetch(:family, "Swiss").to_s.downcase
          font_family_id = FONT_FAMILY_IDS[font_family] || 2

          styles_hash = {
            fonts: [{ name: font_name, sz: font_size, family: font_family_id }],
            fills: [{ pattern: "none" }, { pattern: "gray125" }],
            borders: [{}],
            cell_xfs: [
              { num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 },
              { num_fmt_id: 164, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: 1 },
              { num_fmt_id: 165, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: 1 }
            ],
            xf_entries: [
              { num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0 },
              { num_fmt_id: 164, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: 1 },
              { num_fmt_id: 165, font_id: 0, fill_id: 0, border_id: 0, apply_number_format: 1 }
            ],
            num_fmts: [
              { id: 164, code: 'yyyy\\-mm\\-dd' },
              { id: 165, code: 'yyyy\\-mm\\-dd hh:mm:ss' }
            ]
          }
          styles = Xlsxrb::Elements::Styles.new(styles_hash)

          elements_sheets = @worksheets.map(&:to_xlsxrb)
          sst_array = @sst.empty? ? nil : @sst.keys

          Xlsxrb::Elements::Workbook.new(
            sheets: elements_sheets,
            styles: styles,
            shared_strings: sst_array
          )
        end

        # Builds an Xlsxtream::Workbook adapter instance from an Xlsxrb::Elements::Workbook.
        #
        # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
        # @param output [Object, nil]
        # @param options [Hash{Symbol => untyped}]
        # @return [Workbook]
        def self.from_xlsxrb(xlsxrb_workbook, output = nil, options = {})
          wb_source = xlsxrb_workbook.respond_to?(:load) ? xlsxrb_workbook.load : xlsxrb_workbook
          sio = output || StringIO.new
          adapter = new(sio, options)

          wb_source.sheets.each do |sheet|
            col_opts = (sheet.columns.map { |col| { width_pixels: col.width } } if sheet.columns && !sheet.columns.empty?)

            ws_opts = col_opts ? { columns: col_opts } : {}
            adapter.write_worksheet(sheet.name, ws_opts) do |ws|
              sheet.rows.each do |row|
                cell_values = row.cells.map { |c| c&.value }
                ws << cell_values
              end
            end
          end

          adapter.close
          adapter
        end

        private

        def build_worksheet(name = nil, options = {})
          if name.is_a?(Hash) && options.empty?
            options = name
            name = nil
          end

          use_sst = options.fetch(:use_shared_strings, @options[:use_shared_strings])
          auto_format = options.fetch(:auto_format, @options[:auto_format])
          columns = options.fetch(:columns, @options[:columns])
          retain_rows = options.fetch(:retain_rows, @options.fetch(:retain_rows, true))
          sst = use_sst ? @sst : nil

          sheet_id = @worksheets.size + 1
          name = name || options[:name] || "Sheet#{sheet_id}"

          @writer.add_file "xl/worksheets/sheet#{sheet_id}.xml"

          worksheet = Worksheet.new(
            @writer,
            id: sheet_id,
            name: name,
            sst: sst,
            auto_format: auto_format,
            columns: columns,
            retain_rows: retain_rows
          )
          @worksheets << worksheet

          worksheet
        end

        def write_root_rels
          @writer.add_file "_rels/.rels"
          @writer << XML.header
          @writer << XML.strip(<<-XML)
            <Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">
              <Relationship Id="rId1" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Target="xl/workbook.xml"/>
            </Relationships>
          XML
        end

        def write_workbook
          rid = String.new("rId0")
          @writer.add_file "xl/workbook.xml"
          @writer << XML.header
          @writer << XML.strip(<<-XML)
            <workbook xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" xmlns:r="http://schemas.openxmlformats.org/officeDocument/2006/relationships">
              <workbookPr date1904="false"/>
              <sheets>
          XML
          @worksheets.each do |worksheet|
            @writer << %(<sheet name="#{XML.escape_attr worksheet.name}" sheetId="#{worksheet.id}" r:id="#{rid.next!}"/>)
          end
          @writer << XML.strip(<<-XML)
              </sheets>
            </workbook>
          XML
        end

        def write_styles
          font_options = @options.fetch(:font, {})
          font_size = font_options.fetch(:size, 12).to_s
          font_name = font_options.fetch(:name, "Calibri").to_s
          font_family = font_options.fetch(:family, "Swiss").to_s.downcase
          font_family_id = FONT_FAMILY_IDS[font_family] or raise Error,
                                                                 "Invalid font family #{font_family}, must be one of " \
                                                                 "#{FONT_FAMILY_IDS.keys.map(&:inspect).join(", ")}"

          @writer.add_file "xl/styles.xml"
          @writer << XML.header
          @writer << XML.strip(<<-XML)
            <styleSheet xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main">
              <numFmts count="2">
                <numFmt numFmtId="164" formatCode="yyyy\\-mm\\-dd"/>
                <numFmt numFmtId="165" formatCode="yyyy\\-mm\\-dd hh:mm:ss"/>
              </numFmts>
              <fonts count="1">
                <font>
                  <sz val="#{XML.escape_attr font_size}"/>
                  <name val="#{XML.escape_attr font_name}"/>
                  <family val="#{font_family_id}"/>
                </font>
              </fonts>
              <fills count="2">
                <fill>
                  <patternFill patternType="none"/>
                </fill>
                <fill>
                  <patternFill patternType="gray125"/>
                </fill>
              </fills>
              <borders count="1">
                <border/>
              </borders>
              <cellStyleXfs count="1">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0"/>
              </cellStyleXfs>
              <cellXfs count="3">
                <xf numFmtId="0" fontId="0" fillId="0" borderId="0" xfId="0"/>
                <xf numFmtId="164" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
                <xf numFmtId="165" fontId="0" fillId="0" borderId="0" xfId="0" applyNumberFormat="1"/>
              </cellXfs>
              <cellStyles count="1">
                <cellStyle name="Normal" xfId="0" builtinId="0"/>
              </cellStyles>
              <dxfs count="0"/>
              <tableStyles count="0" defaultTableStyle="TableStyleMedium9" defaultPivotStyle="PivotStyleLight16"/>
            </styleSheet>
          XML
        end

        def write_sst
          @writer.add_file "xl/sharedStrings.xml"
          @writer << XML.header
          @writer << %(<sst xmlns="http://schemas.openxmlformats.org/spreadsheetml/2006/main" count="#{@sst.references}" uniqueCount="#{@sst.size}">)
          @sst.each_key do |string|
            @writer << "<si><t>#{XML.escape_value string}</t></si>"
          end
          @writer << "</sst>"
        end

        def write_workbook_rels
          rid = String.new("rId0")
          @writer.add_file "xl/_rels/workbook.xml.rels"
          @writer << XML.header
          @writer << '<Relationships xmlns="http://schemas.openxmlformats.org/package/2006/relationships">'
          @worksheets.each do |worksheet|
            @writer << %(<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Target="worksheets/sheet#{worksheet.id}.xml"/>)
          end
          @writer << %(<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Target="styles.xml"/>)
          @writer << %(<Relationship Id="#{rid.next!}" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Target="sharedStrings.xml"/>) unless @sst.empty?
          @writer << "</Relationships>"
        end

        def write_content_types
          @writer.add_file "[Content_Types].xml"
          @writer << XML.header
          @writer << XML.strip(<<-XML)
            <Types xmlns="http://schemas.openxmlformats.org/package/2006/content-types">
              <Default Extension="xml" ContentType="application/xml"/>
              <Default Extension="rels" ContentType="application/vnd.openxmlformats-package.relationships+xml"/>
              <Override PartName="/xl/workbook.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"/>
              <Override PartName="/xl/styles.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"/>
          XML
          @writer << '<Override PartName="/xl/sharedStrings.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"/>' unless @sst.empty?
          @worksheets.each do |worksheet|
            @writer << %(<Override PartName="/xl/worksheets/sheet#{worksheet.id}.xml" ContentType="application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"/>)
          end
          @writer << "</Types>"
        end
      end
    end
  end
end
