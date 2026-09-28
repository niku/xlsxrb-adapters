# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # PivotTableCacheDefinition class.
      class PivotTableCacheDefinition
        include OptionsParser

        # @return [PivotTable]
        attr_reader :pivot_table

        # @param pivot_table [PivotTable]
        #: (PivotTable pivot_table) -> void
        def initialize(pivot_table)
          @pivot_table = pivot_table
        end

        # @return [Integer]
        #: () -> Integer
        def index
          pivot_table.sheet.workbook.pivot_tables.index(pivot_table) || 0
        end

        # @return [String]
        #: () -> String
        def pn
          format(PIVOT_TABLE_CACHE_DEFINITION_PN, index + 1)
        end

        # @return [Integer]
        #: () -> Integer
        def cache_id
          index + 1
        end

        # @return [String]
        #: () -> String
        def rId
          pivot_table.relationships.for(self).Id
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << '<pivotCacheDefinition xmlns="' << XML_NS << '" xmlns:r="' << XML_NS_R << '" invalid="1" refreshOnLoad="1" recordCount="0">'
          str << '<cacheSource type="worksheet">'
          str << '<worksheetSource ref="' << pivot_table.range.to_s << '" sheet="' << pivot_table.data_sheet.name << '"/>'
          str << "</cacheSource>"
          str << '<cacheFields count="' << pivot_table.header_cells_count.to_s << '">'
          pivot_table.header_cells.each do |cell|
            str << '<cacheField name="' << cell.clean_value << '" numFmtId="0">'
            str << '<sharedItems count="0">'
            str << "</sharedItems>"
            str << "</cacheField>"
          end
          str << "</cacheFields>"
          str << "</pivotCacheDefinition>"
        end
      end

      # PivotTable class representing an Excel pivot table.
      class PivotTable
        include OptionsParser

        # @return [Array[String]]
        attr_accessor :no_subtotals_on_headers

        # @return [Hash{String => Symbol}]
        attr_reader :sort_on_headers

        # @return [Symbol]
        attr_reader :grand_totals

        # @return [Hash{Symbol => untyped}]
        attr_accessor :style_info

        # @return [String]
        attr_reader :ref

        # @return [String]
        attr_reader :name

        # @return [Worksheet]
        attr_reader :sheet

        # @return [Worksheet, nil]
        attr_writer :data_sheet

        # @return [Boolean]
        attr_accessor :use_auto_formatting

        # @return [Boolean]
        attr_accessor :apply_width_height_formats

        # @return [String, nil]
        attr_reader :range

        # @return [Array[String]]
        attr_reader :rows

        # @return [Array[String]]
        attr_reader :columns

        # @return [Array[Hash{Symbol => untyped}]]
        attr_reader :data

        # @return [Array[String]]
        attr_reader :pages

        # @param ref [String]
        # @param range [String]
        # @param sheet [Worksheet]
        # @param options [Hash{Symbol => untyped}]
        #: (String ref, String range, Worksheet sheet, ?Hash[Symbol, untyped] options) ?{ (PivotTable) -> void } -> void
        def initialize(ref, range, sheet, options = {})
          @ref = ref
          self.range = range
          @sheet = sheet
          @sheet.workbook.pivot_tables << self
          @name = "PivotTable#{index + 1}"
          @data_sheet = nil
          @rows = []
          @columns = []
          @data = []
          @pages = []
          @subtotal = nil
          @no_subtotals_on_headers = []
          @grand_totals = :both
          @sort_on_headers = {}
          @style_info = {}
          @use_auto_formatting = true
          @apply_width_height_formats = true
          parse_options(options)
          yield self if block_given?
        end

        # @param headers [Hash{String => Symbol}, Array[String], nil]
        # @return [Hash{String => Symbol}]
        #: (Hash[String, Symbol] | Array[String] | nil headers) -> Hash[String, Symbol]
        def sort_on_headers=(headers)
          headers ||= {}
          headers = headers.to_h { |h| [h, :ascending] } if headers.is_a?(Array)
          @sort_on_headers = headers
        end

        # @param value [Symbol]
        # @return [Symbol]
        #: (Symbol value) -> Symbol
        def grand_totals=(value)
          RestrictionValidator.validate("PivotTable.grand_totals", %i[both row_only col_only none], value)
          @grand_totals = value
        end

        # @return [Worksheet]
        #: () -> Worksheet
        def data_sheet
          @data_sheet || @sheet
        end

        # @param v [String]
        # @return [void]
        #: (String v) -> void
        def range=(v)
          DataTypeValidator.validate("#{self.class}.range", [String], v)
          @range = v if v.is_a?(String)
        end

        # @param v [Array[String]]
        # @return [void]
        #: (Array[String] v) -> void
        def rows=(v)
          DataTypeValidator.validate("#{self.class}.rows", [Array], v)
          v.each do |ref|
            DataTypeValidator.validate("#{self.class}.rows[]", [String], ref)
          end
          @rows = v
        end

        # @param v [Array[String]]
        # @return [void]
        #: (Array[String] v) -> void
        def columns=(v)
          DataTypeValidator.validate("#{self.class}.columns", [Array], v)
          v.each do |ref|
            DataTypeValidator.validate("#{self.class}.columns[]", [String], ref)
          end
          @columns = v
        end

        # @param v [Array[String, Hash{Symbol => untyped}]]
        # @return [void]
        #: (Array[String | Hash[Symbol, untyped]] v) -> void
        def data=(v)
          DataTypeValidator.validate("#{self.class}.data", [Array], v)
          @data = []
          v.each do |data_field|
            data_field = { ref: data_field } if data_field.is_a?(String)
            DataTypeValidator.validate("#{self.class}.data[]", [Hash], data_field)
            data_field.each do |key, val|
              if key == :num_fmt
                DataTypeValidator.validate("#{self.class}.data[]", [Integer], val)
              else
                DataTypeValidator.validate("#{self.class}.data[]", [String], val)
              end
            end
            @data << data_field
          end
        end

        # @param v [Array[String]]
        # @return [void]
        #: (Array[String] v) -> void
        def pages=(v)
          DataTypeValidator.validate("#{self.class}.pages", [Array], v)
          v.each do |ref|
            DataTypeValidator.validate("#{self.class}.pages[]", [String], ref)
          end
          @pages = v
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @sheet.workbook.pivot_tables.index(self) || 0
        end

        # @return [String]
        #: () -> String
        def pn
          format(PIVOT_TABLE_PN, index + 1)
        end

        # @return [String]
        #: () -> String
        def rels_pn
          format(PIVOT_TABLE_RELS_PN, index + 1)
        end

        # @return [PivotTableCacheDefinition]
        #: () -> PivotTableCacheDefinition
        def cache_definition
          @cache_definition ||= PivotTableCacheDefinition.new(self)
        end

        # @return [Relationships]
        #: () -> Relationships
        def relationships
          r = Relationships.new
          r << Relationship.new(cache_definition, PIVOT_TABLE_CACHE_DEFINITION_R, "../#{cache_definition.pn}")
          r
        end

        # @return [Array[String]]
        #: () -> Array[String]
        def header_cell_refs
          Caxlsx.range_to_a(header_range).first || []
        end

        # @return [Array[Cell]]
        #: () -> Array[Cell]
        def header_cells
          res = data_sheet[header_range]
          cells = res.is_a?(Array) ? res : [res]
          cells.grep(Cell) #: Array[Cell]
        end

        # @return [Array[untyped]]
        #: () -> Array[untyped]
        def header_cell_values
          header_cells.map(&:value)
        end

        # @return [Integer]
        #: () -> Integer
        def header_cells_count
          header_cells.count
        end

        # @param value [untyped]
        # @return [Integer, nil]
        #: (untyped value) -> Integer?
        def header_index_of(value)
          header_cell_values.index(value)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << '<pivotTableDefinition xmlns="' << XML_NS << '" name="' << name << '" cacheId="' << cache_definition.cache_id.to_s << '"'
          str << ' dataOnRows="1"' if data.size <= 1
          str << ' rowGrandTotals="0"' if %i[col_only none].include?(grand_totals)
          str << ' colGrandTotals="0"' if %i[row_only none].include?(grand_totals)
          str << ' applyNumberFormats="0" applyBorderFormats="0" applyFontFormats="0" applyPatternFormats="0" applyAlignmentFormats="0"'
          str << ' applyWidthHeightFormats="' << (@apply_width_height_formats ? "1" : "0") << '"'
          str << ' dataCaption="Data" showMultipleLabel="0" showMemberPropertyTips="0"'
          str << ' useAutoFormatting="' << (@use_auto_formatting ? "1" : "0") << '"'
          str << ' indent="0" compact="0" compactData="0" gridDropZones="1" multipleFieldFilters="0">'

          str << '<location firstDataCol="1" firstDataRow="1" firstHeaderRow="1" ref="' << ref << '"/>'
          str << '<pivotFields count="' << header_cells_count.to_s << '">'

          header_cell_values.each do |cell_value|
            subtotal = !no_subtotals_on_headers.include?(cell_value)
            sorttype = sort_on_headers[cell_value]
            str << pivot_field_for(cell_value, subtotal, sorttype)
          end

          str << "</pivotFields>"
          if rows.empty?
            str << '<rowFields count="1"><field x="-2"/></rowFields>'
            str << '<rowItems count="2"><i><x/></i> <i i="1"><x v="1"/></i></rowItems>'
          else
            str << '<rowFields count="' << rows.size.to_s << '">'
            rows.each do |row_value|
              str << '<field x="' << header_index_of(row_value).to_s << '"/>'
            end
            str << "</rowFields>"
            str << '<rowItems count="' << rows.size.to_s << '">'
            rows.size.times do
              str << "<i/>"
            end
            str << "</rowItems>"
          end

          if columns.empty?
            if data.size > 1
              str << '<colFields count="1"><field x="-2"/></colFields>'
              str << "<colItems count=\"#{data.size}\">"
              str << "<i><x/></i>"
              (data.size - 1).times do |i|
                str << "<i i=\"#{i + 1}\"><x v=\"#{i + 1}\"/></i>"
              end
              str << "</colItems>"
            else
              str << '<colItems count="1"><i/></colItems>'
            end
          elsif data.size > 1
            str << '<colFields count="' << (columns.size + 1).to_s << '">'
            columns.each do |column_value|
              str << '<field x="' << header_index_of(column_value).to_s << '"/>'
            end
            str << '<field x="-2"/></colFields>'
            str << "<colItems count=\"#{data.size}\">"
            str << "<i><x/></i>"
            (data.size - 1).times do |i|
              str << "<i i=\"#{i + 1}\"><x v=\"#{i + 1}\"/></i>"
            end
            str << "</colItems>"
          else
            str << '<colFields count="' << columns.size.to_s << '">'
            columns.each do |column_value|
              str << '<field x="' << header_index_of(column_value).to_s << '"/>'
            end
            str << "</colFields>"
          end

          unless pages.empty?
            str << '<pageFields count="' << pages.size.to_s << '">'
            pages.each do |page_value|
              str << '<pageField fld="' << header_index_of(page_value).to_s << '"/>'
            end
            str << "</pageFields>"
          end

          unless data.empty?
            str << "<dataFields count=\"#{data.size}\">"
            data.each do |datum_value|
              subtotal_name = datum_value[:subtotal] || "sum"
              subtotal_name = "count" if datum_value[:subtotal] == "countNums"
              field_name = datum_value[:name] || "#{subtotal_name.capitalize} of #{datum_value[:ref]}"
              str << "<dataField name='#{field_name}' fld='#{header_index_of(datum_value[:ref])}' baseField='0' baseItem='0'"
              str << " numFmtId='#{datum_value[:num_fmt]}'" if datum_value[:num_fmt]
              str << " subtotal='#{datum_value[:subtotal]}' " if datum_value[:subtotal]
              str << "/>"
            end
            str << "</dataFields>"
          end

          unless style_info.empty?
            str << "<pivotTableStyleInfo"
            style_info.each do |k, v|
              str << " " << k.to_s << '="' << v.to_s << '"'
            end
            str << " />"
          end
          str << "</pivotTableDefinition>"
        end

        private

        # @param cell_ref [untyped]
        # @param subtotal [Boolean]
        # @param sorttype [Symbol, nil]
        # @return [String]
        #: (untyped cell_ref, bool subtotal, Symbol? sorttype) -> String
        def pivot_field_for(cell_ref, subtotal, sorttype)
          attributes = %w[compact="0" outline="0" subtotalTop="0" showAll="0" includeNewItemsInFilter="1"]
          items_tag = '<items count="1"><item t="default"/></items>'
          include_items_tag = false

          if rows.include?(cell_ref)
            attributes << 'axis="axisRow"'
            attributes << "sortType=\"#{sorttype == :descending ? "descending" : "ascending"}\"" if sorttype
            if subtotal
              include_items_tag = true
            else
              attributes << 'defaultSubtotal="0"'
            end
          elsif columns.include?(cell_ref)
            attributes << 'axis="axisCol"'
            attributes << "sortType=\"#{sorttype == :descending ? "descending" : "ascending"}\"" if sorttype
            if subtotal
              include_items_tag = true
            else
              attributes << 'defaultSubtotal="0"'
            end
          elsif pages.include?(cell_ref)
            attributes << 'axis="axisPage"'
            include_items_tag = true
          elsif data_refs.include?(cell_ref)
            attributes << 'dataField="1"'
          end

          "<pivotField #{attributes.join(" ")}>#{items_tag if include_items_tag}</pivotField>"
        end

        # @return [Array[untyped]]
        #: () -> Array[untyped]
        def data_refs
          data.map { |h| h[:ref] }
        end

        # @return [String]
        #: () -> String
        def header_range
          (range || "").gsub(/^(\w+?)(\d+):(\w+?)\d+$/, '\1\2:\3\2')
        end
      end

      # Collection of pivot tables in a worksheet.
      class PivotTables < SimpleTypedList
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          super(PivotTable)
          @worksheet = worksheet
        end

        # @return [Array[Relationship]]
        #: () -> Array[Relationship]
        def relationships
          return [] if empty?

          map { |pivot_table| Relationship.new(pivot_table, PIVOT_TABLE_R, "../#{pivot_table.pn}") }
        end
      end
    end
  end
end
