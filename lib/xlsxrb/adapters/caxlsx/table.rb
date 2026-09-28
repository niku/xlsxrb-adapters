# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # The table style info class manages style attributes for defined tables.
      class TableStyleInfo
        include OptionsParser
        include SerializedAttributes
        include Accessors

        boolean_attr_accessor :show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes
        serializable_attributes :show_first_column, :show_last_column, :show_row_stripes, :show_column_stripes, :name

        # @return [String]
        attr_accessor :name

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          initialize_defaults
          @name = "TableStyleMedium9"
          parse_options(options)
        end

        # @return [void]
        #: () -> void
        def initialize_defaults
          %w[show_first_column show_last_column show_row_stripes show_column_stripes].each do |attr|
            send(:"#{attr}=", 0)
          end
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("tableStyleInfo", str)
        end
      end

      # Table representation for a worksheet.
      class Table
        include OptionsParser

        # @return [String]
        attr_reader :ref

        # @return [String]
        attr_reader :name

        # @return [untyped]
        attr_reader :style

        # @return [Worksheet]
        attr_reader :sheet

        # @param ref [String]
        # @param sheet [Worksheet]
        # @param options [Hash{Symbol => untyped}]
        #: (String ref, Worksheet sheet, ?Hash[Symbol, untyped] options) ?{ (Table) -> void } -> void
        def initialize(ref, sheet, options = {})
          @ref = ref
          @sheet = sheet
          @style = nil
          @sheet.workbook.tables << self
          @table_style_info = TableStyleInfo.new(options[:style_info]) if options[:style_info]
          @name = "Table#{index + 1}"
          parse_options(options)
          yield self if block_given?
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @sheet.workbook.tables.index(self) || 0
        end

        # @return [String]
        #: () -> String
        def pn
          format(TABLE_PN, index + 1)
        end

        # @return [String]
        #: () -> String
        def rId
          @sheet.relationships.for(self).Id
        end

        # @param v [String]
        # @return [void]
        #: (String v) -> void
        def name=(v)
          DataTypeValidator.validate(:table_name, [String], v)
          @name = v if v.is_a?(String)
        end

        # @return [TableStyleInfo]
        #: () -> TableStyleInfo
        def table_style_info
          @table_style_info ||= TableStyleInfo.new
        end
        alias style_info table_style_info

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << '<table xmlns="' << XML_NS << '" id="' << (index + 1).to_s << '" name="' << @name << '" displayName="' << @name.gsub(/\s/, "_") << '" '
          str << 'ref="' << @ref << '" totalsRowShown="0">'
          str << '<autoFilter ref="' << @ref << '"/>'
          str << '<tableColumns count="' << header_cells.length.to_s << '">'
          header_cells.each_with_index do |cell, idx|
            str << '<tableColumn id ="' << (idx + 1).to_s << '" name="' << cell.clean_value << '"/>'
          end
          str << "</tableColumns>"
          table_style_info.to_xml_string(str)
          str << "</table>"
        end

        private

        # @return [Array[Cell]]
        #: () -> Array[Cell]
        def header_cells
          header = @ref.gsub(/^(\w+?)(\d+):(\w+?)\d+$/, '\1\2:\3\2')
          res = @sheet[header]
          res.is_a?(Array) ? res : [res]
        end
      end

      # Collection of tables in a worksheet.
      class Tables < SimpleTypedList
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          super(Table)
          @worksheet = worksheet
        end

        # @return [Array[Relationship]]
        #: () -> Array[Relationship]
        def relationships
          return [] if empty?

          map { |table| Relationship.new(table, TABLE_R, "../#{table.pn}") }
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<tableParts count='#{size}'>"
          each { |table| str << "<tablePart r:id='#{table.rId}'/>" }
          str << "</tableParts>"
        end
      end
    end
  end
end
