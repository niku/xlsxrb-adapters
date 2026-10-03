# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # A BookView defines the display properties for a workbook.
      class WorkbookView
        include SerializedAttributes
        include OptionsParser
        include Accessors

        unsigned_int_attr_accessor :x_window, :y_window, :window_width, :window_height,
                                   :tab_ratio, :first_sheet, :active_tab

        boolean_attr_accessor :minimized, :show_horizontal_scroll, :show_vertical_scroll,
                              :show_sheet_tabs, :auto_filter_date_grouping

        serializable_attributes :visibility, :minimized,
                                :show_horizontal_scroll, :show_vertical_scroll,
                                :show_sheet_tabs, :tab_ratio, :first_sheet, :active_tab,
                                :x_window, :y_window, :window_width, :window_height,
                                :auto_filter_date_grouping

        # @return [Symbol, nil]
        attr_reader :visibility

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) ?{ (WorkbookView) -> void } -> void
        def initialize(options = {})
          parse_options(options)
          yield self if block_given?
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def visibility=(v)
          Caxlsx.validate_view_visibility(v)
          @visibility = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<workbookView "
          serialized_attributes(str)
          str << "></workbookView>"
        end
      end

      # Collection of WorkbookView objects.
      class WorkbookViews < SimpleTypedList
        #: () -> void
        def initialize
          super(WorkbookView)
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<bookViews>"
          each { |view| view.to_xml_string(str) }
          str << "</bookViews>"
        end
      end

      # Defined name in workbook.
      class DefinedName
        include SerializedAttributes
        include OptionsParser
        include Accessors

        string_attr_accessor :short_cut_key, :status_bar, :help, :description, :custom_menu, :comment, :name, :formula
        boolean_attr_accessor :workbook_parameter, :publish_to_server, :xlm, :vb_proceedure, :function, :hidden

        serializable_attributes :short_cut_key, :status_bar, :help, :description, :custom_menu, :comment,
                                :workbook_parameter, :publish_to_server, :xlm, :vb_proceedure, :function, :hidden, :local_sheet_id

        # @return [Integer, nil]
        attr_reader :local_sheet_id

        # @param formula [String]
        # @param options [Hash{Symbol => untyped}]
        #: (String formula, ?Hash[Symbol, untyped] options) -> void
        def initialize(formula, options = {})
          @formula = formula
          parse_options(options)
        end

        # @param value [Integer]
        # @return [Integer]
        #: (Integer value) -> Integer
        def local_sheet_id=(value)
          Caxlsx.validate_unsigned_int(value)
          @local_sheet_id = value
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          raise ArgumentError, "you must specify the name for this defined name." unless name

          str << '<definedName name="' << name << '" '
          serialized_attributes(str)
          str << ">" << @formula.to_s << "</definedName>"
        end
      end

      # Collection of DefinedName objects.
      class DefinedNames < SimpleTypedList
        #: () -> void
        def initialize
          super(DefinedName)
        end

        # @param formula [String]
        # @param options [Hash{Symbol => untyped}]
        # @return [DefinedName]
        #: (String formula, Hash[Symbol, untyped] options) -> DefinedName
        def add_defined_name(formula, options)
          self << DefinedName.new(formula, options)
          last
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<definedNames>"
          each { |defined_name| defined_name.to_xml_string(str) }
          str << "</definedNames>"
        end
      end

      # Shared strings table serialization.
      class SharedStringsTable
        # @return [Integer]
        attr_reader :count

        # @return [Hash{untyped => Integer}]
        attr_reader :unique_cells

        # @return [Symbol]
        attr_reader :xml_space

        # @param cells [Array[Cell]]
        # @param xml_space [Symbol]
        #: (Array[untyped] cells, ?Symbol xml_space) -> void
        def initialize(cells, xml_space = :preserve)
          @index = 0
          @xml_space = xml_space
          @unique_cells = {}
          @shared_xml_string = +""
          shareable_cells = cells.flatten.select { |cell| cell.is_a?(Cell) && (cell.plain_string? || cell.contains_rich_text?) }
          @count = shareable_cells.size
          resolve(shareable_cells)
        end

        # @return [Integer]
        #: () -> Integer
        def unique_count
          @unique_cells.size
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          Caxlsx.sanitize(@shared_xml_string)
          str << '<?xml version="1.0" encoding="UTF-8"?><sst xmlns="' << XML_NS << '"'
          str << ' count="' << @count.to_s << '" uniqueCount="' << unique_count.to_s << '"'
          str << ' xml:space="' << xml_space.to_s << '">' << @shared_xml_string << "</sst>"
        end

        private

        # @param cells [Array[Cell]]
        # @return [void]
        #: (Array[Cell] cells) -> void
        def resolve(cells)
          cells.each do |cell|
            cell_hash = cell.value
            if (idx = @unique_cells[cell_hash])
              cell.ssti = idx
            else
              cell.ssti = @index
              @shared_xml_string << "<si>" << CellSerializer.run_xml_string(cell) << "</si>"
              @unique_cells[cell_hash] = @index
              @index += 1
            end
          end
        end
      end

      # Workbook representation.
      class Workbook
        BOLD_FONT_MULTIPLIER = 1.5
        FONT_SCALE_DIVISOR = 10.0

        # @return [SimpleTypedList]
        attr_reader :worksheets

        # @return [SimpleTypedList]
        attr_reader :drawings

        # @return [SimpleTypedList]
        attr_reader :charts

        # @return [SimpleTypedList]
        attr_reader :images

        # @return [SimpleTypedList]
        attr_reader :tables

        # @return [SimpleTypedList]
        attr_reader :pivot_tables

        # @return [Boolean, nil]
        attr_reader :use_shared_strings

        # @return [Boolean, nil]
        attr_reader :is_reversed

        # @return [Boolean]
        attr_reader :escape_formulas

        # @return [Boolean]
        attr_reader :secure_formulas

        # @return [Boolean]
        attr_reader :use_autowidth

        # @return [Float]
        attr_reader :bold_font_multiplier

        # @return [Float]
        attr_reader :font_scale_divisor

        # @return [Boolean, nil]
        attr_accessor :styles_applied

        @date1904 = false

        class << self
          # @return [Boolean]
          #: () -> bool
          def date1904 # rubocop:disable Style/TrivialAccessors
            @date1904
          end

          # @param v [Boolean]
          # @return [Boolean]
          #: (bool v) -> bool
          def date1904=(v)
            Caxlsx.validate_boolean(v)
            @date1904 = v
          end
        end

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) ?{ (Workbook) -> void } -> void
        def initialize(options = {})
          @styles = Styles.new
          @worksheets = SimpleTypedList.new(Worksheet)
          @drawings = SimpleTypedList.new(Drawing)
          @charts = SimpleTypedList.new(Chart)
          @images = SimpleTypedList.new(Pic)
          @tables = SimpleTypedList.new(Table)
          @pivot_tables = SimpleTypedList.new(PivotTable)
          @use_autowidth = true
          @bold_font_multiplier = BOLD_FONT_MULTIPLIER
          @font_scale_divisor = FONT_SCALE_DIVISOR
          @use_shared_strings = false
          @is_reversed = false
          @theme = nil
          @views = nil
          @defined_names = nil
          @styled_cells = nil
          @styles_applied = false

          self.escape_formulas = options[:escape_formulas].nil? ? Caxlsx.escape_formulas : options[:escape_formulas]
          self.date1904 = !options[:date1904].nil? && options[:date1904]
          yield self if block_given?
        end

        # @return [Boolean]
        #: () -> bool
        def date1904
          self.class.date1904
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def date1904=(v)
          self.class.date1904 = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def use_shared_strings=(v)
          Caxlsx.validate_boolean(v)
          @use_shared_strings = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def is_reversed=(v)
          Caxlsx.validate_boolean(v)
          @is_reversed = v
        end

        # @param value [Boolean]
        # @return [Boolean]
        #: (bool value) -> bool
        def escape_formulas=(value)
          Caxlsx.validate_boolean(value)
          @escape_formulas = value
        end

        # @param value [Boolean]
        # @return [Boolean]
        #: (bool value) -> bool
        def secure_formulas=(value)
          Caxlsx.validate_boolean(value)
          @secure_formulas = value
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (?bool v) -> bool
        def use_autowidth=(v = true)
          Caxlsx.validate_boolean(v)
          @use_autowidth = v
        end

        # @param v [Float]
        # @return [Float]
        #: (Float v) -> Float
        def bold_font_multiplier=(v)
          Caxlsx.validate_float(v)
          @bold_font_multiplier = v
        end

        # @param v [Float]
        # @return [Float]
        #: (Float v) -> Float
        def font_scale_divisor=(v)
          Caxlsx.validate_float(v)
          @font_scale_divisor = v
        end

        # @return [WorkbookViews]
        #: () -> WorkbookViews
        def views
          @views ||= WorkbookViews.new
        end

        # @return [DefinedNames]
        #: () -> DefinedNames
        def defined_names
          @defined_names ||= DefinedNames.new
        end

        # @return [Array[Comments]]
        #: () -> Array[Comments]
        def comments
          worksheets.map(&:comments).compact
        end

        # @return [Styles]
        #: () ?{ (Styles) -> void } -> Styles
        def styles
          yield @styles if block_given?
          @styles
        end

        # @return [Theme]
        #: () -> Theme
        def theme
          @theme ||= Theme.new
        end

        # @return [Set[Cell]]
        #: () -> Set[Cell]
        def styled_cells
          @styled_cells ||= Set.new
        end

        # @return [Boolean]
        #: () -> bool
        def apply_styles
          return false unless @styled_cells

          styled_cells.each do |cell|
            current_style = styles.style_index[cell.style]
            new_style = if current_style
                          Caxlsx.hash_deep_merge(current_style, cell.raw_style)
                        else
                          cell.raw_style
                        end
            cell.style = styles.add_style(new_style)
          end
          self.styles_applied = true
        end

        # @param name [String]
        # @return [Worksheet, nil]
        #: (String name) -> Worksheet?
        def sheet_by_name(name)
          encoded_name = Caxlsx.coder.encode(name)
          @worksheets.find { |sheet| sheet.name == encoded_name }
        end

        # @param index [Integer]
        # @param options [Hash{Symbol => untyped}]
        # @return [Worksheet]
        #: (?Integer index, ?Hash[Symbol, untyped] options) ?{ (Worksheet) -> void } -> Worksheet
        def insert_worksheet(index = 0, options = {})
          worksheet = Worksheet.new(self, options)
          @worksheets.delete_at(@worksheets.size - 1)
          @worksheets.insert(index, worksheet)
          yield worksheet if block_given?
          worksheet
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Worksheet]
        #: (?Hash[Symbol, untyped] options) ?{ (Worksheet) -> void } -> Worksheet
        def add_worksheet(options = {})
          worksheet = Worksheet.new(self, options)
          yield worksheet if block_given?
          worksheet
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [WorkbookView]
        #: (?Hash[Symbol, untyped] options) -> WorkbookView
        def add_view(options = {})
          views << WorkbookView.new(options)
          views.last
        end

        # @param formula [String]
        # @param options [Hash{Symbol => untyped}]
        # @return [DefinedName]
        #: (String formula, Hash[Symbol, untyped] options) -> DefinedName
        def add_defined_name(formula, options)
          defined_names << DefinedName.new(formula, options)
          defined_names.last
        end

        # @return [Relationships]
        #: () -> Relationships
        def relationships
          r = Relationships.new
          @worksheets.each do |sheet|
            r << Relationship.new(sheet, WORKSHEET_R, format(WORKSHEET_PN, r.size + 1))
          end
          pivot_tables.each_with_index do |pivot_table, idx|
            r << Relationship.new(pivot_table.cache_definition, PIVOT_TABLE_CACHE_DEFINITION_R, format(PIVOT_TABLE_CACHE_DEFINITION_PN, idx + 1))
          end
          r << Relationship.new(self, STYLES_R, STYLES_PN)
          r << Relationship.new(self, THEME_R, THEME_PN)
          r << Relationship.new(self, SHARED_STRINGS_R, SHARED_STRINGS_PN) if use_shared_strings
          r
        end

        # @return [SharedStringsTable]
        #: () -> SharedStringsTable
        def shared_strings
          SharedStringsTable.new(worksheets.map(&:cells), xml_space)
        end

        # @return [Symbol]
        #: () -> Symbol
        def xml_space
          @xml_space ||= :preserve
        end

        # @param space [Symbol]
        # @return [Symbol]
        #: (Symbol space) -> Symbol
        def xml_space=(space)
          RestrictionValidator.validate(:xml_space, %i[preserve default], space)
          @xml_space = space
        end

        # @param cell_def [String]
        # @return [Cell, Array[Cell], nil]
        #: (String cell_def) -> (Cell | Array[Cell] | nil)
        def [](cell_def)
          sheet_name = cell_def.split("!").first if cell_def.include?("!")
          worksheet = worksheets.find { |s| s.name == sheet_name }
          raise ArgumentError, "Unknown Sheet" unless sheet_name && worksheet.is_a?(Worksheet)

          worksheet[cell_def.gsub(/.+!/, "")]
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          add_worksheet(name: "Sheet1") if worksheets.empty?
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << '<workbook xmlns="' << XML_NS << '" xmlns:r="' << XML_NS_R << '">'
          str << '<workbookPr date1904="' << self.class.date1904.to_s << '"/>'
          views.to_xml_string(str)
          str << "<sheets>"
          if is_reversed
            worksheets.reverse_each { |sheet| sheet.to_sheet_node_xml_string(str) }
          else
            worksheets.each { |sheet| sheet.to_sheet_node_xml_string(str) }
          end
          str << "</sheets>"
          defined_names.to_xml_string(str)
          unless pivot_tables.empty?
            str << "<pivotCaches>"
            pivot_tables.each do |pivot_table|
              str << '<pivotCache cacheId="' << pivot_table.cache_definition.cache_id.to_s << '" r:id="' << pivot_table.cache_definition.rId << '"/>'
            end
            str << "</pivotCaches>"
          end
          str << "</workbook>"
        end

        # Converts Workbook to an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          apply_styles

          elements_sheets = @worksheets.map(&:to_xlsxrb)
          styles_hash = @styles.to_xlsxrb_hash

          dns = if @defined_names && !@defined_names.empty?
                  @defined_names.map do |dn|
                    {
                      name: dn.name,
                      value: dn.formula,
                      local_sheet_id: dn.local_sheet_id,
                      hidden: dn.hidden
                    }.compact
                  end
                end
          merged_unmapped = {}
          merged_unmapped[:workbook_properties] = { date1904: true } if date1904

          sst_strings = if use_shared_strings
                          shared_strings.unique_cells.keys.map(&:to_s)
                        else
                          []
                        end

          Xlsxrb::Elements::Workbook.new(
            sheets: elements_sheets,
            shared_strings: sst_strings,
            styles: styles_hash,
            unmapped_data: merged_unmapped,
            defined_names: dns
          )
        end
      end
    end
  end
end
