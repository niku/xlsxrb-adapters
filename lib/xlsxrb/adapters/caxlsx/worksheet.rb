# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # BorderCreator helper for drawing borders on cell ranges.
      class BorderCreator
        # @param worksheet [Worksheet]
        # @param cells [Array[Cell]]
        # @param edges [Symbol, Array[Symbol], nil]
        # @param style [Symbol, nil]
        # @param color [String, nil]
        #: (worksheet: Worksheet, cells: Array[Cell], ?edges: Symbol | Array[Symbol] | nil, ?style: Symbol | nil, ?color: String | nil) -> void
        def initialize(worksheet:, cells:, edges: nil, style: nil, color: nil)
          @worksheet = worksheet
          @cells = cells
          @style = style || :thin
          @color = color || "000000"

          @edges = if edges.nil? || edges == :all
                     Border::EDGES
                   elsif !edges.is_a?(Array)
                     raise ArgumentError, "Invalid edges provided, #{edges}"
                   else
                     edges.map { |x| x&.to_sym }.uniq
                   end

          return if (@edges - Border::EDGES).empty?

          raise ArgumentError, "Invalid edges provided, #{edges}"
        end

        # @return [void]
        #: () -> void
        def draw
          if @cells.size == 1
            @worksheet.add_style(
              first_cell,
              { border: { style: @style, color: @color, edges: @edges } }
            )
          else
            @edges.each do |edge|
              @worksheet.add_style(
                border_cells[edge],
                { border: { style: @style, color: @color, edges: [edge] } }
              )
            end
          end
        end

        private

        # @return [Hash{Symbol => String}]
        #: () -> Hash[Symbol, String]
        def border_cells
          {
            top: "#{first_cell}:#{last_col}#{first_row}",
            right: "#{last_col}#{first_row}:#{last_cell}",
            bottom: "#{first_col}#{last_row}:#{last_cell}",
            left: "#{first_cell}:#{first_col}#{last_row}"
          }
        end

        # @return [String]
        #: () -> String
        def first_cell
          @first_cell ||= @cells.first&.r || "A1"
        end

        # @return [String]
        #: () -> String
        def last_cell
          @last_cell ||= @cells.last&.r || "A1"
        end

        # @return [String]
        #: () -> String
        def first_row
          @first_row ||= first_cell[/\d+/] || "1"
        end

        # @return [String]
        #: () -> String
        def first_col
          @first_col ||= first_cell[/\D+/] || "A"
        end

        # @return [String]
        #: () -> String
        def last_row
          @last_row ||= last_cell[/\d+/] || "1"
        end

        # @return [String]
        #: () -> String
        def last_col
          @last_col ||= last_cell[/\D+/] || "A"
        end
      end

      # Dimension node for worksheet.
      class Dimension
        class << self
          # @return [String]
          #: () -> String
          def default_first
            @default_first ||= "A1"
          end

          # @return [String]
          #: () -> String
          def default_last
            @default_last ||= "AA200"
          end
        end

        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          @worksheet = worksheet
        end

        # @return [String]
        #: () -> String
        def sqref
          "#{first_cell_reference}:#{last_cell_reference}"
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if worksheet.rows.empty?

          str << '<dimension ref="' << sqref << '"></dimension>'
        end

        # @return [String]
        #: () -> String
        def first_cell_reference
          dimension_reference(worksheet.rows.first&.first, Dimension.default_first)
        end

        # @return [String]
        #: () -> String
        def last_cell_reference
          dimension_reference(worksheet.rows.last&.last, Dimension.default_last)
        end

        private

        # @param cell [Cell, untyped]
        # @param default [String]
        # @return [String]
        #: (untyped cell, String default) -> String
        def dimension_reference(cell, default)
          return default unless cell.respond_to?(:r)

          cell.r
        end
      end

      # OutlinePr properties.
      class OutlinePr
        include OptionsParser
        include Accessors
        include SerializedAttributes

        serializable_attributes :summary_below, :summary_right, :apply_styles
        boolean_attr_accessor :summary_below, :summary_right, :apply_styles

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          parse_options(options)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<outlinePr "
          serialized_attributes(str)
          str << "/>"
        end
      end

      # Sheet formatting properties.
      class SheetFormatPr
        include SerializedAttributes
        include OptionsParser
        include Accessors

        serializable_attributes :base_col_width, :default_col_width, :default_row_height,
                                :custom_height, :zero_height, :thick_top, :thick_bottom,
                                :outline_level_row, :outline_level_col

        float_attr_accessor :default_col_width, :default_row_height
        boolean_attr_accessor :custom_height, :zero_height, :thick_top, :thick_bottom
        unsigned_int_attr_accessor :base_col_width, :outline_level_row, :outline_level_col

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @base_col_width = 8
          @default_row_height = 18
          parse_options(options)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<sheetFormatPr "
          serialized_attributes(str)
          str << "/>"
        end
      end

      # SheetCalcPr calculation properties.
      class SheetCalcPr
        include OptionsParser
        include SerializedAttributes
        include Accessors

        boolean_attr_accessor :full_calc_on_load
        serializable_attributes :full_calc_on_load

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @full_calc_on_load = true
          parse_options(options)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<sheetCalcPr "
          serialized_attributes(str)
          str << "/>"
        end
      end

      # SheetPr sheet properties element.
      class SheetPr
        include OptionsParser
        include Accessors
        include SerializedAttributes

        serializable_attributes :sync_horizontal, :sync_vertical, :transition_evaluation,
                                :transition_entry, :published, :filter_mode,
                                :enable_format_conditions_calculation, :code_name, :sync_ref

        boolean_attr_accessor :sync_horizontal, :sync_vertical, :transition_evaluation,
                              :transition_entry, :published, :filter_mode,
                              :enable_format_conditions_calculation
        string_attr_accessor :code_name, :sync_ref

        # @return [Worksheet]
        attr_reader :worksheet

        # @return [Color, nil]
        attr_reader :tab_color

        # @param worksheet [Worksheet]
        # @param options [Hash{Symbol => untyped}]
        #: (Worksheet worksheet, ?Hash[Symbol, untyped] options) -> void
        def initialize(worksheet, options = {})
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          @worksheet = worksheet
          @outline_pr = nil
          @page_setup_pr = nil
          @tab_color = nil
          parse_options(options)
        end

        # @return [PageSetUpPr]
        #: () -> PageSetUpPr
        def page_setup_pr
          @page_setup_pr ||= PageSetUpPr.new
        end

        # @return [OutlinePr]
        #: () -> OutlinePr
        def outline_pr
          @outline_pr ||= OutlinePr.new
        end

        # @param v [String]
        # @return [void]
        #: (String v) -> void
        def tab_color=(v)
          @tab_color = Color.new(rgb: v)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          update_properties
          str << "<sheetPr "
          serialized_attributes(str)
          str << ">"
          tab_color&.to_xml_string(str, "tabColor")
          outline_pr.to_xml_string(str) if @outline_pr
          page_setup_pr.to_xml_string(str)
          str << "</sheetPr>"
        end

        private

        # @return [void]
        #: () -> void
        def update_properties
          page_setup_pr.fit_to_page = worksheet.fit_to_page?
          return if worksheet.auto_filter.columns.empty?

          self.filter_mode = true
          self.enable_format_conditions_calculation = true
        end
      end

      # MergedCells collection.
      class MergedCells < SimpleTypedList
        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          super(String)
        end

        # @param cells [String, Array[Cell]]
        # @return [void]
        #: (String | Array[Cell] cells) -> void
        def add(cells)
          if cells.is_a?(String)
            self << cells
          elsif cells.is_a?(Array)
            self << Caxlsx.cell_range(cells, false)
          end
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<mergeCells count='#{size}'>"
          each { |merged_cell| str << "<mergeCell ref='#{merged_cell}'></mergeCell>" }
          str << "</mergeCells>"
        end
      end

      # WorksheetDrawing connector.
      class WorksheetDrawing
        # @return [Worksheet]
        attr_reader :worksheet

        # @return [Drawing, nil]
        attr_reader :drawing

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          @worksheet = worksheet
          @drawing = nil
        end

        # @param chart_type [Class]
        # @param options [Hash{Symbol => untyped}]
        # @return [Chart]
        #: (Class chart_type, Hash[Symbol, untyped] options) -> Chart
        def add_chart(chart_type, options)
          @drawing ||= Drawing.new(worksheet)
          drawing = @drawing
          raise "Drawing missing" unless drawing

          drawing.add_chart(chart_type, options)
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Pic]
        #: (Hash[Symbol, untyped] options) -> Pic
        def add_image(options)
          @drawing ||= Drawing.new(worksheet)
          drawing = @drawing
          raise "Drawing missing" unless drawing

          drawing.add_image(options)
        end

        # @return [Boolean]
        #: () -> bool
        def has_drawing?
          @drawing.is_a?(Drawing)
        end

        # @return [Relationship, nil]
        #: () -> Relationship?
        def relationship
          return unless has_drawing?

          drawing = @drawing
          return unless drawing

          Relationship.new(self, DRAWING_R, "../#{drawing.pn}")
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return unless has_drawing?

          rel = relationship
          return unless rel

          str << "<drawing r:id='#{rel.Id}'/>"
        end
      end

      # SheetData serializer.
      class SheetData
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          @worksheet = worksheet
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<sheetData>"
          worksheet.rows.each_with_index do |row, idx|
            row.to_xml_string(idx, str)
          end
          str << "</sheetData>"
        end
      end

      # The Worksheet class represents a worksheet in the workbook.
      class Worksheet
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :sheet_id, :state

        # @return [Workbook]
        attr_reader :workbook

        # @return [Integer]
        attr_reader :sheet_id

        # @return [Boolean]
        attr_reader :escape_formulas

        # @return [Boolean]
        attr_reader :secure_formulas

        # @return [Boolean, nil]
        attr_accessor :preserve_spaces

        # @param wb [Workbook]
        # @param options [Hash{Symbol => untyped}]
        #: (Workbook wb, ?Hash[Symbol, untyped] options) ?{ (Worksheet) -> void } -> void
        def initialize(wb, options = {})
          self.workbook = wb
          @sheet_protection = nil
          @rows = SimpleTypedList.new(Row)
          @page_margins = nil
          @page_setup = nil
          @print_options = nil
          @header_footer = nil
          @row_breaks = RowBreaks.new
          @col_breaks = ColBreaks.new
          @auto_filter = nil
          @column_info = nil
          @sheet_format_pr = nil
          @sheet_view = nil
          @sheet_calc_pr = nil
          @sheet_pr = nil
          @dimension = nil
          @worksheet_drawing = nil
          @worksheet_comments = nil
          @tables = nil
          @pivot_tables = nil
          @hyperlinks = nil
          @protected_ranges = nil
          @conditional_formattings = nil
          @data_validations = nil
          @merged_cells = nil

          initialize_page_options(options)
          parse_options(options)
          @escape_formulas = wb.escape_formulas unless defined?(@escape_formulas)
          @secure_formulas = wb.secure_formulas unless defined?(@secure_formulas)
          @workbook.worksheets << self
          @sheet_id = index + 1
          yield self if block_given?
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [void]
        #: (Hash[Symbol, untyped] options) -> void
        def initialize_page_options(options)
          @page_margins = PageMargins.new(options[:page_margins]) if options[:page_margins]
          @page_setup = PageSetup.new(options[:page_setup]) if options[:page_setup]
          @print_options = PrintOptions.new(options[:print_options]) if options[:print_options]
          @header_footer = HeaderFooter.new(options[:header_footer]) if options[:header_footer]
        end

        # @return [String]
        #: () -> String
        def name
          @name ||= "Sheet#{index + 1}"
        end

        # @param name [String]
        # @return [void]
        #: (String name) -> void
        def name=(name)
          validate_sheet_name(name)
          @name = name
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

        # @param sheet_state [Symbol]
        # @return [void]
        #: (Symbol sheet_state) -> void
        def state=(sheet_state)
          RestrictionValidator.validate(:worksheet_state, %i[visible hidden very_hidden], sheet_state)
          @state = sheet_state
        end

        # @return [Symbol]
        #: () -> Symbol
        def state
          @state ||= :visible
        end

        # @return [SheetCalcPr]
        #: () -> SheetCalcPr
        def sheet_calc_pr
          @sheet_calc_pr ||= SheetCalcPr.new
        end

        # @return [SheetProtection]
        #: () ?{ (SheetProtection) -> void } -> SheetProtection
        def sheet_protection
          @sheet_protection ||= SheetProtection.new
          yield @sheet_protection if block_given?
          @sheet_protection
        end

        # @return [SheetView]
        #: () ?{ (SheetView) -> void } -> SheetView
        def sheet_view
          @sheet_view ||= SheetView.new
          yield @sheet_view if block_given?
          @sheet_view
        end

        # @return [SheetFormatPr]
        #: () -> SheetFormatPr
        def sheet_format_pr
          @sheet_format_pr ||= SheetFormatPr.new
        end

        # @return [Tables]
        #: () -> Tables
        def tables
          @tables ||= Tables.new(self)
        end

        # @return [PivotTables]
        #: () -> PivotTables
        def pivot_tables
          @pivot_tables ||= PivotTables.new(self)
        end

        # @return [ColBreaks]
        #: () -> ColBreaks
        def col_breaks
          @col_breaks ||= ColBreaks.new
        end

        # @return [RowBreaks]
        #: () -> RowBreaks
        def row_breaks
          @row_breaks ||= RowBreaks.new
        end

        # @return [WorksheetHyperlinks]
        #: () -> WorksheetHyperlinks
        def hyperlinks
          @hyperlinks ||= WorksheetHyperlinks.new(self)
        end

        # @return [Comments, nil]
        #: () -> Comments?
        def comments
          worksheet_comments.comments if worksheet_comments.has_comments?
        end

        # @return [SimpleTypedList]
        #: () -> SimpleTypedList
        def rows
          @rows ||= SimpleTypedList.new(Row)
        end

        # @yieldparam row_idx [Integer]
        # @yieldparam col_idx [Integer]
        # @return [Array[Array[Cell]]]
        #: () ?{ (Integer, Integer) -> void } -> Array[Array[Cell]]
        def cols(&)
          @rows.transpose(&)
        end

        # @return [AutoFilter]
        #: () -> AutoFilter
        def auto_filter
          @auto_filter ||= AutoFilter.new(self)
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def auto_filter=(v)
          DataTypeValidator.validate("#{self.class}.auto_filter", String, v)
          auto_filter.range = v
        end

        # @return [Boolean]
        #: () -> bool
        def fit_to_page?
          return false unless @page_setup

          page_setup.fit_to_page?
        end

        # @return [Cols]
        #: () -> Cols
        def column_info
          @column_info ||= Cols.new(self)
        end

        # @return [PageMargins]
        #: () ?{ (PageMargins) -> void } -> PageMargins
        def page_margins
          @page_margins ||= PageMargins.new
          yield @page_margins if block_given?
          @page_margins
        end

        # @return [PageSetup]
        #: () ?{ (PageSetup) -> void } -> PageSetup
        def page_setup
          @page_setup ||= PageSetup.new
          yield @page_setup if block_given?
          @page_setup
        end

        # @return [PrintOptions]
        #: () ?{ (PrintOptions) -> void } -> PrintOptions
        def print_options
          @print_options ||= PrintOptions.new
          yield @print_options if block_given?
          @print_options
        end

        # @return [HeaderFooter]
        #: () ?{ (HeaderFooter) -> void } -> HeaderFooter
        def header_footer
          @header_footer ||= HeaderFooter.new
          yield @header_footer if block_given?
          @header_footer
        end

        # @return [Array[Cell]]
        #: () -> Array[Cell]
        def cells
          rows.flatten
        end

        # @param cells [String, Array[Cell]]
        # @return [void]
        #: (String | Array[Cell] cells) -> void
        def merge_cells(cells)
          merged_cells.add(cells)
        end

        # @param cells [String, Array[Cell]]
        # @return [ProtectedRange]
        #: (String | Array[Cell] cells) -> ProtectedRange
        def protect_range(cells)
          protected_ranges.add_range(cells)
        end

        # @return [Dimension]
        #: () -> Dimension
        def dimension
          @dimension ||= Dimension.new(self)
        end

        # @return [SheetPr]
        #: () -> SheetPr
        def sheet_pr
          @sheet_pr ||= SheetPr.new(self)
        end

        # @return [String]
        #: () -> String
        def pn
          format(WORKSHEET_PN, index + 1)
        end

        # @return [String]
        #: () -> String
        def rels_pn
          format(WORKSHEET_RELS_PN, index + 1)
        end

        # @return [String]
        #: () -> String
        def rId
          @workbook.relationships.for(self).Id
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @workbook.worksheets.index(self) || 0
        end

        # @return [Drawing, nil]
        #: () -> Drawing?
        def drawing
          worksheet_drawing.drawing
        end

        # @param values [Array[untyped]]
        # @param options [Hash{Symbol => untyped}]
        # @return [Row]
        #: (?Array[untyped] values, ?Hash[Symbol, untyped] options) ?{ (Row) -> void } -> Row
        def add_row(values = [], options = {})
          row = Row.new(self, values, options)
          update_column_info(row, options.delete(:widths))
          yield row if block_given?
          row
        end
        alias << add_row

        # @param cells [String]
        # @param rules [Array[untyped], Hash{Symbol => untyped}]
        # @return [ConditionalFormattings]
        #: (String cells, Array[untyped] | Hash[Symbol, untyped] rules) -> ConditionalFormattings
        def add_conditional_formatting(cells, rules)
          cf = ConditionalFormatting.new(sqref: cells)
          cf.add_rules(rules)
          conditional_formattings << cf
          conditional_formattings
        end

        # @param cells [String]
        # @param data_validation [Hash{Symbol => untyped}]
        # @return [void]
        #: (String cells, Hash[Symbol, untyped] data_validation) -> void
        def add_data_validation(cells, data_validation)
          dv = DataValidation.new(data_validation)
          dv.sqref = cells
          data_validations << dv
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [WorksheetHyperlink]
        #: (?Hash[Symbol, untyped] options) -> WorksheetHyperlink
        def add_hyperlink(options = {})
          hyperlinks.add(options)
        end

        # @param chart_type [Class]
        # @param options [Hash{Symbol => untyped}]
        # @return [Chart]
        #: (Class chart_type, ?Hash[Symbol, untyped] options) ?{ (Chart) -> void } -> Chart
        def add_chart(chart_type, options = {})
          chart = worksheet_drawing.add_chart(chart_type, options)
          yield chart if block_given?
          chart
        end

        # @param ref [String]
        # @param options [Hash{Symbol => untyped}]
        # @return [Table]
        #: (String ref, ?Hash[Symbol, untyped] options) ?{ (Table) -> void } -> Table
        def add_table(ref, options = {})
          tables << Table.new(ref, self, options)
          yield tables.last if block_given?
          tables.last
        end

        # @param ref [String]
        # @param range [String]
        # @param options [Hash{Symbol => untyped}]
        # @return [PivotTable]
        #: (String ref, String range, ?Hash[Symbol, untyped] options) ?{ (PivotTable) -> void } -> PivotTable
        def add_pivot_table(ref, range, options = {})
          pivot_tables << PivotTable.new(ref, range, self, options)
          yield pivot_tables.last if block_given?
          pivot_tables.last
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Comment]
        #: (?Hash[Symbol, untyped] options) -> Comment
        def add_comment(options = {})
          worksheet_comments.add_comment(options)
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Pic]
        #: (?Hash[Symbol, untyped] options) ?{ (Pic) -> void } -> Pic
        def add_image(options = {})
          image = worksheet_drawing.add_image(options)
          yield image if block_given?
          image
        end

        # @param cell [String, Cell]
        # @return [void]
        #: (String | Cell cell) -> void
        def add_page_break(cell)
          DataTypeValidator.validate(:worksheet_page_break, [String, Cell], cell)
          col_idx, row_idx = if cell.is_a?(String)
                               Caxlsx.name_to_indices(cell)
                             else
                               cell.pos
                             end
          col_breaks.add_break(id: col_idx) if col_idx.positive?
          row_breaks.add_break(id: row_idx)
        end

        # @param widths [Array[Numeric, nil]]
        # @return [void]
        #: (*(Numeric | nil) widths) -> void
        def column_widths(*widths)
          widths.each_with_index do |value, idx|
            next if value.nil?

            Caxlsx.validate_unsigned_numeric(value)
            find_or_create_column_info(idx).width = value
          end
        end

        # @param styles [Array[Integer, nil]]
        # @return [void]
        #: (*(Integer | nil) styles) -> void
        def column_styles(*styles)
          styles.each_with_index do |style, idx|
            next if style.nil?

            Caxlsx.validate_unsigned_int(style)
            find_or_create_column_info(idx).style = style
          end
        end

        # @param index [Integer]
        # @param style [Integer]
        # @param options [Hash{Symbol => untyped}]
        # @return [void]
        #: (Integer index, Integer style, ?Hash[Symbol, untyped] options) -> void
        def col_style(index, style, options = {})
          offset = options.delete(:row_offset) || 0
          cells_list = @rows[(offset..)].map { |row| row[index] }.flatten.compact
          cells_list.each { |cell| cell.style = style }
        end

        # @param index [Integer]
        # @param style [Integer]
        # @param options [Hash{Symbol => untyped}]
        # @return [void]
        #: (Integer index, Integer style, ?Hash[Symbol, untyped] options) -> void
        def row_style(index, style, options = {})
          offset = options.delete(:col_offset) || 0
          cells_list = cols[(offset..)].map { |col| col[index] }.flatten.compact
          cells_list.each { |cell| cell.style = style }
        end

        # @param cell_refs [String, Array[String]]
        # @param styles [Array[untyped]]
        # @return [void]
        #: (String | Array[String] cell_refs, *untyped styles) -> void
        def add_style(cell_refs, *styles)
          cell_refs = [cell_refs] unless cell_refs.is_a?(Array)

          cell_refs.each do |cell_ref|
            item = self[cell_ref]
            cells_list = item.is_a?(Array) ? item : [item]
            cells_list.each do |cell|
              styles.each do |style|
                cell.add_style(style)
              end
            end
          end
        end

        # @param cell_refs [String, Array[String]]
        # @param options [Hash{Symbol => untyped}, Array[Symbol], Symbol, nil]
        # @return [void]
        #: (String | Array[String] cell_refs, ?Hash[Symbol, untyped] | Array[Symbol] | Symbol | nil options) -> void
        def add_border(cell_refs, options = nil)
          border_edges = nil
          border_style = nil
          border_color = nil

          if options.is_a?(Hash)
            border_edges = options[:edges]
            border_style = options[:style]
            border_color = options[:color]
          else
            border_edges = options
          end

          cell_refs = [cell_refs] unless cell_refs.is_a?(Array)

          cell_refs.each do |cell_ref|
            item = self[cell_ref]
            cells_list = item.is_a?(Array) ? item : [item]
            BorderCreator.new(worksheet: self, cells: cells_list, edges: border_edges, style: border_style, color: border_color).draw
          end
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_sheet_node_xml_string(str = +"")
          add_autofilter_defined_name_to_workbook
          str << "<sheet "
          serialized_attributes(str)
          str << 'name="' << name << '" '
          str << 'r:id="' << rId << '"></sheet>'
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          add_autofilter_defined_name_to_workbook
          auto_filter.apply if auto_filter.range
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << worksheet_node
          serializable_parts.each do |item|
            item&.to_xml_string(str)
          end
          str << "</worksheet>"
        end

        # @return [Relationships]
        #: () -> Relationships
        def relationships
          r = Relationships.new
          r.concat([tables.relationships,
                    worksheet_comments.relationships,
                    hyperlinks.relationships,
                    worksheet_drawing.relationship,
                    pivot_tables.relationships].flatten.compact)
        end

        # @param cell_def [String, Integer]
        # @return [Cell, Array[Cell], Row, nil]
        #: (String | Integer cell_def) -> (Cell | Array[Cell] | Row | nil)
        def [](cell_def)
          return rows[cell_def] if cell_def.is_a?(Integer)

          parts = cell_def.split(":").map { |part| name_to_cell(part) }
          if parts.size == 1
            parts.first
          elsif parts.size > 2 || parts.first.nil?
            raise ArgumentError, format(ERR_CELL_REFERENCE_INVALID, cell_def)
          else
            first, last = parts
            range(first, last)
          end
        end

        # @param name [String]
        # @return [Cell, nil]
        #: (String name) -> Cell?
        def name_to_cell(name)
          col_idx, row_idx = *Caxlsx.name_to_indices(name)
          r = rows[row_idx]
          r ? r[col_idx] : nil
        end

        # @return [Styles]
        #: () -> Styles
        def styles
          workbook.styles
        end

        # @param start_index [Integer]
        # @param end_index [Integer]
        # @param level [Integer]
        # @param collapsed [Boolean]
        # @return [void]
        #: (Integer start_index, Integer end_index, ?Integer level, ?bool collapsed) -> void
        def outline_level_rows(start_index, end_index, level = 1, collapsed = true)
          outline(rows, start_index..end_index, level, collapsed)
        end

        # @param start_index [Integer]
        # @param end_index [Integer]
        # @param level [Integer]
        # @param collapsed [Boolean]
        # @return [void]
        #: (Integer start_index, Integer end_index, ?Integer level, ?bool collapsed) -> void
        def outline_level_columns(start_index, end_index, level = 1, collapsed = true)
          outline(column_info, start_index..end_index, level, collapsed)
        end

        # @return [ProtectedRanges]
        #: () -> ProtectedRanges
        def protected_ranges
          @protected_ranges ||= ProtectedRanges.new(self)
        end

        # @return [ConditionalFormattings]
        #: () -> ConditionalFormattings
        def conditional_formattings
          @conditional_formattings ||= ConditionalFormattings.new(self)
        end

        # @return [DataValidations]
        #: () -> DataValidations
        def data_validations
          @data_validations ||= DataValidations.new(self)
        end

        # @return [MergedCells]
        #: () -> MergedCells
        def merged_cells
          @merged_cells ||= MergedCells.new(self)
        end

        # @return [String]
        #: () -> String
        def worksheet_node
          "<worksheet xmlns=\"#{XML_NS}\" xmlns:r=\"#{XML_NS_R}\" xml:space=\"#{xml_space}\">"
        end

        # @return [SheetData]
        #: () -> SheetData
        def sheet_data
          @sheet_data ||= SheetData.new(self)
        end

        # @return [WorksheetDrawing]
        #: () -> WorksheetDrawing
        def worksheet_drawing
          @worksheet_drawing ||= WorksheetDrawing.new(self)
        end

        # @return [WorksheetComments]
        #: () -> WorksheetComments
        def worksheet_comments
          @worksheet_comments ||= WorksheetComments.new(self)
        end

        # @param index [Integer]
        # @return [Col]
        #: (Integer index) -> Col
        def find_or_create_column_info(index)
          column_info[index] ||= Col.new(index + 1, index + 1)
        end

        # Converts worksheet to an immutable Xlsxrb::Elements::Worksheet.
        #
        # @return [Xlsxrb::Elements::Worksheet]
        #: () -> Xlsxrb::Elements::Worksheet
        def to_xlsxrb
          sorted_rows = rows.compact.each_with_index.map do |row, r_idx|
            row.to_xlsxrb(r_idx)
          end

          facade_meta = {}
          facade_meta[:merge_cells] = merged_cells.to_a if @merged_cells && !@merged_cells.empty?
          facade_meta[:auto_filter] = auto_filter.range if @auto_filter&.range

          if @page_margins
            facade_meta[:page_margins] = {
              left: @page_margins.left,
              right: @page_margins.right,
              top: @page_margins.top,
              bottom: @page_margins.bottom,
              header: @page_margins.header,
              footer: @page_margins.footer
            }.compact
          end

          if @page_setup
            facade_meta[:page_setup] = {
              orientation: @page_setup.orientation,
              paper_size: @page_setup.paper_size,
              scale: @page_setup.scale,
              fit_to_width: @page_setup.fit_to_width,
              fit_to_height: @page_setup.fit_to_height,
              fit_to_page: @page_setup.fit_to_page
            }.compact
          end

          if @print_options
            facade_meta[:print_options] = {
              grid_lines: @print_options.grid_lines,
              headings: @print_options.headings,
              horizontal_centered: @print_options.horizontal_centered,
              vertical_centered: @print_options.vertical_centered
            }.compact
          end

          if @header_footer
            facade_meta[:header_footer] = {
              different_first: @header_footer.different_first,
              different_odd_even: @header_footer.different_odd_even,
              odd_header: @header_footer.odd_header,
              odd_footer: @header_footer.odd_footer,
              even_header: @header_footer.even_header,
              even_footer: @header_footer.even_footer,
              first_header: @header_footer.first_header,
              first_footer: @header_footer.first_footer
            }.compact
          end

          if @sheet_protection
            facade_meta[:sheet_protection] = {
              sheet: @sheet_protection.sheet,
              objects: @sheet_protection.objects,
              scenarios: @sheet_protection.scenarios,
              format_cells: @sheet_protection.format_cells,
              format_columns: @sheet_protection.format_columns,
              format_rows: @sheet_protection.format_rows,
              insert_columns: @sheet_protection.insert_columns,
              insert_rows: @sheet_protection.insert_rows,
              insert_hyperlinks: @sheet_protection.insert_hyperlinks,
              delete_columns: @sheet_protection.delete_columns,
              delete_rows: @sheet_protection.delete_rows,
              select_locked_cells: @sheet_protection.select_locked_cells,
              sort: @sheet_protection.sort,
              auto_filter: @sheet_protection.auto_filter,
              pivot_tables: @sheet_protection.pivot_tables,
              select_unlocked_cells: @sheet_protection.select_unlocked_cells,
              password: @sheet_protection.password
            }.compact
          end

          if @sheet_view
            facade_meta[:sheet_view] = {
              show_grid_lines: @sheet_view.show_grid_lines,
              show_row_col_headers: @sheet_view.show_row_col_headers,
              show_ruler: @sheet_view.show_ruler,
              view: @sheet_view.view,
              tab_selected: @sheet_view.tab_selected,
              zoom_scale: @sheet_view.zoom_scale
            }.compact
            if @sheet_view.pane
              facade_meta[:freeze_pane] = {
                x_split: @sheet_view.pane.x_split,
                y_split: @sheet_view.pane.y_split,
                top_left_cell: @sheet_view.pane.top_left_cell,
                state: @sheet_view.pane.state
              }.compact
            end
          end

          cf_data = if @conditional_formattings && !@conditional_formattings.empty?
                      @conditional_formattings.flat_map do |cf|
                        sqref = cf.sqref
                        cf.rules.map do |rule|
                          h = {
                            sqref: sqref,
                            type: rule.type,
                            operator: rule.operator,
                            formula: rule.formula,
                            format_id: rule.dxfId,
                            dxf_id: rule.dxfId,
                            priority: rule.priority,
                            stop_if_true: rule.stopIfTrue,
                            above_average: rule.aboveAverage,
                            percent: rule.percent,
                            bottom: rule.bottom,
                            text: rule.text,
                            time_period: rule.timePeriod,
                            rank: rule.rank,
                            std_dev: rule.stdDev,
                            equal_average: rule.equalAverage
                          }
                          h[:color_scale] = rule.color_scale if rule.instance_variable_get(:@color_scale)
                          h[:data_bar] = rule.data_bar if rule.instance_variable_get(:@data_bar)
                          h[:icon_set] = rule.icon_set if rule.instance_variable_get(:@icon_set)
                          h.compact
                        end
                      end
                    else
                      []
                    end

          dv_data = if @data_validations && !@data_validations.empty?
                      @data_validations.map do |dv|
                        {
                          sqref: dv.sqref,
                          formula1: dv.formula1,
                          formula2: dv.formula2,
                          type: dv.type,
                          operator: dv.operator,
                          allow_blank: dv.allowBlank,
                          show_input_message: dv.showInputMessage,
                          show_error_message: dv.showErrorMessage,
                          error_style: dv.errorStyle,
                          error_title: dv.errorTitle,
                          error: dv.error,
                          prompt_title: dv.promptTitle,
                          prompt: dv.prompt
                        }.compact
                      end
                    else
                      []
                    end

          facade_meta[:data_validations] = dv_data if dv_data && !dv_data.empty?
          facade_meta[:conditional_formatting] = cf_data if cf_data && !cf_data.empty?

          if @worksheet_comments&.has_comments?
            comms = @worksheet_comments.comments
            facade_meta[:comments] = comms.map do |c|
              { cell: c.ref, ref: c.ref, author: c.author, text: c.text }
            end
          end

          if @hyperlinks && !@hyperlinks.empty?
            facade_meta[:hyperlinks] = @hyperlinks.map do |h|
              { ref: h.ref, location: h.location, display: h.display, tooltip: h.tooltip, target: h.target }
            end
          end

          if @worksheet_drawing&.has_drawing?
            drw = @worksheet_drawing.drawing
            if drw
              facade_meta[:images] = drw.images.map do |img|
                {
                  image_src: img.image_src,
                  name: img.name,
                  descr: img.descr,
                  width: img.width,
                  height: img.height,
                  start_at: [img.anchor.from.col, img.anchor.from.row]
                }.compact
              end
            end
          end

          cols_data = if @column_info.nil? || @column_info.empty?
                        []
                      else
                        @column_info.map do |col|
                          col_unmapped = {}
                          col_unmapped[:style_index] = col.style if col.style&.positive?
                          Xlsxrb::Elements::Column.new(
                            index: [(col.min || 1) - 1, 0].max,
                            width: col.width,
                            hidden: col.hidden || false,
                            custom_width: col.custom_width || false,
                            unmapped_data: col_unmapped
                          )
                        end
                      end

          merged_unmapped = {}
          merged_unmapped[:facade] = facade_meta unless facade_meta.empty?

          drw = @worksheet_drawing&.drawing
          charts_data = drw ? drw.charts.map(&:to_chart_options) : []

          Xlsxrb::Elements::Worksheet.new(
            name: name,
            rows: sorted_rows,
            columns: cols_data,
            charts: charts_data,
            conditional_formatting: cf_data,
            data_validations: dv_data,
            unmapped_data: merged_unmapped
          )
        end

        # Builds a Worksheet instance from an immutable Xlsxrb::Elements::Worksheet.
        #
        # @param xlsxrb_sheet [Xlsxrb::Elements::Worksheet]

        private

        # @return [Symbol, String]
        #: () -> (Symbol | String)
        def xml_space
          workbook.xml_space
        end

        # @param collection [untyped]
        # @param range [Range[Integer]]
        # @param level [Integer]
        # @param collapsed [Boolean]
        # @return [void]
        #: (untyped collection, Range[Integer] range, ?Integer level, ?bool collapsed) -> void
        def outline(collection, range, level = 1, collapsed = true)
          range.each do |idx|
            item = collection[idx]
            if item
              item.outline_level = level
              item.hidden = collapsed
            end
            sheet_view.show_outline_symbols = true
          end
        end

        # @param name [String]
        # @return [void]
        #: (String name) -> void
        def validate_sheet_name(name)
          DataTypeValidator.validate(:worksheet_name, String, name)
          raise ArgumentError, ERR_SHEET_NAME_EMPTY if name.empty?

          char_length = name.encode("utf-16")[1..].encode("utf-16").bytesize / 2
          raise ArgumentError, format(ERR_SHEET_NAME_TOO_LONG, name) if char_length > WORKSHEET_MAX_NAME_LENGTH && Caxlsx.validate_sheet_name_length

          raise ArgumentError, format(ERR_SHEET_NAME_CHARACTER_FORBIDDEN, name) if WORKSHEET_NAME_FORBIDDEN_CHARS.any? { |char| name.include?(char) }

          name_encoded = Caxlsx.coder.encode(name)
          sheet_names = @workbook.worksheets.reject { |s| s == self }.map(&:name)
          return unless sheet_names.include?(name_encoded)

          raise ArgumentError, format(ERR_DUPLICATE_SHEET_NAME, name)
        end

        # @return [Array[untyped]]
        #: () -> Array[untyped]
        def serializable_parts
          [sheet_pr, dimension, sheet_view, sheet_format_pr, column_info,
           sheet_data, sheet_calc_pr, @sheet_protection, protected_ranges,
           auto_filter, merged_cells, conditional_formattings,
           data_validations, hyperlinks, print_options, page_margins,
           page_setup, header_footer, row_breaks, col_breaks, worksheet_drawing, worksheet_comments,
           tables]
        end

        # @param first [Cell]
        # @param last [Cell]
        # @return [Array[Cell]]
        #: (Cell first, Cell last) -> Array[Cell]
        def range(first, last)
          cells_result = []
          first_row = first.row&.row_index || 0
          last_row = last.row&.row_index || 0
          first_col = first.index || 0
          last_col = last.index || 0

          rows[(first_row..last_row)].each do |r|
            r[(first_col..last_col)].each do |c|
              cells_result << c
            end
          end
          cells_result
        end

        # @param v [Workbook]
        # @return [Workbook]
        #: (Workbook v) -> Workbook
        def workbook=(v)
          DataTypeValidator.validate("Worksheet.workbook", Workbook, v)
          @workbook = v
        end

        # @param row [Row]
        # @param widths [Array[Numeric, Symbol, nil], nil]
        # @return [void]
        #: (Row row, ?Array[Numeric | Symbol | nil]? widths) -> void
        def update_column_info(row, widths = nil)
          use_auto = workbook.use_autowidth
          return unless use_auto || widths

          row.cells.each_with_index do |cell, idx|
            w = widths ? widths[idx] : nil
            col = find_or_create_column_info(idx)
            if w == :ignore
              col.ignore_cell_width = true
            elsif w == :auto
              col.ignore_cell_width = false
              col.width = nil
            elsif w.is_a?(Numeric)
              col.width = w
            elsif !col.width.nil?
              next
            else
              col.update_width(cell, nil, use_auto)
            end
          end
        end

        # @return [void]
        #: () -> void
        def add_autofilter_defined_name_to_workbook
          return unless auto_filter.range

          workbook.defined_names.add_defined_name(
            auto_filter.defined_name,
            name: "_xlnm._FilterDatabase",
            local_sheet_id: index,
            hidden: true
          )
        end
      end
    end
  end
end
