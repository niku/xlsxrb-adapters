# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "stringio"
require "xlsxrb"
require_relative "worksheet"
require_relative "row"
require_relative "cell"
require_relative "formula"
require_relative "color"
require_relative "styles"
require_relative "defined_names"
require_relative "shared_strings"

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a root container compatible with RubyXL root object.
      class Root
        attr_reader :workbook #: Workbook

        # @param workbook [Workbook]
        #: (Workbook workbook) -> void
        def initialize(workbook)
          @workbook = workbook
        end

        # @return [Array[Object]]
        #: () -> Array[untyped]
        def collect_related_objects
          objs = []
          container = @workbook.shared_strings_container
          objs << container if container && !container.empty?
          objs
        end
      end

      # Represents a mutable workbook wrapper compatible with RubyXL::Workbook.
      class Workbook
        [Enumerable].each { |m| include m }

        APPLICATION = "Microsoft Macintosh Excel"
        APPVERSION  = "12.0000"

        DATE1904 = DateTime.new(1904, 1, 1)
        DATE1899 = DateTime.new(1899, 12, 30) # 1899-12-31 - 1 day
        MARCH_1_1900 = 61
        SHEET_NAME_FORBIDDEN_CHARS = %r{[\\/:?*\[\]]}
        SHEET_NAME_FORBIDDEN_NAMES = %w[History].freeze

        attr_accessor :worksheets #: Array[Worksheet]
        attr_accessor :shared_strings #: Array[String]
        attr_accessor :shared_strings_container #: SharedStringsTable
        attr_accessor :styles #: Hash[untyped, untyped]
        attr_accessor :unmapped_data #: Hash[untyped, untyped]
        attr_accessor :stylesheet #: Stylesheet
        attr_accessor :creator #: String?
        attr_accessor :modifier #: String?
        attr_accessor :created_at #: Time?
        attr_accessor :modified_at #: Time?
        attr_accessor :title #: String?
        attr_accessor :company #: String
        attr_accessor :application #: String
        attr_accessor :appversion #: String
        attr_accessor :date1904 #: bool
        attr_accessor :defined_names #: DefinedNames

        # Returns root container compatible with RubyXL::Workbook#root.
        #
        # @return [Root]
        #: () -> Root
        def root
          Root.new(self)
        end

        # Creates a new Workbook from an existing Xlsxrb::Elements::Workbook.
        #
        # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
        # @return [Workbook]
        #: (Xlsxrb::Elements::Workbook xlsxrb_workbook) -> Workbook
        def self.from_xlsxrb(xlsxrb_workbook)
          wb = xlsxrb_workbook.respond_to?(:load) ? xlsxrb_workbook.load : xlsxrb_workbook

          adapter_sheets = wb.sheets.map do |sheet|
            ws = Worksheet.new(
              sheet_name: sheet.name,
              columns: sheet.columns || [],
              charts: sheet.charts || [],
              unmapped_data: sheet.unmapped_data || {}
            )
            ws.data_validations = (sheet.data_validations || sheet.unmapped_data&.dig(:facade, :data_validations) || []).dup

            # Load column styles and widths into ColumnRanges (cols)
            if sheet.columns && !sheet.columns.empty?
              sheet.columns.each do |c|
                cr = ws.cols.get_range(c.index)
                cr.width = c.width if c.width
                cr.custom_width = c.custom_width if c.custom_width
                cr.hidden = c.hidden if c.hidden
                cr.style_index = c.style_index if c.style_index
              end
            end

            # Load merged_cells if present in unmapped_data or worksheet
            if sheet.unmapped_data && sheet.unmapped_data[:merged_cells]
              ws.merged_cells = MergedCells.new
              sheet.unmapped_data[:merged_cells].each do |mc_ref|
                ws.merged_cells << MergedCell.new(ref: mc_ref)
              end
            end

            sheet.rows.each do |row|
              adapter_row = ws.add_row(
                row.index,
                height: row.height,
                hidden: row.hidden,
                custom_height: row.custom_height,
                outline_level: row.outline_level,
                style_index: row.style_index || 0
              )

              row.cells.each do |cell|
                f_val = if cell.formula
                          Formula.new(cell.formula.is_a?(Xlsxrb::Elements::Formula) ? cell.formula.expression : cell.formula.to_s)
                        end
                adapter_cell = Cell.new(
                  worksheet: ws,
                  row: cell.row_index,
                  column: cell.column_index,
                  value: cell.value,
                  formula: f_val,
                  style_index: cell.style_index
                )
                adapter_row[cell.column_index] = adapter_cell
              end
            end

            ws
          end

          instance = new(adapter_sheets)
          instance.shared_strings = wb.shared_strings || []
          instance.shared_strings.each { |s| instance.shared_strings_container.add(s) }
          instance.styles = wb.styles || {}
          instance.stylesheet = Stylesheet.from_xlsxrb(wb.styles || {})
          instance.unmapped_data = wb.unmapped_data || {}

          # Load metadata / core properties if present
          if wb.unmapped_data
            cp = wb.unmapped_data[:core_properties]
            if cp.is_a?(Hash)
              instance.creator = cp[:creator]&.to_s
              instance.modifier = cp[:modifier]&.to_s
              instance.created_at = cp[:created_at] if cp[:created_at].is_a?(Time)
              instance.modified_at = cp[:modified_at] if cp[:modified_at].is_a?(Time)
              instance.title = cp[:title]&.to_s
            end
            if wb.unmapped_data[:defined_names].is_a?(Array)
              instance.defined_names = DefinedNames.new
              wb.unmapped_data[:defined_names].each do |dn|
                instance.defined_names << DefinedName.new(dn)
              end
            end
          end

          instance
        end

        # @param worksheets [Array<Worksheet>]
        # @param shared_strings [Array<String>]
        # @param styles [Hash]
        # @param unmapped_data [Hash]
        # @param creator [String, nil]
        # @param modifier [String, nil]
        # @param created_at [Time, nil]
        # @param modified_at [Time, nil]
        # @param company [String]
        # @param application [String]
        # @param appversion [String]
        # @param date1904 [Boolean]
        #: (?Array[Worksheet] worksheets, ?shared_strings: Array[String], ?styles: Hash[untyped, untyped], ?unmapped_data: Hash[untyped, untyped], ?creator: String?, ?modifier: String?, ?created_at: Time?, ?modified_at: Time?, ?company: String, ?application: String, ?appversion: String, ?date1904: bool) -> void
        def initialize(worksheets = [], shared_strings: [], styles: {}, unmapped_data: {},
                       creator: nil, modifier: nil, created_at: nil, modified_at: nil,
                       company: "", application: APPLICATION, appversion: APPVERSION, date1904: false)
          @worksheets = if worksheets.nil? || worksheets.empty?
                          [Worksheet.new(workbook: self, sheet_name: "Sheet1")]
                        else
                          worksheets.dup
                        end
          @worksheets.each { |ws| ws.workbook = self }
          @shared_strings = shared_strings.dup
          @shared_strings_container = SharedStringsTable.new
          @shared_strings.each { |s| @shared_strings_container.add(s) }
          @styles = styles.dup
          @stylesheet = Stylesheet.from_xlsxrb(styles)
          @unmapped_data = unmapped_data.dup
          @creator = creator
          @modifier = modifier
          @created_at = created_at || Time.now
          @modified_at = modified_at || Time.now
          @title = nil
          @company = company
          @application = application
          @appversion = appversion
          @date1904 = date1904
          @defined_names = DefinedNames.new
        end

        # Iterates through worksheets.
        #
        # @yieldparam worksheet [Worksheet]
        # @return [Enumerator, self]
        #: () ?{ (Worksheet) -> void } -> (self | Enumerator[Worksheet, void])
        def each(&)
          return to_enum(:each) unless block_given?

          @worksheets.each(&)
          self
        end

        # Accesses a worksheet by index (Integer) or sheet name (String/Symbol).
        #
        # @param idx_or_name [Integer, String, Symbol]
        # @return [Worksheet, nil]
        #: (Integer | String | Symbol idx_or_name) -> Worksheet?
        def [](idx_or_name)
          case idx_or_name
          when Integer
            @worksheets[idx_or_name]
          when String, Symbol
            target_name = idx_or_name.to_s
            @worksheets.find { |ws| ws.sheet_name == target_name }
          end
        end

        # Adds a new worksheet with the given name.
        # Enforces Excel limits: max 31 chars, no forbidden chars [ ] * ? / \.
        #
        # @param name [String, nil]
        # @return [Worksheet]
        #: (?String? name) -> Worksheet
        def add_worksheet(name = nil)
          if name.nil? || name.to_s.empty?
            i = 1
            loop do
              candidate = "Sheet#{i}"
              unless self[candidate]
                name = candidate
                break
              end
              i += 1
            end
          end
          validate_sheet_name!(name)

          sheet = Worksheet.new(workbook: self, sheet_name: name)
          @worksheets << sheet
          sheet
        end

        # Validates worksheet name against Excel limits.
        #
        # @param name [String]
        # @return [void]
        #: (String name) -> void
        def validate_sheet_name!(name)
          raise "Worksheet name '#{name}' contains forbidden characters" if name =~ SHEET_NAME_FORBIDDEN_CHARS
          raise "Worksheet name '#{name}' is forbidden" if SHEET_NAME_FORBIDDEN_NAMES.any? { |f| f.casecmp?(name) }
        end
        private :validate_sheet_name!

        # --- Stylesheet Accessors and Convenience Methods ---

        # @return [Array<XF>]
        #: () -> Array[XF]
        def cell_xfs
          @stylesheet.cell_xfs
        end

        # @return [Array<Font>]
        #: () -> Array[Font]
        def fonts
          @stylesheet.fonts
        end

        # @return [Array<Fill>]
        #: () -> Array[Fill]
        def fills
          @stylesheet.fills
        end

        # @return [Array<Border>]
        #: () -> Array[Border]
        def borders
          @stylesheet.borders
        end

        # Returns fill RGB color for a given XF.
        #
        # @param xf [XF]
        # @return [String]
        #: (XF xf) -> String
        def get_fill_color(xf)
          fill = fills[xf.fill_id || 0]
          fill&.fill_color || "ffffff"
        end

        # @param new_fill [Fill]
        # @param old_xf [XF]
        # @return [XF]
        #: (Fill new_fill, XF old_xf) -> XF
        def register_new_fill(new_fill, old_xf)
          new_xf = old_xf.dup
          new_xf.apply_fill = true
          fill_id = fills.find_index { |f| f == new_fill }
          if fill_id.nil?
            fill_id = fills.size
            fills << new_fill
          end
          new_xf.fill_id = fill_id
          new_xf
        end

        # @param new_font [Font]
        # @param old_xf [XF]
        # @return [XF]
        #: (Font new_font, XF old_xf) -> XF
        def register_new_font(new_font, old_xf)
          new_xf = old_xf.dup
          new_xf.apply_font = true
          font_id = fonts.find_index { |f| f == new_font }
          if font_id.nil?
            font_id = fonts.size
            fonts << new_font
          end
          new_xf.font_id = font_id
          new_xf
        end

        # @param new_xf [XF]
        # @return [Integer]
        #: (XF new_xf) -> Integer
        def register_new_xf(new_xf)
          xf_id = cell_xfs.find_index { |x| x == new_xf }
          if xf_id.nil?
            xf_id = cell_xfs.size
            cell_xfs << new_xf
          end
          xf_id
        end

        # Modifies alignment in stylesheet for a given style_index.
        #
        # @param style_index [Integer, nil]
        # @yieldparam alignment [Alignment]
        # @return [Integer]
        #: (Integer? style_index) { (Alignment) -> void } -> Integer
        def modify_alignment(style_index, &)
          old_xf = cell_xfs[style_index || 0] || cell_xfs[0]
          new_xf = old_xf.dup
          new_xf.alignment = old_xf.alignment ? old_xf.alignment.dup : Alignment.new

          yield(new_xf.alignment)
          new_xf.apply_alignment = true

          register_new_xf(new_xf)
        end

        # Modifies fill in stylesheet for a given style_index.
        #
        # @param style_index [Integer, nil]
        # @param rgb [String]
        # @return [Integer]
        #: (Integer? style_index, String rgb) -> Integer
        def modify_fill(style_index, rgb)
          old_xf = cell_xfs[style_index || 0] || cell_xfs[0]
          new_fill = Fill.new(pattern_type: "solid", fg_color: rgb)
          new_xf = register_new_fill(new_fill, old_xf)
          register_new_xf(new_xf)
        end

        # Modifies border in stylesheet for a given style_index.
        #
        # @param style_index [Integer, nil]
        # @param direction [Symbol, String]
        # @param weight [String]
        # @param diagonals [untyped]
        # @return [Integer]
        #: (Integer? style_index, Symbol | String direction, String weight, ?untyped diagonals) -> Integer
        def modify_border(style_index, direction, weight, diagonals = nil)
          old_xf = cell_xfs[style_index || 0] || cell_xfs[0]
          new_xf = old_xf.dup
          old_border = borders[new_xf.border_id || 0] || borders[0]
          new_border = old_border.dup

          new_border.set_edge_style(direction, weight, diagonals)

          border_id = borders.find_index { |b| b == new_border }
          if border_id.nil?
            border_id = borders.size
            borders << new_border
          end
          new_border_id = border_id
          new_xf.border_id = new_border_id
          new_xf.apply_border = true

          register_new_xf(new_xf)
        end

        # Modifies border color in stylesheet for a given style_index.
        #
        # @param style_index [Integer, nil]
        # @param direction [Symbol, String]
        # @param color [String]
        # @return [Integer]
        #: (Integer? style_index, Symbol | String direction, String color) -> Integer
        def modify_border_color(style_index, direction, color)
          old_xf = cell_xfs[style_index || 0] || cell_xfs[0]
          new_xf = old_xf.dup
          old_border = borders[new_xf.border_id || 0] || borders[0]
          new_border = old_border.dup

          new_border.set_edge_color(direction, color)

          border_id = borders.find_index { |b| b == new_border }
          if border_id.nil?
            border_id = borders.size
            borders << new_border
          end
          new_border_id = border_id
          new_xf.border_id = new_border_id
          new_xf.apply_border = true

          register_new_xf(new_xf)
        end

        # --- Date & Numeric Conversion ---

        # @return [DateTime]
        #: () -> DateTime
        def base_date
          @date1904 ? DATE1904 : DATE1899
        end

        # Converts Date, DateTime, or Time into Excel numeric serial date.
        #
        # @param date [Date, DateTime, Time]
        # @return [Float]
        #: (Date | DateTime | Time date) -> Float
        def date_to_num(date)
          case date
          when DateTime, Date
            (date.ajd - base_date.ajd).to_f
          when Time
            ((date.to_r - base_date.to_time.to_r) / 86_400).to_f
          else
            date.to_f
          end
        end

        # Converts Excel numeric serial date into DateTime.
        #
        # @param num [Numeric, nil]
        # @return [DateTime, nil]
        #: (Numeric? num) -> (Date | DateTime)?
        def num_to_date(num)
          return nil if num.nil?

          n = num.to_f
          n += 1 if n < MARCH_1_1900 && !@date1904

          dateparts = n.divmod(1)
          base_date + dateparts[0] + Rational((dateparts[1] * 86_400).round(6), 86_400)
        end

        # --- Defined Names ---

        # @param name [String]
        # @param reference [String]
        # @return [DefinedName]
        #: (String name, String reference) -> DefinedName
        def define_new_name(name, reference)
          @defined_names ||= DefinedNames.new
          entry = DefinedName.new(name: name, reference: reference)
          @defined_names << entry
          entry
        end

        # @param name [String]
        # @return [DefinedName, nil]
        #: (String name) -> DefinedName?
        def get_defined_name(name)
          @defined_names&.find { |n| n.name == name }
        end

        # Calculate password hash from string for use in 'password' fields.
        # https://www.openoffice.org/sc/excelfileformat.pdf
        #
        # @param pwd [String]
        # @return [String]
        #: (String pwd) -> String
        def password_hash(pwd)
          hsh = 0
          pwd.reverse.each_char do |c|
            hsh ^= c.ord
            hsh <<= 1
            hsh -= 0x7fff if hsh > 0x7fff
          end

          (hsh ^ pwd.length ^ 0xCE4B).to_s(16)
        end

        # Converts the mutable adapter model into an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          elements_sheets = @worksheets.map(&:to_xlsxrb)

          styles_hash = @stylesheet.to_xlsxrb_hash

          merged_unmapped = @unmapped_data.dup
          facade_meta = (merged_unmapped[:facade] || {}).dup

          if @creator || @modifier || @created_at || @modified_at || @title
            created_str = @created_at&.utc&.strftime("%Y-%m-%dT%H:%M:%SZ")
            modified_str = @modified_at&.utc&.strftime("%Y-%m-%dT%H:%M:%SZ")
            facade_meta[:core_properties] = {
              creator: @creator,
              last_modified_by: @modifier,
              created: created_str,
              modified: modified_str,
              title: @title
            }.compact
          end
          if @defined_names && !@defined_names.empty?
            facade_meta[:defined_names] = @defined_names.map do |dn|
              { name: dn.name, value: dn.reference }
            end
          end
          merged_unmapped[:facade] = facade_meta unless facade_meta.empty?

          sst = @shared_strings_container
          sst_strings = sst && !sst.empty? ? sst.strings.map(&:to_s) : @shared_strings

          Xlsxrb::Elements::Workbook.new(
            sheets: elements_sheets,
            shared_strings: sst_strings,
            styles: styles_hash,
            unmapped_data: merged_unmapped
          )
        end

        # Writes workbook out to an XLSX file.
        # Trims sheet names exceeding 31 characters to 31 characters as per Excel compatibility.
        #
        # @param filepath [String, IO]
        # @return [void]
        #: (String | IO filepath) -> void
        def write(filepath)
          @worksheets.each do |ws|
            validate_sheet_name!(ws.sheet_name)
            ws.sheet_name = ws.sheet_name[0..30] if ws.sheet_name && ws.sheet_name.length > 31
          end
          Xlsxrb.write(filepath, to_xlsxrb)
        end

        # Exports workbook as a binary StringIO stream.
        #
        # @return [StringIO]
        #: () -> StringIO
        def stream
          @worksheets.each do |ws|
            validate_sheet_name!(ws.sheet_name)
            ws.sheet_name = ws.sheet_name[0..30] if ws.sheet_name && ws.sheet_name.length > 31
          end
          binary = Xlsxrb.write(to_xlsxrb)
          io = StringIO.new(binary)
          io.binmode
          io
        end

        # Inspect representation.
        #
        # @return [String]
        #: () -> String
        def inspect
          "#<#{self.class.name} worksheets_count=#{@worksheets.size}>"
        end
      end
    end
  end
end
