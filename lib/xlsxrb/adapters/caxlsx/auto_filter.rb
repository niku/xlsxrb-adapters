# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # This class expresses a filter criteria value.
      class Filter
        # @return [untyped]
        attr_accessor :val

        # @param value [untyped]
        #: (untyped value) -> void
        def initialize(value)
          @val = value
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<filter val='#{@val}' />"
        end
      end

      # This collection is used to express a group of dates or times used in an AutoFilter criteria.
      class DateGroupItem
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :date_time_grouping, :year, :month, :day, :hour, :minute, :second

        DATE_TIME_GROUPING = %w[year month day hour minute second].freeze

        # @return [String]
        attr_reader :date_time_grouping

        # @return [Integer, String]
        attr_reader :year

        # @return [Integer, nil]
        attr_reader :month

        # @return [Integer, nil]
        attr_reader :day

        # @return [Integer, nil]
        attr_reader :hour

        # @return [Integer, nil]
        attr_reader :minute

        # @return [Integer, nil]
        attr_reader :second

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          raise ArgumentError, "You must specify a year for date time grouping" unless options[:year]
          raise ArgumentError, "You must specify a date_time_grouping when creating a DateGroupItem for auto filter" unless options[:date_time_grouping]

          parse_options(options)
        end

        # @param value [Integer, String]
        # @return [Integer, String]
        #: (Integer | String value) -> (Integer | String)
        def year=(value)
          RegexValidator.validate("DateGroupItem.year", /\d{4}/, value)
          @year = value
        end

        # @param value [Integer]
        # @return [Integer]
        #: (Integer value) -> Integer
        def month=(value)
          RangeValidator.validate("DateGroupItem.month", 0, 12, value)
          @month = value
        end

        # @param value [Integer]
        # @return [Integer]
        #: (Integer value) -> Integer
        def day=(value)
          RangeValidator.validate("DateGroupItem.day", 0, 31, value)
          @day = value
        end

        # @param value [Integer]
        # @return [Integer]
        #: (Integer value) -> Integer
        def hour=(value)
          RangeValidator.validate("DateGroupItem.hour", 0, 23, value)
          @hour = value
        end

        # @param value [Integer]
        # @return [Integer]
        #: (Integer value) -> Integer
        def minute=(value)
          RangeValidator.validate("DateGroupItem.minute", 0, 59, value)
          @minute = value
        end

        # @param value [Integer]
        # @return [Integer]
        #: (Integer value) -> Integer
        def second=(value)
          RangeValidator.validate("DateGroupItem.second", 0, 59, value)
          @second = value
        end

        # @param grouping [Symbol, String]
        # @return [String]
        #: (Symbol | String grouping) -> String
        def date_time_grouping=(grouping)
          RestrictionValidator.validate("DateGroupItem.date_time_grouping", DATE_TIME_GROUPING, grouping.to_s)
          @date_time_grouping = grouping.to_s
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("dateGroupItem", str)
        end
      end

      # Filters for an auto filter column.
      class Filters
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :blank, :calendar_type

        CALENDAR_TYPES = %w[gregorian gregorianUs gregorianMeFrench gregorianArabic hijri hebrew taiwan japan thai korea saka gregorianXlitEnglish gregorianXlitFrench none].freeze

        # @return [Boolean, nil]
        attr_reader :blank

        # @return [String, nil]
        attr_reader :calendar_type

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          parse_options(options)
        end

        # @param cell [Cell, nil]
        # @return [Boolean]
        #: (Cell? cell) -> bool
        def apply(cell)
          return false unless cell

          filter_items.each do |filter|
            return false if cell.value == filter.val
          end
          true
        end

        # @return [Array[Filter]]
        #: () -> Array[Filter]
        def filter_items
          @filter_items ||= []
        end

        # @return [Array[DateGroupItem]]
        #: () -> Array[DateGroupItem]
        def date_group_items
          @date_group_items ||= []
        end

        # @param calendar [String]
        # @return [String]
        #: (String calendar) -> String
        def calendar_type=(calendar)
          RestrictionValidator.validate("Filters.calendar_type", CALENDAR_TYPES, calendar)
          @calendar_type = calendar
        end

        # @param use_blank [Boolean]
        # @return [Boolean]
        #: (bool use_blank) -> bool
        def blank=(use_blank)
          Caxlsx.validate_boolean(use_blank)
          @blank = use_blank
        end

        # @param values [Array[untyped]]
        # @return [Array[untyped]]
        #: (Array[untyped] values) -> Array[untyped]
        def filter_items=(values)
          values.each do |value|
            filter_items << Filter.new(value)
          end
        end

        # @param options [Array[Hash{Symbol => untyped}]]
        # @return [Array[Hash{Symbol => untyped}]]
        #: (Array[Hash[Symbol, untyped]] options) -> Array[Hash[Symbol, untyped]]
        def date_group_items=(options)
          options.each do |date_group|
            raise ArgumentError, "date_group_items should be an array of hashes specifying the options for each date_group_item" unless date_group.is_a?(Hash)

            date_group_items << DateGroupItem.new(date_group)
          end
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<filters "
          serialized_attributes(str)
          str << ">"
          filter_items.each { |filter| filter.to_xml_string(str) }
          date_group_items.each { |date_group_item| date_group_item.to_xml_string(str) }
          str << "</filters>"
        end
      end

      # The filterColumn collection identifies a particular column in the AutoFilter range.
      class FilterColumn
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :col_id, :hidden_button, :show_button

        FILTERS = [:filters].freeze

        # @return [Integer]
        attr_reader :col_id

        # @return [Filters]
        attr_reader :filter

        # @param col_id [Integer, Cell]
        # @param filter_type [Symbol]
        # @param options [Hash{Symbol => untyped}]
        #: (Integer | Cell col_id, Symbol filter_type, ?Hash[Symbol, untyped] options) ?{ (Filters) -> void } -> void
        def initialize(col_id, filter_type, options = {})
          RestrictionValidator.validate("FilterColumn.filter", FILTERS, filter_type)
          self.col_id = col_id
          parse_options(options)
          @filter = Caxlsx.const_get(Caxlsx.camel(filter_type)).new(options)
          yield @filter if block_given?
        end

        # @return [Boolean]
        #: () -> bool
        def show_button
          return @show_button if defined?(@show_button)

          true
        end

        # @return [Boolean]
        #: () -> bool
        def hidden_button
          @hidden_button ||= false
        end

        # @param column_index [Integer, Cell]
        # @return [Integer]
        #: (Integer | Cell column_index) -> Integer
        def col_id=(column_index)
          column_index = column_index.col if column_index.is_a?(Cell)
          Caxlsx.validate_unsigned_int(column_index)
          @col_id = column_index
        end

        # @param row [Row]
        # @param offset [Integer]
        # @return [void]
        #: (Row row, Integer offset) -> void
        def apply(row, offset)
          row.hidden = @filter.apply(row.cells[offset + col_id.to_i])
        end

        # @param hidden [Boolean]
        # @return [Boolean]
        #: (bool hidden) -> bool
        def hidden_button=(hidden)
          Caxlsx.validate_boolean(hidden)
          @hidden_button = hidden
        end

        # @param show [Boolean]
        # @return [Boolean]
        #: (bool show) -> bool
        def show_button=(show)
          Caxlsx.validate_boolean(show)
          @show_button = show
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<filterColumn "
          serialized_attributes(str)
          str << ">"
          @filter.to_xml_string(str)
          str << "</filterColumn>"
        end
      end

      # This class represents an individual sort condition belonging to the sort state of an auto filter.
      class SortCondition
        # @return [Integer]
        attr_reader :column_index

        # @return [Symbol]
        attr_reader :order

        # @return [Array[untyped]]
        attr_reader :custom_list

        # @param column_index [Integer]
        # @param order [Symbol]
        # @param custom_list [Array[untyped]]
        #: (column_index: Integer, order: Symbol, custom_list: Array[untyped]) -> void
        def initialize(column_index:, order:, custom_list:)
          Caxlsx.validate_int(column_index)
          @column_index = column_index

          RestrictionValidator.validate("SortCondition.order", %i[asc desc], order)
          @order = order

          DataTypeValidator.validate(:sort_condition_custom_list, Array, custom_list)
          @custom_list = custom_list
        end

        # @param ref [String]
        # @param column_index [Integer]
        # @return [String]
        #: (String ref, Integer column_index) -> String
        def ref_to_single_column(ref, column_index)
          first_cell, last_cell = ref.split(":")
          start_point = Caxlsx.name_to_indices(first_cell)

          first_row = first_cell[/\d+/]
          last_row = last_cell[/\d+/]

          first_column = Caxlsx.col_ref(column_index + start_point.first)
          "#{first_column}#{first_row}:#{first_column}#{last_row}"
        end

        # @param str [String]
        # @param ref [String]
        # @return [String]
        #: (String str, String ref) -> String
        def to_xml_string(str, ref)
          col_ref = ref_to_single_column(ref, column_index)

          str << "<sortCondition "
          str << "descending='1' " if order == :desc
          str << "ref='#{col_ref}' "
          str << "customList='#{custom_list.join(",")}' " unless custom_list.empty?
          str << "/>"
        end
      end

      # This class performs sorting on a range in a worksheet.
      class SortState
        # @return [SimpleTypedList]
        #: () -> SimpleTypedList
        def sort_conditions
          @sort_conditions ||= SimpleTypedList.new(SortCondition)
        end

        # @param auto_filter [AutoFilter]
        #: (AutoFilter auto_filter) -> void
        def initialize(auto_filter)
          @auto_filter = auto_filter
        end

        # @param column_index [Integer]
        # @param order [Symbol]
        # @param custom_list [Array[untyped]]
        # @return [SortCondition]
        #: (column_index: Integer, ?order: Symbol, ?custom_list: Array[untyped]) -> SortCondition
        def add_sort_condition(column_index:, order: :asc, custom_list: [])
          sort_conditions << SortCondition.new(column_index: column_index, order: order, custom_list: custom_list)
          sort_conditions.last
        end

        # @param str [String]
        # @return [String]
        #: (String str) -> String
        def increment_cell_value(str)
          letter = str[/[A-Za-z]+/]
          number = str[/\d+/].to_i
          "#{letter}#{number + 1}"
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if sort_conditions.empty?

          ref = @auto_filter.range
          return unless ref

          first_cell, last_cell = ref.split(":")
          ref_adj = "#{increment_cell_value(first_cell)}:#{last_cell}"

          str << "<sortState xmlns:xlrd2='http://schemas.microsoft.com/office/spreadsheetml/2017/richdata2' ref='#{ref_adj}'>"
          sort_conditions.each { |sort_condition| sort_condition.to_xml_string(str, ref_adj) }
          str << "</sortState>"
        end
      end

      # This class represents an auto filter range in a worksheet.
      class AutoFilter
        # @return [Worksheet]
        attr_reader :worksheet

        # @return [Boolean]
        attr_reader :sort_on_generate

        # @return [String, nil]
        attr_accessor :range

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          @worksheet = worksheet
          @sort_on_generate = true
        end

        # @return [String, nil]
        #: () -> String?
        def defined_name
          return unless range

          cells = range.split(":").map do |name|
            cell = worksheet.name_to_cell(name)
            next cell if cell

            col_index, row_index = *Caxlsx.name_to_indices(name)
            worksheet.add_row while worksheet.rows[row_index].nil?

            row = worksheet.rows[row_index]
            row.add_cell while row[col_index].nil?

            row[col_index]
          end

          Caxlsx.cell_range(cells)
        end

        # @return [SimpleTypedList]
        #: () -> SimpleTypedList
        def columns
          @columns ||= SimpleTypedList.new(FilterColumn)
        end

        # @param col_id [Integer, Cell]
        # @param filter_type [Symbol]
        # @param options [Hash{Symbol => untyped}]
        # @return [FilterColumn]
        #: (Integer | Cell col_id, Symbol filter_type, ?Hash[Symbol, untyped] options) -> FilterColumn
        def add_column(col_id, filter_type, options = {})
          columns << FilterColumn.new(col_id, filter_type, options)
          columns.last
        end

        # @return [void]
        #: () -> void
        def apply
          return unless range

          first_cell, last_cell = range.split(":")
          start_point = Caxlsx.name_to_indices(first_cell)
          end_point = Caxlsx.name_to_indices(last_cell)
          rows = worksheet.rows[(start_point.last + 1)..end_point.last] || []

          if !sort_state.sort_conditions.empty? && sort_on_generate
            sort_conditions = sort_state.sort_conditions
            sorted_rows = rows.sort do |row1, row2|
              comparison = 0

              sort_conditions.each do |condition|
                cell_value_row1 = row1.cells[condition.column_index + start_point.first]&.value
                cell_value_row2 = row2.cells[condition.column_index + start_point.first]&.value
                custom_list = condition.custom_list
                comparison = if cell_value_row1.nil? || cell_value_row2.nil?
                               cell_value_row1.nil? ? 1 : -1
                             elsif custom_list.empty?
                               condition.order == :asc ? (cell_value_row1 <=> cell_value_row2 || 0) : (cell_value_row2 <=> cell_value_row1 || 0)
                             else
                               index1 = custom_list.index(cell_value_row1) || custom_list.size
                               index2 = custom_list.index(cell_value_row2) || custom_list.size

                               condition.order == :asc ? index1 <=> index2 : index2 <=> index1
                             end

                break unless comparison.zero?
              end

              comparison
            end
            insert_index = start_point.last + 1

            sorted_rows.each do |row|
              worksheet.rows[insert_index] = row
              insert_index += 1
            end
          end

          column_offset = start_point.first
          columns.each do |column|
            rows.each do |row|
              next if row.hidden

              column.apply(row, column_offset)
            end
          end
        end

        # @return [SortState]
        #: () -> SortState
        def sort_state
          @sort_state ||= SortState.new(self)
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def sort_on_generate=(v)
          Caxlsx.validate_boolean(v)
          @sort_on_generate = v
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return unless range

          str << "<autoFilter ref='#{range}'>"
          columns.each { |filter_column| filter_column.to_xml_string(str) }
          @sort_state&.to_xml_string(str)
          str << "</autoFilter>"
        end
      end
    end
  end
end
