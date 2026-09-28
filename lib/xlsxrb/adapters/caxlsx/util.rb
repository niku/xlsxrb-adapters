# frozen_string_literal: true

# rbs_inline: enabled

require_relative "constants"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Coder fallback for HTML entities encoding
      class EntityCoder
        #: (String? str) -> String
        def encode(str)
          return "" if str.nil?

          str.to_s
             .gsub("&", "&amp;")
             .gsub("<", "&lt;")
             .gsub(">", "&gt;")
             .gsub('"', "&quot;")
             .gsub("'", "&apos;")
        end
      end

      # Module functions matching Axlsx helper methods
      class << self
        # @return [Boolean]
        #: () -> bool
        def escape_formulas
          !defined?(@escape_formulas) || @escape_formulas.nil? ? true : @escape_formulas
        end

        # @param value [Boolean]
        # @return [Boolean]
        #: (bool value) -> bool
        def escape_formulas=(value)
          validate_boolean(value)
          @escape_formulas = value
        end

        # @return [Boolean]
        #: () -> bool
        def secure_formulas
          escape_formulas
        end

        # @param value [Boolean]
        # @return [Boolean]
        #: (bool value) -> bool
        def secure_formulas=(value)
          self.escape_formulas = value
        end

        # @return [Boolean]
        #: () -> bool
        def trust_input
          !defined?(@trust_input) || @trust_input.nil? ? false : @trust_input
        end

        # @param value [Boolean]
        # @return [Boolean]
        #: (bool value) -> bool
        def trust_input=(value)
          validate_boolean(value)
          @trust_input = value
        end

        # @return [Boolean]
        #: () -> bool
        def validate_sheet_name_length
          !defined?(@validate_sheet_name_length) || @validate_sheet_name_length.nil? ? true : @validate_sheet_name_length
        end

        # @param value [Boolean]
        # @return [Boolean]
        #: (bool value) -> bool
        def validate_sheet_name_length=(value)
          validate_boolean(value)
          @validate_sheet_name_length = value
        end

        # Returns HTML entity coder
        # @return [Object]
        #: () -> untyped
        def coder
          @coder ||= begin
            require "htmlentities"
            ::HTMLEntities.new
          rescue LoadError
            EntityCoder.new
          end
        end

        # Sanitizes strings by removing control characters
        # @param str [String]
        # @return [String]
        #: (untyped str) -> String
        def sanitize(str)
          return "" if str.nil?

          str.to_s.gsub(/[\x00-\x08\x0B\x0C\x0E-\x1F\x7F]/, "")
        end

        # Converts a column index (0-based) into an Excel column letter reference (e.g. 0 -> 'A')
        # @param index [Integer]
        # @return [String]
        #: (Integer index) -> String
        def col_ref(index)
          chars = []
          i = index
          while i >= 26
            chars << ((i % 26) + 65).chr
            i = (i / 26) - 1
          end
          chars << (i + 65).chr
          chars.reverse.join
        end

        # Converts a row index (0-based) into an Excel 1-based row number
        # @param index [Integer]
        # @return [String]
        #: (Integer index) -> String
        def row_ref(index)
          (index + 1).to_s
        end

        # Returns the cell reference (e.g. 'A1') for a given column and row index
        # @param col_index [Integer]
        # @param row_index [Integer]
        # @return [String]
        #: (Integer col_index, Integer row_index) -> String
        def cell_r(col_index, row_index)
          "#{col_ref(col_index)}#{row_ref(row_index)}"
        end

        # Converts an Excel cell reference (e.g. 'A1' or 'AB12') into 0-based [col_index, row_index]
        # @param name [String]
        # @return [Array<Integer>]
        #: (String name) -> Array[Integer]
        def name_to_indices(name)
          raise ArgumentError, "invalid cell name" if name.nil? || name.to_s.length < 2

          col_str = name[/\A[A-Z]+/i]
          row_str = name[/\d+\z/]
          raise ArgumentError, "invalid cell name" if col_str.nil? || row_str.nil?

          col_index = 0
          col_str.upcase.each_byte do |b|
            col_index = (col_index * 26) + (b - 64)
          end
          [col_index - 1, row_str.to_i - 1]
        end

        # Converts a range definition into a 2D array of coordinate strings
        # @param range [String]
        # @return [Array<Array<String>>]
        #: (String range) -> Array[Array[String]]
        def range_to_a(range)
          parts = range.split(":")
          raise ArgumentError, "invalid range" unless parts.size == 2

          first_col, first_row = *name_to_indices(parts.first)
          last_col, last_row = *name_to_indices(parts.last)

          (first_row..last_row).map do |r|
            (first_col..last_col).map do |c|
              cell_r(c, r)
            end
          end
        end

        # Returns a formula range string for a list of cells
        # @param cells [Array]
        # @param absolute [Boolean]
        # @return [String]
        #: (Array[untyped] cells, ?bool absolute) -> String
        def cell_range(cells, absolute = true)
          return "" if cells.empty?

          sorted = sort_cells(cells)
          first = sorted.first
          last = sorted.last
          sheet = first.row.worksheet.name

          f_ref = absolute ? first.r_abs : first.r
          l_ref = absolute ? last.r_abs : last.r

          if first == last
            "'#{sheet}'!#{f_ref}"
          else
            "'#{sheet}'!#{f_ref}:#{l_ref}"
          end
        end

        # Sorts a list of cells in reading order (top to bottom, left to right)
        # @param cells [Array]
        # @return [Array]
        #: (Array[untyped] cells) -> Array[untyped]
        def sort_cells(cells)
          cells.sort do |a, b|
            if a.row.row_index == b.row.row_index
              a.index <=> b.index
            else
              a.row.row_index <=> b.row.row_index
            end
          end
        end

        # Converts various types into a boolean
        # @param v [Object]
        # @return [Boolean]
        #: (untyped v) -> bool
        def booleanize(v)
          case v
          when true, 1, "1", "true"
            true
          when false, 0, "0", "false", nil
            false
          else
            str = v.to_s
            return true if str == "true"
            return false if str == "false"

            !v.nil?
          end
        end

        # Recursively merges two hashes
        # @param hash [Hash]
        # @param other_hash [Hash]
        # @return [Hash]
        #: (Hash[untyped, untyped] hash, Hash[untyped, untyped] other_hash) -> Hash[untyped, untyped]
        def hash_deep_merge(hash, other_hash)
          target = hash.dup
          other_hash.each do |k, v|
            target[k] = if target[k].is_a?(Hash) && v.is_a?(Hash)
                          hash_deep_merge(target[k], v)
                        else
                          v
                        end
          end
          target
        end

        # Returns a hash mapping instance variable names (without leading @) as strings to their values
        # @param obj [Object]
        # @return [Hash]
        #: (untyped obj) -> Hash[String, untyped]
        def instance_values_for(obj)
          res = {}
          obj.instance_variables.each do |iv|
            key = iv.to_s.delete_prefix("@")
            res[key] = obj.instance_variable_get(iv)
          end
          res
        end

        # Converts snake_case to CamelCase or camelCase
        # @param s [String, Symbol]
        # @param all_caps [Boolean]
        # @return [String]
        #: (?String | Symbol s, ?bool all_caps) -> String
        def camel(s = "", all_caps = true)
          str = s.to_s
          str = str.capitalize if all_caps
          str.gsub(/_(.)/) { Regexp.last_match(1).upcase }
        end

        # --- Validation helper methods ---

        #: (untyped v) -> void
        def validate_boolean(v)
          raise ArgumentError, "Invalid boolean: #{v.inspect}" unless VALID_BOOLEAN_VALUES.include?(v)
        end

        #: (untyped v) -> void
        def validate_string(v)
          raise ArgumentError, "Invalid string: #{v.inspect}" unless v.is_a?(String)
        end

        #: (untyped v) -> void
        def validate_int(v)
          raise ArgumentError, "Invalid integer: #{v.inspect}" unless v.is_a?(Integer)
        end

        #: (untyped v) -> void
        def validate_unsigned_int(v)
          raise ArgumentError, "Invalid unsigned integer: #{v.inspect}" unless v.is_a?(Integer) && v >= 0
        end

        #: (untyped v) -> void
        def validate_float(v)
          raise ArgumentError, "Invalid float: #{v.inspect}" unless v.is_a?(Float) || v.is_a?(Numeric)
        end

        #: (untyped v) -> void
        def validate_unsigned_numeric(v)
          raise ArgumentError, "Invalid unsigned numeric: #{v.inspect}" unless v.is_a?(Numeric) && v >= 0
        end

        #: (untyped v) -> void
        def validate_family(v)
          raise ArgumentError, "Invalid font family: #{v.inspect}" unless VALID_FAMILY_VALUES.cover?(v.to_i)
        end

        #: (untyped v) -> void
        def validate_cell_u(v)
          raise ArgumentError, "Invalid cell underline: #{v.inspect}" unless VALID_CELL_U_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_pattern_type(v)
          raise ArgumentError, "Invalid pattern type: #{v.inspect}" unless VALID_PATTERN_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_gradient_type(v)
          raise ArgumentError, "Invalid gradient type: #{v.inspect}" unless VALID_GRADIENT_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_horizontal_alignment(v)
          raise ArgumentError, "Invalid horizontal alignment: #{v.inspect}" unless VALID_HORIZONTAL_ALIGNMENT_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_vertical_alignment(v)
          raise ArgumentError, "Invalid vertical alignment: #{v.inspect}" unless VALID_VERTICAL_ALIGNMENT_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_page_orientation(v)
          raise ArgumentError, "Invalid page orientation: #{v.inspect}" unless VALID_PAGE_ORIENTATION_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_pane_type(v)
          raise ArgumentError, "Invalid pane type: #{v.inspect}" unless VALID_PANE_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_split_state_type(v)
          raise ArgumentError, "Invalid split state type: #{v.inspect}" unless VALID_SPLIT_STATE_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_view_visibility(v)
          raise ArgumentError, "Invalid view visibility: #{v.inspect}" unless VALID_VIEW_VISIBILITY_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_data_validation_type(v)
          raise ArgumentError, "Invalid data validation type: #{v.inspect}" unless VALID_DATA_VALIDATION_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_data_validation_operator(v)
          raise ArgumentError, "Invalid data validation operator: #{v.inspect}" unless VALID_DATA_VALIDATION_OPERATOR_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_data_validation_error_style(v)
          raise ArgumentError, "Invalid data validation error style: #{v.inspect}" unless VALID_DATA_VALIDATION_ERROR_STYLE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_conditional_formatting_type(v)
          raise ArgumentError, "Invalid conditional formatting type: #{v.inspect}" unless VALID_CONDITIONAL_FORMATTING_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_conditional_formatting_operator(v)
          raise ArgumentError, "Invalid conditional formatting operator: #{v.inspect}" unless VALID_CONDITIONAL_FORMATTING_OPERATOR_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_conditional_formatting_value_object_type(v)
          raise ArgumentError, "Invalid cfvo type: #{v.inspect}" unless VALID_CONDITION_FORMATTING_VALUE_OBJECT_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_icon_set(v)
          raise ArgumentError, "Invalid icon set: #{v.inspect}" unless VALID_ICON_SET_VALUES.include?(v.to_s)
        end

        #: (untyped v) -> void
        def validate_time_period_type(v)
          raise ArgumentError, "Invalid time period type: #{v.inspect}" unless VALID_TIME_PERIOD_TYPE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_scatter_style(v)
          raise ArgumentError, "Invalid scatter style: #{v.inspect}" unless VALID_SCATTER_STYLE_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_marker_symbol(v)
          raise ArgumentError, "Invalid marker symbol: #{v.inspect}" unless VALID_MARKER_SYMBOL_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_display_blanks_as(v)
          raise ArgumentError, "Invalid display blank as: #{v.inspect}" unless VALID_DISPLAY_BLANK_AS_VALUES.include?(v.to_sym)
        end

        #: (untyped v) -> void
        def validate_scale_10_400(v)
          raise ArgumentError, "Scale must be between 10 and 400" unless v.is_a?(Numeric) && v.between?(10, 400)
        end

        #: (untyped v) -> void
        def validate_scale_0_10_400(v)
          raise ArgumentError, "Scale must be 0 or between 10 and 400" unless v.is_a?(Numeric) && (v.zero? || v.between?(10, 400))
        end

        #: (untyped v) -> void
        def validate_angle(v)
          raise ArgumentError, format(ERR_ANGLE, v) unless v.is_a?(Numeric) && v.between?(-90, 90)
        end

        #: (untyped v) -> void
        def validate_number_with_unit(v)
          raise ArgumentError, "Invalid number with unit: #{v.inspect}" unless v.is_a?(String) && v.match?(/\A-?[0-9]+(\.[0-9]+)?(mm|cm|in|pt|pc|pi)?\z/)
        end
      end

      # SimpleTypedList is a typed collection inheriting from Array
      # @rbs inherits Array[untyped]
      class SimpleTypedList < Array
        attr_reader :allowed_types #: Array[Class]
        attr_reader :serialize_as #: String?
        attr_accessor :locked_at #: Integer?

        DESTRUCTIVE = %w[
          replace insert collect! map! pop delete_if
          reverse! shift shuffle! slice! sort! uniq!
          unshift zip flatten! fill drop drop_while
          clear
        ].freeze

        DESTRUCTIVE.each do |m|
          undef_method m if method_defined?(m)
        end

        alias == equal?
        alias eql? equal?

        # @param type [Class, Array<Class>]
        # @param serialize_as [String, nil]
        # @param start_size [Integer]
        #: (Class | Array[Class] type, ?String? serialize_as, ?Integer start_size) -> void
        def initialize(type, serialize_as = nil, start_size = 0)
          super(start_size)
          @allowed_types = type.is_a?(Array) ? type : [type]
          @allowed_types.each do |t|
            raise ArgumentError, "All members of type must be Class objects" unless t.is_a?(Class)
          end
          @serialize_as = serialize_as
          @locked_at = nil
        end

        #: (untyped item) -> (Integer | untyped)
        def <<(item)
          DataTypeValidator.validate :SimpleTypedList_push, @allowed_types, item unless @allowed_types.size == 1 && item.is_a?(@allowed_types.first)
          super
          size - 1
        end

        #: (*untyped values) -> self
        def push(*values)
          values.each { |v| self << v }
          self
        end

        #: (*untyped others) -> self
        def concat(*others)
          others.each do |arr|
            arr.each { |item| self << item }
          end
          self
        end

        #: (untyped item) -> untyped
        def delete(item)
          return unless include?(item)
          raise ArgumentError, "Item is protected and cannot be deleted" if protected?(index(item))

          super
        end

        #: (Integer idx) -> untyped
        def delete_at(idx)
          raise ArgumentError, "Item is protected and cannot be deleted" if protected?(idx)

          super
        end

        #: (Integer idx, untyped item) -> untyped
        def []=(idx, item)
          DataTypeValidator.validate :SimpleTypedList_insert, @allowed_types, item unless @allowed_types.size == 1 && item.is_a?(@allowed_types.first)
          raise ArgumentError, "Item is protected and cannot be changed" if protected?(idx)

          super
        end

        #: (Integer idx) -> bool
        def protected?(idx)
          return false unless @locked_at.is_a?(Integer)

          idx < @locked_at
        end

        #: () -> self
        def lock
          @locked_at = size
          self
        end

        #: () -> self
        def unlock
          @locked_at = nil
          self
        end

        #: () ?{ (Integer, Integer) -> untyped } -> Array[Array[untyped]]
        def transpose(&)
          return clone if empty?

          row_count = size
          max_cols = map { |row| row.respond_to?(:cells) ? row.cells.size : row.size }.max || 0
          result = Array.new(max_cols) { Array.new(row_count) }

          row_count.times do |r_idx|
            row = self[r_idx]
            cells = row.respond_to?(:cells) ? row.cells : row
            max_cols.times do |c_idx|
              result[c_idx][r_idx] = if cells.size > c_idx
                                       cells[c_idx]
                                     elsif block_given?
                                       yield(c_idx, r_idx)
                                     end
            end
          end
          result
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          tag = @serialize_as || "items"
          str << "<#{tag} count=\"#{size}\">"
          each { |item| item.to_xml_string(str) if item.respond_to?(:to_xml_string) }
          str << "</#{tag}>"
          str
        end
      end

      # Mixin for option parsing
      module OptionsParser
        #: (Hash[Symbol, untyped] options) -> void
        def parse_options(options = {})
          options.each do |k, v|
            setter = :"#{k}="
            send(setter, v) if respond_to?(setter)
          end
        end
      end

      # Mixin for serialized attributes
      module SerializedAttributes
        def self.included(base)
          base.extend ClassMethods
        end

        # Class methods for SerializedAttributes
        module ClassMethods
          #: (*Symbol attrs) -> void
          def serializable_attributes(*attrs)
            @serializable_attrs ||= []
            @serializable_attrs.concat(attrs)
          end

          #: () -> Array[Symbol]
          def serializable_attrs
            @serializable_attrs || []
          end
        end

        #: (String str) -> String
        def serialized_attributes(str = +"")
          self.class.serializable_attrs.each do |attr|
            val = send(attr)
            next if val.nil?

            str << " #{attr}=\"#{val}\""
          end
          str
        end

        #: (String tag, String str, ?Hash[Symbol, untyped] extra_attrs) ?{ () -> void } -> String
        def serialized_tag(tag, str = +"", extra_attrs = {})
          str << "<#{tag}"
          extra_attrs.each do |k, v|
            str << " #{k}=\"#{v}\"" unless v.nil?
          end
          serialized_attributes(str)
          if block_given?
            str << ">"
            yield
            str << "</#{tag}>"
          else
            str << "/>"
          end
          str
        end
      end

      # Mixin for validating accessor helpers
      module Accessors
        def self.included(base)
          base.extend ClassMethods
        end

        # Class methods for Accessors
        module ClassMethods
          #: (*Symbol names) -> void
          def string_attr_accessor(*names)
            validated_attr_accessor(names, :validate_string)
          end

          #: (*Symbol names) -> void
          def unsigned_int_attr_accessor(*names)
            validated_attr_accessor(names, :validate_unsigned_int)
          end

          #: (*Symbol names) -> void
          def float_attr_accessor(*names)
            validated_attr_accessor(names, :validate_float)
          end

          #: (*Symbol names) -> void
          def boolean_attr_accessor(*names)
            validated_attr_accessor(names, :validate_boolean)
          end

          #: (Array[Symbol] symbols, Symbol validator) -> void
          def validated_attr_accessor(symbols, validator)
            symbols.each do |symbol|
              attr_reader symbol

              define_method(:"#{symbol}=") do |val|
                Caxlsx.send(validator, val) unless val.nil?
                val = Caxlsx.booleanize(val) if validator == :validate_boolean && !val.nil?
                instance_variable_set(:"@#{symbol}", val)
              end
            end
          end
        end
      end

      # Validators
      module DataTypeValidator
        # Validates an object against allowed types
        # @param name [Symbol, String]
        # @param types [Class, Array<Class>]
        # @param v [Object]
        #: (Symbol | String name, Class | Array[Class] types, untyped v) -> void
        def self.validate(name, types, v)
          allowed = types.is_a?(Array) ? types : [types]
          return if allowed.any? { |t| v.is_a?(t) }

          raise ArgumentError, "#{name} must be one of #{allowed.map(&:name).join(", ")}, got #{v.class.name}"
        end
      end

      module RestrictionValidator
        # Validates that a value is in an allowed collection
        # @param name [Symbol, String]
        # @param allowed [Array, Range]
        # @param v [Object]
        #: (Symbol | String name, Array[untyped] | Range[untyped] allowed, untyped v) -> void
        def self.validate(name, allowed, v)
          return if allowed.include?(v)

          raise ArgumentError, "#{name} must be one of #{allowed.inspect}, got #{v.inspect}"
        end
      end

      module RangeValidator
        # Validates that a value is within a numeric range
        # @param name [Symbol, String]
        # @param min [Numeric]
        # @param max [Numeric]
        # @param v [Numeric]
        #: (Symbol | String name, Numeric min, Numeric max, untyped v) -> void
        def self.validate(name, min, max, v)
          raise ArgumentError, "#{name} must be between #{min} and #{max}, got #{v.inspect}" unless v.is_a?(Numeric) && v.between?(min, max)
        end
      end

      module RegexValidator
        # Validates that a value matches a regular expression
        # @param name [Symbol, String]
        # @param regex [Regexp]
        # @param v [Object]
        #: (Symbol | String name, Regexp regex, untyped v) -> void
        def self.validate(name, regex, v)
          raise ArgumentError, "#{name} must match #{regex.inspect}, got #{v.inspect}" unless v.to_s.match?(regex)
        end
      end
    end
  end
end
