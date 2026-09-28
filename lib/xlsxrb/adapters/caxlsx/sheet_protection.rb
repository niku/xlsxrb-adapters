# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # The Protected Range class represents a set of cells in the worksheet.
      class ProtectedRange
        include OptionsParser
        include SerializedAttributes

        serializable_attributes :sqref, :name

        # @return [String, nil]
        attr_reader :sqref

        # @return [String, nil]
        attr_reader :name

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) ?{ (ProtectedRange) -> void } -> void
        def initialize(options = {})
          parse_options(options)
          yield self if block_given?
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def sqref=(v)
          Caxlsx.validate_string(v)
          @sqref = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def name=(v)
          Caxlsx.validate_string(v)
          @name = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("protectedRange", str)
        end
      end

      # A collection of ranges that should be protected in the worksheet.
      class ProtectedRanges < SimpleTypedList
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "You must provide a worksheet" unless worksheet.is_a?(Worksheet)

          super(ProtectedRange)
          @worksheet = worksheet
        end

        # @param cells [String, SimpleTypedList, Array[Cell]]
        # @return [ProtectedRange]
        #: (String | SimpleTypedList | Array[Cell] cells) -> ProtectedRange
        def add_range(cells)
          sqref = if cells.is_a?(String)
                    cells
                  elsif cells.is_a?(SimpleTypedList) || cells.is_a?(Array)
                    Caxlsx.cell_range(cells, false)
                  end
          self << ProtectedRange.new(sqref: sqref, name: "Range#{size}")
          last
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<protectedRanges>"
          each { |range| range.to_xml_string(str) }
          str << "</protectedRanges>"
        end
      end

      # The SheetProtection object manages worksheet protection options per sheet.
      class SheetProtection
        include OptionsParser
        include SerializedAttributes
        include Accessors

        boolean_attr_accessor :sheet, :objects, :scenarios, :format_cells, :format_columns, :format_rows,
                              :insert_columns, :insert_rows, :insert_hyperlinks, :delete_columns, :delete_rows,
                              :select_locked_cells, :sort, :auto_filter, :pivot_tables, :select_unlocked_cells

        serializable_attributes :sheet, :objects, :scenarios, :format_cells, :format_columns, :format_rows,
                                :insert_columns, :insert_rows, :insert_hyperlinks, :delete_columns, :delete_rows,
                                :select_locked_cells, :sort, :auto_filter, :pivot_tables, :select_unlocked_cells,
                                :salt, :password

        # @return [String, nil]
        attr_reader :salt_value

        # @return [String, nil]
        attr_reader :password

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @objects = @scenarios = @select_locked_cells = @select_unlocked_cells = false
          @sheet = @format_cells = @format_rows = @format_columns = @insert_columns = @insert_rows = @insert_hyperlinks = @delete_columns = @delete_rows = @sort = @auto_filter = @pivot_tables = true
          @password = nil
          parse_options(options)
        end

        # @param v [String, nil]
        # @return [void]
        #: (String? v) -> void
        def password=(v)
          return if v.nil?

          @password = create_password_hash(v)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("sheetProtection", str)
        end

        private

        # @param password [String]
        # @return [String]
        #: (String password) -> String
        def create_password_hash(password)
          encoded_password = encode_password(password)

          password_as_hex = [encoded_password].pack("v")
          password_as_string = password_as_hex.unpack1("H*").upcase

          password_as_string[2..3] + password_as_string[0..1]
        end

        # @param password [String]
        # @return [Integer]
        #: (String password) -> Integer
        def encode_password(password)
          i = 0
          chars = password.chars
          count = chars.size

          chars.map! do |char|
            i += 1
            c_val = (char.unpack1("c") || 0) << i
            low_15 = c_val & 0x7fff
            high_15 = (c_val & (0x7fff << 15)) >> 15
            low_15 | high_15
          end

          encoded_password = 0x0000
          chars.each { |c| encoded_password ^= c }
          encoded_password ^= count
          encoded_password ^ 0xCE4B
        end
      end
    end
  end
end
