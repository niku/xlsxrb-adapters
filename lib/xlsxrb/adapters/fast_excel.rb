# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require_relative "fast_excel/constants"
require_relative "fast_excel/formula"
require_relative "fast_excel/url"
require_relative "fast_excel/format"
require_relative "fast_excel/worksheet"
require_relative "fast_excel/workbook"

module Xlsxrb
  module Adapters
    # FastExcel adapter module providing a drop-in compatible interface for fast_excel
    # (FFI / libxlsxwriter wrapper) without hijacking the global ::FastExcel namespace by default.
    module FastExcel
      # Opens or creates a new workbook compatible with FastExcel.open.
      #
      # @param filename [String, Pathname, nil] File path to save workbook to. If nil, creates a temporary file.
      # @param constant_memory [Boolean] Whether to enforce sequential constant-memory writing constraints.
      # @param default_format [Hash{Symbol => untyped}, nil] Optional default format options.
      # @yield [workbook]
      # @yieldparam workbook [Workbook]
      # @return [Workbook, untyped]
      #: (?String? filename, ?constant_memory: bool, ?default_format: Hash[Symbol, untyped]?) ?{ (Workbook) -> untyped } -> untyped
      def self.open(filename = nil, constant_memory: false, default_format: nil, &block)
        wb = Workbook.new(filename, constant_memory: constant_memory, default_format: default_format)

        if block
          begin
            yield wb
          ensure
            wb.close
          end
        else
          wb
        end
      end

      # Converts a Time or DateTime object into an Excel 1900 date system serial number.
      #
      # @param time [Time, DateTime, Numeric]
      # @param offset [Numeric, nil] Timezone offset in seconds (defaults to time.utc_offset)
      # @return [Float]
      #: (untyped time, ?Numeric? offset) -> Float
      def self.date_num(time, offset = nil)
        offset ||= if time.respond_to?(:utc_offset)
                     time.utc_offset
                   elsif Time.respond_to?(:zone) && !Time.zone.nil?
                     Time.zone.utc_offset
                   else
                     0
                   end

        time_f = if time.respond_to?(:to_time)
                   time.to_time.to_f
                 else
                   time.to_f
                 end

        (time_f / XLSX_DATE_DAY) + XLSX_DATE_EPOCH_DIFF + (offset.to_f / XLSX_DATE_DAY)
      end

      # Converts a Date or DateTime object into a Datetime struct.
      #
      # @param time [Date, DateTime, Time]
      # @return [Datetime]
      #: (untyped time) -> Datetime
      def self.lxw_datetime(time)
        Datetime.new(
          time.year,
          time.month,
          time.day,
          time.respond_to?(:hour) ? time.hour : 0,
          if time.respond_to?(:minute)
            time.minute
          else
            (time.respond_to?(:min) ? time.min : 0)
          end,
          if time.respond_to?(:second)
            time.second
          else
            (time.respond_to?(:sec) ? time.sec : 0)
          end
        )
      end

      # Converts a Time object into a Datetime struct.
      #
      # @param time [Time]
      # @return [Datetime]
      #: (Time time) -> Datetime
      def self.lxw_time(time)
        Datetime.new(
          time.year,
          time.month,
          time.day,
          time.hour,
          time.min,
          time.sec
        )
      end

      # Converts a color name, hex string, symbol, or hex number into an integer RGB hex number.
      #
      # @param value [String, Symbol, Numeric]
      # @return [Integer]
      #: (untyped value) -> Integer
      def self.color_to_hex(value)
        orig_value = value
        val_str = value.to_s if value.is_a?(Symbol)
        val_str = value if value.is_a?(String)

        if val_str
          sym = val_str.to_sym
          if EXTRA_COLORS.key?(sym)
            return EXTRA_COLORS[sym]
          elsif COLOR_ENUM.find(sym)
            return COLOR_ENUM.find(sym)
          elsif COLOR_ENUM.find(:"color_#{sym}")
            return COLOR_ENUM.find(:"color_#{sym}")
          elsif val_str =~ /^#?(0x)?([\da-f]){6}$/i
            cleaned = val_str.sub("#", "")
            return cleaned.start_with?("0x") ? cleaned.to_i(16) : "0x#{cleaned}".to_i(16)
          else
            raise ArgumentError, "Unknown color value #{orig_value.inspect}, expected hex string or color name"
          end
        end

        return value.to_i if value.is_a?(Numeric)

        raise ArgumentError, "Can not use #{value.class} (#{value.inspect}) for color value, expected String or Hex Number"
      end

      # Debug printing helper matching FastExcel.print_ffi_obj
      #
      # @param value [Object]
      # @param do_print [Boolean]
      # @param offset [String]
      # @param deep [Boolean]
      # @return [String, nil]
      def self.print_ffi_obj(value, do_print: true, offset: "", _deep: false)
        result = value.class.to_s
        if value.respond_to?(:members)
          value.members.each do |key|
            fval = value[key]
            result += "\n#{offset}* #{key}: #{fval.inspect}"
          end
        end

        if do_print
          puts result
          nil
        else
          result
        end
      end

      # Converts an adapter Workbook to an immutable Xlsxrb::Elements::Workbook.
      #
      # @param obj [Workbook]
      # @return [Xlsxrb::Elements::Workbook]
      #: (Workbook obj) -> Xlsxrb::Elements::Workbook
      def self.to_xlsxrb(obj)
        obj.to_xlsxrb
      end

      # Converts an Xlsxrb::Elements::Workbook into a FastExcel Workbook adapter instance.
      #
      # @param xlsxrb_workbook [Xlsxrb::Elements::Workbook]
      # @param filename [String, nil]
      # @param constant_memory [Boolean]
      # @return [Workbook]
      #: (Xlsxrb::Elements::Workbook xlsxrb_workbook, ?String? filename, ?constant_memory: bool) -> Workbook
      def self.from_xlsxrb(xlsxrb_workbook, filename = nil, constant_memory: false)
        Workbook.from_xlsxrb(xlsxrb_workbook, filename, constant_memory: constant_memory)
      end
    end
  end
end
