# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Writexlsx
      # Manages standard Excel palette colors and RGB color conversions
      class Colors
        COLORS = {
          aqua: 0x0F,
          cyan: 0x0F,
          black: 0x08,
          blue: 0x0C,
          brown: 0x10,
          magenta: 0x0E,
          fuchsia: 0x0E,
          gray: 0x17,
          grey: 0x17,
          green: 0x11,
          lime: 0x0B,
          navy: 0x12,
          orange: 0x35,
          pink: 0x21,
          purple: 0x14,
          red: 0x0A,
          silver: 0x16,
          white: 0x09,
          yellow: 0x0D,
          automatic: 0x40
        }.freeze

        # Standard Excel 56-color palette (indices 8..63)
        DEFAULT_PALETTE = [
          [0x00, 0x00, 0x00], # 8
          [0xFF, 0xFF, 0xFF], # 9
          [0xFF, 0x00, 0x00], # 10
          [0x00, 0xFF, 0x00], # 11
          [0x00, 0x00, 0xFF], # 12
          [0xFF, 0xFF, 0x00], # 13
          [0xFF, 0x00, 0xFF], # 14
          [0x00, 0xFF, 0xFF], # 15
          [0x80, 0x00, 0x00], # 16
          [0x00, 0x80, 0x00], # 17
          [0x00, 0x00, 0x80], # 18
          [0x80, 0x80, 0x00], # 19
          [0x80, 0x00, 0x80], # 20
          [0x00, 0x80, 0x80], # 21
          [0xC0, 0xC0, 0xC0], # 22
          [0x80, 0x80, 0x80], # 23
          [0x99, 0x99, 0xFF], # 24
          [0x99, 0x33, 0x66], # 25
          [0xFF, 0xFF, 0xCC], # 26
          [0xCC, 0xFF, 0xFF], # 27
          [0x66, 0x00, 0x66], # 28
          [0xFF, 0x80, 0x80], # 29
          [0x00, 0x66, 0xCC], # 30
          [0xCC, 0xCC, 0xFF], # 31
          [0x00, 0x00, 0x80], # 32
          [0xFF, 0x00, 0xFF], # 33
          [0xFF, 0xFF, 0x00], # 34
          [0x00, 0xFF, 0xFF], # 35
          [0x80, 0x00, 0x80], # 36
          [0x80, 0x00, 0x00], # 37
          [0x00, 0x80, 0x80], # 38
          [0x00, 0x00, 0xFF], # 39
          [0x00, 0xCC, 0xFF], # 40
          [0xCC, 0xFF, 0xFF], # 41
          [0xCC, 0xFF, 0xCC], # 42
          [0xFF, 0xFF, 0x99], # 43
          [0x99, 0xCC, 0xFF], # 44
          [0xFF, 0x99, 0xCC], # 45
          [0xCC, 0x99, 0xFF], # 46
          [0xFF, 0xCC, 0x99], # 47
          [0x33, 0x66, 0xFF], # 48
          [0x33, 0xCC, 0xCC], # 49
          [0x99, 0xCC, 0x00], # 50
          [0xFF, 0xCC, 0x00], # 51
          [0xFF, 0x99, 0x00], # 52
          [0xFF, 0x66, 0x00], # 53
          [0x66, 0x66, 0x99], # 54
          [0x96, 0x96, 0x96], # 55
          [0x00, 0x33, 0x66], # 56
          [0x33, 0x99, 0x66], # 57
          [0x00, 0x33, 0x00], # 58
          [0x33, 0x33, 0x00], # 59
          [0x99, 0x33, 0x00], # 60
          [0x99, 0x33, 0x66], # 61
          [0x33, 0x33, 0x99], # 62
          [0x33, 0x33, 0x33]  # 63
        ].freeze

        # @return [Array<Array<Integer>>]
        attr_reader :palette

        #: () -> void
        def initialize
          @palette = DEFAULT_PALETTE.map(&:dup)
        end

        # Converts a color argument into an integer color index (0..63) or hex string
        #
        # @param color_code [Integer, String, Symbol, nil]
        # @return [Integer, String]
        #: (untyped color_code) -> untyped
        def color(color_code = nil)
          return 0x7FFF if color_code.nil?

          case color_code
          when Integer
            if color_code.negative? || color_code > 63
              0x7FFF
            elsif color_code < 8
              color_code + 8
            else
              color_code
            end
          when Symbol
            COLORS[color_code.downcase] || 0x7FFF
          when String
            clean = color_code.strip
            if clean =~ /^#?[0-9a-fA-F]{6}$/
              clean.start_with?("#") ? clean : "##{clean}"
            else
              sym = clean.downcase.to_sym
              COLORS[sym] || 0x7FFF
            end
          else
            0x7FFF
          end
        end

        # Converts a color representation to an OpenXML 6-character hex string ("RRGGBB")
        #
        # @param color_val [Integer, String, Symbol, nil]
        # @return [String, nil]
        #: (untyped color_val) -> String?
        def to_hex(color_val)
          return nil if color_val.nil?

          if color_val.is_a?(Integer)
            if color_val.between?(8, 63)
              r, g, b = @palette[color_val - 8] || DEFAULT_PALETTE[color_val - 8]
              return [r, g, b].map { |val| val.to_s(16).rjust(2, "0").upcase }.join
            elsif color_val.between?(0, 0xFFFFFF)
              return color_val.to_s(16).rjust(6, "0").upcase
            end
            return nil
          end

          str = color_val.to_s.strip
          if str =~ /^#?([0-9a-fA-F]{6})$/
            return ::Regexp.last_match(1).upcase
          elsif str =~ /^#?([0-9a-fA-F]{3})$/
            chars = ::Regexp.last_match(1).chars
            return "#{chars[0]}#{chars[0]}#{chars[1]}#{chars[1]}#{chars[2]}#{chars[2]}".upcase
          end

          sym = str.downcase.to_sym
          if COLORS.key?(sym)
            idx = COLORS[sym]
            if idx.between?(8, 63)
              r, g, b = @palette[idx - 8] || DEFAULT_PALETTE[idx - 8]
              return [r, g, b].map { |val| val.to_s(16).rjust(2, "0").upcase }.join
            end
          end

          if defined?(Xlsxrb.to_hex_color) && (core_hex = Xlsxrb.to_hex_color(color_val))
            return core_hex
          end

          return Xlsxrb::StyleBuilder::CSS_COLORS[sym] if defined?(Xlsxrb::StyleBuilder::CSS_COLORS) && Xlsxrb::StyleBuilder::CSS_COLORS[sym]

          nil
        end

        # Class-level helper to resolve color to hex
        #
        # @param color_val [Integer, String, Symbol, nil]
        # @return [String, nil]
        #: (untyped color_val) -> String?
        def self.to_hex(color_val)
          @default_instance ||= new
          @default_instance.to_hex(color_val)
        end
      end
    end
  end
end
