# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Color helper and validation compatible with RubyXL::Color.
      class Color
        COLOR_REGEX = /\A(?:[a-f0-9]{6}|[a-f0-9]{8})\Z/i

        # Validates that color is a 6- or 8-character hex string without leading #.
        #
        # @param color [Object]
        # @return [true]
        #: (untyped color) -> true
        def self.validate_color(color)
          return true if color.is_a?(String) && color.match?(COLOR_REGEX)

          raise "Invalid color code: #{color.inspect}. Color must be a 6 or 8-digit hex string without #"
        end

        # Converts String, Hash, or Color to a Color instance.
        #
        # @param color [Object]
        # @return [Color, nil]
        #: (untyped color) -> Color?
        def self.from(color)
          case color
          when Color
            color
          when String
            new(rgb: color)
          when Hash
            rgb_val = color[:rgb] || color[:color]
            new(rgb: rgb_val&.to_s, theme: color[:theme]&.to_i, tint: color[:tint]&.to_f)
          end
        end

        attr_accessor :rgb #: String?
        attr_accessor :theme #: Integer?
        attr_accessor :tint #: Float?

        # @param rgb [String, nil]
        # @param theme [Integer, nil]
        # @param tint [Float, nil]
        #: (?rgb: String?, ?theme: Integer?, ?tint: Float?) -> void
        def initialize(rgb: nil, theme: nil, tint: nil)
          @rgb = rgb
          @theme = theme
          @tint = tint
        end

        # @return [String?]
        #: () -> String?
        def to_s
          @rgb
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(Color)

          (@rgb&.downcase == other.rgb&.downcase) &&
            (@theme == other.theme) &&
            (@tint == other.tint)
        end
      end

      # Compatibility class for RubyXL::RgbColor
      class RgbColor
        attr_accessor :r #: Integer?
        attr_accessor :g #: Integer?
        attr_accessor :b #: Integer?
        attr_accessor :a #: Integer?

        # @return [String]
        #: () -> String
        def to_s
          str = format("%<r>02x%<g>02x%<b>02x", r: @r || 0, g: @g || 0, b: @b || 0)
          str += format("%<a>02x", a: @a) if @a
          str
        end
      end
    end
  end
end
