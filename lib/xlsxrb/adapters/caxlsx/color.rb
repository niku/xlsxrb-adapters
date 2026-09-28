# frozen_string_literal: true

# rbs_inline: enabled

require_relative "util"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Represents a color used for fonts, fills, and borders
      class Color
        include OptionsParser
        include SerializedAttributes

        attr_reader :auto #: bool?
        attr_reader :rgb #: String?
        attr_reader :tint #: Float?

        serializable_attributes :auto, :rgb, :tint

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @rgb = "FF000000"
          parse_options options
        end

        # Converts Color, Hash, String or object responding to rgb to a Color instance.
        #
        # @param val [Object]
        # @return [Color, nil]
        #: (untyped val) -> Color?
        def self.from(val)
          return nil if val.nil?
          return val if val.is_a?(Color)

          if val.is_a?(Hash)
            rgb = val[:rgb] || val["rgb"] || val[:color] || val["color"]
            tint = val[:tint] || val["tint"]
            auto = val[:auto] || val["auto"]
            return nil if rgb.nil? && tint.nil? && auto.nil?

            opts = {}
            opts[:rgb] = rgb if rgb
            opts[:tint] = tint if tint
            opts[:auto] = auto unless auto.nil?
            return new(opts)
          end

          return new(rgb: val.rgb) if val.respond_to?(:rgb) && !val.is_a?(String)

          new(rgb: val.to_s)
        end

        #: (untyped v) -> void
        def auto=(v)
          Caxlsx.validate_boolean(v)
          @auto = Caxlsx.booleanize(v)
        end

        #: (untyped v) -> void
        def rgb=(v)
          v = v[:rgb] || v["rgb"] if v.is_a?(Hash)
          v = v.rgb if v.respond_to?(:rgb) && !v.is_a?(String)
          return if v.nil?

          Caxlsx.validate_string(v)
          v = v.to_s.upcase
          v *= 3 if v.size == 2
          v = v.rjust(8, "FF")
          raise ArgumentError, "Invalid color rgb value: #{v}." unless /[0-9A-F]{8}/.match?(v)

          @rgb = v
        end

        #: (untyped v) -> void
        def tint=(v)
          Caxlsx.validate_float(v)
          @tint = v.to_f
        end

        #: (?String str, ?String tag_name) -> String
        def to_xml_string(str = +"", tag_name = "color")
          serialized_tag(tag_name.to_s, str)
        end
      end
    end
  end
end
