# frozen_string_literal: true

# rbs_inline: enabled

require_relative "constants"
require_relative "colors"

module Xlsxrb
  module Adapters
    module Writexlsx
      # Represents cell formatting properties in WriteXLSX
      class Format
        # @return [String, nil]
        attr_accessor :font

        # @return [Numeric, nil]
        attr_accessor :size

        # @return [Integer, String, nil]
        attr_accessor :font_color

        # @return [Integer, Boolean, nil]
        attr_accessor :bold

        # @return [Integer, Boolean, nil]
        attr_accessor :italic

        # @return [Integer, Boolean, nil]
        attr_accessor :underline

        # @return [Integer, Boolean, nil]
        attr_accessor :font_strikeout

        # @return [Integer, Boolean, nil]
        attr_accessor :font_outline

        # @return [Integer, Boolean, nil]
        attr_accessor :font_shadow

        # @return [Integer, nil]
        attr_accessor :font_script

        # @return [Integer, nil]
        attr_accessor :font_family

        # @return [Integer, nil]
        attr_accessor :font_charset

        # @return [Integer, nil]
        attr_accessor :font_condense

        # @return [Integer, nil]
        attr_accessor :font_extend

        # @return [Integer, nil]
        attr_accessor :text_h_align
        alias align text_h_align
        alias align= text_h_align=

        # @return [Integer, nil]
        attr_accessor :text_v_align
        alias valign text_v_align
        alias valign= text_v_align=

        # @return [Integer, nil]
        attr_accessor :rotation

        # @return [Integer, Boolean, nil]
        attr_accessor :text_wrap

        # @return [Integer, Boolean, nil]
        attr_accessor :text_justlast

        # @return [Integer, nil]
        attr_accessor :indent

        # @return [Integer, Boolean, nil]
        attr_accessor :shrink

        # @return [Integer, nil]
        attr_accessor :reading_order

        # @return [String, nil]
        attr_accessor :num_format

        # @return [Integer, nil]
        attr_accessor :num_format_index

        # @return [Integer, String, nil]
        attr_accessor :pattern

        # @return [Integer, String, nil]
        attr_accessor :bg_color

        # @return [Integer, String, nil]
        attr_accessor :fg_color

        # @return [Integer, Boolean, nil]
        attr_accessor :has_fill

        # @return [Integer, Symbol, nil]
        attr_accessor :bottom
        alias border bottom

        # @return [Integer, Symbol, nil]
        attr_accessor :top

        # @return [Integer, Symbol, nil]
        attr_accessor :left

        # @return [Integer, Symbol, nil]
        attr_accessor :right

        # @return [Integer, String, nil]
        attr_accessor :bottom_color
        alias border_color bottom_color

        # @return [Integer, String, nil]
        attr_accessor :top_color

        # @return [Integer, String, nil]
        attr_accessor :left_color

        # @return [Integer, String, nil]
        attr_accessor :right_color

        # @return [Integer, nil]
        attr_accessor :diag_border

        # @return [Integer, String, nil]
        attr_accessor :diag_color

        # @return [Integer, nil]
        attr_accessor :diag_type

        # @return [Integer, Boolean, nil]
        attr_accessor :locked

        # @return [Integer, Boolean, nil]
        attr_accessor :hidden

        # @return [String, nil]
        attr_accessor :hyperlink

        # @return [Integer, Boolean, nil]
        attr_accessor :quote_prefix

        # @return [Integer, nil]
        attr_accessor :xf_index

        # @return [Integer, nil]
        attr_accessor :dxf_index

        # @return [Integer, nil]
        attr_accessor :xf_id

        # @return [Integer, nil]
        attr_accessor :theme

        # @param formats [Object, nil]
        # @param params [Hash{Symbol => untyped}]
        #: (?untyped formats, ?Hash[Symbol, untyped] params) -> void
        def initialize(formats = nil, params = {})
          @formats = formats
          @colors = Colors.new

          @font = "Calibri"
          @size = 11
          @font_color = nil
          @bold = nil
          @italic = nil
          @underline = nil
          @font_strikeout = nil
          @font_outline = nil
          @font_shadow = nil
          @font_script = nil
          @font_family = nil
          @font_charset = nil
          @font_condense = nil
          @font_extend = nil

          @text_h_align = nil
          @text_v_align = nil
          @rotation = nil
          @text_wrap = nil
          @text_justlast = nil
          @indent = nil
          @shrink = nil
          @reading_order = nil

          @num_format = nil
          @num_format_index = nil

          @pattern = nil
          @bg_color = nil
          @fg_color = nil
          @has_fill = nil

          @bottom = nil
          @top = nil
          @left = nil
          @right = nil
          @bottom_color = nil
          @top_color = nil
          @left_color = nil
          @right_color = nil
          @diag_border = nil
          @diag_color = nil
          @diag_type = nil

          @locked = nil
          @hidden = nil
          @hyperlink = nil
          @quote_prefix = nil
          @xf_index = nil
          @dxf_index = nil
          @xf_id = nil
          @theme = nil

          set_format_properties(params) unless params.empty?
        end

        # Copies attributes from another Format
        #
        # @param other [Format]
        # @return [self]
        #: (Format other) -> self
        def copy(other)
          other.instance_variables.each do |var|
            next if %i[@formats @colors].include?(var)

            val = other.instance_variable_get(var)
            instance_variable_set(var, val)
          end
          self
        end

        # Sets multiple properties on format
        #
        # @param properties [Array<Hash{Symbol => untyped}>]
        # @return [self]
        def set_format_properties(*properties)
          properties.each do |property|
            next unless property.is_a?(Hash)

            property.each do |key, value|
              setter = :"set_#{key}"
              if respond_to?(setter)
                send(setter, value)
              elsif respond_to?(:"#{key}=")
                send(:"#{key}=", value)
              end
            end
          end
          self
        end

        # Resolves color code through Colors helper
        #
        # @param color_code [untyped]
        # @return [Integer, String]
        #: (untyped color_code) -> untyped
        def color(color_code)
          @colors.color(color_code)
        end

        # Font setters & aliases
        def set_font(val)
          @font = val&.to_s
        end

        def set_size(val)
          @size = val
        end

        def set_color(val)
          @font_color = color(val)
        end

        def set_bold(val = 1)
          @bold = val
        end

        def set_italic(val = 1)
          @italic = val
        end

        def set_underline(val = 1)
          @underline = val
        end

        def set_font_strikeout(val = 1)
          @font_strikeout = val
        end
        alias set_strikeout set_font_strikeout
        alias strikeout font_strikeout
        alias strikeout= font_strikeout=

        def set_font_outline(val = 1)
          @font_outline = val
        end

        def set_font_shadow(val = 1)
          @font_shadow = val
        end

        def set_font_script(val)
          @font_script = val
        end

        def set_font_family(val)
          @font_family = val
        end

        def set_font_charset(val)
          @font_charset = val
        end

        def set_font_condense(val = 1)
          @font_condense = val
        end

        def set_font_extend(val = 1)
          @font_extend = val
        end

        # Alignment setters & aliases
        def set_text_h_align(val)
          @text_h_align = val
        end

        def set_text_v_align(val)
          @text_v_align = val
        end

        def set_align(location)
          return unless location

          loc = location.to_s.downcase
          case loc
          when "left" then set_text_h_align(1)
          when "centre", "center" then set_text_h_align(2)
          when "right" then set_text_h_align(3)
          when "fill" then set_text_h_align(4)
          when "justify" then set_text_h_align(5)
          when "center_across", "centre_across", "merge" then set_text_h_align(6)
          when "distributed", "equal_space", "justify_distributed" then set_text_h_align(7)
          when "top" then set_text_v_align(1)
          when "vcentre", "vcenter" then set_text_v_align(2)
          when "bottom" then set_text_v_align(3)
          when "vjustify" then set_text_v_align(4)
          when "vdistributed", "vequal_space" then set_text_v_align(5)
          end
        end

        def set_valign(location)
          set_align(location)
        end

        def set_center_across(_val = 1)
          set_text_h_align(6)
        end

        def set_merge(_val = 1)
          set_text_h_align(6)
        end

        def set_rotation(rot)
          r = rot.to_i
          if r == 270
            @rotation = 255
          elsif r.between?(-90, 90)
            @rotation = r.negative? ? -r + 90 : r
          else
            raise "Rotation #{rot} outside range: -90 <= angle <= 90"
          end
        end

        def set_text_wrap(val = 1)
          @text_wrap = val
        end

        def set_text_justlast(val = 1)
          @text_justlast = val
        end

        def set_indent(val = 1)
          @indent = val
        end

        def set_shrink(val = 1)
          @shrink = val
        end

        def set_reading_order(val)
          @reading_order = val
        end

        # Number format setters
        def set_num_format(fmt)
          @num_format = fmt&.to_s
        end

        def set_num_format_index(val)
          @num_format_index = val
        end

        # Fill setters
        def set_fg_color(val)
          @fg_color = color(val)
          @has_fill = 1
        end

        def set_bg_color(val)
          @bg_color = color(val)
          @has_fill = 1
        end

        def set_pattern(val)
          @pattern = val
          @has_fill = 1
        end

        def set_has_fill(val = 1)
          @has_fill = val
        end

        # Border setters
        def set_border(val)
          set_bottom(val)
          set_top(val)
          set_left(val)
          set_right(val)
        end

        def set_border_color(val)
          c = color(val)
          @bottom_color = c
          @top_color = c
          @left_color = c
          @right_color = c
        end

        def set_bottom(val)
          @bottom = val
        end

        def set_top(val)
          @top = val
        end

        def set_left(val)
          @left = val
        end

        def set_right(val)
          @right = val
        end

        def set_bottom_color(val)
          @bottom_color = color(val)
        end

        def set_top_color(val)
          @top_color = color(val)
        end

        def set_left_color(val)
          @left_color = color(val)
        end

        def set_right_color(val)
          @right_color = color(val)
        end

        def set_diag_border(val)
          @diag_border = val
        end

        def set_diag_color(val)
          @diag_color = color(val)
        end

        def set_diag_type(val)
          @diag_type = val
        end

        # Protection setters
        def set_locked(val = 1)
          @locked = val
        end

        def set_hidden(val = 1)
          @hidden = val
        end

        # Hyperlink & misc setters
        def set_hyperlink(val)
          @hyperlink = val
          set_underline(1)
          set_color("blue")
        end

        def set_quote_prefix(val = 1)
          @quote_prefix = val
        end

        def set_xf_index(val)
          @xf_index = val
        end

        def set_dxf_index(val)
          @dxf_index = val
        end

        def set_xf_id(val)
          @xf_id = val
        end

        def set_theme(val)
          @theme = val
        end

        # Predicates matching write_xlsx methods
        def bold?
          [1, true].include?(@bold)
        end

        def italic?
          [1, true].include?(@italic)
        end

        def underline?
          !@underline.nil? && @underline != 0 && @underline != false
        end

        def strikeout?
          [1, true].include?(@font_strikeout)
        end

        def outline?
          [1, true].include?(@font_outline)
        end

        def shadow?
          [1, true].include?(@font_shadow)
        end

        def has_border?
          !@bottom.nil? || !@top.nil? || !@left.nil? || !@right.nil?
        end

        def has_fill?
          !@has_fill.nil? || !@bg_color.nil? || !@fg_color.nil? || !@pattern.nil?
        end

        def has_font?
          bold? || italic? || underline? || strikeout? || !@font_color.nil? || (@size && @size != 11) || (@font && @font != "Calibri")
        end

        def force_text_format?
          @num_format == "@" || @num_format_index == 49
        end

        # Returns border style proxy
        def border_style
          self
        end

        # Hash-like property access
        def [](key)
          k = key.to_sym
          respond_to?(k) ? send(k) : nil
        end

        def []=(key, val)
          setter = :"#{key}="
          send(setter, val) if respond_to?(setter)
        end

        # Converts format properties to a style hash for xlsxrb
        #
        # @return [Hash{Symbol => untyped}]
        def to_style_hash
          f = {
            name: @font || "Calibri",
            sz: @size || 11,
            bold: bold? || nil,
            italic: italic? || nil,
            underline: underline? || nil,
            strike: strikeout? || nil,
            color: Colors.to_hex(@font_color)
          }.compact

          h_align = case @text_h_align
                    when 1 then "left"
                    when 2 then "center"
                    when 3 then "right"
                    when 4 then "fill"
                    when 5 then "justify"
                    when 6 then "centerContinuous"
                    when 7 then "distributed"
                    end

          v_align = case @text_v_align
                    when 1 then "top"
                    when 2 then "center"
                    when 3 then "bottom"
                    when 4 then "justify"
                    when 5 then "distributed"
                    end

          align = {
            horizontal: h_align,
            vertical: v_align,
            wrap_text: [1, true].include?(@text_wrap) || nil,
            text_rotation: @rotation,
            indent: @indent,
            shrink_to_fit: [1, true].include?(@shrink) || nil
          }.compact

          {
            font: f,
            alignment: align
          }
        end

        # Converts format font properties to a hash suitable for Xlsxrb rich text and style builder
        #
        # @return [Hash{Symbol => untyped}]
        #: () -> Hash[Symbol, untyped]
        def to_font_hash
          {
            name: @font,
            sz: @size,
            bold: bold? || nil,
            italic: italic? || nil,
            underline: underline? || nil,
            strike: strikeout? || nil,
            color: Colors.to_hex(@font_color)
          }.compact
        end

        alias to_xlsxrb_style to_style_hash
      end
    end
  end
end
