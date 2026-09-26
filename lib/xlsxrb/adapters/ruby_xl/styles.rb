# frozen_string_literal: true

# rbs_inline: enabled

require_relative "color"

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a Font definition.
      class Font
        attr_accessor :name #: String
        attr_accessor :size #: Float | Integer
        attr_accessor :color #: Color?
        attr_accessor :bold #: bool?
        attr_accessor :italic #: bool?
        attr_accessor :underline #: String | bool | nil
        attr_accessor :strikethrough #: bool?

        # @param name [String]
        # @param size [Float, Integer]
        # @param color [Color, String, nil]
        # @param bold [Boolean, nil]
        # @param italic [Boolean, nil]
        # @param underline [String, Boolean, nil]
        # @param strikethrough [Boolean, nil]
        #: (?name: String, ?size: (Float | Integer), ?color: (Color | String)?, ?bold: bool?, ?italic: bool?, ?underline: (String | bool)?, ?strikethrough: bool?) -> void
        def initialize(name: "Verdana", size: 10, color: nil, bold: nil, italic: nil, underline: nil, strikethrough: nil)
          @name = name
          @size = size
          @color = Color.from(color)
          @bold = bold
          @italic = italic
          @underline = underline
          @strikethrough = strikethrough
        end

        #: (String new_name) -> void
        def set_name(new_name)
          @name = new_name
        end

        #: () -> String
        def get_name
          @name
        end

        #: (Float | Integer new_size) -> void
        def set_size(new_size)
          @size = new_size
        end

        #: () -> (Float | Integer)
        def get_size
          @size
        end

        #: ((String | Color | Hash[untyped, untyped])? new_color) -> void
        def set_rgb_color(new_color)
          @color = Color.from(new_color)
        end

        #: () -> String?
        def get_rgb_color
          @color&.rgb
        end

        #: (bool? val) -> void
        def set_bold(val)
          @bold = val ? true : nil
        end

        #: () -> bool?
        def is_bold
          @bold
        end

        #: (bool? val) -> void
        def set_italic(val)
          @italic = val ? true : nil
        end

        #: () -> bool?
        def is_italic
          @italic
        end

        #: (String | bool | nil val) -> void
        def set_underline(val)
          @underline = val ? "single" : nil
        end

        #: () -> bool?
        def is_underlined
          @underline ? true : nil
        end

        #: (bool? val) -> void
        def set_strikethrough(val)
          @strikethrough = val ? true : nil
        end

        #: () -> bool?
        def is_strikethrough
          @strikethrough
        end

        # @return [Font]
        #: () -> Font
        def dup
          self.class.new(
            name: @name,
            size: @size,
            color: @color ? Color.new(rgb: @color.rgb, theme: @color.theme, tint: @color.tint) : nil,
            bold: @bold,
            italic: @italic,
            underline: @underline,
            strikethrough: @strikethrough
          )
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(Font)

          @name == other.name &&
            @size == other.size &&
            @color == other.color &&
            @bold == other.bold &&
            @italic == other.italic &&
            @underline == other.underline &&
            @strikethrough == other.strikethrough
        end

        # Converts to xlsxrb font Hash.
        # @return [Hash]
        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          h = { name: @name, size: @size, sz: @size }
          h[:bold] = true if @bold
          h[:italic] = true if @italic
          h[:underline] = @underline if @underline
          h[:strike] = true if @strikethrough
          h[:color] = @color.rgb if @color&.rgb
          h
        end
      end

      # Represents a Fill definition.
      class Fill
        attr_accessor :pattern_type #: String
        attr_accessor :fg_color #: Color?
        attr_accessor :bg_color #: Color?

        # @param pattern_type [String]
        # @param fg_color [Color, String, Hash, nil]
        # @param bg_color [Color, String, Hash, nil]
        #: (?pattern_type: String, ?fg_color: (Color | String | Hash[untyped, untyped])?, ?bg_color: (Color | String | Hash[untyped, untyped])?) -> void
        def initialize(pattern_type: "none", fg_color: nil, bg_color: nil)
          @pattern_type = pattern_type
          @fg_color = Color.from(fg_color)
          @bg_color = Color.from(bg_color)
        end

        # Returns fill RGB color.
        # @return [String?]
        #: () -> String?
        def fill_color
          @fg_color&.rgb
        end

        # @return [Fill]
        #: () -> Fill
        def dup
          self.class.new(
            pattern_type: @pattern_type,
            fg_color: @fg_color ? Color.new(rgb: @fg_color.rgb, theme: @fg_color.theme, tint: @fg_color.tint) : nil,
            bg_color: @bg_color ? Color.new(rgb: @bg_color.rgb, theme: @bg_color.theme, tint: @bg_color.tint) : nil
          )
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(Fill)

          @pattern_type == other.pattern_type &&
            @fg_color == other.fg_color &&
            @bg_color == other.bg_color
        end

        # Converts to xlsxrb fill Hash.
        # @return [Hash]
        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          h = { pattern: @pattern_type }
          h[:fg_color] = @fg_color.rgb if @fg_color&.rgb
          h[:bg_color] = @bg_color.rgb if @bg_color&.rgb
          h
        end
      end

      # Represents a Border definition.
      class Border
        # Single border edge.
        class Edge
          attr_accessor :style #: String?
          attr_accessor :color #: Color?

          # @param style [String, nil]
          # @param color [Color, String, Hash, nil]
          #: (?style: String?, ?color: (Color | String | Hash[untyped, untyped])?) -> void
          def initialize(style: nil, color: nil)
            @style = style
            @color = Color.from(color)
          end

          # @return [Edge]
          #: () -> Edge
          def dup
            Edge.new(
              style: @style,
              color: @color ? Color.new(rgb: @color.rgb, theme: @color.theme, tint: @color.tint) : nil
            )
          end

          # @param other [untyped]
          # @return [bool]
          #: (untyped other) -> bool
          def ==(other)
            return false unless other.is_a?(Edge)

            @style == other.style && @color == other.color
          end
        end

        attr_accessor :left #: Edge
        attr_accessor :right #: Edge
        attr_accessor :top #: Edge
        attr_accessor :bottom #: Edge
        attr_accessor :diagonal #: Edge
        attr_accessor :diagonal_up #: bool?
        attr_accessor :diagonal_down #: bool?

        #: () -> void
        def initialize
          @left = Edge.new
          @right = Edge.new
          @top = Edge.new
          @bottom = Edge.new
          @diagonal = Edge.new
          @diagonal_up = nil
          @diagonal_down = nil
        end

        # @param direction [Symbol, String]
        # @return [Edge]
        #: (Symbol | String direction) -> Edge
        def edge(direction)
          case direction.to_sym
          when :left then @left
          when :right then @right
          when :top then @top
          when :bottom then @bottom
          when :diagonal then @diagonal
          else raise "Invalid border direction: #{direction}"
          end
        end

        # @param direction [Symbol, String]
        # @return [String?]
        #: (Symbol | String direction) -> String?
        def get_edge_style(direction)
          edge(direction).style
        end

        # @param direction [Symbol, String]
        # @param weight [String]
        # @param diagonals [untyped]
        # @return [void]
        #: (Symbol | String direction, String weight, ?untyped diagonals) -> void
        def set_edge_style(direction, weight, diagonals = nil)
          edge(direction).style = weight
          return unless direction.to_sym == :diagonal && diagonals

          @diagonal_up = true if diagonals[:up]
          @diagonal_down = true if diagonals[:down]
        end

        # @param direction [Symbol, String]
        # @return [String?]
        #: (Symbol | String direction) -> String?
        def get_edge_color(direction)
          edge(direction).color&.rgb
        end

        # @param direction [Symbol, String]
        # @param color [String, Color, Hash, nil]
        # @return [void]
        #: (Symbol | String direction, (String | Color | Hash[untyped, untyped])? color) -> void
        def set_edge_color(direction, color)
          c = Color.from(color)
          edge(direction).color = c
        end

        # @return [Border]
        #: () -> Border
        def dup
          b = self.class.new
          b.left = @left.dup
          b.right = @right.dup
          b.top = @top.dup
          b.bottom = @bottom.dup
          b.diagonal = @diagonal.dup
          b.diagonal_up = @diagonal_up
          b.diagonal_down = @diagonal_down
          b
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(Border)

          @left == other.left &&
            @right == other.right &&
            @top == other.top &&
            @bottom == other.bottom &&
            @diagonal == other.diagonal &&
            @diagonal_up == other.diagonal_up &&
            @diagonal_down == other.diagonal_down
        end

        # Converts to xlsxrb border Hash.
        # @return [Hash]
        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          h = {}
          %i[left right top bottom diagonal].each do |side|
            e = edge(side)
            next unless e.style

            side_h = { style: e.style }
            side_h[:color] = e.color.rgb if e.color&.rgb
            h[side] = side_h
          end
          h[:diagonal_up] = true if @diagonal_up
          h[:diagonal_down] = true if @diagonal_down
          h
        end
      end

      # Represents cell alignment.
      class Alignment
        attr_accessor :horizontal #: String?
        attr_accessor :vertical #: String?
        attr_accessor :wrap_text #: bool?
        attr_accessor :shrink_to_fit #: bool?
        attr_accessor :indent #: Integer?
        attr_accessor :text_rotation #: Integer?

        # @param horizontal [String, nil]
        # @param vertical [String, nil]
        # @param wrap_text [Boolean, nil]
        # @param shrink_to_fit [Boolean, nil]
        # @param indent [Integer, nil]
        # @param text_rotation [Integer, nil]
        #: (?horizontal: String?, ?vertical: String?, ?wrap_text: bool?, ?shrink_to_fit: bool?, ?indent: Integer?, ?text_rotation: Integer?) -> void
        def initialize(horizontal: nil, vertical: nil, wrap_text: nil, shrink_to_fit: nil, indent: nil, text_rotation: nil)
          @horizontal = horizontal
          @vertical = vertical
          @wrap_text = wrap_text
          @shrink_to_fit = shrink_to_fit
          @indent = indent
          @text_rotation = text_rotation
        end

        # @return [Alignment]
        #: () -> Alignment
        def dup
          self.class.new(
            horizontal: @horizontal,
            vertical: @vertical,
            wrap_text: @wrap_text,
            shrink_to_fit: @shrink_to_fit,
            indent: @indent,
            text_rotation: @text_rotation
          )
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(Alignment)

          @horizontal == other.horizontal &&
            @vertical == other.vertical &&
            @wrap_text == other.wrap_text &&
            @shrink_to_fit == other.shrink_to_fit &&
            @indent == other.indent &&
            @text_rotation == other.text_rotation
        end

        # Converts to xlsxrb alignment Hash.
        # @return [Hash]
        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          h = {}
          h[:horizontal] = @horizontal if @horizontal
          h[:vertical] = @vertical if @vertical
          h[:wrap_text] = @wrap_text if @wrap_text
          h[:shrink_to_fit] = @shrink_to_fit if @shrink_to_fit
          h[:indent] = @indent if @indent
          h[:text_rotation] = @text_rotation if @text_rotation
          h
        end
      end

      # Represents a format record (xf).
      class XF
        attr_accessor :num_fmt_id #: Integer
        attr_accessor :font_id #: Integer
        attr_accessor :fill_id #: Integer
        attr_accessor :border_id #: Integer
        attr_accessor :xf_id #: Integer
        attr_accessor :alignment #: Alignment?
        attr_accessor :apply_number_format #: bool?
        attr_accessor :apply_font #: bool?
        attr_accessor :apply_fill #: bool?
        attr_accessor :apply_border #: bool?
        attr_accessor :apply_alignment #: bool?

        # @param num_fmt_id [Integer]
        # @param font_id [Integer]
        # @param fill_id [Integer]
        # @param border_id [Integer]
        # @param xf_id [Integer]
        # @param alignment [Alignment, nil]
        #: (?num_fmt_id: Integer, ?font_id: Integer, ?fill_id: Integer, ?border_id: Integer, ?xf_id: Integer, ?alignment: Alignment?) -> void
        def initialize(num_fmt_id: 0, font_id: 0, fill_id: 0, border_id: 0, xf_id: 0, alignment: nil)
          @num_fmt_id = num_fmt_id
          @font_id = font_id
          @fill_id = fill_id
          @border_id = border_id
          @xf_id = xf_id
          @alignment = alignment
          @apply_number_format = nil
          @apply_font = nil
          @apply_fill = nil
          @apply_border = nil
          @apply_alignment = nil
        end

        # @return [XF]
        #: () -> XF
        def dup
          xf = self.class.new(
            num_fmt_id: @num_fmt_id,
            font_id: @font_id,
            fill_id: @fill_id,
            border_id: @border_id,
            xf_id: @xf_id,
            alignment: @alignment&.dup
          )
          xf.apply_number_format = @apply_number_format
          xf.apply_font = @apply_font
          xf.apply_fill = @apply_fill
          xf.apply_border = @apply_border
          xf.apply_alignment = @apply_alignment
          xf
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(XF)

          @num_fmt_id == other.num_fmt_id &&
            @font_id == other.font_id &&
            @fill_id == other.fill_id &&
            @border_id == other.border_id &&
            @alignment == other.alignment
        end

        # Converts to xlsxrb xf Hash.
        # @return [Hash]
        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          h = {
            num_fmt_id: @num_fmt_id,
            font_id: @font_id,
            fill_id: @fill_id,
            border_id: @border_id,
            xf_id: @xf_id
          }
          h[:alignment] = @alignment.to_xlsxrb_hash if @alignment
          h
        end
      end

      # Represents a single number format definition.
      class NumberFormat
        attr_accessor :num_fmt_id #: Integer?
        attr_accessor :format_code #: String?

        # @param num_fmt_id [Integer, nil]
        # @param format_code [String, nil]
        #: (?num_fmt_id: Integer?, ?format_code: String?) -> void
        def initialize(num_fmt_id: nil, format_code: nil)
          @num_fmt_id = num_fmt_id
          @format_code = format_code
        end

        # @return [bool]
        #: () -> bool
        def is_date_format?
          !!(@format_code.to_s.gsub(/("[^"]*"|\[[^\]]*\]|[\\_*].)/i, "") =~ /[dmyhs]/i)
        end
      end

      # Collection of number formats.
      # @rbs inherits Array[NumberFormat]
      class NumberFormats < Array
        # @param format_id [Integer]
        # @return [NumberFormat, nil]
        #: (Integer format_id) -> NumberFormat?
        def find_by_format_id(format_id)
          find { |fmt| fmt.num_fmt_id == format_id }
        end

        # Accesses format code by ID or index.
        #
        # @param key [Integer]
        # @return [String, NumberFormat, nil]
        #: (Integer key) -> untyped
        def [](key)
          fmt = find_by_format_id(key)
          return fmt.format_code if fmt

          super
        end

        DEFAULT_NUMBER_FORMATS = new([
                                       NumberFormat.new(num_fmt_id: 1, format_code: "0"),
                                       NumberFormat.new(num_fmt_id: 2, format_code: "0.00"),
                                       NumberFormat.new(num_fmt_id: 3, format_code: "#, ##0"),
                                       NumberFormat.new(num_fmt_id: 4, format_code: "#, ##0.00"),
                                       NumberFormat.new(num_fmt_id: 5, format_code: "$#, ##0_);($#, ##0)"),
                                       NumberFormat.new(num_fmt_id: 6, format_code: "$#, ##0_);[Red]($#, ##0)"),
                                       NumberFormat.new(num_fmt_id: 7, format_code: "$#, ##0.00_);($#, ##0.00)"),
                                       NumberFormat.new(num_fmt_id: 8, format_code: "$#, ##0.00_);[Red]($#, ##0.00)"),
                                       NumberFormat.new(num_fmt_id: 9, format_code: "0%"),
                                       NumberFormat.new(num_fmt_id: 10, format_code: "0.00%"),
                                       NumberFormat.new(num_fmt_id: 11, format_code: "0.00E+00"),
                                       NumberFormat.new(num_fmt_id: 12, format_code: "# ?/?"),
                                       NumberFormat.new(num_fmt_id: 13, format_code: "# ??/??"),
                                       NumberFormat.new(num_fmt_id: 14, format_code: "m/d/yyyy"),
                                       NumberFormat.new(num_fmt_id: 15, format_code: "d-mmm-yy"),
                                       NumberFormat.new(num_fmt_id: 16, format_code: "d-mmm"),
                                       NumberFormat.new(num_fmt_id: 17, format_code: "mmm-yy"),
                                       NumberFormat.new(num_fmt_id: 18, format_code: "h:mm AM/PM"),
                                       NumberFormat.new(num_fmt_id: 19, format_code: "h:mm:ss AM/PM"),
                                       NumberFormat.new(num_fmt_id: 20, format_code: "h:mm"),
                                       NumberFormat.new(num_fmt_id: 21, format_code: "h:mm:ss"),
                                       NumberFormat.new(num_fmt_id: 22, format_code: "m/d/yyyy h:mm"),
                                       NumberFormat.new(num_fmt_id: 37, format_code: "#, ##0_);(#, ##0)"),
                                       NumberFormat.new(num_fmt_id: 38, format_code: "#, ##0_);[Red](#, ##0)"),
                                       NumberFormat.new(num_fmt_id: 39, format_code: "#, ##0.00_);(#, ##0.00)"),
                                       NumberFormat.new(num_fmt_id: 40, format_code: "#, ##0.00_);[Red](#, ##0.00)"),
                                       NumberFormat.new(num_fmt_id: 45, format_code: "mm:ss"),
                                       NumberFormat.new(num_fmt_id: 46, format_code: "[h]:mm:ss"),
                                       NumberFormat.new(num_fmt_id: 47, format_code: "mm:ss.0"),
                                       NumberFormat.new(num_fmt_id: 48, format_code: "##0.0E+0"),
                                       NumberFormat.new(num_fmt_id: 49, format_code: "@")
                                     ]).freeze
      end

      # Stylesheet managing fonts, fills, borders, number formats, and XFs.
      class Stylesheet
        attr_accessor :fonts #: Array[Font]
        attr_accessor :fills #: Array[Fill]
        attr_accessor :borders #: Array[Border]
        attr_accessor :cell_xfs #: Array[XF]
        attr_accessor :number_formats #: NumberFormats

        #: () -> void
        def initialize
          @fonts = [Font.new]
          @fills = [
            Fill.new(pattern_type: "none"),
            Fill.new(pattern_type: "gray125")
          ]
          @borders = [Border.new]
          @cell_xfs = [XF.new]
          @number_formats = NumberFormats.new
        end

        # Returns number format by ID, checking default and custom formats.
        #
        # @param format_id [Integer]
        # @return [NumberFormat, nil]
        #: (Integer format_id) -> NumberFormat?
        def get_number_format_by_id(format_id)
          NumberFormats::DEFAULT_NUMBER_FORMATS.find_by_format_id(format_id) ||
            @number_formats&.find_by_format_id(format_id)
        end

        # Registers a custom number format code and returns its ID.
        #
        # @param format_code [String]
        # @return [Integer]
        #: (String format_code) -> Integer
        def register_number_format(format_code)
          all_formats = NumberFormats::DEFAULT_NUMBER_FORMATS + @number_formats
          existing = all_formats.find { |fmt| fmt.format_code == format_code }
          return existing.num_fmt_id if existing&.num_fmt_id

          max_fmt_id = 163
          all_formats.each do |fmt|
            max_fmt_id = fmt.num_fmt_id if fmt.num_fmt_id && fmt.num_fmt_id > max_fmt_id
          end

          id = max_fmt_id + 1
          @number_formats << NumberFormat.new(num_fmt_id: id, format_code: format_code)
          id
        end

        # Builds a default Stylesheet from raw xlsxrb styles hash if present.
        #
        # @param raw_styles [Hash]
        # @return [Stylesheet]
        #: (Hash[untyped, untyped] raw_styles) -> Stylesheet
        def self.from_xlsxrb(raw_styles)
          sheet = new
          return sheet if raw_styles.nil? || raw_styles.empty?

          # Parse fonts
          raw_fonts = raw_styles[:fonts]
          if raw_fonts.is_a?(Array) && !raw_fonts.empty?
            sheet.fonts = raw_fonts.map do |rf|
              Font.new(
                name: rf[:name] || "Calibri",
                size: rf[:size] || rf[:sz] || 11,
                color: rf[:color],
                bold: rf[:bold],
                italic: rf[:italic],
                underline: rf[:underline],
                strikethrough: rf[:strike]
              )
            end
          end

          # Parse fills
          raw_fills = raw_styles[:fills]
          if raw_fills.is_a?(Array) && !raw_fills.empty?
            sheet.fills = raw_fills.map do |rf|
              Fill.new(
                pattern_type: rf[:pattern] || "none",
                fg_color: rf[:fg_color],
                bg_color: rf[:bg_color]
              )
            end
          end

          # Parse borders
          raw_borders = raw_styles[:borders]
          if raw_borders.is_a?(Array) && !raw_borders.empty?
            sheet.borders = raw_borders.map do |rb|
              b = Border.new
              %i[left right top bottom diagonal].each do |side|
                edge_data = rb[side]
                if edge_data.is_a?(Hash)
                  b.set_edge_style(side, edge_data[:style]) if edge_data[:style]
                  b.set_edge_color(side, edge_data[:color]) if edge_data[:color]
                end
              end
              b.diagonal_up = rb[:diagonal_up]
              b.diagonal_down = rb[:diagonal_down]
              b
            end
          end

          # Parse num_fmts
          raw_num_fmts = raw_styles[:num_fmts]
          if raw_num_fmts.is_a?(Hash)
            raw_num_fmts.each { |id, code| sheet.number_formats << NumberFormat.new(num_fmt_id: id.to_i, format_code: code.to_s) }
          elsif raw_num_fmts.is_a?(Array)
            raw_num_fmts.each do |nf|
              sheet.number_formats << NumberFormat.new(num_fmt_id: nf[:num_fmt_id].to_i, format_code: nf[:format_code].to_s) if nf.is_a?(Hash)
            end
          end

          # Parse cell_xfs
          raw_xfs = raw_styles[:cell_xfs] || raw_styles[:xf_entries]
          if raw_xfs.is_a?(Array) && !raw_xfs.empty?
            sheet.cell_xfs = raw_xfs.map do |rx|
              align = nil
              if rx[:alignment].is_a?(Hash)
                align = Alignment.new(
                  horizontal: rx[:alignment][:horizontal],
                  vertical: rx[:alignment][:vertical],
                  wrap_text: rx[:alignment][:wrap_text],
                  shrink_to_fit: rx[:alignment][:shrink_to_fit],
                  indent: rx[:alignment][:indent],
                  text_rotation: rx[:alignment][:text_rotation]
                )
              end
              XF.new(
                num_fmt_id: rx[:num_fmt_id].to_i,
                font_id: rx[:font_id].to_i,
                fill_id: rx[:fill_id].to_i,
                border_id: rx[:border_id].to_i,
                xf_id: rx[:xf_id].to_i,
                alignment: align
              )
            end
          end

          sheet
        end

        # Serializes into xlsxrb styles hash.
        # @return [Hash]
        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          xfs = @cell_xfs.map(&:to_xlsxrb_hash)
          {
            fonts: @fonts.map(&:to_xlsxrb_hash),
            fills: @fills.map(&:to_xlsxrb_hash),
            borders: @borders.map(&:to_xlsxrb_hash),
            num_fmts: @number_formats.map { |nf| { num_fmt_id: nf.num_fmt_id, format_code: nf.format_code } },
            cell_xfs: xfs,
            xf_entries: xfs
          }
        end
      end
    end
  end
end
