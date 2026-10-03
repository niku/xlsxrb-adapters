# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      # Font property container matching Roo::Font.
      class Font
        attr_accessor :bold, :italic, :underline

        # @param bold [Boolean]
        # @param italic [Boolean]
        # @param underline [Boolean]
        #: (?bold: bool, ?italic: bool, ?underline: bool) -> void
        def initialize(bold: false, italic: false, underline: false)
          @bold = bold
          @italic = italic
          @underline = underline
        end

        # Returns whether the font is bold.
        #: () -> bool
        def bold?
          !@bold.nil? && @bold != false
        end

        # Returns whether the font is italic.
        #: () -> bool
        def italic?
          !@italic.nil? && @italic != false
        end

        # Returns whether the font is underlined.
        #: () -> bool
        def underline?
          !@underline.nil? && @underline != false
        end
      end
    end
  end
end
