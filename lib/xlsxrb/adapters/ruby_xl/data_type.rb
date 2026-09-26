# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents OpenXML cell data types compatible with RubyXL::DataType.
      module DataType
        SHARED_STRING = "s"
        RAW_STRING    = "str"
        INLINE_STRING = "inlineStr"
        ERROR         = "e"
        BOOLEAN       = "b"
        NUMBER        = "n"
        DATE          = "d"
      end
    end
  end
end
