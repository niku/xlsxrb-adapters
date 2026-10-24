# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      ERROR_VALUES = %w[#N/A #REF! #NAME? #DIV/0! #NULL! #VALUE! #NUM!].to_set.freeze #: Set[String]
      class Error < StandardError; end
      class HeaderRowNotFoundError < Error; end
      class FileNotFound < Error; end
    end
  end
end
