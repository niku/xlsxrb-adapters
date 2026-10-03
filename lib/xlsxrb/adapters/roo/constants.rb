# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      ERROR_VALUES = %w[#N/A #REF! #NAME? #DIV/0! #NULL! #VALUE! #NUM!].to_set.freeze #: Set[String]
      class HeaderRowNotFoundError < StandardError
      end
    end
  end
end
