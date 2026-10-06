# frozen_string_literal: true

# rbs_inline: enabled

require "xlsxrb"
require_relative "adapters/version"
require_relative "adapters/ruby_xl"
require_relative "adapters/caxlsx"
require_relative "adapters/roo"
require_relative "adapters/xsv"
require_relative "adapters/creek"
require_relative "adapters/fast_excel"
require_relative "adapters/write_xlsx"
require_relative "adapters/xlsxtream"
require_relative "adapters/simple_xlsx_reader"

module Xlsxrb
  # Namespace for compatibility adapters helping gradual migration to xlsxrb.
  module Adapters
  end
end
