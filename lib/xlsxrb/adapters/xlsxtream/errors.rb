# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Xlsxtream
      # Base exception for Xlsxtream errors.
      class Error < StandardError; end

      # Deprecation exception for Xlsxtream.
      class Deprecation < StandardError; end
    end
  end
end
