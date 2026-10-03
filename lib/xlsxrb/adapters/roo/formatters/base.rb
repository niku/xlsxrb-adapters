# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      module Formatters
        # Helper methods for formatter modules.
        module Base
          # Converts integer seconds into HH:MM:SS string.
          #
          # @param content [Numeric, nil]
          # @return [String]
          #: (Numeric? content) -> String
          def integer_to_timestring(content)
            return "00:00:00" if content.nil?

            sec = content.to_i
            h = sec / 3600
            m = (sec % 3600) / 60
            s = sec % 60
            Kernel.format("%<h>02d:%<m>02d:%<s>02d", h: h, m: m, s: s)
          end
        end
      end
    end
  end
end
