# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      class Excelx
        # 1-based [row, column] coordinate representation matching Roo::Excelx::Coordinate.
        class Coordinate < ::Array #[Integer]
          # @param row [Integer] 1-based row index.
          # @param column [Integer] 1-based column index.
          #: (Integer row, Integer column) -> void
          def initialize(row, column)
            super()
            self << row << column
            freeze
          end

          # 1-based row number.
          #: () -> Integer
          def row
            self[0]
          end

          # 1-based column number.
          #: () -> Integer
          def column
            self[1]
          end
        end
      end
    end
  end
end
