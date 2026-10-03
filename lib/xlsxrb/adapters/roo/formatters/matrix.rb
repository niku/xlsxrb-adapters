# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      module Formatters
        # Matrix export formatter matching Roo::Formatters::Matrix.
        module Matrix
          # Returns a Matrix from the whole sheet or a rectangular area.
          #
          # @param from_row [Integer, nil]
          # @param from_column [Integer, nil]
          # @param to_row [Integer, nil]
          # @param to_column [Integer, nil]
          # @param sheet [String, Integer]
          # @return [Object]
          #: (?Integer? from_row, ?Integer? from_column, ?Integer? to_row, ?Integer? to_column, ?sheet: untyped) -> untyped
          def to_matrix(from_row = nil, from_column = nil, to_row = nil, to_column = nil, sheet = default_sheet, sheet_kw: nil)
            require "matrix"

            target_sheet = sheet_kw || sheet || default_sheet
            return ::Matrix.empty unless first_row(target_sheet)

            f_row = from_row || first_row(target_sheet)
            t_row = to_row || last_row(target_sheet)
            f_col = from_column || first_column(target_sheet)
            t_col = to_column || last_column(target_sheet)

            rows_data = f_row.upto(t_row).map do |r|
              f_col.upto(t_col).map do |c|
                cell(r, c, target_sheet)
              end
            end

            ::Matrix.rows(rows_data)
          end
        end
      end
    end
  end
end
