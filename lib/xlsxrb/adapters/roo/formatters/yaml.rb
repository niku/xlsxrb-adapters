# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Roo
      module Formatters
        # YAML export formatter matching Roo::Formatters::YAML.
        module YAML
          # Exports a rectangular area of cells to YAML string.
          #
          # @param prefix [Hash]
          # @param from_row [Integer, nil]
          # @param from_column [Integer, nil]
          # @param to_row [Integer, nil]
          # @param to_column [Integer, nil]
          # @param sheet [String, Integer]
          # @return [String]
          #: (?Hash[untyped, untyped] prefix, ?Integer? from_row, ?Integer? from_column, ?Integer? to_row, ?Integer? to_column, ?untyped sheet) -> String
          def to_yaml(prefix = {}, from_row = nil, from_column = nil, to_row = nil, to_column = nil, sheet = default_sheet)
            target_sheet = sheet || default_sheet
            return "" unless first_row(target_sheet)

            f_row = from_row || first_row(target_sheet)
            t_row = to_row || last_row(target_sheet)
            f_col = from_column || first_column(target_sheet)
            t_col = to_column || last_column(target_sheet)

            result = String.new("--- \n")
            f_row.upto(t_row) do |r|
              f_col.upto(t_col) do |c|
                next if empty?(r, c, target_sheet)

                result << "cell_#{r}_#{c}: \n"
                prefix.each do |k, v|
                  result << "  #{k}: #{v} \n"
                end
                result << "  row: #{r} \n"
                result << "  col: #{c} \n"
                result << "  celltype: #{celltype(r, c, target_sheet)} \n"
                val = cell(r, c, target_sheet)
                val = integer_to_timestring(val) if celltype(r, c, target_sheet) == :time
                result << "  value: #{val} \n"
              end
            end

            result
          end
        end
      end
    end
  end
end
