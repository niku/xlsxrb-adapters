# frozen_string_literal: true

# rbs_inline: enabled

require "rexml/document"

module Xlsxrb
  module Adapters
    module Roo
      module Formatters
        # XML export formatter matching Roo::Formatters::XML.
        module XML
          # Returns an XML representation of all sheets.
          #
          # @return [String]
          #: () -> String
          def to_xml
            doc = REXML::Document.new
            doc << REXML::XMLDecl.new("1.0", "UTF-8")
            root = doc.add_element("spreadsheet")

            sheets.each do |sheet_name|
              sheet_el = root.add_element("sheet", { "name" => sheet_name })
              f_row = first_row(sheet_name)
              l_row = last_row(sheet_name)
              f_col = first_column(sheet_name)
              l_col = last_column(sheet_name)

              next unless f_row && l_row && f_col && l_col

              f_row.upto(l_row) do |r|
                f_col.upto(l_col) do |c|
                  next if empty?(r, c, sheet_name)

                  val = cell(r, c, sheet_name)
                  ctype = celltype(r, c, sheet_name)

                  cell_el = sheet_el.add_element("cell", {
                                                   "row" => r.to_s,
                                                   "column" => c.to_s,
                                                   "type" => ctype.to_s
                                                 })
                  cell_el.text = val.to_s
                end
              end
            end

            out = +""
            doc.write(out)
            out
          end
        end
      end
    end
  end
end
