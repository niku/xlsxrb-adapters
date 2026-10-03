# frozen_string_literal: true

# rbs_inline: enabled

require_relative "excelx"

module Xlsxrb
  module Adapters
    module Roo
      # Spreadsheet opener matching Roo::Spreadsheet.
      class Spreadsheet
        CLASS_FOR_EXTENSION = {
          xlsx: Excelx,
          xlsm: Excelx
        }.freeze #: Hash[Symbol, singleton(Excelx)]

        class << self
          # Opens a spreadsheet file detecting parser class by extension.
          #
          # @param path [String, Pathname, File, IO]
          # @param options [Hash]
          # @return [Excelx]
          #: (untyped path, ?Hash[Symbol, untyped] options) -> Excelx
          def open(path, options = {})
            target_path = path.respond_to?(:path) ? path.path : path
            extension = extension_for(target_path, options)

            klass = CLASS_FOR_EXTENSION[extension]
            unless klass
              raise ArgumentError,
                    "Can't detect the type of #{path} - please use the :extension option to declare its type."
            end

            instance = klass.new(path, options)
            if block_given?
              yield instance
            else
              instance
            end
          end

          # Extracts extension symbol from path or options.
          #
          # @param path [String, Pathname, Object]
          # @param options [Hash]
          # @return [Symbol]
          #: (untyped path, Hash[Symbol, untyped] options) -> Symbol
          def extension_for(path, options)
            case (extension = options.delete(:extension))
            when ::Symbol
              options[:file_warning] = :ignore
              extension
            when ::String
              options[:file_warning] = :ignore
              extension.delete(".").downcase.to_sym
            else
              parsed = path.to_s.split("?").first.to_s
              File.extname(parsed).delete(".").downcase.to_sym
            end
          end
        end
      end

      CLASS_FOR_EXTENSION = Spreadsheet::CLASS_FOR_EXTENSION
    end
  end
end
