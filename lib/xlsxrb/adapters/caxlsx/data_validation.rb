# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # Data validation rules for cells in a worksheet.
      class DataValidation
        include OptionsParser

        CHILD_ELEMENTS = %i[formula1 formula2].freeze

        # @return [String, nil]
        attr_reader :formula1

        # @return [String, nil]
        attr_reader :formula2

        # @return [Boolean, nil]
        attr_reader :allowBlank

        # @return [String, nil]
        attr_reader :error

        # @return [Symbol, nil]
        attr_reader :errorStyle

        # @return [String, nil]
        attr_reader :errorTitle

        # @return [Symbol, nil]
        attr_reader :operator

        # @return [String, nil]
        attr_reader :prompt

        # @return [String, nil]
        attr_reader :promptTitle

        # @return [Boolean, nil]
        attr_reader :showDropDown
        alias hideDropDown showDropDown

        # @return [Boolean, nil]
        attr_reader :showErrorMessage

        # @return [Boolean, nil]
        attr_reader :showInputMessage

        # @return [String, nil]
        attr_reader :sqref

        # @return [Symbol, nil]
        attr_reader :type

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @allowBlank = @showErrorMessage = @showInputMessage = false
          @showDropDown = false
          @errorStyle = :stop
          @type = :none
          parse_options(options)
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def formula1=(v)
          Caxlsx.validate_string(v)
          @formula1 = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def formula2=(v)
          Caxlsx.validate_string(v)
          @formula2 = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def allowBlank=(v)
          Caxlsx.validate_boolean(v)
          @allowBlank = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def error=(v)
          Caxlsx.validate_string(v)
          @error = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def errorStyle=(v)
          Caxlsx.validate_data_validation_error_style(v)
          @errorStyle = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def errorTitle=(v)
          Caxlsx.validate_string(v)
          @errorTitle = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def operator=(v)
          Caxlsx.validate_data_validation_operator(v)
          @operator = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def prompt=(v)
          Caxlsx.validate_string(v)
          @prompt = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def promptTitle=(v)
          Caxlsx.validate_string(v)
          @promptTitle = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def showDropDown=(v)
          Caxlsx.validate_boolean(v)
          @showDropDown = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def hideDropDown=(v)
          Caxlsx.validate_boolean(v)
          @showDropDown = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def showErrorMessage=(v)
          Caxlsx.validate_boolean(v)
          @showErrorMessage = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def showInputMessage=(v)
          Caxlsx.validate_boolean(v)
          @showInputMessage = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def sqref=(v)
          Caxlsx.validate_string(v)
          @sqref = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def type=(v)
          Caxlsx.validate_data_validation_type(v)
          @type = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          valid_attributes = get_valid_attributes
          h = Caxlsx.instance_values_for(self).select { |key, _| valid_attributes.include?(key.to_sym) && !CHILD_ELEMENTS.include?(key.to_sym) }

          str << "<dataValidation "
          h.each_with_index do |key_value, index|
            str << " " unless index.zero?
            str << key_value.first << '="' << Caxlsx.booleanize(key_value.last).to_s << '"'
          end
          str << ">"
          str << "<formula1>" << formula1 << "</formula1>" if formula1 && valid_attributes.include?(:formula1)
          str << "<formula2>" << formula2 << "</formula2>" if formula2 && valid_attributes.include?(:formula2)
          str << "</dataValidation>"
        end

        private

        # @return [Array[Symbol]]
        #: () -> Array[Symbol]
        def get_valid_attributes
          attributes = %i[allowBlank error errorStyle errorTitle prompt promptTitle showErrorMessage showInputMessage sqref type]

          if %i[whole decimal data time date textLength].include?(@type)
            attributes << :operator << :formula1
            attributes << :formula2 if %i[between notBetween].include?(@operator)
          elsif @type == :list
            attributes << :showDropDown << :formula1
          elsif @type == :custom
            attributes << :formula1
          end

          attributes
        end
      end

      # Collection of DataValidation objects for a worksheet.
      class DataValidations < SimpleTypedList
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          super(DataValidation)
          @worksheet = worksheet
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<dataValidations count='#{size}'>"
          each { |item| item.to_xml_string(str) }
          str << "</dataValidations>"
        end
      end
    end
  end
end
