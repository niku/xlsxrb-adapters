# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Represents a defined name in a workbook.
      class DefinedName
        attr_accessor :name #: String
        attr_accessor :reference #: String
        attr_accessor :local_sheet_id #: Integer?

        # @param params [Hash]
        #: (?Hash[Symbol, untyped] params) -> void
        def initialize(params = {})
          @name = params[:name].to_s
          @reference = (params[:reference] || params[:value]).to_s
          @local_sheet_id = params[:local_sheet_id]
        end

        # @param other [untyped]
        # @return [bool]
        #: (untyped other) -> bool
        def ==(other)
          return false unless other.is_a?(DefinedName)

          @name == other.name && @reference == other.reference
        end
      end

      # Collection of defined names.
      # @rbs inherits Array[DefinedName]
      class DefinedNames < Array
        # @param other_array [Array<DefinedName>]
        #: (?Array[DefinedName] other_array) -> void
        def initialize(other_array = [])
          super()
          other_array.each { |dn| self << dn }
        end
      end
    end
  end
end
