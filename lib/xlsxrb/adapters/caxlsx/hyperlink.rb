# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # A worksheet hyperlink object.
      class WorksheetHyperlink
        include OptionsParser
        include Accessors
        include SerializedAttributes

        string_attr_accessor :display, :location, :tooltip
        serializable_attributes :display, :tooltip, :ref

        # @return [String, nil]
        attr_reader :ref

        # @return [Worksheet]
        attr_reader :worksheet

        # @return [Symbol]
        attr_accessor :target

        # @param worksheet [Worksheet]
        # @param options [Hash{Symbol => untyped}]
        #: (Worksheet worksheet, ?Hash[Symbol, untyped] options) ?{ (WorksheetHyperlink) -> void } -> void
        def initialize(worksheet, options = {})
          DataTypeValidator.validate("Hyperlink.worksheet", [Worksheet], worksheet)
          @worksheet = worksheet
          @target = :external
          parse_options(options)
          yield self if block_given?
        end

        # @param target [Symbol]
        # @return [Symbol]
        #: (Symbol target) -> Symbol

        # @param cell_reference [String, Cell]
        # @return [String]
        #: (String | Cell cell_reference) -> String
        def ref=(cell_reference)
          cell_reference = cell_reference.r if cell_reference.is_a?(Cell)
          Caxlsx.validate_string(cell_reference)
          @ref = cell_reference
        end

        # @return [Relationship, nil]
        #: () -> Relationship?
        def relationship
          return unless @target == :external

          Relationship.new(self, HYPERLINK_R, location, target_mode: :External)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<hyperlink "
          serialized_attributes(str, location_or_id, false)
          str << "/>"
        end

        # @return [Hash{Symbol => untyped}]
        #: () -> Hash[Symbol, untyped]
        def location_or_id
          if @target == :external
            rel = relationship
            rel ? { "r:id": rel.Id } : {}
          else
            { location: Caxlsx.coder.encode(location) }
          end
        end
      end

      # A collection of hyperlink objects for a worksheet.
      class WorksheetHyperlinks < SimpleTypedList
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          DataTypeValidator.validate("Hyperlinks.worksheet", [Worksheet], worksheet)
          @worksheet = worksheet
          super(WorksheetHyperlink)
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [WorksheetHyperlink]
        #: (Hash[Symbol, untyped] options) -> WorksheetHyperlink
        def add(options)
          self << WorksheetHyperlink.new(@worksheet, options)
          last
        end

        # @return [Array[Relationship]]
        #: () -> Array[Relationship]
        def relationships
          return [] if empty?

          map(&:relationship).compact
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return if empty?

          str << "<hyperlinks>"
          each { |hyperlink| hyperlink.to_xml_string(str) }
          str << "</hyperlinks>"
        end
      end
    end
  end
end
