# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # A relationship defines a reference between package parts.
      class Relationship
        class << self
          # @return [Hash{Array[untyped] => String}]
          def ids_cache
            Thread.current[:axlsx_relationship_ids_cache] ||= {}
          end

          # @return [void]
          #: () -> void
          def initialize_ids_cache
            Thread.current[:axlsx_relationship_ids_cache] = {}
          end

          # @return [void]
          #: () -> void
          def clear_ids_cache
            Thread.current[:axlsx_relationship_ids_cache] = nil
          end

          # @return [String]
          #: () -> String
          def next_free_id
            "rId#{ids_cache.size + 1}"
          end
        end

        # @return [String]
        attr_reader :Id

        # @return [String]
        attr_reader :Target

        # @return [String]
        attr_reader :Type

        # @return [Symbol, nil]
        attr_reader :TargetMode

        # @return [untyped]
        attr_reader :source_obj

        # @param source_obj [untyped]
        # @param type [String]
        # @param target [String]
        # @param options [Hash{Symbol => untyped}]
        #: (untyped source_obj, String type, String target, ?Hash[Symbol, untyped] options) -> void
        def initialize(source_obj, type, target, options = {})
          @source_obj = source_obj
          self.Target = target
          self.Type = type
          self.TargetMode = options[:target_mode] if options[:target_mode]
          @Id = (self.class.ids_cache[ids_cache_key] ||= self.class.next_free_id)
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def Target=(v)
          Caxlsx.validate_string(v)
          @Target = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def Type=(v)
          Caxlsx.validate_relationship_type(v)
          @Type = v
        end

        # @param v [Symbol]
        # @return [Symbol]
        #: (Symbol v) -> Symbol
        def TargetMode=(v)
          RestrictionValidator.validate("Relationship.TargetMode", %i[External Internal], v)
          @TargetMode = v
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          h = Caxlsx.instance_values_for(self).except("source_obj")
          str << "<Relationship "
          h.each_with_index do |key_value, idx|
            str << " " unless idx.zero?
            str << key_value.first.to_s << '="' << Caxlsx.coder.encode(key_value.last.to_s) << '"'
          end
          str << "/>"
        end

        # @return [Array[untyped]]
        #: () -> Array[untyped]
        def ids_cache_key
          key = [source_obj, self.Type, self.TargetMode]
          key << self.Target if self.TargetMode == :External
          key
        end
      end

      # Relationships collection based on SimpleTypedList.
      class Relationships < SimpleTypedList
        #: () -> void
        def initialize
          super(Relationship)
        end

        # @param source_obj [untyped]
        # @return [Relationship, nil]
        #: (untyped source_obj) -> Relationship?
        def for(source_obj)
          find { |rel| rel.source_obj == source_obj }
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << '<Relationships xmlns="' << RELS_R << '">'
          each { |rel| rel.to_xml_string(str) }
          str << "</Relationships>"
        end
      end
    end
  end
end
