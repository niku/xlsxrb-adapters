# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Lightweight wrapper for conditional formatting providing both RubyXL object-style access
      # (.sqref, .type) and Hash-style access ([:sqref], [:type]).
      class ConditionalFormatting
        attr_accessor :sqref, :type, :operator, :formula, :priority, :stop_if_true,
                      :style_index, :format_id, :dxf_id, :rules, :unmapped_attributes

        def initialize(attributes = {})
          attrs = attributes.is_a?(Hash) ? attributes : {}
          @sqref = attrs[:sqref] || attrs["sqref"]
          @type = attrs[:type] || attrs["type"]
          @operator = attrs[:operator] || attrs["operator"]
          @formula = attrs[:formula] || attrs["formula"]
          @priority = attrs[:priority] || attrs["priority"]
          @stop_if_true = attrs[:stop_if_true] || attrs["stop_if_true"]
          @style_index = attrs[:style_index] || attrs["style_index"]
          @format_id = attrs[:format_id] || attrs["format_id"] || attrs[:dxf_id] || attrs["dxf_id"]
          @dxf_id = @format_id
          @rules = attrs[:rules] || attrs["rules"]
          known_keys = %i[sqref type operator formula priority stop_if_true style_index format_id dxf_id rules]
          @unmapped_attributes = attrs.reject do |k, _|
            sym = k.to_s.to_sym
            known_keys.include?(sym)
          end
        end

        def [](key)
          k = key.to_s.to_sym
          if respond_to?(k) && !%i[[] []= to_h to_xlsxrb].include?(k)
            public_send(k)
          else
            @unmapped_attributes[key]
          end
        end

        def []=(key, val)
          setter = :"#{key}="
          if respond_to?(setter)
            public_send(setter, val)
          else
            @unmapped_attributes[key] = val
          end
        end

        def to_h
          h = {
            sqref: @sqref,
            type: @type,
            operator: @operator,
            formula: @formula,
            priority: @priority,
            stop_if_true: @stop_if_true,
            style_index: @style_index,
            format_id: @format_id,
            dxf_id: @dxf_id,
            rules: @rules
          }.compact
          h.merge!(@unmapped_attributes) if @unmapped_attributes && !@unmapped_attributes.empty?
          h
        end
        alias to_xlsxrb to_h
      end
    end
  end
end
