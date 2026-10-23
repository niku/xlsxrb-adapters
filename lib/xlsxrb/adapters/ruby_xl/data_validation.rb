# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module RubyXL
      # Lightweight wrapper for data validation providing both RubyXL object-style access
      # (.type, .sqref) and Hash-style access ([:type], [:sqref]).
      class DataValidation
        attr_accessor :type, :sqref, :formula1, :formula2, :operator, :allow_blank,
                      :show_input_message, :show_error_message, :error_title, :error,
                      :prompt_title, :prompt, :unmapped_attributes

        def initialize(attributes = {})
          attrs = attributes.is_a?(Hash) ? attributes : {}
          @type = attrs[:type] || attrs["type"]
          @sqref = attrs[:sqref] || attrs["sqref"]
          @formula1 = attrs[:formula1] || attrs["formula1"]
          @formula2 = attrs[:formula2] || attrs["formula2"]
          @operator = attrs[:operator] || attrs["operator"]
          @allow_blank = attrs[:allow_blank] || attrs["allow_blank"]
          @show_input_message = attrs[:show_input_message] || attrs["show_input_message"]
          @show_error_message = attrs[:show_error_message] || attrs["show_error_message"]
          @error_title = attrs[:error_title] || attrs["error_title"]
          @error = attrs[:error] || attrs["error"]
          @prompt_title = attrs[:prompt_title] || attrs["prompt_title"]
          @prompt = attrs[:prompt] || attrs["prompt"]
          known_keys = %i[type sqref formula1 formula2 operator allow_blank show_input_message
                          show_error_message error_title error prompt_title prompt]
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
            type: @type,
            sqref: @sqref,
            formula1: @formula1,
            formula2: @formula2,
            operator: @operator,
            allow_blank: @allow_blank,
            show_input_message: @show_input_message,
            show_error_message: @show_error_message,
            error_title: @error_title,
            error: @error,
            prompt_title: @prompt_title,
            prompt: @prompt
          }.compact
          h.merge!(@unmapped_attributes) if @unmapped_attributes && !@unmapped_attributes.empty?
          h
        end
        alias to_xlsxrb to_h
      end
    end
  end
end
