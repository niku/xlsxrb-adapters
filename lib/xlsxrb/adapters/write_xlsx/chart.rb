# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Writexlsx
      # Represents a chart object in WriteXLSX
      class Chart
        # @return [Object]
        attr_accessor :workbook

        # @return [String, Symbol]
        attr_accessor :type

        # @return [String, Symbol, nil]
        attr_accessor :subtype

        # @return [Array<Hash{Symbol => untyped}>]
        attr_accessor :series

        # @return [Hash{Symbol => untyped}]
        attr_accessor :title

        # @return [Hash{Symbol => untyped}]
        attr_accessor :x_axis

        # @return [Hash{Symbol => untyped}]
        attr_accessor :y_axis

        # @return [Hash{Symbol => untyped}]
        attr_accessor :legend

        # @return [Integer, nil]
        attr_accessor :style_id

        # @return [Numeric, nil]
        attr_accessor :width

        # @return [Numeric, nil]
        attr_accessor :height

        # @return [Numeric, nil]
        attr_accessor :x_scale

        # @return [Numeric, nil]
        attr_accessor :y_scale

        # @return [Object, nil]
        attr_accessor :palette

        # @return [String, nil]
        attr_accessor :name

        # @return [Integer, nil]
        attr_accessor :index

        # @return [Boolean, nil]
        attr_accessor :embedded

        # @param workbook_or_subtype [Object, nil]
        # @param options [Hash{Symbol => untyped}]
        #: (?untyped workbook_or_subtype, ?Hash[Symbol, untyped] options) -> void
        def initialize(workbook_or_subtype = nil, options = {})
          if workbook_or_subtype.is_a?(Hash)
            options = workbook_or_subtype
            @workbook = nil
            @subtype = options[:subtype]
          elsif workbook_or_subtype.is_a?(String) || workbook_or_subtype.is_a?(Symbol)
            @subtype = workbook_or_subtype
            @workbook = nil
          else
            @workbook = workbook_or_subtype
            @subtype = options[:subtype]
          end

          @type = options[:type] || default_chart_type
          @series = []
          @title = {}
          @x_axis = {}
          @y_axis = {}
          @legend = {}
          @style_id = 2
          @width = 480
          @height = 288
          @x_scale = 1
          @y_scale = 1
        end

        def default_chart_type
          :column
        end

        # Adds a data series to the chart
        #
        # @param options [Hash{Symbol => untyped}]
        # @return [Hash{Symbol => untyped}]
        #: (Hash[Symbol, untyped] options) -> Hash[Symbol, untyped]
        def add_series(options = {})
          s = {
            categories: options[:categories],
            values: options[:values],
            name: options[:name]
          }.compact
          s[:line] = options[:line] if options[:line]
          s[:fill] = options[:fill] if options[:fill]
          @series << s
          s
        end

        # Sets the chart title options
        #
        # @param options [Hash{Symbol => untyped}, String]
        # @return [Hash{Symbol => untyped}]
        #: (untyped options) -> Hash[Symbol, untyped]
        def set_title(options)
          opts = options.is_a?(Hash) ? options : { name: options.to_s }
          @title = opts
        end

        # Sets the X axis options
        #
        # @param options [Hash{Symbol => untyped}, String]
        # @return [Hash{Symbol => untyped}]
        #: (untyped options) -> Hash[Symbol, untyped]
        def set_x_axis(options)
          opts = options.is_a?(Hash) ? options : { name: options.to_s }
          @x_axis = opts
        end

        # Sets the Y axis options
        #
        # @param options [Hash{Symbol => untyped}, String]
        # @return [Hash{Symbol => untyped}]
        #: (untyped options) -> Hash[Symbol, untyped]
        def set_y_axis(options)
          opts = options.is_a?(Hash) ? options : { name: options.to_s }
          @y_axis = opts
        end

        # Sets the legend options
        #
        # @param options [Hash{Symbol => untyped}]
        # @return [Hash{Symbol => untyped}]
        #: (Hash[Symbol, untyped] options) -> Hash[Symbol, untyped]
        def set_legend(options)
          @legend = options
        end

        # Sets the chart style preset ID
        #
        # @param style_id [Integer]
        # @return [Integer]
        #: (Integer style_id) -> Integer
        def set_style(style_id)
          @style_id = style_id
        end

        # Sets the dimensions / scaling of the chart
        #
        # @param options [Hash{Symbol => untyped}]
        # @return [void]
        #: (Hash[Symbol, untyped] options) -> void
        def set_size(options = {})
          @width = options[:width] if options[:width]
          @height = options[:height] if options[:height]
          @x_scale = options[:x_scale] if options[:x_scale]
          @y_scale = options[:y_scale] if options[:y_scale]
        end

        # Compiles this chart into an options hash for xlsxrb
        #
        # @return [Hash{Symbol => untyped}]
        #: () -> Hash[Symbol, untyped]
        def to_chart_options
          chart_type = case @type.to_s.downcase
                       when "col", "column" then :col
                       when "bar" then :bar
                       when "line" then :line
                       when "pie" then :pie
                       when "scatter" then :scatter
                       when "area" then :area
                       when "radar" then :radar
                       when "doughnut" then :doughnut
                       else @type.to_sym
                       end

          title_str = @title[:name] || @title[:title]
          cat_title = @x_axis[:name] || @x_axis[:title]
          val_title = @y_axis[:name] || @y_axis[:title]

          mapped_series = @series.map do |s|
            series_hash = s.dup
            series_hash[:cat_ref] ||= series_hash[:categories]
            series_hash[:val_ref] ||= series_hash[:values]
            series_hash
          end

          {
            type: chart_type,
            title: title_str,
            series: mapped_series,
            cat_axis_title: cat_title,
            val_axis_title: val_title,
            legend: @legend.empty? ? nil : (@legend[:position] || "r"),
            style: @style_id
          }.compact
        end

        # Factory method to create appropriate Chart subclass instance
        #
        # @param type [String, Symbol]
        # @param subtype [String, Symbol, nil]
        # @return [Chart]
        #: (untyped type, ?untyped subtype) -> Chart
        def self.factory(type, subtype = nil)
          t = type.to_s.sub(/^Chart::/i, "").capitalize
          klass = case t
                  when "Area" then Area
                  when "Bar" then Bar
                  when "Column", "Col" then Column
                  when "Line" then Line
                  when "Pie" then Pie
                  when "Scatter" then Scatter
                  when "Radar" then Radar
                  when "Doughnut" then Doughnut
                  when "Stock" then Stock
                  else Chart
                  end
          chart = klass.allocate
          chart.send(:initialize, nil, type: type, subtype: subtype)
          chart
        end

        # Subclasses for WriteXLSX chart types
        class Area < Chart
          def default_chart_type = :area
        end

        class Bar < Chart
          def default_chart_type = :bar
        end

        class Column < Chart
          def default_chart_type = :column
        end

        class Line < Chart
          def default_chart_type = :line
        end

        class Pie < Chart
          def default_chart_type = :pie
        end

        class Scatter < Chart
          def default_chart_type = :scatter
        end

        class Radar < Chart
          def default_chart_type = :radar
        end

        class Doughnut < Chart
          def default_chart_type = :doughnut
        end

        class Stock < Chart
          def default_chart_type = :stock
        end
      end
    end
  end
end
