# frozen_string_literal: true

# rbs_inline: enabled

require "tempfile"
require "fileutils"
require_relative "styles"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Subclass of StreamWriter that resolves integer style IDs to themselves
      # and supports custom style definitions from Caxlsx::Styles.
      class CaxlsxStreamWriter < Xlsxrb::StreamWriter
        # @return [Hash{Symbol => untyped}, nil]
        attr_accessor :custom_styles_hash

        # Initializes writer with integer style ID pass-through.
        #
        # @param target [String, IO, StringIO] Destination target.
        # @param opts [Hash{Symbol => untyped}] Writer options.
        def initialize(target, **opts)
          super
          @style_name_to_id.default_proc = ->(_h, k) { k.is_a?(Integer) ? k : (Integer(k, exception: false) || nil) }
        end

        # Closes the writer, injecting custom styles if configured.
        #
        # @return [void]
        def close
          raise ArgumentError, "Workbook must contain at least one sheet (Excel limitation)" if @strict_excel_mode && @sheets.empty? && @current_sheet.nil?

          flush_current_sheet

          styles_definition = @custom_styles_hash || {
            fonts: @style_writer.fonts.dup,
            fills: @style_writer.fills.dup,
            borders: @style_writer.borders.dup,
            xf_entries: @style_writer.xf_entries.dup,
            num_fmts: @style_writer.num_fmts.dup,
            dxfs: @dxfs || []
          }

          resolved_names = resolve_defined_names(@defined_names, @sheets)

          wb_writer = Xlsxrb::Ooxml::WorkbookWriter.new(
            sheets: @sheets,
            shared_strings: @sst,
            shared_strings_index: @sst_index,
            styles: styles_definition,
            defined_names: resolved_names.empty? ? nil : resolved_names,
            core_properties: @core_properties.empty? ? nil : @core_properties,
            app_properties: @app_properties.empty? ? nil : @app_properties,
            custom_properties: @custom_properties.empty? ? nil : @custom_properties,
            workbook_protection: @workbook_protection,
            workbook_properties: @workbook_properties
          )
          wb_writer.write_package_parts(@zip)

          @zip.close
          @io.close if @owns_io && !@io.closed?
        ensure
          cleanup!
        end
      end

      # High-performance streaming worksheet providing Caxlsx DSL ergonomics
      # while streaming row data directly into zip output with O(1) memory.
      class StreamingWorksheet
        # @return [StreamingWorkbook]
        attr_reader :workbook

        # @return [String]
        attr_reader :name

        # @return [Integer]
        attr_reader :row_count

        # @param workbook [StreamingWorkbook]
        # @param name [String]
        # @param stream_writer [CaxlsxStreamWriter]
        def initialize(workbook, name, stream_writer)
          @workbook = workbook
          @name = name
          @stream_writer = stream_writer
          @row_count = 0
          @auto_filter_proxy = nil
          @sheet_view_proxy = nil
        end

        # Appends a row of cells to the active sheet in constant memory.
        #
        # @param values [Array<Object>] Row cell values.
        # @param options [Hash{Symbol => untyped}] Row options (e.g. :style, :height, :offset).
        # @return [self]
        def add_row(values = [], options = {})
          offset = options[:offset]
          if offset&.positive?
            values = Array.new(offset, nil) + values
            options[:style] = Array.new(offset, nil) + options[:style] if options[:style].is_a?(Array)
          end

          style = options[:style]
          style = nil if style.is_a?(Array) && style.compact.empty?

          @stream_writer.row(
            values,
            styles: style,
            height: options[:height],
            hidden: options[:hidden] || false,
            custom_height: options[:custom_height] || !options[:height].nil?,
            outline_level: options[:outline_level]
          )
          @row_count += 1
          self
        end
        alias << add_row

        # Rows list compatibility stub (streaming worksheets do not retain rows in memory).
        #
        # @return [Array]
        def rows
          []
        end

        # Merges a range of cells.
        #
        # @param range [String, Array<String, Cell>] Cell range or start/end cells.
        # @return [void]
        def merge_cells(range)
          ref = if range.is_a?(Array)
                  r1 = range.first.respond_to?(:r) ? range.first.r : range.first.to_s
                  r2 = range.last.respond_to?(:r) ? range.last.r : range.last.to_s
                  "#{r1}:#{r2}"
                else
                  range.to_s
                end
          @stream_writer.merge(ref)
        end

        # Sets column widths for successive columns.
        #
        # @param widths [Array<Numeric, nil>] Widths in character units.
        # @return [void]
        def column_widths(*widths)
          widths.flatten.each_with_index do |w, idx|
            @stream_writer.column(idx, width: w) if w
          end
        end

        # Sets or returns the auto filter range.
        #
        # @param range [String, nil]
        # @return [String, Object]
        def auto_filter=(range)
          @stream_writer.auto_filter(range.to_s)
        end

        # @return [Object] Auto-filter proxy supporting `.range = ...`
        def auto_filter
          @auto_filter ||= Class.new do
            def initialize(writer)
              @writer = writer
            end

            def range=(val)
              @writer.auto_filter(val.to_s)
            end
          end.new(@stream_writer)
        end

        # Adds a conditional formatting rule.
        #
        # @param cells [String] Cell range (e.g. "A1:A10").
        # @param rules [Array<Hash>, Hash] Formatting rules.
        # @return [void]
        def add_conditional_formatting(cells, rules)
          rule_list = rules.is_a?(Array) ? rules : [rules]
          rule_list.each do |rule|
            opts = rule.is_a?(Hash) ? rule : {}
            @stream_writer.conditional_format(cells, **opts)
          end
        end

        # Adds a data validation rule.
        #
        # @param cells [String] Cell range.
        # @param data_validation [Hash{Symbol => untyped}]
        # @return [void]
        def add_data_validation(cells, data_validation)
          opts = data_validation.dup
          opts[:allow_blank] = opts.delete(:allowBlank) if opts.key?(:allowBlank)
          opts[:show_input_message] = opts.delete(:showInputMessage) if opts.key?(:showInputMessage)
          opts[:show_error_message] = opts.delete(:showErrorMessage) if opts.key?(:showErrorMessage)
          opts[:error_style] = opts.delete(:errorStyle) if opts.key?(:errorStyle)
          opts[:error_title] = opts.delete(:errorTitle) if opts.key?(:errorTitle)
          opts[:prompt_title] = opts.delete(:promptTitle) if opts.key?(:promptTitle)
          @stream_writer.validate_data(cells, **opts)
        end

        # Configures sheet view and freeze panes.
        #
        # @yield [pane]
        # @return [Object]
        def sheet_view
          @sheet_view ||= Class.new do
            def initialize(writer)
              @writer = writer
            end

            def pane
              pane_obj = Struct.new(:top_left_cell, :state, :x_split, :y_split, :active_pane).new
              yield pane_obj if block_given?
              @writer.freeze_pane(row: pane_obj.y_split, col: pane_obj.x_split) if pane_obj.y_split || pane_obj.x_split
              pane_obj
            end
          end.new(@stream_writer)
        end
      end

      # High-performance streaming workbook managing worksheets and styles.
      class StreamingWorkbook
        # @return [StreamingPackage]
        attr_reader :package

        # @return [Styles]
        attr_reader :styles

        # @return [Array<StreamingWorksheet>]
        attr_reader :worksheets

        # @return [Boolean]
        attr_reader :date1904

        # @param package [StreamingPackage]
        # @param stream_writer [CaxlsxStreamWriter]
        def initialize(package, stream_writer)
          @package = package
          @stream_writer = stream_writer
          @styles = Styles.new
          @worksheets = []
          @date1904 = false
        end

        # @param v [Boolean]
        def date1904=(v)
          Caxlsx.validate_boolean(v)
          @date1904 = v
          @stream_writer.instance_variable_set(:@date1904, v)
          @stream_writer.workbook_property(:date1904, v)
        end

        # Adds a new streaming worksheet to the workbook.
        #
        # @param options [Hash{Symbol => untyped}] Worksheet options.
        # @yield [worksheet]
        # @yieldparam worksheet [StreamingWorksheet]
        # @return [StreamingWorksheet]
        def add_worksheet(options = {}, &)
          name = options[:name] || "Sheet#{@worksheets.size + 1}"
          streaming_sheet = nil

          if block_given?
            @stream_writer.sheet(name) do |_proxy|
              streaming_sheet = StreamingWorksheet.new(self, name, @stream_writer)
              @worksheets << streaming_sheet
              yield streaming_sheet
            end
          else
            @stream_writer.sheet(name)
            streaming_sheet = StreamingWorksheet.new(self, name, @stream_writer)
            @worksheets << streaming_sheet
          end

          streaming_sheet
        end
      end

      # High-performance streaming package that provides Caxlsx DSL ergonomics
      # while streaming row data directly into zip output with O(1) memory.
      class StreamingPackage
        # @return [StreamingWorkbook]
        attr_reader :workbook

        # Opens a streaming package with block syntax and automatic close.
        #
        # @example
        #   StreamingPackage.open("large.xlsx") do |pkg|
        #     pkg.workbook.add_worksheet(name: "Data") do |sheet|
        #       1_000_000.times { |i| sheet.add_row [i, "Row #{i}"] }
        #     end
        #   end
        #
        # @param target [String, IO, StringIO, nil] File path or writable IO.
        # @param options [Hash{Symbol => untyped}]
        # @yield [package]
        # @return [StreamingPackage, void]
        def self.open(target = nil, **)
          pkg = new(target, **)
          if block_given?
            begin
              yield pkg
            ensure
              pkg.close
            end
          else
            pkg
          end
        end

        # Initializes a streaming package.
        #
        # @param target [String, IO, StringIO, nil] File path or writable IO.
        # @param options [Hash{Symbol => untyped}]
        def initialize(target = nil, **_options)
          if target
            @target = target
            @owns_tempfile = false
            @tempfile = nil
          else
            @tempfile = Tempfile.new(["xlsxrb_stream", ".xlsx"], binmode: true)
            @target = @tempfile.path
            @owns_tempfile = true
          end

          @stream_writer = CaxlsxStreamWriter.new(@target)
          @workbook = StreamingWorkbook.new(self, @stream_writer)
          @closed = false

          return unless block_given?

          begin
            yield self
          ensure
            close
          end
        end

        # Serializes the streaming package to disk.
        #
        # @param output_path [String, nil] Destination file path.
        # @param _options [Hash, Boolean]
        # @param _secondary_options [Hash, nil]
        # @return [Boolean]
        def serialize(output_path = nil, _options = {}, _secondary_options = nil)
          close
          FileUtils.cp(@target, output_path) if output_path && output_path != @target
          true
        end

        # Serializes the workbook to a StringIO instance.
        #
        # @param _old_confirm_valid [Boolean, nil]
        # @param _confirm_valid [Boolean]
        # @param _password [String, nil]
        # @return [StringIO]
        def to_stream(_old_confirm_valid = nil, _confirm_valid: false, _password: nil)
          close
          content = File.binread(@target)
          StringIO.new(content)
        end

        # Finalizes and closes streaming output.
        #
        # @return [void]
        def close
          return if @closed

          styles_hash = @workbook.styles.to_xlsxrb_hash
          @stream_writer.custom_styles_hash = styles_hash if styles_hash
          @stream_writer.close
          @closed = true
        ensure
          if @owns_tempfile && @tempfile && @closed
            begin
              @tempfile.close
            rescue StandardError
              nil
            end
          end
        end
      end
    end
  end
end
