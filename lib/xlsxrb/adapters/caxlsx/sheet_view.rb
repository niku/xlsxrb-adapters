# frozen_string_literal: true

# rbs_inline: enabled

require_relative "util"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Selection definition
      class Selection
        include OptionsParser
        include SerializedAttributes

        attr_reader :pane #: Symbol?
        attr_accessor :active_cell #: String?
        attr_accessor :active_cell_id #: Integer?
        attr_accessor :sqref #: String?

        serializable_attributes :pane, :active_cell, :active_cell_id, :sqref

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          parse_options options
        end

        #: (untyped v) -> void
        def pane=(v)
          Caxlsx.validate_pane_type(v)
          @pane = v.to_sym
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("selection", str)
        end
      end

      # Pane definition (split/freeze)
      class Pane
        include OptionsParser
        include SerializedAttributes

        attr_accessor :x_split #: Integer?
        attr_accessor :y_split #: Integer?
        attr_accessor :top_left_cell #: String?
        attr_reader :active_pane #: Symbol?
        attr_reader :state #: Symbol?

        serializable_attributes :x_split, :y_split, :top_left_cell, :active_pane, :state

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @x_split = nil
          @y_split = nil
          @top_left_cell = nil
          @active_pane = nil
          @state = nil
          parse_options options
        end

        #: (untyped v) -> void
        def active_pane=(v)
          Caxlsx.validate_pane_type(v)
          @active_pane = Caxlsx.camel(v.to_s, false)
        end

        #: (untyped v) -> void
        def state=(v)
          Caxlsx.validate_split_state_type(v)
          @state = Caxlsx.camel(v.to_s, false)
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("pane", str)
        end
      end

      # Sheet view settings
      class SheetView
        include OptionsParser
        include SerializedAttributes

        attr_reader :selections #: SimpleTypedList
        attr_accessor :show_grid_lines #: bool?
        attr_accessor :show_row_col_headers #: bool?
        attr_accessor :show_zeros #: bool?
        attr_accessor :right_to_left #: bool?
        attr_accessor :tab_selected #: bool?
        attr_accessor :show_formulas #: bool?
        attr_accessor :zoom_scale #: Integer?
        attr_accessor :zoom_scale_normal #: Integer?
        attr_reader :view #: Symbol?
        attr_accessor :top_left_cell #: String?
        attr_accessor :color_id #: Integer?
        attr_accessor :default_grid_color #: bool?
        attr_accessor :show_outline_symbols #: bool?
        attr_accessor :show_ruler #: bool?
        attr_accessor :show_white_space #: bool?
        attr_accessor :window_protection #: bool?
        attr_accessor :workbook_view_id #: Integer?
        attr_accessor :zoom_scale_page_layout_view #: Integer?
        attr_accessor :zoom_scale_sheet_layout_view #: Integer?

        serializable_attributes :show_grid_lines, :show_row_col_headers, :show_zeros,
                                :right_to_left, :tab_selected, :show_formulas, :zoom_scale,
                                :view, :top_left_cell, :workbook_view_id

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @selections = SimpleTypedList.new(Selection)
          @pane = nil
          @show_grid_lines = true
          @show_row_col_headers = true
          @tab_selected = nil
          @view = :normal
          @workbook_view_id = 0
          parse_options options
        end

        # Accesses or defines the pane
        # @yieldparam pane [Pane]
        # @return [Pane]
        #: () ?{ (Pane) -> void } -> Pane
        def pane
          @pane ||= Pane.new
          yield @pane if block_given?
          @pane
        end

        #: (Symbol pane_name, ?Hash[Symbol, untyped] options) -> Selection
        def add_selection(pane_name, options = {})
          sel = Selection.new(options.merge(pane: pane_name))
          @selections << sel
          sel
        end

        #: (untyped v) -> void
        def view=(v)
          Caxlsx.validate_sheet_view_type(v)
          @view = v.to_sym
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<sheetViews><sheetView"
          serialized_attributes(str)
          str << ">"
          @pane&.to_xml_string(str)
          @selections.each { |s| s.to_xml_string(str) }
          str << "</sheetView></sheetViews>"
          str
        end

        #: () -> Hash[Symbol, untyped]
        def to_xlsxrb_hash
          res = {}
          res[:show_grid_lines] = @show_grid_lines unless @show_grid_lines.nil?
          res[:show_row_col_headers] = @show_row_col_headers unless @show_row_col_headers.nil?
          res[:zoom_scale] = @zoom_scale if @zoom_scale
          res[:right_to_left] = @right_to_left unless @right_to_left.nil?
          res
        end

        #: () -> Hash[Symbol, untyped]?
        def to_freeze_pane_hash
          return nil unless @pane && %i[frozen frozen_split].include?(@pane.state)

          {
            col: @pane.x_split || 0,
            row: @pane.y_split || 0,
            state: :frozen
          }
        end

        #: () -> Hash[Symbol, untyped]?
        def to_split_pane_hash
          return nil unless @pane && @pane.state == :split

          {
            x_split: @pane.x_split || 0,
            y_split: @pane.y_split || 0,
            top_left_cell: @pane.top_left_cell,
            state: :split
          }
        end
      end
    end
  end
end
