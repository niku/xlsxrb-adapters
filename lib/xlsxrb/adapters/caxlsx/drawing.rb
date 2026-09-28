# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # Mime type utility functions for images.
      module MimeTypeUtils
        EXTENSION_MAP = {
          ".jpg" => "image/jpeg",
          ".jpeg" => "image/jpeg",
          ".png" => "image/png",
          ".gif" => "image/gif"
        }.freeze

        class << self
          # @param filepath [String]
          # @return [String]
          #: (String filepath) -> String
          def get_mime_type(filepath)
            ext = File.extname(filepath).downcase
            EXTENSION_MAP[ext] || "application/octet-stream"
          end

          # @param uri_str [String]
          # @return [String]
          #: (String uri_str) -> String
          def get_mime_type_from_uri(uri_str)
            ext = File.extname(URI.parse(uri_str).path).downcase
            EXTENSION_MAP[ext] || "image/jpeg"
          end
        end
      end

      # Picture locking properties.
      class PictureLocking
        include OptionsParser
        include SerializedAttributes
        include Accessors

        boolean_attr_accessor :noGrp, :noSelect, :noRot, :noChangeAspect,
                              :noMove, :noResize, :noEditPoints, :noAdjustHandles,
                              :noChangeArrowheads, :noChangeShapeType

        serializable_attributes :noGrp, :noSelect, :noRot, :noChangeAspect,
                                :noMove, :noResize, :noEditPoints, :noAdjustHandles,
                                :noChangeArrowheads, :noChangeShapeType

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @noChangeAspect = true
          parse_options(options)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("a:picLocks", str)
        end
      end

      # A drawing hyperlink object.
      class Hyperlink
        include SerializedAttributes
        include OptionsParser

        serializable_attributes :invalid_url, :action, :end_snd, :highlight_click, :history, :tgt_frame, :tooltip

        # @return [Pic]
        attr_reader :parent

        # @return [String, nil]
        attr_accessor :href

        # @return [String, nil]
        attr_accessor :invalid_url
        alias invalidUrl invalid_url
        alias invalidUrl= invalid_url=

        # @return [String, nil]
        attr_accessor :action

        # @return [Boolean, nil]
        attr_reader :end_snd
        alias endSnd end_snd

        # @return [Boolean, nil]
        attr_reader :highlight_click
        alias highlightClick highlight_click

        # @return [Boolean, nil]
        attr_reader :history

        # @return [String, nil]
        attr_accessor :tgt_frame
        alias tgtFrame tgt_frame
        alias tgtFrame= tgt_frame=

        # @return [String, nil]
        attr_accessor :tooltip

        # @param parent [Pic]
        # @param options [Hash{Symbol => untyped}]
        #: (Pic parent, ?Hash[Symbol, untyped] options) ?{ (Hyperlink) -> void } -> void
        def initialize(parent, options = {})
          DataTypeValidator.validate("Hyperlink.parent", [Pic], parent)
          @parent = parent
          parse_options(options)
          yield self if block_given?
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def end_snd=(v)
          Caxlsx.validate_boolean(v)
          @end_snd = v
        end
        alias endSnd= end_snd=

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def highlight_click=(v)
          Caxlsx.validate_boolean(v)
          @highlight_click = v
        end
        alias highlightClick= highlight_click=

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def history=(v)
          Caxlsx.validate_boolean(v)
          @history = v
        end

        # @return [Relationship]
        #: () -> Relationship
        def relationship
          Relationship.new(self, HYPERLINK_R, href.to_s, target_mode: :External)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          serialized_tag("a:hlinkClick", str, { "r:id": relationship.Id, "xmlns:r": XML_NS_R })
        end
      end

      # Defines a point in the worksheet that drawing anchors attach to.
      class Marker
        include OptionsParser

        # @return [Integer]
        attr_reader :col

        # @return [Integer]
        attr_reader :colOff

        # @return [Integer]
        attr_reader :row

        # @return [Integer]
        attr_reader :rowOff

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @col = @colOff = @row = @rowOff = 0
          parse_options(options)
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def col=(v)
          Caxlsx.validate_unsigned_int(v)
          @col = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def colOff=(v)
          Caxlsx.validate_int(v)
          @colOff = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def row=(v)
          Caxlsx.validate_unsigned_int(v)
          @row = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def rowOff=(v)
          Caxlsx.validate_int(v)
          @rowOff = v
        end

        # @param col [Integer, String, Cell, Array[Integer]]
        # @param row [Integer]
        # @return [void]
        #: (Integer | String | Cell | Array[Integer] col, ?Integer row) -> void
        def coord(col, row = 0)
          coordinates = parse_coord_args(col, row)
          self.col = coordinates[0]
          self.row = coordinates[1]
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          %i[col colOff row rowOff].each do |k|
            str << "<xdr:" << k.to_s << ">" << send(k).to_s << "</xdr:" << k.to_s << ">"
          end
          str
        end

        private

        # @param x [Integer, String, Cell, Array[Integer]]
        # @param y [Integer]
        # @return [Array[Integer]]
        #: (Integer | String | Cell | Array[Integer] x, ?Integer y) -> Array[Integer]
        def parse_coord_args(x, y = 0)
          case x
          when String
            x, y = *Caxlsx.name_to_indices(x)
          when Cell
            x, y = *x.pos
          when Array
            x, y = *x
          end
          [x.to_i, y.to_i]
        end
      end

      # Picture object in a drawing.
      class Pic
        include OptionsParser

        ALLOWED_MIME_TYPES = %w[image/jpeg image/png image/gif].freeze

        # @return [String, nil]
        attr_reader :name

        # @return [String, nil]
        attr_reader :descr

        # @return [String, nil]
        attr_reader :image_src

        # @return [OneCellAnchor, TwoCellAnchor]
        attr_reader :anchor

        # @return [PictureLocking]
        attr_reader :picture_locking

        # @return [Hyperlink, nil]
        attr_reader :hyperlink

        # @return [Integer, nil]
        attr_reader :opacity

        # @return [Boolean, nil]
        attr_reader :remote

        # @param anchor [OneCellAnchor, TwoCellAnchor]
        # @param options [Hash{Symbol => untyped}]
        #: (OneCellAnchor | TwoCellAnchor anchor, ?Hash[Symbol, untyped] options) ?{ (Pic) -> void } -> void
        def initialize(anchor, options = {})
          @anchor = anchor
          @picture_locking = PictureLocking.new(options)
          @name = ""
          @remote = false
          @image_src = nil
          @descr = ""
          @hyperlink = nil
          @opacity = nil
          parse_options(options)
          yield self if block_given?
        end

        # @param v [String]
        # @param options [Hash{Symbol => untyped}]
        # @return [Hyperlink]
        #: (String v, ?Hash[Symbol, untyped] options) -> Hyperlink
        def hyperlink=(v, options = {})
          options[:href] = v
          if hyperlink.is_a?(Hyperlink)
            options.each do |key, val|
              hyperlink.send(:"#{key}=", val) if hyperlink.respond_to?(:"#{key}=")
            end
          else
            @hyperlink = Hyperlink.new(self, options)
          end
          @hyperlink
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def image_src=(v)
          Caxlsx.validate_string(v)
          if remote?
            RestrictionValidator.validate("Pic.image_src", ALLOWED_MIME_TYPES, MimeTypeUtils.get_mime_type_from_uri(v))
          else
            RestrictionValidator.validate("Pic.image_src", ALLOWED_MIME_TYPES, MimeTypeUtils.get_mime_type(v))
            raise ArgumentError, "File does not exist" unless File.exist?(v)
          end

          @image_src = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def name=(v)
          Caxlsx.validate_string(v)
          @name = v
        end

        # @param v [String]
        # @return [String]
        #: (String v) -> String
        def descr=(v)
          Caxlsx.validate_string(v)
          @descr = v
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def remote=(v)
          Caxlsx.validate_boolean(v)
          @remote = v
        end

        # @return [Boolean]
        #: () -> bool
        def remote?
          remote == 1 || remote.to_s == "true"
        end

        # @return [String, nil]
        #: () -> String?
        def file_name
          File.basename(image_src) unless remote? || image_src.nil?
        end

        # @return [String, nil]
        #: () -> String?
        def extname
          File.extname(image_src).delete(".") unless image_src.nil?
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @anchor.drawing.worksheet.workbook.images.index(self) || 0
        end

        # @return [String]
        #: () -> String
        def pn
          format(IMAGE_PN, index + 1, extname)
        end

        # @return [Relationship]
        #: () -> Relationship
        def relationship
          if remote?
            Relationship.new(self, IMAGE_R, image_src.to_s, target_mode: :External)
          else
            Relationship.new(self, IMAGE_R, "../#{pn}")
          end
        end

        # @return [Integer, nil]
        #: () -> Integer?
        def width
          return unless @anchor.is_a?(OneCellAnchor)

          @anchor.width
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def width=(v)
          use_one_cell_anchor unless @anchor.is_a?(OneCellAnchor)
          @anchor.width = v
        end

        # @return [Integer, nil]
        #: () -> Integer?
        def height
          return unless @anchor.is_a?(OneCellAnchor)

          @anchor.height
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def height=(v)
          use_one_cell_anchor unless @anchor.is_a?(OneCellAnchor)
          @anchor.height = v
        end

        # @param x [Integer]
        # @param y [Integer, nil]
        # @return [Marker]
        #: (Integer x, ?Integer? y) -> Marker
        def start_at(x, y = nil)
          @anchor.start_at(x, y || 0)
          @anchor.from
        end

        # @param x [Integer]
        # @param y [Integer, nil]
        # @return [Marker]
        #: (Integer x, ?Integer? y) -> Marker
        def end_at(x, y = nil)
          use_two_cell_anchor unless @anchor.is_a?(TwoCellAnchor)
          @anchor.end_at(x, y || 0)
          @anchor.to
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<xdr:pic>"
          str << "<xdr:nvPicPr>"
          str << '<xdr:cNvPr id="2" name="' << name.to_s << '" descr="' << descr.to_s << '">'
          hyperlink.to_xml_string(str) if hyperlink.is_a?(Hyperlink)
          str << "</xdr:cNvPr><xdr:cNvPicPr>"
          picture_locking.to_xml_string(str)
          str << "</xdr:cNvPicPr></xdr:nvPicPr>"
          str << "<xdr:blipFill>"
          str << relationship_xml_portion
          str << "<a:alphaModFix amt=\"#{opacity}\"/>" if opacity
          str << "</a:blip>"
          str << "<a:stretch><a:fillRect/></a:stretch></xdr:blipFill><xdr:spPr>"
          str << '<a:xfrm><a:off x="0" y="0"/><a:ext cx="2336800" cy="2161540"/></a:xfrm>'
          str << '<a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic>'
        end

        private

        # @return [String]
        #: () -> String
        def relationship_xml_portion
          if remote?
            +"<a:blip xmlns:r=\"" << XML_NS_R << '" r:link="' << relationship.Id << '">'
          else
            +"<a:blip xmlns:r=\"" << XML_NS_R << '" r:embed="' << relationship.Id << '">'
          end
        end

        # @return [void]
        #: () -> void
        def use_one_cell_anchor
          return if @anchor.is_a?(OneCellAnchor)

          new_anchor = OneCellAnchor.new(@anchor.drawing, start_at: [@anchor.from.col, @anchor.from.row])
          swap_anchor(new_anchor)
        end

        # @return [void]
        #: () -> void
        def use_two_cell_anchor
          return if @anchor.is_a?(TwoCellAnchor)

          new_anchor = TwoCellAnchor.new(@anchor.drawing, start_at: [@anchor.from.col, @anchor.from.row])
          swap_anchor(new_anchor)
        end

        # @param new_anchor [OneCellAnchor, TwoCellAnchor]
        # @return [void]
        #: (OneCellAnchor | TwoCellAnchor new_anchor) -> void
        def swap_anchor(new_anchor)
          new_anchor.drawing.anchors.delete(new_anchor)
          idx = @anchor.drawing.anchors.index(@anchor)
          @anchor.drawing.anchors[idx] = new_anchor if idx
          new_anchor.instance_variable_set(:@object, @anchor.object)
          @anchor = new_anchor
        end
      end

      # Graphic frame container for chart objects.
      class GraphicFrame
        # @return [Chart]
        attr_reader :chart

        # @return [TwoCellAnchor]
        attr_reader :anchor

        # @param anchor [TwoCellAnchor]
        # @param chart_type [Class]
        # @param options [Hash{Symbol => untyped}]
        #: (TwoCellAnchor anchor, Class chart_type, Hash[Symbol, untyped] options) -> void
        def initialize(anchor, chart_type, options)
          @anchor = anchor
          @chart = chart_type.new(self, options)
        end

        # @return [String]
        #: () -> String
        def rId
          @anchor.drawing.relationships.for(chart).Id
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<xdr:graphicFrame>"
          str << "<xdr:nvGraphicFramePr>"
          str << '<xdr:cNvPr id="' << @anchor.drawing.index.to_s << '" name="item_' << @anchor.drawing.index.to_s << '"/>'
          str << "<xdr:cNvGraphicFramePr/>"
          str << "</xdr:nvGraphicFramePr>"
          str << "<xdr:xfrm>"
          str << '<a:off x="0" y="0"/>'
          str << '<a:ext cx="0" cy="0"/>'
          str << "</xdr:xfrm>"
          str << "<a:graphic>"
          str << '<a:graphicData uri="' << XML_NS_C << '">'
          str << '<c:chart xmlns:c="' << XML_NS_C << '" xmlns:r="' << XML_NS_R << '" r:id="' << rId << '"/>'
          str << "</a:graphicData>"
          str << "</a:graphic>"
          str << "</xdr:graphicFrame>"
        end
      end

      # Single cell anchor for images.
      class OneCellAnchor
        include OptionsParser

        # @return [Marker]
        attr_reader :from

        # @return [Pic]
        attr_reader :object

        # @return [Drawing]
        attr_reader :drawing

        # @return [Integer]
        attr_reader :width

        # @return [Integer]
        attr_reader :height

        # @param drawing [Drawing]
        # @param options [Hash{Symbol => untyped}]
        #: (Drawing drawing, ?Hash[Symbol, untyped] options) -> void
        def initialize(drawing, options = {})
          @drawing = drawing
          @width = 0
          @height = 0
          drawing.anchors << self
          @from = Marker.new
          parse_options(options)
          start_at(*options[:start_at]) if options[:start_at]
          @object = Pic.new(self, options)
        end

        # @param x [Integer, String, Cell, Array[Integer]]
        # @param y [Integer]
        # @return [void]
        #: (Integer | String | Cell | Array[Integer] x, ?Integer y) -> void
        def start_at(x, y = 0)
          from.coord(x, y)
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def height=(v)
          Caxlsx.validate_unsigned_int(v)
          @height = v
        end

        # @param v [Integer]
        # @return [Integer]
        #: (Integer v) -> Integer
        def width=(v)
          Caxlsx.validate_unsigned_int(v)
          @width = v
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @drawing.anchors.index(self) || 0
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<xdr:oneCellAnchor>"
          str << "<xdr:from>"
          from.to_xml_string(str)
          str << "</xdr:from>"
          str << '<xdr:ext cx="' << ext[:cx].to_s << '" cy="' << ext[:cy].to_s << '"/>'
          @object.to_xml_string(str)
          str << "<xdr:clientData/>"
          str << "</xdr:oneCellAnchor>"
        end

        private

        # @return [Hash{Symbol => Integer}]
        #: () -> Hash[Symbol, Integer]
        def ext
          cy = @height * 914_400 / 96
          cx = @width * 914_400 / 96
          { cy: cy, cx: cx }
        end
      end

      # Two cell anchor for drawings and charts.
      class TwoCellAnchor
        include OptionsParser

        # @return [Marker]
        attr_reader :from

        # @return [Marker]
        attr_reader :to

        # @return [Pic, GraphicFrame, untyped]
        attr_reader :object

        # @return [Drawing]
        attr_reader :drawing

        # @param drawing [Drawing]
        # @param options [Hash{Symbol => untyped}]
        #: (Drawing drawing, ?Hash[Symbol, untyped] options) -> void
        def initialize(drawing, options = {})
          @drawing = drawing
          drawing.anchors << self
          @from = Marker.new
          @to = Marker.new(col: 5, row: 10)
          parse_options(options)

          start_at(*options[:start_at]) if options[:start_at]
          end_at(*options[:end_at]) if options[:end_at]
        end

        # @param x [Integer, String, Cell, Array[Integer]]
        # @param y [Integer, nil]
        # @return [void]
        #: (Integer | String | Cell | Array[Integer] x, ?Integer? y) -> void
        def start_at(x, y = nil)
          from.coord(x, y || 0)
        end

        # @param x [Integer, String, Cell, Array[Integer]]
        # @param y [Integer, nil]
        # @return [void]
        #: (Integer | String | Cell | Array[Integer] x, ?Integer? y) -> void
        def end_at(x, y = nil)
          to.coord(x, y || 0)
        end

        # @param chart_type [Class]
        # @param options [Hash{Symbol => untyped}]
        # @return [Chart]
        #: (Class chart_type, Hash[Symbol, untyped] options) -> Chart
        def add_chart(chart_type, options)
          @object = GraphicFrame.new(self, chart_type, options)
          @object.chart
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Pic]
        #: (?Hash[Symbol, untyped] options) -> Pic
        def add_pic(options = {})
          @object = Pic.new(self, options)
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @drawing.anchors.index(self) || 0
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << "<xdr:twoCellAnchor>"
          str << "<xdr:from>"
          from.to_xml_string(str)
          str << "</xdr:from>"
          str << "<xdr:to>"
          to.to_xml_string(str)
          str << "</xdr:to>"
          object&.to_xml_string(str)
          str << "<xdr:clientData/>"
          str << "</xdr:twoCellAnchor>"
        end
      end

      # Canvas for charts and images in a worksheet.
      class Drawing
        # @return [Worksheet]
        attr_reader :worksheet

        # @return [SimpleTypedList]
        attr_reader :anchors

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          DataTypeValidator.validate("Drawing.worksheet", Worksheet, worksheet)
          @worksheet = worksheet
          @worksheet.workbook.drawings << self
          @anchors = SimpleTypedList.new([TwoCellAnchor, OneCellAnchor])
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Pic]
        #: (?Hash[Symbol, untyped] options) -> Pic
        def add_image(options = {})
          if options[:end_at]
            TwoCellAnchor.new(self, options).add_pic(options)
          else
            OneCellAnchor.new(self, options)
          end
          @anchors.last.object
        end

        # @param chart_type [Class]
        # @param options [Hash{Symbol => untyped}]
        # @return [Chart]
        #: (Class chart_type, ?Hash[Symbol, untyped] options) -> Chart
        def add_chart(chart_type, options = {})
          TwoCellAnchor.new(self, options)
          @anchors.last.add_chart(chart_type, options)
        end

        # @return [Array[Chart]]
        #: () -> Array[Chart]
        def charts
          frames = @anchors.select { |a| a.object.is_a?(GraphicFrame) }
          frames.map { |a| a.object.chart }
        end

        # @return [Array[Hyperlink]]
        #: () -> Array[Hyperlink]
        def hyperlinks
          links = images.select { |a| a.hyperlink.is_a?(Hyperlink) }
          links.map(&:hyperlink).compact
        end

        # @return [Array[Pic]]
        #: () -> Array[Pic]
        def images
          pics = @anchors.select { |a| a.object.is_a?(Pic) }
          pics.map(&:object)
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @worksheet.workbook.drawings.index(self) || 0
        end

        # @return [String]
        #: () -> String
        def pn
          format(DRAWING_PN, index + 1)
        end

        # @return [String]
        #: () -> String
        def rels_pn
          format(DRAWING_RELS_PN, index + 1)
        end

        # @return [Array[untyped]]
        #: () -> Array[untyped]
        def child_objects
          charts + images + hyperlinks
        end

        # @return [Relationships]
        #: () -> Relationships
        def relationships
          r = Relationships.new
          child_objects.each { |child| r << child.relationship }
          r
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>'
          str << '<xdr:wsDr xmlns:xdr="' << XML_NS_XDR << '" xmlns:a="' << XML_NS_A << '">'
          anchors.each { |anchor| anchor.to_xml_string(str) }
          str << "</xdr:wsDr>"
        end
      end
    end
  end
end
