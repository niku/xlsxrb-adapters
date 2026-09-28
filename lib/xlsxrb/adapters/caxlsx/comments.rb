# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # A VmlShape is used to position and render a comment.
      class VmlShape
        include OptionsParser
        include Accessors

        unsigned_int_attr_accessor :row, :column, :left_column, :left_offset, :top_row, :top_offset,
                                   :right_column, :right_offset, :bottom_row, :bottom_offset

        boolean_attr_accessor :visible

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) ?{ (VmlShape) -> void } -> void
        def initialize(options = {})
          @row = @column = @left_column = @top_row = @right_column = @bottom_row = 0
          @left_offset = 15
          @top_offset = 2
          @right_offset = 50
          @bottom_offset = 5
          @visible = true
          @id = Array.new(8) { rand(65..89).chr }.join
          parse_options(options)
          yield self if block_given?
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << <<~XML

            <v:shape id="#{@id}" type="#_x0000_t202" fillcolor="#ffffa1 [80]" o:insetmode="auto"
              style="visibility:#{@visible ? "visible" : "hidden"}">
              <v:fill color2="#ffffa1 [80]"/>
              <v:shadow on="t" obscured="t"/>
              <v:path o:connecttype="none"/>
              <v:textbox style='mso-fit-text-with-word-wrap:t'>
               <div style='text-align:left'></div>
              </v:textbox>

              <x:ClientData ObjectType="Note">
               <x:MoveWithCells/>
               <x:SizeWithCells/>
               <x:Anchor>#{left_column}, #{left_offset}, #{top_row}, #{top_offset}, #{right_column}, #{right_offset}, #{bottom_row}, #{bottom_offset}</x:Anchor>
               <x:AutoFill>False</x:AutoFill>
               <x:Row>#{row}</x:Row>
               <x:Column>#{column}</x:Column>
               #{"<x:Visible/>" if @visible}
              </x:ClientData>
             </v:shape>
          XML
        end
      end

      # A vml drawing used for comments in Excel.
      class VmlDrawing
        # @return [Comments]
        attr_reader :comments

        # @param comments [Comments]
        #: (Comments comments) -> void
        def initialize(comments)
          raise ArgumentError, "you must provide a comments object" unless comments.is_a?(Comments)

          @comments = comments
        end

        # @return [String]
        #: () -> String
        def pn
          format(VML_DRAWING_PN, @comments.worksheet.index + 1)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << <<~XML
            <xml xmlns:v="urn:schemas-microsoft-com:vml"
             xmlns:o="urn:schemas-microsoft-com:office:office"
             xmlns:x="urn:schemas-microsoft-com:office:excel">
             <o:shapelayout v:ext="edit">
              <o:idmap v:ext="edit" data="#{@comments.worksheet.index + 1}"/>
             </o:shapelayout>
             <v:shapetype id="_x0000_t202" coordsize="21600,21600" o:spt="202"
              path="m0,0l0,21600,21600,21600,21600,0xe">
              <v:stroke joinstyle="miter"/>
              <v:path gradientshapeok="t" o:connecttype="rect"/>
             </v:shapetype>
          XML
          @comments.each { |comment| comment.vml_shape.to_xml_string(str) }
          str << "</xml>"
        end
      end

      # A comment is the text data for a comment.
      class Comment
        include OptionsParser
        include Accessors

        string_attr_accessor :text, :author
        boolean_attr_accessor :visible

        # @return [Comments]
        attr_reader :comments

        # @return [String, nil]
        attr_reader :ref

        # @param comments [Comments]
        # @param options [Hash{Symbol => untyped}]
        #: (Comments comments, ?Hash[Symbol, untyped] options) ?{ (Comment) -> void } -> void
        def initialize(comments, options = {})
          raise ArgumentError, "A comment needs a parent comments object" unless comments.is_a?(Comments)

          @visible = true
          @comments = comments
          parse_options(options)
          yield self if block_given?
        end

        # @return [VmlShape]
        #: () -> VmlShape
        def vml_shape
          @vml_shape ||= initialize_vml_shape
        end

        # @return [Integer, nil]
        #: () -> Integer?
        def author_index
          @comments.authors.index(author)
        end

        # @param v [String, Cell]
        # @return [void]
        #: (String | Cell v) -> void
        def ref=(v)
          DataTypeValidator.validate(:comment_ref, [String, Cell], v)
          @ref = v if v.is_a?(String)
          @ref = v.r if v.is_a?(Cell)
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          auth = @comments.authors[author_index || 0]
          str << '<comment ref="' << ref.to_s << '" authorId="' << (author_index || 0).to_s << '">'
          str << "<text>"
          unless auth.to_s.empty?
            str << '<r><rPr><b/><color indexed="81"/></rPr>'
            str << "<t>" << ::CGI.escapeHTML(auth.to_s) << ":\n</t></r>"
          end
          str << "<r>"
          str << '<rPr><color indexed="81"/></rPr>'
          str << "<t>" << ::CGI.escapeHTML(text.to_s) << "</t></r></text>"
          str << "</comment>"
        end

        private

        # @return [VmlShape]
        #: () -> VmlShape
        def initialize_vml_shape
          pos = Caxlsx.name_to_indices(ref || "A1")
          VmlShape.new(row: pos[1], column: pos[0], visible: @visible) do |vml|
            vml.left_column = vml.column
            vml.right_column = vml.column + 2
            vml.top_row = vml.row
            vml.bottom_row = vml.row + 4
          end
        end
      end

      # Comments collection for a worksheet.
      class Comments < SimpleTypedList
        # @return [VmlDrawing]
        attr_reader :vml_drawing

        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "you must provide a worksheet" unless worksheet.is_a?(Worksheet)

          super(Comment)
          @worksheet = worksheet
          @vml_drawing = VmlDrawing.new(self)
        end

        # @return [Integer]
        #: () -> Integer
        def index
          @worksheet.index
        end

        # @return [String]
        #: () -> String
        def pn
          format(COMMENT_PN, index + 1)
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Comment]
        #: (Hash[Symbol, untyped] options) ?{ (Comment) -> void } -> Comment
        def add_comment(options = {})
          raise ArgumentError, "Comment require an author" unless options[:author]
          raise ArgumentError, "Comment requires text" unless options[:text]
          raise ArgumentError, "Comment requires ref" unless options[:ref]

          self << Comment.new(self, options)
          yield last if block_given?
          last
        end

        # @return [Array[String]]
        #: () -> Array[String]
        def authors
          map { |comment| comment.author.to_s }.uniq.sort
        end

        # @return [Array[Relationship]]
        #: () -> Array[Relationship]
        def relationships
          [Relationship.new(self, VML_DRAWING_R, "../#{vml_drawing.pn}"),
           Relationship.new(self, COMMENT_R, "../#{pn}")]
        end

        # @param str [String]
        # @return [String]
        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << '<comments xmlns="' << XML_NS << '"><authors>'
          authors.each do |author|
            str << "<author>" << author.to_s << "</author>"
          end
          str << "</authors><commentList>"
          each do |comment|
            comment.to_xml_string(str)
          end
          str << "</commentList></comments>"
        end
      end

      # A wrapper class for comments that defines its worksheet serialization.
      class WorksheetComments
        # @return [Worksheet]
        attr_reader :worksheet

        # @param worksheet [Worksheet]
        #: (Worksheet worksheet) -> void
        def initialize(worksheet)
          raise ArgumentError, "You must provide a worksheet" unless worksheet.is_a?(Worksheet)

          @worksheet = worksheet
        end

        # @return [Comments]
        #: () -> Comments
        def comments
          @comments ||= Comments.new(worksheet)
        end

        # @param options [Hash{Symbol => untyped}]
        # @return [Comment]
        #: (Hash[Symbol, untyped] options) -> Comment
        def add_comment(options = {})
          comments.add_comment(options)
        end

        # @return [Array[Relationship]]
        #: () -> Array[Relationship]
        def relationships
          return [] unless has_comments?

          comments.relationships
        end

        # @return [Boolean]
        #: () -> bool
        def has_comments?
          !comments.empty?
        end

        # @return [String, nil]
        #: () -> String?
        def drawing_rId
          comments.relationships.find { |r| r.Type == VML_DRAWING_R }&.Id
        end

        # @param str [String]
        # @return [String, nil]
        #: (?String str) -> String?
        def to_xml_string(str = +"")
          return unless has_comments?

          str << "<legacyDrawing r:id='#{drawing_rId}' />"
        end
      end
    end
  end
end
