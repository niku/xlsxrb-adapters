# frozen_string_literal: true

# rbs_inline: enabled

require "time"
require_relative "util"

module Xlsxrb
  module Adapters
    module Caxlsx
      # Core properties for the package
      class Core
        include OptionsParser

        attr_accessor :creator #: String?
        attr_accessor :created #: Time?
        attr_accessor :title #: String?
        attr_accessor :subject #: String?
        attr_accessor :description #: String?
        attr_accessor :category #: String?
        attr_accessor :keywords #: String?
        attr_accessor :last_modified_by #: String?
        attr_accessor :modified #: Time?

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @creator = "caxlsx"
          @created = Time.now
          @last_modified_by = @creator
          @modified = @created
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          time_str = @created&.utc&.strftime("%Y-%m-%dT%H:%M:%SZ") || Time.now.utc.strftime("%Y-%m-%dT%H:%M:%SZ")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << "<cp:coreProperties xmlns:cp=\"#{CORE_NS}\" xmlns:dc=\"#{CORE_NS_DC}\" xmlns:dcterms=\"#{CORE_NS_DCT}\" xmlns:dcmitype=\"#{CORE_NS_DCMIT}\" xmlns:xsi=\"#{CORE_NS_XSI}\">"
          str << "<dc:creator>#{Caxlsx.coder.encode(@creator)}</dc:creator>" if @creator
          str << "<cp:lastModifiedBy>#{Caxlsx.coder.encode(@last_modified_by)}</cp:lastModifiedBy>" if @last_modified_by
          str << "<dcterms:created xsi:type=\"dcterms:W3CDTF\">#{time_str}</dcterms:created>"
          str << "<dcterms:modified xsi:type=\"dcterms:W3CDTF\">#{time_str}</dcterms:modified>"
          str << "<dc:title>#{Caxlsx.coder.encode(@title)}</dc:title>" if @title
          str << "<dc:subject>#{Caxlsx.coder.encode(@subject)}</dc:subject>" if @subject
          str << "<dc:description>#{Caxlsx.coder.encode(@description)}</dc:description>" if @description
          str << "</cp:coreProperties>"
          str
        end
      end

      # App extended properties for the package
      class App
        include OptionsParser

        attr_accessor :company #: String?
        attr_accessor :manager #: String?
        attr_accessor :application #: String
        attr_accessor :app_version #: String
        attr_accessor :doc_security #: Integer?
        attr_accessor :scale_crop #: bool?
        attr_accessor :links_up_to_date #: bool?
        attr_accessor :shared_doc #: bool?
        attr_accessor :hyperlinks_changed #: bool?
        attr_accessor :hyperlink_base #: String?

        alias Company company
        alias Company= company=
        alias Application application
        alias AppVersion app_version
        alias AppVersion= app_version=
        alias Manager manager
        alias Manager= manager=
        alias DocSecurity doc_security
        alias DocSecurity= doc_security=
        alias ScaleCrop scale_crop
        alias ScaleCrop= scale_crop=
        alias LinksUpToDate links_up_to_date
        alias LinksUpToDate= links_up_to_date=
        alias SharedDoc shared_doc
        alias SharedDoc= shared_doc=
        alias HyperlinksChanged hyperlinks_changed
        alias HyperlinksChanged= hyperlinks_changed=
        alias HyperLinksChanged= hyperlinks_changed=
        alias HyperlinkBase hyperlink_base
        alias HyperlinkBase= hyperlink_base=

        # Dummy accessors for caxlsx compatibility
        attr_accessor :lines, :words, :characters, :characters_with_spaces, :pages, :paragraphs, :slides, :notes, :hidden_slides, :m_m_clips, :presentation_format, :template, :total_time
        alias Lines lines
        alias Lines= lines=
        alias Words words
        alias Words= words=
        alias Characters characters
        alias Characters= characters=
        alias CharactersWithSpaces characters_with_spaces
        alias CharactersWithSpaces= characters_with_spaces=
        alias Pages pages
        alias Pages= pages=
        alias Paragraphs paragraphs
        alias Paragraphs= paragraphs=
        alias Slides slides
        alias Slides= slides=
        alias Notes notes
        alias Notes= notes=
        alias HiddenSlides hidden_slides
        alias HiddenSlides= hidden_slides=
        alias MMClips m_m_clips
        alias MMClips= m_m_clips=
        alias PresentationFormat presentation_format
        alias PresentationFormat= presentation_format=
        alias Template template
        alias Template= template=
        alias TotalTime total_time
        alias TotalTime= total_time=

        # @param options [Hash]
        #: (?Hash[Symbol, untyped] options) -> void
        def initialize(options = {})
          @application = "caxlsx"
          @app_version = "1.0"
          @company = ""
          parse_options options
        end

        #: (?String str) -> String
        def to_xml_string(str = +"")
          str << '<?xml version="1.0" encoding="UTF-8"?>'
          str << "<Properties xmlns=\"#{APP_NS}\" xmlns:vt=\"#{APP_NS_VT}\">"
          str << "<Application>#{@application}</Application>"
          str << "<AppVersion>#{@app_version}</AppVersion>"
          str << "<Company>#{Caxlsx.coder.encode(@company)}</Company>" if @company && !@company.empty?
          str << "<Manager>#{Caxlsx.coder.encode(@manager)}</Manager>" if @manager
          str << "</Properties>"
          str
        end
      end
    end
  end
end
