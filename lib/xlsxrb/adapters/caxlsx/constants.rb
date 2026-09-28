# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # Version of the caxlsx API targeted by this adapter
      VERSION = "4.5.0"

      # Error messages
      ERR_SHEET_NAME_EMPTY = "Your worksheet name cannot be empty"
      ERR_SHEET_NAME_TOO_LONG = "Your worksheet name '%s' is too long. Worksheet names must be %d characters or less"
      ERR_SHEET_NAME_CHARACTER_FORBIDDEN = "Your worksheet name '%s' contains a character that is forbidden. Worksheet names cannot contain %s"
      ERR_DUPLICATE_SHEET_NAME = "There is already a worksheet in this workbook named '%s'"
      ERR_CELL_REFERENCE_INVALID = "Invalid cell reference: %s. Cell references must be in the format 'A1' or 'A1:B2'"
      ERR_CELL_REFERENCE_MISSING_CELL = "Could not find cell: %s in range: %s"
      ERR_INTEGERISH = "Expected an integer-like value, got %s"
      ERR_RANGE = "Value %s is out of range (%s)"
      ERR_TYPE = "Invalid type: %s (expected %s)"
      ERR_ANGLE = "Angle must be between -90 and 90, got %s"
      ERR_RESTRICTION = "Value %s is not in allowed list: %s"
      ERR_REGEX = "Value %s does not match pattern: %s"
      ERR_INVALID_BORDER_ID = "Invalid border id: %s"
      ERR_INVALID_BORDER_OPTIONS = "Invalid border options: %s"

      # XML Namespaces
      XML_NS = "http://schemas.openxmlformats.org/spreadsheetml/2006/main"
      XML_NS_A = "http://schemas.openxmlformats.org/drawingml/2006/main"
      XML_NS_C = "http://schemas.openxmlformats.org/drawingml/2006/chart"
      XML_NS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships"
      XML_NS_T = "http://schemas.openxmlformats.org/package/2006/content-types"
      XML_NS_XDR = "http://schemas.openxmlformats.org/drawingml/2006/spreadsheetDrawing"

      SCHEMA_BASE = "http://schemas.openxmlformats.org"
      SML_XSD = "#{SCHEMA_BASE}/spreadsheetml/2006/main".freeze
      RELS_XSD = "#{SCHEMA_BASE}/package/2006/relationships".freeze
      DRAWING_XSD = "#{SCHEMA_BASE}/drawingml/2006/main".freeze
      APP_XSD = "#{SCHEMA_BASE}/officeDocument/2006/extended-properties".freeze
      CORE_XSD = "#{SCHEMA_BASE}/package/2006/metadata/core-properties".freeze
      CONTENT_TYPES_XSD = "#{SCHEMA_BASE}/package/2006/content-types".freeze

      # Core Properties Namespaces
      CORE_NS = "http://schemas.openxmlformats.org/package/2006/metadata/core-properties"
      CORE_NS_DC = "http://purl.org/dc/elements/1.1/"
      CORE_NS_DCMIT = "http://purl.org/dc/dcmitype/"
      CORE_NS_DCT = "http://purl.org/dc/terms/"
      CORE_NS_XSI = "http://www.w3.org/2001/XMLSchema-instance"
      APP_NS = "http://schemas.openxmlformats.org/officeDocument/2006/extended-properties"
      APP_NS_VT = "http://schemas.openxmlformats.org/officeDocument/2006/docPropsVTypes"

      # Content Types & Part Names
      WORKBOOK_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml"
      WORKBOOK_PN = "xl/workbook.xml"
      WORKBOOK_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument"
      WORKBOOK_RELS_PN = "xl/_rels/workbook.xml.rels"

      WORKSHEET_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml"
      WORKSHEET_MAX_NAME_LENGTH = 31
      WORKSHEET_NAME_FORBIDDEN_CHARS = ["[", "]", "*", "/", "\\", "?", ":"].freeze
      WORKSHEET_PN = "worksheets/sheet%d.xml"
      WORKSHEET_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet"
      WORKSHEET_RELS_PN = "worksheets/_rels/sheet%d.xml.rels"

      STYLES_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml"
      STYLES_PN = "styles.xml"
      STYLES_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles"

      THEME_CT = "application/vnd.openxmlformats-officedocument.theme+xml"
      THEME_PN = "theme/theme1.xml"
      THEME_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme"
      THEME_XSD = ""

      SHARED_STRINGS_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml"
      SHARED_STRINGS_PN = "sharedStrings.xml"
      SHARED_STRINGS_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings"

      CORE_CT = "application/vnd.openxmlformats-package.core-properties+xml"
      CORE_PN = "docProps/core.xml"
      CORE_R = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties"

      APP_CT = "application/vnd.openxmlformats-officedocument.extended-properties+xml"
      APP_PN = "docProps/app.xml"
      APP_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties"

      DRAWING_CT = "application/vnd.openxmlformats-officedocument.drawing+xml"
      DRAWING_PN = "drawings/drawing%d.xml"
      DRAWING_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/drawing"
      DRAWING_RELS_PN = "drawings/_rels/drawing%d.xml.rels"

      CHART_CT = "application/vnd.openxmlformats-officedocument.drawingml.chart+xml"
      CHART_PN = "charts/chart%d.xml"
      CHART_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/chart"

      COMMENT_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.comments+xml"
      COMMENT_PN = "comments%d.xml"
      COMMENT_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/comments"
      COMMENT_R_NULL = "http://purl.oclc.org/ooxml/officeDocument/relationships/comments"

      VML_DRAWING_CT = "application/vnd.openxmlformats-officedocument.vmlDrawing"
      VML_DRAWING_PN = "drawings/vmlDrawing%d.vml"
      VML_DRAWING_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/vmlDrawing"

      TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.table+xml"
      TABLE_PN = "tables/table%d.xml"
      TABLE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/table"

      PIVOT_TABLE_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotTable+xml"
      PIVOT_TABLE_PN = "pivotTables/pivotTable%d.xml"
      PIVOT_TABLE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotTable"
      PIVOT_TABLE_RELS_PN = "pivotTables/_rels/pivotTable%d.xml.rels"
      PIVOT_TABLE_CACHE_DEFINITION_CT = "application/vnd.openxmlformats-officedocument.spreadsheetml.pivotCacheDefinition+xml"
      PIVOT_TABLE_CACHE_DEFINITION_PN = "pivotCache/pivotCacheDefinition%d.xml"
      PIVOT_TABLE_CACHE_DEFINITION_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/pivotCacheDefinition"

      CONTENT_TYPES_PN = "[Content_Types].xml"
      RELS_CT = "application/vnd.openxmlformats-package.relationships+xml"
      RELS_EX = "rels"
      RELS_PN = "_rels/.rels"
      RELS_R = "http://schemas.openxmlformats.org/package/2006/relationships"

      XML_CT = "application/xml"
      XML_EX = "xml"

      IMAGE_PN = "media/image%d.%s"
      IMAGE_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"
      HYPERLINK_R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink"

      JPEG_CT = "image/jpeg"
      JPEG_EXS = %w[jpeg jpg].freeze
      GIF_CT = "image/gif"
      GIF_EX = "gif"
      PNG_CT = "image/png"
      PNG_EX = "png"

      DIGITAL_SIGNATURE_XML_CT = "application/vnd.openxmlformats-package.digital-signature-xmlsignature+xml"
      DIGITAL_SIGNATURE_ORIGIN_CT = "application/vnd.openxmlformats-package.digital-signature-origin+xml"
      DIGITAL_SIGNATURE_CERTIFICATE_CT = "application/vnd.openxmlformats-package.digital-signature-certificate"
      DIGITAL_SIGNATURE_R = "http://schemas.openxmlformats.org/package/2006/relationships/digital-signature/origin"
      DIGITAL_SIGNATURE_NS = "http://schemas.openxmlformats.org/package/2006/digital-signature"

      CONTROL_CHARS = '\x00-\x08\x0B\x0C\x0E-\x1F\x7F'
      ENCODING = "UTF-8"

      # Formula prefixes
      FORMULA_PREFIX = "="
      ARRAY_FORMULA_PREFIX = "{="
      ARRAY_FORMULA_SUFFIX = "}"
      SECONDARY_FORMULA_PREFIXES = ["+", "-", "@"].freeze

      # Style constants
      STYLE_THIN_BORDER = 1
      STYLE_DATE = 2
      NUM_FMT_PERCENT = 9
      NUM_FMT_YYYYMMDD = 100
      NUM_FMT_YYYYMMDDHHMMSS = 101

      # Valid collections
      VALID_BOOLEAN_CLASSES = [TrueClass, FalseClass, Integer, String, Symbol].freeze
      VALID_BOOLEAN_VALUES = [true, false, 1, 0, "1", "0", "true", "false", true, false].freeze
      VALID_CELL_U_VALUES = %i[none single double singleAccounting doubleAccounting].freeze
      VALID_HORIZONTAL_ALIGNMENT_VALUES = %i[general left center right fill justify centerContinuous distributed].freeze
      VALID_VERTICAL_ALIGNMENT_VALUES = %i[top center bottom justify distributed].freeze
      VALID_PAGE_ORIENTATION_VALUES = %i[default landscape portrait].freeze
      VALID_PANE_TYPE_VALUES = %i[bottom_left bottom_right top_left top_right].freeze
      VALID_SPLIT_STATE_TYPE_VALUES = %i[frozen frozen_split split].freeze
      VALID_VIEW_VISIBILITY_VALUES = %i[visible hidden very_hidden].freeze
      VALID_SHEET_VIEW_TYPE_VALUES = %i[normal page_break_preview page_layout].freeze
      VALID_PATTERN_TYPE_VALUES = %i[none solid mediumGray darkGray lightGray darkHorizontal darkVertical darkDown darkUp darkGrid darkTrellis lightHorizontal lightVertical lightDown lightUp lightGrid lightTrellis gray125 gray0625].freeze
      VALID_GRADIENT_TYPE_VALUES = %i[linear path].freeze
      VALID_FAMILY_VALUES = (1..5)
      VALID_DATA_VALIDATION_TYPE_VALUES = %i[custom data decimal list none textLength date time whole].freeze
      VALID_DATA_VALIDATION_OPERATOR_VALUES = %i[lessThan lessThanOrEqual equal notEqual greaterThanOrEqual greaterThan between notBetween].freeze
      VALID_DATA_VALIDATION_ERROR_STYLE_VALUES = %i[information stop warning].freeze
      VALID_CONDITIONAL_FORMATTING_TYPE_VALUES = %i[expression cellIs colorScale dataBar iconSet top10 uniqueValues duplicateValues containsText notContainsText beginsWith endsWith containsBlanks notContainsBlanks containsErrors notContainsErrors timePeriod aboveAverage].freeze
      VALID_CONDITIONAL_FORMATTING_OPERATOR_VALUES = %i[lessThan lessThanOrEqual equal notEqual greaterThanOrEqual greaterThan between notBetween containsText notContains beginsWith endsWith].freeze
      VALID_CONDITION_FORMATTING_VALUE_OBJECT_TYPE_VALUES = %i[num percent max min formula percentile].freeze
      VALID_ICON_SET_VALUES = %w[3Arrows 3ArrowsGray 3Flags 3TrafficLights1 3TrafficLights2 3Signs 3Symbols 3Symbols2 4Arrows 4ArrowsGray 4RedToBlack 4Rating 4TrafficLights 5Arrows 5ArrowsGray 5Rating 5Quarters].freeze
      VALID_TIME_PERIOD_TYPE_VALUES = %i[today yesterday tomorrow last7Days thisMonth lastMonth nextMonth thisWeek lastWeek nextWeek].freeze
      VALID_SCATTER_STYLE_VALUES = %i[none line lineMarker marker smooth smoothMarker].freeze
      VALID_MARKER_SYMBOL_VALUES = %i[default circle dash diamond dot picture plus square star triangle x].freeze
      VALID_DISPLAY_BLANK_AS_VALUES = %i[gap span zero].freeze
      VALID_TABLE_ELEMENT_TYPE_VALUES = %i[wholeTable headerRow totalRow firstColumn lastColumn firstRowStripe secondRowStripe firstColumnStripe secondColumnStripe firstHeaderCell lastHeaderCell firstTotalCell lastTotalCell firstSubtotalColumn secondSubtotalColumn thirdSubtotalColumn firstSubtotalRow secondSubtotalRow thirdSubtotalRow blankRow firstColumnSubheading secondColumnSubheading thirdColumnSubheading firstRowSubheading secondRowSubheading thirdRowSubheading pageFieldLabels pageFieldValues].freeze

      # Regular Expressions
      NUMERIC_REGEX = /\A-?[0-9]+\z/
      SAFE_FLOAT_REGEX = /\A-?[0-9]*\.[0-9]+\z/
      MAYBE_FLOAT_REGEX = /\A-?[0-9]+(?:\.[0-9]+)?e(?<exp>[+-]?[0-9]+)\z/i
      ISO_8601_REGEX = /\A\d{4}-\d{2}-\d{2}T\d{2}:\d{2}:\d{2}(?:\.\d+)?(?:Z|[+-]\d{2}:\d{2})?\z/
    end
  end
end
