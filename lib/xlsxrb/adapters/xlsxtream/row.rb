# frozen_string_literal: true

# rbs_inline: enabled

require "date"
require "time"
require_relative "xml"

module Xlsxrb
  module Adapters
    module Xlsxtream
      # Serializes a row of cell values into OpenXML markup.
      class Row
        ENCODING = Encoding.find("UTF-8")

        NUMBER_PATTERN = /\A-?[0-9]+(\.[0-9]+)?\z/
        DATE_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}\z/
        TIME_PATTERN = /\A[0-9]{4}-[0-9]{2}-[0-9]{2}T[0-9]{2}:[0-9]{2}(?::[0-9]{2}(?:\.[0-9]{1,9})?)?(?:Z|[+-][0-9]{2}:[0-9]{2})?\z/

        TRUE_STRING = "true"
        FALSE_STRING = "false"

        DATE_STYLE = 1
        TIME_STYLE = 2

        # @param row [Enumerable<Object>] Row values.
        # @param rownum [Integer] Row index (1-based).
        # @param options [Hash{Symbol => untyped}] Row formatting options.
        #: (untyped row, Integer rownum, ?Hash[Symbol, untyped] options) -> void
        def initialize(row, rownum, options = {})
          @row = row
          @rownum = rownum
          @sst = options[:sst]
          @auto_format = options[:auto_format]
        end

        # Generates OpenXML <row r="..."> string.
        #
        # @return [String]
        #: () -> String
        def to_xml
          column = String.new("A")
          xml = %(<row r="#{@rownum}">)

          @row.each do |value|
            cid = "#{column}#{@rownum}"
            column.next!

            value = auto_format(value) if @auto_format && value.is_a?(String)

            case value
            when Numeric
              xml << %(<c r="#{cid}" t="n"><v>#{value}</v></c>)
            when TrueClass, FalseClass
              xml << %(<c r="#{cid}" t="b"><v>#{value ? 1 : 0}</v></c>)
            when Time
              xml << %(<c r="#{cid}" s="#{TIME_STYLE}"><v>#{time_to_oa_date(value)}</v></c>)
            when DateTime
              xml << %(<c r="#{cid}" s="#{TIME_STYLE}"><v>#{datetime_to_oa_date(value)}</v></c>)
            when Date
              xml << %(<c r="#{cid}" s="#{DATE_STYLE}"><v>#{date_to_oa_date(value)}</v></c>)
            else
              value = value.to_s

              unless value.empty?
                value = value.encode(ENCODING) if value.encoding != ENCODING

                xml << if @sst
                         %(<c r="#{cid}" t="s"><v>#{@sst[value]}</v></c>)
                       else
                         %(<c r="#{cid}" t="inlineStr"><is><t>#{XML.escape_value(value)}</t></is></c>)
                       end
              end
            end
          end

          xml << "</row>"
        end

        private

        def auto_format(value)
          case value
          when TRUE_STRING
            true
          when FALSE_STRING
            false
          when NUMBER_PATTERN
            value.include?(".") ? value.to_f : value.to_i
          when DATE_PATTERN
            begin
              Date.parse(value)
            rescue StandardError
              value
            end
          when TIME_PATTERN
            begin
              DateTime.parse(value)
            rescue StandardError
              value
            end
          else
            value
          end
        end

        def time_to_oa_date(time)
          ((time.to_f + time.utc_offset) / 86_400) + 25_569
        end

        def datetime_to_oa_date(date)
          if RUBY_ENGINE == "ruby"
            _, jd, df, sf, of = date.marshal_dump
            (jd - 2_415_019) + ((df + of + (sf / 1e9)) / 86_400)
          else
            (date.jd - 2_415_019) + (((date.hour * 3600) + date.sec + date.sec_fraction.to_f) / 86_400)
          end
        end

        def date_to_oa_date(date)
          (date.jd - 2_415_019).to_f
        end
      end
    end
  end
end
