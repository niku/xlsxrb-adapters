# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersVersionTest < Test::Unit::TestCase
  def test_version
    refute_nil ::Xlsxrb::Adapters::VERSION
  end
end
