# frozen_string_literal: true

$LOAD_PATH.unshift File.expand_path("../lib", __dir__)

require "test-unit"
require "fileutils"
require "tmpdir"
require "date"
require "time"

# Require xlsxrb adapters
require "xlsxrb-adapters"

# Require official rubyXL for side-by-side comparison tests
require "rubyXL"
require "rubyXL/convenience_methods"

module TestHelper
  def with_tempfile(ext = ".xlsx")
    Dir.mktmpdir do |dir|
      path = File.join(dir, "test#{ext}")
      yield path
    end
  end
end

module Test
  module Unit
    class TestCase
      include TestHelper
    end
  end
end
