# frozen_string_literal: true

source "https://rubygems.org"

xlsxrb_path = ENV.fetch("XLSXRB_PATH", "../xlsxrb")

gem "xlsxrb", path: xlsxrb_path

# Runtime dependencies must be specified in xlsxrb-adapters.gemspec.
gemspec

# Development and test dependencies are listed flatly in alphabetical order:
# - No `group`: Gemfile is dev-only; grouping does not affect consumers or CI.
# - No `require: false`: Bundler.require is not used; files are required explicitly.
# - No version constraints: Gemfile.lock pins versions.
gem "irb"
gem "rake"
gem "rbs-inline"
gem "rspec"
gem "rubocop"
gem "rubocop-rake"
gem "rubyXL"
gem "steep"
gem "test-unit"
