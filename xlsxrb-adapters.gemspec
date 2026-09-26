# frozen_string_literal: true

require_relative "lib/xlsxrb/adapters/version"

Gem::Specification.new do |spec|
  spec.name = "xlsxrb-adapters"
  spec.version = Xlsxrb::Adapters::VERSION
  spec.authors = ["niku"]
  spec.email = ["10890+niku@users.noreply.github.com"]

  spec.summary = "Adapters for drop-in migration to xlsxrb from other XLSX libraries."
  spec.description = "Provides compatibility adapters (such as RubyXL) to assist in gradual migration to xlsxrb and interoperability testing."
  spec.homepage = "https://github.com/niku/xlsxrb-adapters"
  spec.license = "MIT"
  spec.required_ruby_version = ">= 4.0.0"

  spec.metadata["allowed_push_host"] = "https://rubygems.org"
  spec.metadata["source_code_uri"] = "https://github.com/niku/xlsxrb-adapters"
  spec.metadata["rubygems_mfa_required"] = "true"

  gemspec = File.basename(__FILE__)
  spec.files = IO.popen(%w[git ls-files -z], chdir: __dir__, err: IO::NULL) do |ls|
    ls.readlines("\x0", chomp: true).reject do |f|
      (f == gemspec) ||
        f.start_with?(*%w[
                        bin/ Gemfile .gitignore test/ .github/ .devcontainer/ Rakefile
                      ])
    end
  end
  spec.require_paths = ["lib"]

  spec.add_dependency "xlsxrb"
end
