# frozen_string_literal: true

# Runs the *upstream* test suite of a peer gem (minitest or RSpec) against the
# corresponding xlsxrb adapter by binding the upstream constant to the adapter.
#
# Usage (normally invoked via UpstreamSuite.run from bin/compatibility_runner_*):
#   ruby bin/support/upstream_suite.rb <name>
#
# The upstream checkout is looked up in $XLSXRB_UPSTREAM_DIR (default: ~/src).

require "bundler/setup"
require "open3"
require "rbconfig"

# Per-gem configuration.
#   const:    global constant(s) the upstream tests reference
#   adapter:  adapter module name under Xlsxrb::Adapters
#   features: require paths that must resolve to the adapter (no-op requires)
#   mode:     :replace (bind constant only) or :overlay (load upstream, then
#             override constants one by one; needed when tests use upstream helpers)
#   files:    glob patterns relative to the checkout, or a String path to an
#             allowlist file (one relative test path per line)
#   shims:    names pretended to be already loaded (test-only dependencies that
#             are not installed in the bundle)
module UpstreamSuite
  CONFIG = {
    "simple_xlsx_reader" => { const: %w[SimpleXlsxReader], adapter: "SimpleXlsxReader", features: %w[simple_xlsx_reader],
                              files: ["test/*_test.rb"], exclude: %w[test/performance_test.rb], shims: %w[pry], preload: %w[zip] },
    "xlsxtream" => { const: %w[Xlsxtream], adapter: "Xlsxtream", features: %w[xlsxtream],
                     files: ["test/**/*_test.rb"], shims: [] },
    "xsv" => { const: %w[Xsv], adapter: "Xsv", features: %w[xsv], files: ["test/**/*_test.rb"], shims: [], preload: %w[zip] },
    "fast_excel" => { const: %w[FastExcel], adapter: "FastExcel", features: %w[fast_excel],
                      files: ["test/*_test.rb"], shims: %w[simplecov minitest/reporters] },
    "roo" => { const: %w[Roo], adapter: "Roo", features: %w[roo],
               files: ["test/roo/test_excelx.rb", "test/excelx/**/*.rb"],
               shims: %w[simplecov shoulda minitest/reporters rackup], rackup_stub: true },
    "creek" => { const: %w[Creek], adapter: "Creek", features: %w[creek], files: ["spec/*_spec.rb"],
                 shims: %w[pry], rspec: true, preload: %w[nokogiri] },
    "caxlsx" => { adapter: "Caxlsx", features: %w[axlsx], mode: :overlay, overlay: "Axlsx", require: "axlsx",
                  files: "bin/support/upstream_allowlist/caxlsx.txt",
                  shims: %w[simplecov timecop webmock/minitest ooxml_crypt] },
    "write_xlsx" => { adapter: "WriteXLSX", features: [], mode: :overlay, require: "write_xlsx",
                      overlay: "Writexlsx", also: "WriteXLSX",
                      files: "bin/support/upstream_allowlist/write_xlsx.txt", shims: [] }
  }.freeze

  SUMMARY_MINITEST = /(\d+) runs, (\d+) assertions, (\d+) failures, (\d+) errors, (\d+) skips/
  SUMMARY_RSPEC = /(\d+) examples?, (\d+) failures?/

  module_function

  def upstream_root
    if (dir = ENV["XLSXRB_UPSTREAM_DIR"])
      return File.expand_path(dir)
    end

    candidates = [
      File.expand_path("../../vendor/upstream", __dir__),
      File.expand_path("~/src")
    ]
    candidates.find { |d| Dir.exist?(d) } || candidates.first
  end

  def repo_path(name)
    if (dir = ENV["XLSXRB_UPSTREAM_DIR"])
      return File.join(File.expand_path(dir), name)
    end

    candidates = [
      File.expand_path("../../vendor/upstream/#{name}", __dir__),
      File.expand_path("~/src/#{name}")
    ]
    candidates.find { |d| Dir.exist?(d) } || candidates.first
  end

  # Returns { tests:, assertions:, passed:, failed:, success:, skipped:, note: }
  def run(name)
    cfg = CONFIG.fetch(name)
    repo = repo_path(name)
    unless Dir.exist?(repo)
      if ENV["CI"] || ENV["XLSXRB_REQUIRE_UPSTREAM"]
        return {
          tests: 0,
          assertions: 0,
          passed: 0,
          failed: 1,
          success: false,
          skipped: false,
          note: "#{name} upstream repository not found at #{repo} (required by CI/XLSXRB_REQUIRE_UPSTREAM)"
        }
      end
      return { tests: 0, assertions: 0, passed: 0, failed: 0, success: true, skipped: true, note: "#{repo} not found" }
    end

    out, status = Open3.capture2e(RbConfig.ruby, __FILE__, name)
    parse(out, status, cfg)
  end

  def parse(out, status, cfg)
    if cfg[:rspec] && out =~ SUMMARY_RSPEC
      tests = Regexp.last_match(1).to_i
      failed = Regexp.last_match(2).to_i
      assertions = tests
    elsif out =~ SUMMARY_MINITEST
      tests = Regexp.last_match(1).to_i
      assertions = Regexp.last_match(2).to_i
      failed = Regexp.last_match(3).to_i + Regexp.last_match(4).to_i
    else
      return { tests: 0, assertions: 0, passed: 0, failed: 1, success: false, skipped: false, note: out.lines.last(15).join }
    end
    { tests: tests, assertions: assertions, passed: tests - failed, failed: failed,
      success: status.success? && failed.zero?, skipped: false, note: failed.zero? ? nil : out.lines.last(40).join }
  end

  # ---- child process side -------------------------------------------------

  def gem_lib_dirs(*names)
    names.flat_map do |n|
      spec = Gem.loaded_specs[n] || (begin
        Gem::Specification.find_by_name(n)
      rescue Gem::MissingSpecError
        nil
      end)
      spec ? spec.load_paths : []
    end
  end

  def install_rackup_stub
    require "socket"
    Object.const_set(:WEBrick, Module.new { const_set(:BasicLog, Class.new { def initialize(*); end }) }) unless defined?(::WEBrick)
    handler = Module.new do
      def self.run(app, **opts)
        server = TCPServer.new(opts[:Host] || "0.0.0.0", opts[:Port])
        loop do
          client = server.accept
          client.gets
          while (line = client.gets) && line != "\r\n"; end
          status, headers, body = app.call({})
          payload = body.join
          client.write("HTTP/1.1 #{status} OK\r\n")
          headers.each { |k, v| client.write("#{k}: #{v}\r\n") }
          client.write("Content-Length: #{payload.bytesize}\r\nConnection: close\r\n\r\n#{payload}")
          client.close
        end
      ensure
        server&.close
      end
    end
    Object.const_set(:Rackup, Module.new) unless defined?(::Rackup)
    Rackup.const_set(:Handler, Module.new) unless defined?(::Rackup::Handler)
    Rackup::Handler.const_set(:WEBrick, handler)
  end

  def register_features(_cfg, repo)
    lib = File.join(repo, "lib")
    # Mark every upstream library file as loaded so that `require "<gem>/<part>"` in
    # the tests resolves to the adapter instead of re-opening it with upstream code.
    Dir.glob(File.join(lib, "**/*.rb")).each do |path|
      $LOADED_FEATURES << path
      $LOADED_FEATURES << path.delete_prefix("#{lib}/")
    end
    @registered_features = $LOADED_FEATURES.dup
  end

  def bind_adapter(cfg)
    mod = Xlsxrb::Adapters.const_get(cfg[:adapter])
    if cfg[:mode] == :overlay
      overlay = Object.const_get(cfg[:overlay])
      mod.constants.each do |c|
        overlay.send(:remove_const, c) if overlay.const_defined?(c, false)
        overlay.const_set(c, mod.const_get(c))
      end
      replace_constant(cfg[:also], mod) if cfg[:also]
    else
      cfg[:const].each { |c| replace_constant(c, mod) }
    end
  end

  def replace_constant(name, value)
    Object.send(:remove_const, name) if Object.const_defined?(name, false)
    Object.const_set(name, value)
  end

  def test_files(cfg, repo, project_root)
    files = if ENV["XLSXRB_ALLOWLIST"]
              File.readlines(ENV["XLSXRB_ALLOWLIST"], chomp: true)
                  .reject { |l| l.empty? || l.start_with?("#") }.map { |l| File.join(repo, l) }
            elsif cfg[:files].is_a?(String)
              File.readlines(File.join(project_root, cfg[:files]), chomp: true)
                  .reject { |l| l.empty? || l.start_with?("#") }.map { |l| File.join(repo, l) }
            else
              cfg[:files].flat_map { |g| Dir.glob(File.join(repo, g)) }
            end
    (files - (cfg[:exclude] || []).map { |f| File.join(repo, f) }).sort
  end

  def install_shims(cfg)
    Kernel.define_method(:gem) { |*| true } # upstream helpers call `gem "minitest"`
    cfg[:shims].each { |s| $LOADED_FEATURES << "#{s}.rb" }
    install_rackup_stub if cfg[:rackup_stub]
    if cfg[:shims].include?("simplecov") && !defined?(::SimpleCov)
      Object.const_set(:SimpleCov, Module.new do
        def self.start(*); end

        def self.formatter=(*); end
      end)
    end
    install_reporters_stub if cfg[:shims].include?("minitest/reporters")
    return unless cfg[:shims].include?("timecop")

    Object.const_set(:Timecop, Module.new { def self.freeze(*) = (yield if block_given?) })
  end

  def install_reporters_stub
    reporters = Module.new do
      def self.use!(*); end
      const_set(:DefaultReporter, Class.new { def initialize(*, **); end })
      const_set(:SpecReporter, Class.new { def initialize(*, **); end })
    end
    require "minitest"
    Minitest.const_set(:Reporters, reporters) unless Minitest.const_defined?(:Reporters)
  end

  # Guards against a false pass: if the upstream library itself got loaded, the
  # suite would be testing upstream code instead of the adapter.
  def assert_upstream_not_loaded(cfg, repo)
    return if cfg[:mode] == :overlay

    leaked = ($LOADED_FEATURES - @registered_features).grep(%r{\A#{Regexp.escape(File.join(repo, "lib"))}/})
    abort "upstream library was loaded instead of the adapter: #{leaked.first(3).join(", ")}" unless leaked.empty?
  end

  def child_main(name)
    cfg = CONFIG.fetch(name)
    project_root = File.expand_path("../..", __dir__)
    repo = repo_path(name)

    $LOAD_PATH.unshift(File.join(project_root, "lib"))
    gem_lib_dirs("minitest", "rspec-core", "rspec-expectations", "rspec-support", "rspec-mocks", "rspec", "diff-lcs")
      .each { |d| $LOAD_PATH.unshift(d) unless $LOAD_PATH.include?(d) }
    $LOAD_PATH.unshift(File.join(repo, "test"), File.join(repo, "spec"))
    install_shims(cfg)
    (cfg[:preload] || []).each { |lib| require lib }

    require "xlsxrb/adapters/#{name}"
    if cfg[:mode] == :overlay
      $LOAD_PATH.unshift(File.join(repo, "lib"))
      require cfg[:require]
    else
      register_features(cfg, repo)
    end
    bind_adapter(cfg)

    files = test_files(cfg, repo, project_root)
    abort "no upstream test files found for #{name}" if files.empty?
    Dir.chdir(repo) # keep for the whole process: minitest runs inside at_exit

    if cfg[:rspec]
      require "rspec/core"
      RSpec.configure { |c| c.before(:suite) { UpstreamSuite.assert_upstream_not_loaded(cfg, repo) } }
      exit RSpec::Core::Runner.run(files)
    else
      files.each { |f| require f } # minitest/autorun (from the upstream helper) runs at exit
      assert_upstream_not_loaded(cfg, repo)
    end
  end
end

UpstreamSuite.child_main(ARGV.fetch(0)) if $PROGRAM_NAME == __FILE__
