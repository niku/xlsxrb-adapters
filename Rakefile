# frozen_string_literal: true

require "bundler/gem_tasks"
require "rake/testtask"
require "rubocop/rake_task"

Rake::TestTask.new(:test) do |t|
  t.libs << "test"
  t.libs << "lib"
  t.test_files = FileList["test/**/*_test.rb"]
end

namespace :test do
  Rake::TestTask.new(:caxlsx) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/xlsxrb/adapters/caxlsx_*test.rb"]
  end

  Rake::TestTask.new(:ruby_xl) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/xlsxrb/adapters/ruby_xl*test.rb"]
  end

  Rake::TestTask.new(:roo) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/xlsxrb/adapters/roo*test.rb"]
  end

  Rake::TestTask.new(:xsv) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/xlsxrb/adapters/xsv*test.rb"]
  end

  Rake::TestTask.new(:creek) do |t|
    t.libs << "test"
    t.libs << "lib"
    t.test_files = FileList["test/xlsxrb/adapters/creek*test.rb"]
  end
end

RuboCop::RakeTask.new

desc "Generate RBS signature files from inline annotations"
task :sig do
  files = FileList["lib/**/*.rb"].to_a
  sh "bundle", "exec", "rbs-inline", "--output=sig/generated", "--base=lib", *files
end

desc "Run static type checking with Steep"
task typecheck: :sig do
  sh "bundle exec steep check"
end

desc "Run all compatibility suites (rubyXL, caxlsx, roo, xsv, and creek)"
task compatibility: ["compatibility:ruby_xl", "compatibility:caxlsx", "compatibility:roo", "compatibility:xsv", "compatibility:creek"]

namespace :compatibility do
  desc "Run official rubyXL RSpec compatibility suite"
  task :ruby_xl do
    sh "bundle exec ruby bin/compatibility_runner ruby_xl"
  end

  desc "Run caxlsx compatibility test suite"
  task :caxlsx do
    sh "bundle exec ruby bin/compatibility_runner caxlsx"
  end

  desc "Run roo compatibility test suite"
  task :roo do
    sh "bundle exec ruby bin/compatibility_runner roo"
  end

  desc "Run xsv compatibility test suite"
  task :xsv do
    sh "bundle exec ruby bin/compatibility_runner xsv"
  end

  desc "Run creek compatibility test suite"
  task :creek do
    sh "bundle exec ruby bin/compatibility_runner creek"
  end
end

desc "Run benchmarks (usage: rake benchmark [rows=10000] [cols=10] [runs=3])"
task :benchmark do
  rows = ENV.fetch("rows", "10000")
  cols = ENV.fetch("cols", "10")
  sh "bundle exec ruby benchmark.rb #{rows} #{cols}"
end

task default: %i[rubocop typecheck test]
