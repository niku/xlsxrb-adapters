# frozen_string_literal: true

require "bundler/gem_tasks"
require "rake/testtask"
require "rubocop/rake_task"

Rake::TestTask.new(:test) do |t|
  t.libs << "test"
  t.libs << "lib"
  t.test_files = FileList["test/**/*_test.rb"]
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

desc "Run official rubyXL RSpec compatibility suite"
task :compatibility do
  sh "bundle exec ruby bin/compatibility_runner"
end

desc "Run benchmarks (usage: rake benchmark [rows=10000] [cols=10] [runs=3])"
task :benchmark do
  rows = ENV.fetch("rows", "10000")
  cols = ENV.fetch("cols", "10")
  sh "bundle exec ruby benchmark.rb #{rows} #{cols}"
end

task default: %i[rubocop typecheck test]
