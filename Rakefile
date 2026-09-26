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

task default: %i[rubocop typecheck test]
