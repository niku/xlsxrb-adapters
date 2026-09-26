# frozen_string_literal: true
# rubocop:disable all

require "json"
require "open3"
require "fileutils"
require "optparse"

ROWS = (ARGV[0] || "10000").to_i
COLS = (ARGV[1] || "10").to_i
RUNS = (ENV["RUNS"] || "3").to_i

puts "=" * 80
puts "Benchmarking Excel Adapters (#{ROWS} rows x #{COLS} cols = #{ROWS * COLS} cells)"
puts "Runs per benchmark: #{RUNS} (Median reported, Mean calculated)"
puts "Ruby: #{RUBY_DESCRIPTION}"
puts "=" * 80

RUNNER_SCRIPT = <<~'RUBY'
  require "json"
  require "stringio"

  def measure
    GC.start
    GC.compact if GC.respond_to?(:compact)
    gc_before = GC.stat[:count]
    t0 = Process.clock_gettime(Process::CLOCK_MONOTONIC)

    yield

    t1 = Process.clock_gettime(Process::CLOCK_MONOTONIC)
    gc_after = GC.stat[:count]

    peak_mb = 0.0
    if File.exist?("/proc/self/status")
      status = File.read("/proc/self/status")
      if status =~ /VmHWM:\s+(\d+)\s+kB/i
        peak_mb = $1.to_f / 1024.0
      elsif status =~ /VmRSS:\s+(\d+)\s+kB/i
        peak_mb = $1.to_f / 1024.0
      end
    end

    {
      time: (t1 - t0),
      peak_memory_mb: peak_mb,
      gc_count: (gc_after - gc_before)
    }
  end

  def generate_row(r, cols)
    base = [r + 1, "User #{r + 1}", 123.45, true, "Active", (r + 1) * 10, "Tokyo", 99.9, false, "Item #{r % 50}"]
    if cols <= base.size
      base.first(cols)
    else
      base + Array.new(cols - base.size) { |c| "col_#{c}_#{r}" }
    end
  end

  lib = ARGV[0]
  mode = ARGV[1]
  rows = ARGV[2].to_i
  cols = ARGV[3].to_i
  filename = ARGV[4]

  result = case [lib, mode]
  when ["rubyXL", "write"]
    require "rubyXL"
    measure do
      wb = RubyXL::Workbook.new
      sheet = wb[0]
      sheet.sheet_name = "Data"
      rows.times do |r|
        row_data = generate_row(r, cols)
        row_data.each_with_index do |val, c|
          sheet.add_cell(r, c, val)
        end
      end
      wb.write(filename)
    end
  when ["rubyXL", "read"]
    require "rubyXL"
    measure do
      wb = RubyXL::Parser.parse(filename)
      count = 0
      wb.worksheets.each do |sheet|
        sheet.each do |row|
          next unless row
          row.cells.each do |cell|
            _val = cell&.value
            count += 1
          end
        end
      end
    end
  when ["xlsxrb_adapters", "write"]
    require "xlsxrb"
    require "xlsxrb/adapters/ruby_xl"
    measure do
      wb = Xlsxrb::Adapters::RubyXL::Workbook.new
      sheet = wb[0]
      sheet.sheet_name = "Data"
      rows.times do |r|
        row_data = generate_row(r, cols)
        row_data.each_with_index do |val, c|
          sheet.add_cell(r, c, val)
        end
      end
      wb.write(filename)
    end
  when ["xlsxrb_adapters", "read"]
    require "xlsxrb"
    require "xlsxrb/adapters/ruby_xl"
    measure do
      wb = Xlsxrb::Adapters::RubyXL::Parser.parse(filename)
      count = 0
      wb.worksheets.each do |sheet|
        sheet.each do |row|
          next unless row
          row.cells.each do |cell|
            _val = cell&.value
            count += 1
          end
        end
      end
    end
  when ["xlsxrb_stream", "write"]
    require "xlsxrb"
    measure do
      Xlsxrb.write(filename) do |wb|
        wb.sheet("Data") do |sheet|
          rows.times do |r|
            sheet.row(generate_row(r, cols))
          end
        end
      end
    end
  when ["xlsxrb_stream", "read"]
    require "xlsxrb"
    measure do
      count = 0
      Xlsxrb.read(filename) do |sheet|
        sheet.each do |row|
          row.cells.each do |cell|
            _val = cell.value
            count += 1
          end
        end
      end
    end
  when ["xlsxrb_inmemory", "write"]
    require "xlsxrb"
    measure do
      wb = Xlsxrb.build do |b|
        b.sheet("Data") do |s|
          rows.times do |r|
            s.row(generate_row(r, cols))
          end
        end
      end
      Xlsxrb.write(filename, wb)
    end
  when ["xlsxrb_inmemory", "read"]
    require "xlsxrb"
    measure do
      wb = Xlsxrb.read(filename).load
      count = 0
      wb.sheets.each do |sheet|
        sheet.rows.each do |row|
          row.cells.each do |cell|
            _val = cell.value
            count += 1
          end
        end
      end
    end
  else
    raise "Unknown benchmark target: #{lib} #{mode}"
  end

  puts result.to_json
RUBY

runner_file = "tmp/benchmark_runner.rb"
FileUtils.mkdir_p("tmp")
File.write(runner_file, RUNNER_SCRIPT)

def run_isolated(lib, mode, rows, cols, filename)
  cmd = ["bundle", "exec", "ruby", "-Ilib", "tmp/benchmark_runner.rb", lib, mode, rows.to_s, cols.to_s, filename]
  stdout, stderr, status = Open3.capture3(*cmd)
  unless status.success?
    warn "Failed to run #{lib} #{mode}: #{stderr}"
    return nil
  end
  JSON.parse(stdout.strip, symbolize_names: true)
end

def run_benchmark_series(name, lib, mode, rows, cols, filename, runs)
  print "Running #{name} (#{runs} runs)... "
  $stdout.flush
  results = []
  runs.times do
    File.delete(filename) if File.exist?(filename) && mode == "write"
    res = run_isolated(lib, mode, rows, cols, filename)
    if res
      results << res
      print "#{res[:time].round(3)}s "
      $stdout.flush
    else
      print "ERR "
      $stdout.flush
    end
  end
  puts
  return nil if results.empty?

  times = results.map { |r| r[:time] }.sort
  mems = results.map { |r| r[:peak_memory_mb] }.sort
  gcs = results.map { |r| r[:gc_count] }.sort

  {
    name: name,
    median_time: times[times.size / 2],
    mean_time: times.sum / times.size,
    median_mem: mems[mems.size / 2],
    median_gc: gcs[gcs.size / 2]
  }
end

# 1. Generate reference file
ref_file = "tmp/bench_reference_data.xlsx"
puts "\n[Setup] Generating reference file (#{ROWS} x #{COLS}) for read benchmarks..."
run_isolated("xlsxrb_stream", "write", ROWS, COLS, ref_file)

# 2. Benchmark Write
puts "\n=== Benchmarking Write Performance ==="
write_targets = [
  ["xlsxrb (Streaming)", "xlsxrb_stream"],
  ["xlsxrb (In-Memory)", "xlsxrb_inmemory"],
  ["xlsxrb-adapters (RubyXL)", "xlsxrb_adapters"],
  ["rubyXL (Original)", "rubyXL"]
]

write_results = []
write_targets.each do |name, lib|
  target_file = "tmp/bench_write_#{lib}.xlsx"
  res = run_benchmark_series(name, lib, "write", ROWS, COLS, target_file, RUNS)
  write_results << res if res
  FileUtils.rm_f(target_file)
end

# 3. Benchmark Read
puts "\n=== Benchmarking Read Performance ==="
read_targets = [
  ["xlsxrb (Streaming)", "xlsxrb_stream"],
  ["xlsxrb (In-Memory)", "xlsxrb_inmemory"],
  ["xlsxrb-adapters (RubyXL)", "xlsxrb_adapters"],
  ["rubyXL (Original)", "rubyXL"]
]

read_results = []
read_targets.each do |name, lib|
  res = run_benchmark_series(name, lib, "read", ROWS, COLS, ref_file, RUNS)
  read_results << res if res
end

FileUtils.rm_f(ref_file)
FileUtils.rm_f(runner_file)

# Output Summary Tables
puts "\n" + ("=" * 80)
puts "### Write Performance (#{ROWS * COLS} cells: #{ROWS} rows x #{COLS} cols)"
puts ""
puts "| Library                  | Time (Median) | Time (Mean) | Peak Memory | GC Count |"
puts "| :----------------------- | :------------ | :---------- | :---------- | :------- |"
write_results.each do |r|
  printf "| %-24s | %6.3f s      | %6.3f s    | %7.1f MB  | %6.1f   |\n",
         r[:name], r[:median_time], r[:mean_time], r[:median_mem], r[:median_gc]
end

puts "\n### Read Performance (#{ROWS * COLS} cells: #{ROWS} rows x #{COLS} cols)"
puts ""
puts "| Library                  | Time (Median) | Time (Mean) | Peak Memory | GC Count |"
puts "| :----------------------- | :------------ | :---------- | :---------- | :------- |"
read_results.each do |r|
  printf "| %-24s | %6.3f s      | %6.3f s    | %7.1f MB  | %6.1f   |\n",
         r[:name], r[:median_time], r[:mean_time], r[:median_mem], r[:median_gc]
end
puts "=" * 80
