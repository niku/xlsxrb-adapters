# frozen_string_literal: true
# rubocop:disable all

require "json"
require "open3"
require "fileutils"
require "optparse"

ROWS = (ARGV[0] || "10000").to_i
COLS = (ARGV[1] || "10").to_i
CATEGORY = (ARGV[2] || "all").downcase
RUNS = (ENV["RUNS"] || "3").to_i

puts "=" * 80
puts "Benchmarking Excel Adapters (#{ROWS} rows x #{COLS} cols = #{ROWS * COLS} cells) [Category: #{CATEGORY}]"
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
  when ["caxlsx", "write"]
    require "caxlsx"
    measure do
      p = Axlsx::Package.new
      wb = p.workbook
      wb.add_worksheet(name: "Data") do |sheet|
        rows.times do |r|
          sheet.add_row(generate_row(r, cols))
        end
      end
      p.serialize(filename)
    end
  when ["xlsxrb_adapters_caxlsx", "write"]
    require "xlsxrb"
    require "xlsxrb/adapters/caxlsx"
    measure do
      p = Xlsxrb::Adapters::Caxlsx::Package.new
      wb = p.workbook
      wb.add_worksheet(name: "Data") do |sheet|
        rows.times do |r|
          sheet.add_row(generate_row(r, cols))
        end
      end
      p.serialize(filename)
    end
  when ["xlsxrb_adapters_caxlsx_streaming", "write"]
    require "xlsxrb"
    require "xlsxrb/adapters/caxlsx"
    measure do
      Xlsxrb::Adapters::Caxlsx::StreamingPackage.open(filename) do |p|
        wb = p.workbook
        wb.add_worksheet(name: "Data") do |sheet|
          rows.times do |r|
            sheet.add_row(generate_row(r, cols))
          end
        end
      end
    end
  when ["write_xlsx", "write"]
    require "write_xlsx"
    measure do
      wb = WriteXLSX.new(filename)
      ws = wb.add_worksheet("Data")
      rows.times do |r|
        ws.write_row(r, 0, generate_row(r, cols))
      end
      wb.close
    end
  when ["xlsxrb_adapters_write_xlsx", "write"]
    require "xlsxrb"
    require "xlsxrb/adapters/write_xlsx"
    measure do
      wb = Xlsxrb::Adapters::WriteXLSX::Workbook.new(filename)
      ws = wb.add_worksheet("Data")
      rows.times do |r|
        ws.write_row(r, 0, generate_row(r, cols))
      end
      wb.close
    end
  when ["fast_excel", "write"]
    require "fast_excel"
    measure do
      wb = FastExcel.open(filename)
      ws = wb.add_worksheet("Data")
      rows.times do |r|
        ws.append_row(generate_row(r, cols))
      end
      wb.close
    end
  when ["xlsxrb_adapters_fast_excel", "write"]
    require "xlsxrb"
    require "xlsxrb/adapters/fast_excel"
    measure do
      wb = Xlsxrb::Adapters::FastExcel.open(filename)
      ws = wb.add_worksheet("Data")
      rows.times do |r|
        ws.append_row(generate_row(r, cols))
      end
      wb.close
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
  when ["roo", "read"]
    require "roo"
    measure do
      xlsx = Roo::Excelx.new(filename)
      count = 0
      xlsx.each_row_streaming do |row|
        row.each do |cell|
          _val = cell&.value
          count += 1
        end
      end
    end
  when ["xlsxrb_adapters_roo", "read"]
    require "xlsxrb"
    require "xlsxrb/adapters/roo"
    measure do
      xlsx = Xlsxrb::Adapters::Roo::Excelx.new(filename)
      count = 0
      xlsx.each_row_streaming do |row|
        row.each do |cell|
          _val = cell&.value
          count += 1
        end
      end
    end
  when ["xsv", "read"]
    require "xsv"
    measure do
      x = Xsv.open(filename)
      count = 0
      x.sheets.each do |sheet|
        sheet.each do |row|
          row.each do |cell|
            _val = cell
            count += 1
          end
        end
      end
    end
  when ["xlsxrb_adapters_xsv", "read"]
    require "xlsxrb"
    require "xlsxrb/adapters/xsv"
    measure do
      x = Xlsxrb::Adapters::Xsv.open(filename)
      count = 0
      x.sheets.each do |sheet|
        sheet.each do |row|
          row.each do |cell|
            _val = cell
            count += 1
          end
        end
      end
    end
  when ["creek", "read"]
    require "creek"
    measure do
      c = Creek::Book.new(filename)
      count = 0
      c.sheets.each do |sheet|
        sheet.rows.each do |row|
          row.each_value do |cell|
            _val = cell
            count += 1
          end
        end
      end
      c.close
    end
  when ["xlsxrb_adapters_creek", "read"]
    require "xlsxrb"
    require "xlsxrb/adapters/creek"
    measure do
      c = Xlsxrb::Adapters::Creek::Book.new(filename)
      count = 0
      c.sheets.each do |sheet|
        sheet.rows.each do |row|
          row.each_value do |cell|
            _val = cell
            count += 1
          end
        end
      end
      c.close
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

# 1. Benchmark Write
write_results = {}
if %w[all caxlsx rubyxl write_xlsx fast_excel].include?(CATEGORY)
  puts "\n=== Benchmarking Write Performance ==="
  write_targets = []
  write_targets << ["xlsxrb (Streaming)", "xlsxrb_stream"]
  write_targets << ["xlsxrb (In-Memory)", "xlsxrb_inmemory"]
  if %w[all caxlsx].include?(CATEGORY)
    write_targets << ["xlsxrb-adapters (Caxlsx Streaming)", "xlsxrb_adapters_caxlsx_streaming"]
    write_targets << ["xlsxrb-adapters (Caxlsx)", "xlsxrb_adapters_caxlsx"]
    write_targets << ["caxlsx (Original)", "caxlsx"]
  end
  if %w[all rubyxl].include?(CATEGORY)
    write_targets << ["xlsxrb-adapters (RubyXL)", "xlsxrb_adapters"]
    write_targets << ["rubyXL (Original)", "rubyXL"]
  end
  if %w[all write_xlsx].include?(CATEGORY)
    write_targets << ["xlsxrb-adapters (WriteXLSX)", "xlsxrb_adapters_write_xlsx"]
    write_targets << ["write_xlsx (Original 1.15.1)", "write_xlsx"]
  end
  if %w[all fast_excel].include?(CATEGORY)
    write_targets << ["xlsxrb-adapters (FastExcel)", "xlsxrb_adapters_fast_excel"]
    write_targets << ["fast_excel (Original 0.5.0)", "fast_excel"]
  end

  write_targets.each do |name, lib|
    target_file = "tmp/bench_write_#{lib}.xlsx"
    res = run_benchmark_series(name, lib, "write", ROWS, COLS, target_file, RUNS)
    write_results[lib] = res if res
    FileUtils.rm_f(target_file)
  end
end

# 2. Benchmark Read (for libraries supporting read/parse, e.g. rubyXL, roo, xsv, creek)
read_results = {}
if %w[all rubyxl roo xsv creek].include?(CATEGORY)
  ref_file = "tmp/bench_reference_data.xlsx"
  puts "\n[Setup] Generating reference file (#{ROWS} x #{COLS}) for read benchmarks..."
  run_isolated("xlsxrb_stream", "write", ROWS, COLS, ref_file)

  puts "\n=== Benchmarking Read Performance ==="
  read_targets = [
    ["xlsxrb (Streaming)", "xlsxrb_stream"],
    ["xlsxrb (In-Memory)", "xlsxrb_inmemory"]
  ]
  if %w[all rubyxl].include?(CATEGORY)
    read_targets << ["xlsxrb-adapters (RubyXL)", "xlsxrb_adapters"]
    read_targets << ["rubyXL (Original)", "rubyXL"]
  end
  if %w[all roo].include?(CATEGORY)
    read_targets << ["xlsxrb-adapters (Roo)", "xlsxrb_adapters_roo"]
    read_targets << ["roo (Original 3.0.0)", "roo"]
  end
  if %w[all xsv].include?(CATEGORY)
    read_targets << ["xlsxrb-adapters (Xsv)", "xlsxrb_adapters_xsv"]
    read_targets << ["xsv (Original 1.4.1)", "xsv"]
  end
  if %w[all creek].include?(CATEGORY)
    read_targets << ["xlsxrb-adapters (Creek)", "xlsxrb_adapters_creek"]
    read_targets << ["creek (Original 2.6.3)", "creek"]
  end

  read_targets.each do |name, lib|
    res = run_benchmark_series(name, lib, "read", ROWS, COLS, ref_file, RUNS)
    read_results[lib] = res if res
  end
  FileUtils.rm_f(ref_file)
end

FileUtils.rm_f(runner_file)

def print_table(title, items, note: nil)
  puts "\n### #{title}"
  puts note if note
  puts ""
  puts "| Library                            | Time (Median) | Time (Mean) | Peak Memory | GC Count |"
  puts "| :--------------------------------- | :------------ | :---------- | :---------- | :------- |"
  items.compact.each do |r|
    printf "| %-34s | %6.3f s      | %6.3f s    | %7.1f MB  | %6.1f   |\n",
           r[:name], r[:median_time], r[:mean_time], r[:median_mem], r[:median_gc]
  end
end

puts "\n" + ("=" * 80)
puts "# Benchmark Summary (#{ROWS * COLS} cells: #{ROWS} rows x #{COLS} cols)"

if %w[all caxlsx].include?(CATEGORY)
  puts "\n" + ("-" * 80)
  puts "## Caxlsx Migration Benchmark (caxlsx vs. xlsxrb-adapters vs. xlsxrb)"

  caxlsx_write = [
    write_results["caxlsx"],
    write_results["xlsxrb_adapters_caxlsx"],
    write_results["xlsxrb_adapters_caxlsx_streaming"],
    write_results["xlsxrb_inmemory"],
    write_results["xlsxrb_stream"]
  ]
  print_table("Caxlsx Write Performance", caxlsx_write)
end

if %w[all rubyxl].include?(CATEGORY)
  puts "\n" + ("-" * 80)
  puts "## RubyXL Migration Benchmark (rubyXL vs. xlsxrb-adapters vs. xlsxrb)"

  rubyxl_write = [
    write_results["rubyXL"],
    write_results["xlsxrb_adapters"],
    write_results["xlsxrb_inmemory"],
    write_results["xlsxrb_stream"]
  ]
  print_table("RubyXL Write Performance", rubyxl_write)

  rubyxl_read = [
    read_results["rubyXL"],
    read_results["xlsxrb_adapters"],
    read_results["xlsxrb_inmemory"],
    read_results["xlsxrb_stream"]
  ]
  print_table("RubyXL Read Performance", rubyxl_read)
end

if %w[all roo].include?(CATEGORY)
  puts "\n" + ("-" * 80)
  puts "## Roo Migration Benchmark (roo vs. xlsxrb-adapters vs. xlsxrb)"

  roo_read = [
    read_results["roo"],
    read_results["xlsxrb_adapters_roo"],
    read_results["xlsxrb_inmemory"],
    read_results["xlsxrb_stream"]
  ]
  print_table("Roo Read Performance", roo_read)
end

if %w[all xsv].include?(CATEGORY)
  puts "\n" + ("-" * 80)
  puts "## Xsv Migration Benchmark (xsv vs. xlsxrb-adapters vs. xlsxrb)"

  xsv_read = [
    read_results["xsv"],
    read_results["xlsxrb_adapters_xsv"],
    read_results["xlsxrb_inmemory"],
    read_results["xlsxrb_stream"]
  ]
  print_table("Xsv Read Performance", xsv_read)
end

if %w[all creek].include?(CATEGORY)
  puts "\n" + ("-" * 80)
  puts "## Creek Migration Benchmark (creek vs. xlsxrb-adapters vs. xlsxrb)"

  creek_read = [
    read_results["creek"],
    read_results["xlsxrb_adapters_creek"],
    read_results["xlsxrb_inmemory"],
    read_results["xlsxrb_stream"]
  ]
  print_table("Creek Read Performance", creek_read)
end

if %w[all write_xlsx].include?(CATEGORY)
  puts "\n" + ("-" * 80)
  puts "## WriteXLSX Migration Benchmark (write_xlsx vs. xlsxrb-adapters vs. xlsxrb)"

  write_xlsx_write = [
    write_results["write_xlsx"],
    write_results["xlsxrb_adapters_write_xlsx"],
    write_results["xlsxrb_inmemory"],
    write_results["xlsxrb_stream"]
  ]
  print_table("WriteXLSX Write Performance", write_xlsx_write)
end

if %w[all fast_excel].include?(CATEGORY)
  puts "\n" + ("-" * 80)
  puts "## FastExcel Migration Benchmark (fast_excel vs. xlsxrb-adapters vs. xlsxrb)"

  fast_excel_write = [
    write_results["fast_excel"],
    write_results["xlsxrb_adapters_fast_excel"],
    write_results["xlsxrb_inmemory"],
    write_results["xlsxrb_stream"]
  ]
  print_table("FastExcel Write Performance", fast_excel_write)
end

puts "\n" + ("=" * 80)

