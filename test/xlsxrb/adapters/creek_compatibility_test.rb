# frozen_string_literal: true

require "test_helper"

class XlsxrbAdaptersCreekCompatibilityTest < Test::Unit::TestCase
  FIXTURES = %w[
    escaped.xlsx
    escaped2.xlsx
    large_numbers.xlsx
    sample-with-headers.xlsx
    sample-with-headers_namespaced.xlsx
    sample-with-images.xlsx
    sample-with-one-cell-anchored-images.xlsx
    sample.xlsx
    sample_dates.xlsx
    sample_namespaced.xlsx
  ].freeze

  def fixture(name)
    File.expand_path("../../fixtures/creek/#{name}", __dir__)
  end

  def test_side_by_side_rows_all_fixtures
    FIXTURES.each do |filename|
      path = fixture(filename)
      orig_wb = Creek::Book.new(path, check_file_extension: false)
      adapt_wb = Xlsxrb::Adapters::Creek::Book.new(path, check_file_extension: false)

      assert_equal orig_wb.sheets.size, adapt_wb.sheets.size, "#{filename}: sheet count mismatch"

      orig_wb.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_wb.sheets[idx]

        assert_equal orig_sheet.name, adapt_sheet.name, "#{filename} [#{idx}]: sheet name mismatch"
        assert_equal orig_sheet.state, adapt_sheet.state, "#{filename} [#{orig_sheet.name}]: state mismatch"
        assert_equal orig_sheet.rid, adapt_sheet.rid, "#{filename} [#{orig_sheet.name}]: rid mismatch"
        assert_equal orig_sheet.sheetid, adapt_sheet.sheetid, "#{filename} [#{orig_sheet.name}]: sheetid mismatch"

        orig_rows = orig_sheet.rows.to_a
        adapt_rows = adapt_sheet.rows.to_a

        assert_equal orig_rows.size, adapt_rows.size, "#{filename} [#{orig_sheet.name}]: row count mismatch"
        orig_rows.each_with_index do |orig_row, r_idx|
          adapt_row = adapt_rows[r_idx]
          assert_equal orig_row, adapt_row, "#{filename} [#{orig_sheet.name}] row #{r_idx} mismatch"
        end
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_simple_rows_all_fixtures
    FIXTURES.each do |filename|
      path = fixture(filename)
      orig_wb = Creek::Book.new(path, check_file_extension: false)
      adapt_wb = Xlsxrb::Adapters::Creek::Book.new(path, check_file_extension: false)

      orig_wb.sheets.each_with_index do |orig_sheet, idx|
        adapt_sheet = adapt_wb.sheets[idx]

        orig_rows = orig_sheet.simple_rows.to_a
        adapt_rows = adapt_sheet.simple_rows.to_a

        assert_equal orig_rows.size, adapt_rows.size, "#{filename} [#{orig_sheet.name}]: simple_rows count mismatch"
        orig_rows.each_with_index do |orig_row, r_idx|
          adapt_row = adapt_rows[r_idx]
          assert_equal orig_row, adapt_row, "#{filename} [#{orig_sheet.name}] simple_row #{r_idx} mismatch"
        end
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_rows_with_meta_data
    %w[sample.xlsx sample_dates.xlsx large_numbers.xlsx].each do |filename|
      path = fixture(filename)
      orig_wb = Creek::Book.new(path)
      adapt_wb = Xlsxrb::Adapters::Creek::Book.new(path)

      orig_sheet = orig_wb.sheets.first
      adapt_sheet = adapt_wb.sheets.first

      orig_meta = orig_sheet.rows_with_meta_data.to_a
      adapt_meta = adapt_sheet.rows_with_meta_data.to_a

      assert_equal orig_meta.size, adapt_meta.size, "#{filename}: meta count mismatch"
      orig_meta.each_with_index do |orig_row, r_idx|
        adapt_row = adapt_meta[r_idx]
        assert_equal orig_row, adapt_row, "#{filename} meta row #{r_idx} mismatch"
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_simple_rows_with_meta_data
    %w[sample.xlsx sample_dates.xlsx].each do |filename|
      path = fixture(filename)
      orig_wb = Creek::Book.new(path)
      adapt_wb = Xlsxrb::Adapters::Creek::Book.new(path)

      orig_sheet = orig_wb.sheets.first
      adapt_sheet = adapt_wb.sheets.first

      orig_meta = orig_sheet.simple_rows_with_meta_data.to_a
      adapt_meta = adapt_sheet.simple_rows_with_meta_data.to_a

      assert_equal orig_meta.size, adapt_meta.size, "#{filename}: simple meta count mismatch"
      orig_meta.each_with_index do |orig_row, r_idx|
        adapt_row = adapt_meta[r_idx]
        assert_equal orig_row, adapt_row, "#{filename} simple meta row #{r_idx} mismatch"
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_with_headers
    %w[sample-with-headers.xlsx sample-with-headers_namespaced.xlsx].each do |filename|
      path = fixture(filename)
      orig_wb = Creek::Book.new(path, with_headers: true)
      adapt_wb = Xlsxrb::Adapters::Creek::Book.new(path, with_headers: true)

      orig_rows = orig_wb.sheets.first.simple_rows.to_a
      adapt_rows = adapt_wb.sheets.first.simple_rows.to_a

      assert_equal orig_rows.size, adapt_rows.size, "#{filename}: with_headers row count mismatch"
      orig_rows.each_with_index do |orig_row, r_idx|
        adapt_row = adapt_rows[r_idx]
        assert_equal orig_row, adapt_row, "#{filename} with_headers row #{r_idx} mismatch"
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_images
    %w[sample-with-images.xlsx sample-with-one-cell-anchored-images.xlsx].each do |filename|
      path = fixture(filename)
      orig_wb = Creek::Book.new(path)
      adapt_wb = Xlsxrb::Adapters::Creek::Book.new(path)

      orig_sheet = orig_wb.sheets.first.with_images
      adapt_sheet = adapt_wb.sheets.first.with_images

      orig_rows = orig_sheet.rows.to_a
      adapt_rows = adapt_sheet.rows.to_a

      assert_equal orig_rows.size, adapt_rows.size, "#{filename}: images row count mismatch"
      orig_rows.each_with_index do |orig_row, r_idx|
        adapt_row = adapt_rows[r_idx]
        # Normalize Pathnames to basenames because tmp directories differ between instances
        norm_orig = orig_row.transform_values { |v| v.is_a?(Array) ? v.map { |p| File.basename(p.to_s) } : v }
        norm_adapt = adapt_row.transform_values { |v| v.is_a?(Array) ? v.map { |p| File.basename(p.to_s) } : v }
        assert_equal norm_orig, norm_adapt, "#{filename} images row #{r_idx} mismatch"
      end

      # Test images_at on specific cells
      %w[A2 A4 A10 B3].each do |coord|
        orig_imgs = orig_sheet.images_at(coord)
        adapt_imgs = adapt_sheet.images_at(coord)
        if orig_imgs.nil?
          assert_nil adapt_imgs
        else
          assert_equal(orig_imgs.map { |p| File.basename(p.to_s) }, adapt_imgs.map { |p| File.basename(p.to_s) })
        end
      end

      orig_wb.close
      adapt_wb.close
    end
  end

  def test_side_by_side_shared_strings_xml
    %w[sst.xml sst_namespaced.xml].each do |filename|
      content = File.read(fixture(filename))
      doc = Nokogiri::XML(content)

      orig_dict = Creek::SharedStrings.parse_shared_string_from_document(doc)
      adapt_dict = Xlsxrb::Adapters::Creek::SharedStrings.parse_shared_string_from_document(doc)
      adapt_fast_dict = Xlsxrb::Adapters::Creek::SharedStrings.parse_shared_string_from_document(content)

      assert_equal orig_dict, adapt_dict, "#{filename}: Nokogiri SST dictionary mismatch"
      assert_equal orig_dict, adapt_fast_dict, "#{filename}: Fast-scan SST dictionary mismatch"
    end
  end

  def test_side_by_side_styles_xml
    content = File.read(fixture("styles/first.xml"))
    doc = Nokogiri::XML(content)

    orig_styles = Creek::Styles::StyleTypes.new(doc).call
    adapt_styles = Xlsxrb::Adapters::Creek::Styles::StyleTypes.new(doc).call
    adapt_fast_styles = Xlsxrb::Adapters::Creek::Styles::StyleTypes.new(content).call

    assert_equal orig_styles, adapt_styles, "Styles Nokogiri mismatch"
    assert_equal orig_styles, adapt_fast_styles, "Styles fast-scan mismatch"
  end
end
