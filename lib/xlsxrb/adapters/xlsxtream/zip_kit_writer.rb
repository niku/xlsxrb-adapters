# frozen_string_literal: true

# rbs_inline: enabled

begin
  require "zip_kit"
rescue LoadError
  # zip_kit is optional; fallback to Xlsxrb::Ooxml::ZipWriter
end
require "xlsxrb"

module Xlsxrb
  module Adapters
    module Xlsxtream
      # Streaming ZIP archive writer compatible with Xlsxtream::ZipKitWriter.
      # Writes compressed entries using ZipKit::Streamer if available, or Xlsxrb::Ooxml::ZipWriter.
      class ZipKitWriter
        BUFFER_SIZE = 64 * 1024

        # Wraps an output sink into a ZipKitWriter instance.
        #
        # @param output [String, IO, Object]
        # @return [ZipKitWriter]
        def self.with_output_to(output)
          if output.is_a?(self)
            output
          elsif defined?(::ZipKit::Streamer) && output.is_a?(::ZipKit::Streamer)
            new(output, close: [])
          elsif output.is_a?(String)
            # rubocop:disable-next Style/FileOpen
            file = File.open(output, "wb")
            if defined?(::ZipKit::Streamer)
              streamer = ::ZipKit::Streamer.new(file)
              new(streamer, close: [streamer, file])
            else
              zip_writer = Xlsxrb::Ooxml::ZipWriter.new(file)
              new(zip_writer, close: [zip_writer, file], mode: :xlsxrb)
            end
          elsif output.respond_to?(:<<) || output.respond_to?(:write)
            if defined?(::ZipKit::Streamer)
              streamer = ::ZipKit::Streamer.new(output)
              new(streamer, close: [streamer])
            else
              zip_writer = Xlsxrb::Ooxml::ZipWriter.new(output)
              new(zip_writer, close: [zip_writer], mode: :xlsxrb)
            end
          else
            error = <<~MSG
              An `output` object must be one of:

              * A String containing a path to a file ("workbook.xslx")
              * A ZipKit::Streamer
              * An IO-like object responding to #<< or #write

              but it was a #{output.class}
            MSG
            raise ArgumentError, error
          end
        end

        def initialize(streamer, close: [], mode: :zip_kit)
          @streamer = streamer
          @currently_writing_file_inside_zip = nil
          @buffer = String.new
          @close = close
          @mode = mode
        end

        # Appends data to internal buffer.
        #
        # @param data [String]
        # @return [self]
        #: (String data) -> self
        def <<(data)
          @buffer << data
          flush_buffer if @buffer.size >= BUFFER_SIZE
          self
        end

        # Opens a new compressed file entry in the ZIP package.
        #
        # @param path [String]
        # @return [void]
        #: (String path) -> void
        def add_file(path)
          flush_file
          if @mode == :xlsxrb
            @streamer.start_entry(path)
            @currently_writing_file_inside_zip = :xlsxrb_entry
          else
            @currently_writing_file_inside_zip = @streamer.write_deflated_file(path)
          end
        end

        # Closes current entry and finalized ZIP archive.
        #
        # @return [void]
        #: () -> void
        def close
          flush_file
          @close.each(&:close)
        end

        private

        def flush_buffer
          return if @buffer.empty?

          if @mode == :xlsxrb
            @streamer.write_data(@buffer)
          else
            @currently_writing_file_inside_zip << @buffer
          end
          @buffer.clear
        end

        def flush_file
          return unless @currently_writing_file_inside_zip

          flush_buffer if @buffer.bytesize.positive?
          if @mode == :xlsxrb
            @streamer.finish_entry
          else
            @currently_writing_file_inside_zip.close
          end
          @currently_writing_file_inside_zip = nil
        end
      end
    end
  end
end
