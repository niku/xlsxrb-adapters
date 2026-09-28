# frozen_string_literal: true

# rbs_inline: enabled

module Xlsxrb
  module Adapters
    module Caxlsx
      # Package manages OOXML serialization and validation.
      class Package
        include OptionsParser

        # @return [App]
        attr_reader :app

        # @return [Core]
        attr_reader :core

        # @param options [Hash{Symbol => untyped}]
        #: (?Hash[Symbol, untyped] options) ?{ (Package) -> void } -> void
        def initialize(options = {})
          @workbook = nil
          @core = Core.new
          @app = App.new
          @core.creator = options[:author] if options[:author]
          @core.created = options[:created_at] if options[:created_at]
          parse_options(options)
          yield self if block_given?
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def use_autowidth=(v)
          Caxlsx.validate_boolean(v)
          workbook.use_autowidth = v
        end

        # @return [Boolean, nil]
        #: () -> bool?
        def use_shared_strings
          workbook.use_shared_strings
        end

        # @param v [Boolean]
        # @return [Boolean]
        #: (bool v) -> bool
        def use_shared_strings=(v)
          Caxlsx.validate_boolean(v)
          workbook.use_shared_strings = v
        end

        # @return [Workbook]
        #: () ?{ (Workbook) -> void } -> Workbook
        def workbook
          @workbook ||= Workbook.new
          yield @workbook if block_given?
          @workbook
        end

        # @param workbook [Workbook]
        # @return [Workbook]
        #: (Workbook workbook) -> Workbook
        def workbook=(workbook)
          DataTypeValidator.validate(:Package_workbook, Workbook, workbook)
          @workbook = workbook
        end

        # Serializes the workbook to disk.
        #
        # @param output [String]
        # @param options [Hash{Symbol => untyped}, Boolean]
        # @param secondary_options [Hash{Symbol => untyped}, nil]
        # @return [Boolean]
        #: (String output, ?Hash[Symbol, untyped] | bool options, ?Hash[Symbol, untyped]? secondary_options) -> bool
        def serialize(output, options = {}, secondary_options = nil)
          confirm_valid, _zip_command, password = parse_serialize_options(options, secondary_options)
          return false if confirm_valid && !validate.empty?

          wb_elements = to_xlsxrb
          Xlsxrb.write(output, wb_elements)

          OoxmlCrypt.encrypt_file(output, password, output) if password && !password.empty? && defined?(OoxmlCrypt)

          true
        end

        # Serializes the workbook to a StringIO instance.
        #
        # @param old_confirm_valid [Boolean, nil]
        # @param confirm_valid [Boolean]
        # @param password [String, nil]
        # @return [StringIO, Boolean]
        #: (?bool? old_confirm_valid, ?confirm_valid: bool, ?password: String?) -> (StringIO | false)
        def to_stream(old_confirm_valid = nil, confirm_valid: false, password: nil)
          confirm_valid = old_confirm_valid unless old_confirm_valid.nil?
          return false if confirm_valid && !validate.empty?

          wb_elements = to_xlsxrb
          binary = Xlsxrb.write(wb_elements)

          binary = OoxmlCrypt.encrypt(binary, password) if password && !password.empty? && defined?(OoxmlCrypt)

          io = StringIO.new(binary)
          io.binmode
          io
        end

        # Validates package contents.
        #
        # @return [Array[untyped]]
        #: () -> Array[untyped]
        def validate
          []
        end

        # @param file_name [String]
        # @param password [String]
        # @return [Boolean]
        #: (String file_name, String password) -> bool
        def encrypt(_file_name, _password)
          false
        end

        # Converts Package to an immutable Xlsxrb::Elements::Workbook.
        #
        # @return [Xlsxrb::Elements::Workbook]
        #: () -> Xlsxrb::Elements::Workbook
        def to_xlsxrb
          wb_elements = workbook.to_xlsxrb
          merged_unmapped = (wb_elements.unmapped_data || {}).dup
          facade_meta = (merged_unmapped[:facade] || {}).dup

          created_str = @core.created&.utc&.strftime("%Y-%m-%dT%H:%M:%SZ")
          modified_str = @core.modified&.utc&.strftime("%Y-%m-%dT%H:%M:%SZ")

          core_props = {
            creator: @core.creator,
            last_modified_by: @core.last_modified_by,
            created: created_str,
            modified: modified_str,
            title: @core.title
          }.compact
          facade_meta[:core_properties] = core_props unless core_props.empty?

          app_props = {
            company: @app.company,
            application: @app.application,
            app_version: @app.app_version
          }.compact
          facade_meta[:app_properties] = app_props unless app_props.empty?

          merged_unmapped[:facade] = facade_meta unless facade_meta.empty?

          Xlsxrb::Elements::Workbook.new(
            sheets: wb_elements.sheets,
            shared_strings: wb_elements.shared_strings,
            styles: wb_elements.styles,
            unmapped_data: merged_unmapped
          )
        end

        private

        # @param options [Hash{Symbol => untyped}, Boolean]
        # @param secondary_options [Hash{Symbol => untyped}, nil]
        # @return [Array[untyped]]
        #: (Hash[Symbol, untyped] | bool options, ?Hash[Symbol, untyped]? secondary_options) -> Array[untyped]
        def parse_serialize_options(options, secondary_options = nil)
          if options.is_a?(Hash)
            opts = options.dup
            opts.merge!(secondary_options) if secondary_options
            [opts.fetch(:confirm_valid, false), opts.fetch(:zip_command, nil), opts.fetch(:password, nil)]
          else
            parse_serialize_options((secondary_options || {}).merge(confirm_valid: options), nil)
          end
        end
      end
    end
  end
end
