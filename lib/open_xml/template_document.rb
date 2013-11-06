require 'zip'

module OpenXml


  class TemplateDocument
    attr_reader :template_path, :parts

    def initialize(options)
      @template_path = options[:path]
      @parts = {}
      split_parts
    end

    def split_parts
      Zip::File.new(template_path).each do |f|
        parts[f.name] = f.get_input_stream.read
      end

      parts
    end

    def to_zip_buffer
      Zip::OutputStream.write_buffer do |w|
        parts.each do |k, v|
          w.put_next_entry k
          w.write v
        end

      end

    end

  end

end
