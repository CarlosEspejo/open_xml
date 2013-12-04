require 'zip'
require 'nokogiri'

module OpenXml
  class TemplateDocument
    attr_reader :template_path, :parts, :data

    def initialize(options)
      @template_path = options[:path]
      @parts = {}
      @data = options[:data]
      split_parts
    end

    def to_zip_buffer
      Zip::OutputStream.write_buffer do |w|
        parts.each do |key, value|
          w.put_next_entry key
          w.write value
        end
      end
    end

    def process
      doc = Nokogiri::XML(parts['word/document.xml'])
      doc.xpath('//w:t').each do |node|
        data.each do |key, value|
          node.content = node.content.gsub(key, Array(value).join("\n")) if node.content[/#{key}/]
        end
      end

      parts['word/document.xml'] = doc.to_xml
    end

    private

    def split_parts
      Zip::File.new(template_path).each do |f|
        parts[f.name] = f.get_input_stream.read
      end

      parts
    end
  end
end
