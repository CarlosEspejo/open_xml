require 'zip'

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
        parts.each do |k, v|
          w.put_next_entry k
          w.write v
        end
      end
    end

    def process
      doc = Nokogiri::XML(parts["word/document.xml"])
      doc.xpath("//w:t").each do |node|
        data.each do |k, v|
          node.content = node.content.gsub(k, v.to_s) if node.content[/#{k}/]
        end
      end

      parts["word/document.xml"] = doc.to_xml
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
