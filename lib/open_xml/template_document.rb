require 'zip'
require 'nokogiri'

module OpenXml
  class TemplateDocument
    attr_reader :template_path, :parts

    def initialize(options)
      @template_path = options.fetch(:path)
      @parts = {}
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

    def process(data)
      @parts = @parts_cache.clone

      doc = Nokogiri::XML(parts['word/document.xml'])
      doc.xpath('//w:t').each do |node|
        data.each do |key, value|

          if node.content[/#{key}/]
            values = Array(value)

            if values.size > 1
              values.each do |v|
                br = Nokogiri::XML::Node.new 'w:br', doc
                n = Nokogiri::XML::Node.new 'w:t', doc

                n.content = v.to_s

                node.parent << n
                node.parent << br
              end

              node.remove
            else
              node.content = node.content.gsub(key, values.first.to_s) if values.first
            end

          end

        end
      end

      parts['word/document.xml'] = doc.to_xml
    end

    private

    def split_parts
      Zip::File.new(template_path).each do |f|
        parts[f.name] = f.get_input_stream.read
      end

      @parts_cache = parts.clone
    end
  end
end
