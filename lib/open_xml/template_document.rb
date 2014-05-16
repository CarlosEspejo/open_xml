require 'zip'
require 'nokogiri'

module OpenXml
  class TemplateDocument
    attr_reader :template_path, :parts, :xslt

    def initialize(options)
      @template_path = options.fetch(:path)
      @xslt = options.fetch(:xslt){File.expand_path('../xslt/basic_html.xslt', File.dirname(__FILE__))}
      @parts = {}

      read_files
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
            process_plain_text(node, key, value, doc) unless value[:html]
            process_html(node, key, value, doc) if value[:html]
          end

        end
      end

      parts['word/document.xml'] = doc.to_xml(:indent => 0).gsub("\n","")
    end

    private

    def process_plain_text(node, key, value, doc)
      values = Array(value[:text])

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

    def process_html(node, key, value, doc)
      Array(value[:text]).each do |v|
        html = Nokogiri::HTML(v)
        new_value = xslt.transform(html).children.first

        node.parent.parent << new_value
      end

      node.remove
    end

    def read_files
      Zip::File.new(template_path).each do |f|
        parts[f.name] = f.get_input_stream.read
      end

      @xslt = Nokogiri::XSLT(File.read(xslt))
      @parts_cache = parts.clone
    end
  end
end
