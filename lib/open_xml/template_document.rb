require 'zip'
require 'nokogiri'

module OpenXml
  class TemplateDocument
    attr_reader :template_path, :parts

    def initialize(options)
      @template_path = options.fetch(:path)
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
      # move this into its own method
      content = Nokogiri::XML(parts['[Content_Types].xml'])
      type = Nokogiri::XML::Node.new 'Default', content
      type['ContentType'] = 'application/xhtml+xml'
      type['Extension'] = 'xhtml'
      content.at_xpath('//xmlns:Default').add_next_sibling type
      parts['[Content_Types].xml'] = content.to_xml(indent: 0).gsub("\n","")

      Array(value[:text]).each do |v|
        node.parent.parent.add_next_sibling create_chunk_file(key, v, doc)
      end

      node.remove
    end

    def create_chunk_file(key, content, doc)
      parts["word/#{key}.xhtml"] = "<html><body>#{content}</body></html>"

      relationships = Nokogiri::XML(parts['word/_rels/document.xml.rels'])
      rel = Nokogiri::XML::Node.new 'Relationship', relationships
      rel['Id'] = key
      rel['Type'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk'
      rel['Target'] = "/word/#{key}.xhtml"
      relationships.at_xpath('//xmlns:Relationships') << rel

      parts['word/_rels/document.xml.rels'] = relationships.to_xml(indent: 0).gsub("\n", "")
      chunk = Nokogiri::XML::Node.new 'w:altChunk', doc
      chunk['r:id'] = key
      chunk
    end

    def read_files
      Zip::File.new(template_path).each do |f|
        parts[f.name] = f.get_input_stream.read
      end

      @parts_cache = parts.clone
    end
  end
end
