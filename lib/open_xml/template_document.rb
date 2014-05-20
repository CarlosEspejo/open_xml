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
      register_type 'application/xhtml+xml', 'xhtml'

      doc = Nokogiri::XML(parts['word/document.xml'])
      doc.xpath('//w:t').each do |node|
        data.each do |key, value|

          if node.content[/#{key}/]
            process_plain_text(node, key, value, doc) unless value[:html]
            process_html(node, key, value, doc) if value[:html]
          end

        end
      end

      parts['word/document.xml'] = to_flat_xml doc
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
      previous_sibling = node.parent.parent

      Array(value[:text]).each_with_index do |v, index|
        new_node = create_chunk_file(key, v, doc, index + 1)
        previous_sibling.add_next_sibling new_node
        previous_sibling = new_node
      end

      node.remove
    end

    def create_chunk_file(key, content, doc, index)
      id = "#{key}#{index}"

      parts["word/#{id}.xhtml"] = "<html><body>#{content}</body></html>"
      add_relation id

      chunk = Nokogiri::XML::Node.new 'w:altChunk', doc
      chunk['r:id'] = id
      chunk
    end

    def add_relation(id)
      relationships = Nokogiri::XML(parts['word/_rels/document.xml.rels'])
      rel = Nokogiri::XML::Node.new 'Relationship', relationships
      rel['Id'] = id
      rel['Type'] = 'http://schemas.openxmlformats.org/officeDocument/2006/relationships/aFChunk'
      rel['Target'] = "/word/#{id}.xhtml"

      relationships.at_xpath('//xmlns:Relationships') << rel
      parts['word/_rels/document.xml.rels'] = to_flat_xml relationships
    end

    def register_type(type, extension)
      content = Nokogiri::XML(parts['[Content_Types].xml'])
      node = Nokogiri::XML::Node.new 'Default', content
      node['ContentType'] = type
      node['Extension'] = extension

      content.at_xpath('//xmlns:Default').add_next_sibling node
      parts['[Content_Types].xml'] = to_flat_xml content
    end

    def to_flat_xml(doc)
      doc.to_xml(indent: 0).gsub("\n","")
    end

    def read_files
      Zip::File.new(template_path).each do |f|
        parts[f.name] = f.get_input_stream.read
      end

      @parts_cache = parts.clone
    end
  end
end
