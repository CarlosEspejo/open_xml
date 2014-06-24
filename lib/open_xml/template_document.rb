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
      register_type 'message/rfc822', 'mht'

      doc = Nokogiri::XML(parts['word/document.xml'])
      doc.xpath('//w:t').each do |node|
        data.each do |key, value|

          if node.content[/#{key}/]
            process_plain_text(node, key, value, doc) unless value[:html]
            process_html(node, key, value, doc) if value[:html]
          end

        end
      end

      parts['word/document.xml'] = flatten_xml doc
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
      new_node = create_chunk_file(key, value, doc)
      node.parent.parent.add_next_sibling new_node
      node.remove
    end

    def create_chunk_file(key, content, doc)
      id = key

      parts["word/#{id}.mht"] = build_mht(content)
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
      rel['Target'] = "/word/#{id}.mht"

      relationships.at_xpath('//xmlns:Relationships') << rel
      parts['word/_rels/document.xml.rels'] = flatten_xml relationships
    end

    def register_type(type, extension)
      content = Nokogiri::XML(parts['[Content_Types].xml'])
      node = Nokogiri::XML::Node.new 'Default', content
      node['ContentType'] = type
      node['Extension'] = extension

      content.at_xpath('//xmlns:Default').add_next_sibling node
      parts['[Content_Types].xml'] = flatten_xml content
    end

    def flatten_xml(doc)
      doc.to_xml(indent: 0).gsub("\n","")
    end

    def read_files
      Zip::File.new(template_path).each do |f|
        parts[f.name] = f.get_input_stream.read
      end

      @parts_cache = parts.clone
    end

    def build_mht(content)
      message =<<MESSAGE
MIME-Version: 1.0
Content-Type: multipart/related; boundary=MY-SEPARATOR

--MY-SEPARATOR
Content-Type: text/html; charset=utf-8
Content-Transfer-Encoding: 8bit

#{content[:text]}

MESSAGE


      content.fetch(:images){{}}.each do |key, value|
      message << img_template(key, value)
     end

      message << "\n--MY-SEPARATOR--"
      message
    end

    def img_template(key, value)
      <<IMG

--MY-SEPARATOR
Content-Location: #{key}
Content-Transfer-Encoding: Base64

#{value}

IMG
    end

  end
end
