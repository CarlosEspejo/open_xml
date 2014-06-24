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

      #parts["word/#{id}.xhtml"] = "<html><body>#{content}</body></html>"
      parts["word/#{id}.mht"] = mht_default_text
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

    def mht_default_text
      path = '/Users/cespejo/Pictures/mac_vim_icon_dark.png'
      encoded_image = Base64.encode64(File.read(path))

      message =<<MESSAGE
MIME-Version: 1.0
Content-Type: multipart/related; boundary=boundary-example-1

--boundary-example-1
Content-Type: text/html; charset=utf-8
Content-Transfer-Encoding: 8bit

<html>
  <body>
    <div class="image-meta">
      <h2>Epson Smartcanvas</h2>

      <span class="caption"><p>New for this year, Epson's Smart Canvas watch gets a little artistic with its design.</p></span>



    <p class="credits">
                <time datetime="2014-06-05 15:23:00">June 5, 2014 8:23 AM PDT</time><span class="credit">
                        <strong>Photo by:</strong> Nic Healey/CNET
    <strong> /  Caption by:</strong> <a rel="author" href="/profiles/nichealey/" itemprop="name">Nic Healey</a>                                                            </span>
    </p>
    <p>
      <table border=3>
      <tr><td>1</td><td>2</td></tr>
      <tr><td>3</td><td>4</td></tr>
      </table>
    </p>
  </div>
    <img src='my_image.png' />
  </body>
</html>

--boundary-example-1
Content-Location: my_image.png
Content-Transfer-Encoding: Base64

#{encoded_image}

--boundary-example-1--
MESSAGE

      message
    end

  end
end
