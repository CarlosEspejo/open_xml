require 'spec_helper'
require 'tempfile'
require 'nokogiri'

describe TemplateDocument do

  it "should read the template file and split it into parts" do
    t = TemplateDocument.new(path: template_path)
    t.parts.keys.must_equal template_parts.keys

  end

  it "should create a new file from the template parts" do
    t = Tempfile.new(['output','.docx'])

    temp_doc = TemplateDocument.new(path: template_path)

    IO.write t.path, temp_doc.to_zip_buffer.string

    Zip::File.new(t.path).each do |f|
      temp_doc.parts[f.name].must_equal f.get_input_stream.read
    end

  end

  it "should replace key words in the document xml" do
    t = TemplateDocument.new(path: template_path, data: {"NAME" => "Steve Jobs"})
    t.process
    doc = Nokogiri::XML(t.parts["word/document.xml"])
    doc.xpath('//w:t').text[/Steve Jobs/].must_equal "Steve Jobs"
  end
  
  let(:template_path){"#{File.expand_path('samples', File.dirname(__FILE__))}/template_sample.docx"}

  let(:template_parts) do
    {
      "_rels/.rels" => '',
      "docProps/core.xml" => '',
      "docProps/app.xml" => '',
      "word/document.xml" => '',
      "word/styles.xml" => '',
      "word/fontTable.xml" => '',
      "word/header.xml" => '',
      "word/footer.xml" => '',
      "word/settings.xml" => '',
      "word/_rels/document.xml.rels" => '',
      "[Content_Types].xml" => ''
    }
  end

end
