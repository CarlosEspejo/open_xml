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
    temp = Tempfile.new(['output','.docx'])

    t = TemplateDocument.new(path: template_path, data: {"[NAME]" => "<b>Carlos</b>", "[AGE]" => 30 })
    t.process
    IO.write temp.path, t.to_zip_buffer.string

    processed = TemplateDocument.new(path: temp.path)
    doc = Nokogiri::XML(processed.parts["word/document.xml"])

    doc.xpath('//w:t').text[/\[NAME\]/].must_be_nil
    doc.xpath('//w:t').text[/\[AGE\]/].must_be_nil
  end

  it "should replace one key word with many items" do
    data = {
      "[LIST]" => [
                    "list 1",
                    "list 2",
                    "list 3",
                    "list 4",
                    "list 5"
                  ]
    }

    t = TemplateDocument.new(path: template_path, data: data)
    t.process
    doc = Nokogiri::XML(t.parts["word/document.xml"])

    text = doc.xpath('//w:t').text

    text[/list 1\nlist 2/].wont_be_nil
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
