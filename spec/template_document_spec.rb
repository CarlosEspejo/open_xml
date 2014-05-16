require 'spec_helper'
require 'tempfile'
require 'nokogiri'
require 'pry'

describe TemplateDocument do

  let(:template_path){"#{File.expand_path('samples', File.dirname(__FILE__))}/template_sample.docx"}
  let(:report_path){"#{File.expand_path('samples', File.dirname(__FILE__))}/report.docx"}

  it 'should read the template file and split it into parts' do
    t = TemplateDocument.new(path: template_path)
    t.parts.keys.must_include 'word/document.xml'
  end

  it 'should create a new file from the template parts' do
    t = Tempfile.new(['output', '.docx'])

    temp_doc = TemplateDocument.new(path: template_path)

    IO.write t.path, temp_doc.to_zip_buffer.string

    Zip::File.new(t.path).each do |f|
      temp_doc.parts[f.name].must_equal f.get_input_stream.read
    end

  end

  it 'should replace key words in the document xml' do
    t = TemplateDocument.new(path: template_path)
    t.process({ 'person_name' => {text: '<b>Carlos</b>'}, 'person_age' => {text: 30} })
    doc = Nokogiri::XML(t.parts['word/document.xml'])

    text = doc.xpath('//w:t').text
    text[/Carlos/].wont_be_nil
    text[/30/].wont_be_nil
  end

  it 'should replace one key word with many items' do
    data = {
      'my_list' => {text: [
                    'list 1',
                    'list 2',
                    'list 3',
                    'list 4',
                    'list 5'
                  ]}
    }

    t = TemplateDocument.new(path: template_path)
    t.process(data)

    doc = Nokogiri::XML(t.parts['word/document.xml'])

    text = doc.xpath('//w:t').text

    text[/list 2/].wont_be_nil
  end

  it "should cache the template document" do
    t = TemplateDocument.new(path: template_path)
    t.process({'person_name' => {text: 'steve'}})
    doc = Nokogiri::XML(t.parts['word/document.xml'])

    doc.text[/steve/].wont_be_nil

    t.process({'person_name' => {text: 'carlos'}})
    doc = Nokogiri::XML(t.parts['word/document.xml'])

    doc.text[/carlos/].wont_be_nil
  end

  describe "Converting HTML to WordML" do
    it "should convert <p> tags" do
      t = TemplateDocument.new(path: report_path)
      t.process({'my_content' => {text: '<p>This content should not have paragraph tags</p>', html: true}})

      doc = Nokogiri::XML(t.parts['word/document.xml'])
      text = doc.xpath('//w:p').text
      text[/<p>/].must_be_nil
      text[/This content should/].wont_be_nil
    end

    it "should convert a list of <p> tags" do
      t = TemplateDocument.new(path: report_path)

      data = {
                'my_content' => {text: [
                    '<p>list 1</p>',
                    '<p>list 2</p>',
                    '<p>list 3</p>',
                    '<p>list 4</p>',
                    '<p>list 5</p>'
                  ], html: true}
       }

      t.process(data)
      doc = Nokogiri::XML(t.parts['word/document.xml'])

      text = doc.xpath('//w:p').text
      text[/<p>/].must_be_nil
      text[/list 4/].wont_be_nil
    end

  end
end
