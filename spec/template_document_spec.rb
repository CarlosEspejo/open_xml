require 'spec_helper'
require 'tempfile'
require 'nokogiri'
require 'pry'

describe TemplateDocument do

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
    t.process({ 'person_name' => '<b>Carlos</b>', 'person_age' => 30 })
    doc = Nokogiri::XML(t.parts['word/document.xml'])

    text = doc.xpath('//w:t').text
    text[/Carlos/].wont_be_nil
    text[/30/].wont_be_nil
  end

  it 'should replace one key word with many items' do
    data = {
      'my_list' => [
                    'list 1',
                    'list 2',
                    'list 3',
                    'list 4',
                    'list 5'
                  ]
    }

    t = TemplateDocument.new(path: template_path)
    t.process(data)

    doc = Nokogiri::XML(t.parts['word/document.xml'])

    text = doc.xpath('//w:t').text

    text[/list 2/].wont_be_nil
  end

  it "should cache the template document" do
    t = TemplateDocument.new(path: template_path)
    t.process({'person_name' => 'steve'})
    doc = Nokogiri::XML(t.parts['word/document.xml'])

    doc.text[/steve/].wont_be_nil

    t.process({'person_name' => 'carlos'})
    doc = Nokogiri::XML(t.parts['word/document.xml'])

    doc.text[/carlos/].wont_be_nil
  end

  def template_path
    "#{File.expand_path('samples', File.dirname(__FILE__))}/template_sample.docx"
  end

end
