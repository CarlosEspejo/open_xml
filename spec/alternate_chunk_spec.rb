require 'spec_helper'

describe "Adding HTML through the alternate chunk feature" do

  let(:report_path){"#{File.expand_path('samples', File.dirname(__FILE__))}/report.docx"}

  it "should generate chunk file" do
    t = TemplateDocument.new(path: report_path)
    t.process 'my_content' => {text: '<u>This is underlined</u>', html: true}
    #t.parts['word/my_content.xhtml'].must_equal '<html><body><u>This is underlined</u></body></html>'
  end

  it "should add chunk file to document relationship file" do
    t = TemplateDocument.new(path: report_path)
    t.process 'my_content' => {text: '<u>This is underlined</u>', html: true}
    doc = Nokogiri::XML(t.parts['word/_rels/document.xml.rels'])
    #binding.pry
    #doc.xpath('//xmlns:Relationship/@Id').map{|n| n.text}.must_include 'mycontent'

    IO.write 'test.docx', t.to_zip_buffer.string
  end
end
