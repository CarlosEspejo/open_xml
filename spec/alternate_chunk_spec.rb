require 'spec_helper'

describe "Adding HTML through the alternate chunk feature" do

  let(:report_path){"#{File.expand_path('samples', File.dirname(__FILE__))}/report.docx"}

  it "should generate chunk file" do
    t = TemplateDocument.new(path: report_path)
    t.process 'my_content' => {text: '<u>This is underlined</u>', html: true}
    t.parts['word/my_content1.xhtml'].must_equal '<html><body><u>This is underlined</u></body></html>'
  end

  it "should add chunk file to document relationship file" do
    t = TemplateDocument.new(path: report_path)
    t.process 'my_content' => {text: '<u>This is underlined</u>', html: true}
    doc = Nokogiri::XML(t.parts['word/_rels/document.xml.rels'])
    doc.xpath('//xmlns:Relationship/@Id').map{|n| n.text}.must_include 'my_content1'
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
      t.parts['word/my_content4.xhtml'][/list 4/].wont_be_nil
    end

end
