require 'spec_helper'
require 'base64'

describe "Adding HTML and Images through the alternate chunk feature" do

  let(:report_path){"#{File.expand_path('samples', __dir__)}/report.docx"}
  let(:encoded_img){Base64.encode64(File.read("#{File.expand_path('samples', __dir__)}/caterpillar.jpg"))}

  let(:t){TemplateDocument.new(path: report_path)}

  it "should register MIME html type" do
    t.process 'my_content' => {text: 'empty', html: true}
    doc = Nokogiri::XML(t.parts['[Content_Types].xml'])
    doc.xpath('//xmlns:Default/@ContentType').map(&:value).must_include 'message/rfc822'
  end

  it "should add the chunk id to the rels file" do
    t.process 'my_content' => {text: 'empty', html: true}
    doc = Nokogiri::XML(t.parts['word/_rels/document.xml.rels'])
    doc.xpath('//xmlns:Relationship/@Id').map(&:value).must_include 'my_content'
  end

  it "should generate MIME html file" do
    t.process 'my_content' => {text: '<u>This is underlined</u>', html: true}
    t.parts['word/my_content.mht'].must_match(/#{'<u>This is underlined</u>'}/)
  end

  it "should generate MIME html file with a image" do
    content = '<h1>Look at the image</h1><img src="./image.jpg" />'

    t.process 'my_content' => {text: content, html: true, images: {"./image.jpg" => encoded_img}}
    t.parts['word/my_content.mht'].must_match(/#{'Content-Location: ./image.jpg'}/)
  end

  it "should geneate MIME html file with multiple images" do
    content = '<h1>Look at the images</h1>'
    content << '<img src="/image.jpg" /><br/><br/><img src="/image2.jpg" />'
    t.process 'my_content' => {text: content, html: true, images: {"/image.jpg" => encoded_img, '/image2.jpg' => encoded_img}}
    t.parts['word/my_content.mht'].must_match(/#{'Content-Location: /image.jpg'}/)
    t.parts['word/my_content.mht'].must_match(/#{'Content-Location: /image2.jpg'}/)
    IO.write File.expand_path('~/Downloads/test.docx'), t.to_zip_buffer.string
  end

end
