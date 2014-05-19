#require 'spec_helper'
#require 'nokogiri'
#require 'pry'

#describe "HTML to WordprocessingML transformations" do

  #it "should transform b or strong tags" do
    #w = xslt.transform(to_doc('<strong>word</strong>'))
    #single_line(w).must_equal strong
  #end

  #it "should transform i or em tags" do
    #w = xslt.transform(to_doc('<em>word</em>'))
    #single_line(w).must_equal em
  #end

  #let(:xslt){Nokogiri::XSLT(File.read("#{File.expand_path('../lib/xslt', File.dirname(__FILE__))}/basic_html.xslt"))}
  #let(:strong){'<?xml version="1.0"?><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:b/></w:rPr><w:t xml:space="preserve">word</w:t></w:r>'}
  #let(:em) {'<?xml version="1.0"?><w:r xmlns:w="http://schemas.openxmlformats.org/wordprocessingml/2006/main"><w:rPr><w:i/></w:rPr><w:t xml:space="preserve">word</w:t></w:r>'}

  #def to_doc(text)
    #Nokogiri::HTML(text)
  #end

  #def single_line(doc)
    #doc.to_xml(indent: 0).gsub("\n", "")
  #end
#end
