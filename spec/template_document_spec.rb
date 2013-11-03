require 'spec_helper'

describe TemplateDocument do

  it "should read the key word of the template file" do
    t = TemplateDocument.new(path: './samples/template_sample.docx')
    t.must_respond_to :doc
  end
  
end
