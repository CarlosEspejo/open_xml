require_relative '../lib/open_xml'

doc = OpenXml::TemplateDocument.new(path: "./template.docx", data: {"[SUPERPOWER]" => "Bug Fixing!!!!"})
doc.process

IO.write "./powers.docx", doc.to_zip_buffer.string
