open_xml
========

Library for reading and writing to open xml documents (*but at the moment you can generate word docs from a template*)

## Requirements
* rubyzip
* nokogiri
    
## Usage
Provide a path to a docx with the word [SUPERPOWER] placed anywhere.
```ruby
require 'open_xml'

doc = OpenXml::TemplateDocument.new(path: "[path to template]", data: {"[SUPERPOWER]" => "Bug Fixing!!!!"})
doc.process

IO.write "./powers.docx", doc.to_zip_buffer.string
```

## Todo
  * ~~Implement reading and writing the word zip files~~
  * ~~Create a template word document with formatted key words (bold, 14pt).~~
  * ~~Replace the key words with the supplied plain text content but maintain all the formatting.~~
  * ~~Handle replacing a key with multiple content~~
  * Format html content for wordprocessingML e.x. bold, italic,
    underline
