OpenXml
========

Library for reading and writing to open xml documents (*but at the moment you can generate word docs from a template*)

## Installation

Add this line to your application's Gemfile:

    gem 'open_xml'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install open_xml

   
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
  * Extract these features into a gem
  * Format html content for wordprocessingML e.x. bold, italic,
    underline

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
