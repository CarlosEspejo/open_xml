OpenXml
========

A ruby library for generating word documents that can handle basic html and images too.

## Installation

Add this line to your application's Gemfile:

    gem 'open_xml'

And then execute:

    $ bundle

Or install it yourself as:

    $ gem install open_xml


## Usage
Provide a path to a docx with the text **[SUPERPOWER]** placed anywhere.

```ruby
require 'open_xml'

doc = OpenXml::TemplateDocument.new(path: "[path to template]")
doc.process({"[SUPERPOWER]" => {text: "Bug Fixing!!!!"}})

IO.write "./powers.docx", doc.to_zip_buffer.string
```

HTML content

```ruby
doc = OpenXml::TemplateDocument.new(path: "[path to template]")
doc.process({"[SUPERPOWER]" => {text: "<h1>Bug Fixing!!!!</h1>", html: true}})
```

HTML with images

```ruby
doc = OpenXml::TemplateDocument.new(path: "[path to template]")
doc.process({"[SUPERPOWER]" => {text: "<img src='/powers.png' />", html: true, images: {'/powers.png' => "[Base64 encoded image]"}}})
```

## Todo
  * ~~Implement reading and writing the word zip files~~
  * ~~Create a template word document with formatted key words (bold, 14pt).~~
  * ~~Replace the key words with the supplied plain text content but maintain all the formatting.~~
  * ~~Handle replacing a key with multiple content~~
  * ~~Extract these features into a gem~~
  * ~~Format html content for wordprocessingML e.x. bold, italic,
    underline and handle images~~

## Contributing

1. Fork it
2. Create your feature branch (`git checkout -b my-new-feature`)
3. Commit your changes (`git commit -am 'Add some feature'`)
4. Push to the branch (`git push origin my-new-feature`)
5. Create new Pull Request
