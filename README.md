open_xml
========

Library for reading and writing to open xml documents (*but at the moment you can generate word docs from a template*)

## Requirements
* rubyzip
* nokogiri
    
## Usage
for now do a git clone and run the following

```bash
bundle
ruby bin/sample.rb
```

## Todo
  * ~~Implement reading and writing the word zip files~~
  * ~~Create a template word document with formatted key words (bold, 14pt).~~
  * ~~Replace the key words with the supplied plain text content but maintain all the formatting.~~
  * ~~Handle replacing a key with multiple content~~
  * Format html content for wordprocessingML e.x. bold, italic,
    underline
