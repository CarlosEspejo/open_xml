# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'open_xml/version'

Gem::Specification.new do |spec|
  spec.name          = "open_xml"
  spec.version       = OpenXml::VERSION
  spec.authors       = ["Carlos Espejo"]
  spec.email         = ["carlosespejo@gmail.com"]
  spec.description   = %q{Currently you can only generate word documents from a template word document.}
  spec.summary       = %q{Library for reading and writing to open xml documents (*but at the moment you can generate word docs from a template*)}
  spec.homepage      = "https://github.com/CarlosEspejo/open_xml"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "rubyzip"
  spec.add_development_dependency "nokogiri"
  spec.add_development_dependency "minitest"
  spec.add_development_dependency "pry"
end

