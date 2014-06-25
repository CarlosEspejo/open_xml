# coding: utf-8
lib = File.expand_path('../lib', __FILE__)
$LOAD_PATH.unshift(lib) unless $LOAD_PATH.include?(lib)
require 'open_xml/version'

Gem::Specification.new do |spec|
  spec.name          = "open_xml"
  spec.version       = OpenXml::VERSION
  spec.authors       = ["Carlos Espejo"]
  spec.email         = ["carlosespejo@gmail.com"]
  spec.description   = %q{Generate Word documents from a template, also handle html and images too.}
  spec.summary       = %q{A ruby library for generating word documents that can handle basic html and images too.}
  spec.homepage      = "https://github.com/CarlosEspejo/open_xml"
  spec.license       = "MIT"

  spec.files         = `git ls-files`.split($/)
  spec.executables   = spec.files.grep(%r{^bin/}) { |f| File.basename(f) }
  spec.test_files    = spec.files.grep(%r{^(test|spec|features)/})
  spec.require_paths = ["lib"]

  spec.add_development_dependency "bundler", "~> 1.3"
  spec.add_development_dependency "rake"
  spec.add_development_dependency "minitest"
  spec.add_development_dependency "pry"

  spec.add_runtime_dependency "nokogiri", "~> 1.6"
  spec.add_runtime_dependency "rubyzip", "~> 1.1"

end

