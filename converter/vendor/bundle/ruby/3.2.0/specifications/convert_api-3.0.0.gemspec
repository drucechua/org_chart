# -*- encoding: utf-8 -*-
# stub: convert_api 3.0.0 ruby lib

Gem::Specification.new do |s|
  s.name = "convert_api".freeze
  s.version = "3.0.0"

  s.required_rubygems_version = Gem::Requirement.new(">= 0".freeze) if s.respond_to? :required_rubygems_version=
  s.require_paths = ["lib".freeze]
  s.authors = ["Tomas Rutkauskas".freeze]
  s.bindir = "exe".freeze
  s.date = "2024-09-14"
  s.description = "Convert various files like MS Word, Excel, PowerPoint, Images to PDF and Images. Create PDF and Images from url and raw HTML. Extract and create PowerPoint presentation from PDF. Merge, Encrypt, Split, Repair and Decrypt PDF files. All supported files conversions and manipulations can be found at https://www.convertapi.com/doc/supported-formats".freeze
  s.email = ["support@convertapi.com".freeze]
  s.homepage = "https://github.com/ConvertAPI/convertapi-ruby".freeze
  s.licenses = ["MIT".freeze]
  s.rubygems_version = "3.4.19".freeze
  s.summary = "ConvertAPI client library".freeze

  s.installed_by_version = "3.4.19" if s.respond_to? :installed_by_version

  s.specification_version = 4

  s.add_development_dependency(%q<bundler>.freeze, [">= 0"])
  s.add_development_dependency(%q<rake>.freeze, [">= 12.3.3"])
  s.add_development_dependency(%q<rspec>.freeze, ["~> 3.0"])
end
