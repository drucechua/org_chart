# -*- encoding: utf-8 -*-
# stub: functions_framework 1.6.2 ruby lib

Gem::Specification.new do |s|
  s.name = "functions_framework".freeze
  s.version = "1.6.2"

  s.required_rubygems_version = Gem::Requirement.new(">= 0".freeze) if s.respond_to? :required_rubygems_version=
  s.metadata = { "bug_tracker_uri" => "https://github.com/GoogleCloudPlatform/functions-framework-ruby/issues", "changelog_uri" => "https://googlecloudplatform.github.io/functions-framework-ruby/v1.6.2/file.CHANGELOG.html", "documentation_uri" => "https://googlecloudplatform.github.io/functions-framework-ruby/v1.6.2", "source_code_uri" => "https://github.com/GoogleCloudPlatform/functions-framework-ruby" } if s.respond_to? :metadata=
  s.require_paths = ["lib".freeze]
  s.authors = ["Daniel Azuma".freeze]
  s.date = "1980-01-02"
  s.description = "The Functions Framework is an open source framework for writing lightweight, portable Ruby functions that run in a serverless environment. Functions written to this Framework will run on Google Cloud Functions, Google Cloud Run, or any other Knative-based environment.".freeze
  s.email = ["dazuma@google.com".freeze]
  s.executables = ["functions-framework".freeze, "functions-framework-ruby".freeze]
  s.files = ["bin/functions-framework".freeze, "bin/functions-framework-ruby".freeze]
  s.homepage = "https://github.com/GoogleCloudPlatform/functions-framework-ruby".freeze
  s.licenses = ["Apache-2.0".freeze]
  s.required_ruby_version = Gem::Requirement.new(">= 3.1.0".freeze)
  s.rubygems_version = "3.4.19".freeze
  s.summary = "Functions Framework for Ruby".freeze

  s.installed_by_version = "3.4.19" if s.respond_to? :installed_by_version

  s.specification_version = 4

  s.add_runtime_dependency(%q<cloud_events>.freeze, [">= 0.7.0", "< 2.a"])
  s.add_runtime_dependency(%q<puma>.freeze, [">= 4.3.0", "< 7.a"])
  s.add_runtime_dependency(%q<rack>.freeze, [">= 2.1", "< 4.a"])
end
