Gem::Specification.new do |spec|
  spec.name        = "docx_builder"
  spec.version     = "0.3.6"
  spec.date        = "2013-08-28"
  spec.summary     = "Generate Microsoft Word Office Open XML files"
  spec.description = "Generate and modify Word .docx files programatically"
  spec.authors     = ["Mike Gunderloy", "Mike Welham"]
  spec.email       = "MikeG1@larkfarm.com"
  spec.files       = `git ls-files`.split("\n")
  spec.homepage    = "https://github.com/ffmike/docx_builder"
  spec.add_dependency("nokogiri", ">= 1.5.2")
  spec.add_dependency("rmagick", ">= 2.12.2")
  spec.add_dependency("rubyzip", ">= 0.9.8")
  spec.add_development_dependency("equivalent-xml", ">= 0.2.9")
  spec.license = "MIT"
end
