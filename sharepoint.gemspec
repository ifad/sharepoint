Gem::Specification.new do |gem|
  gem.name           = 'sharepoint'
  gem.version        = '0.1.0'
  gem.authors        = [ 'Antonio Delfin' ]
  gem.email          = [ 'a.delfin@ifad.org' ]
  gem.description    = %q(Ruby client to consume Sharepoint services)
  gem.summary        = %q(Ruby client to consume Sharepoint services)
  gem.homepage       = "https://github.com/ifad/sharepoint"

  gem.files             = `git ls-files`.split("\n")
  gem.require_paths     = ["lib"]

  gem.required_ruby_version = '>= 2.3'

  gem.add_dependency 'ethon'
  gem.add_dependency 'activesupport', '>= 4.0'
  gem.add_dependency 'addressable'

  gem.add_development_dependency 'rake'
  gem.add_development_dependency 'rspec'
  gem.add_development_dependency 'dotenv'
  gem.add_development_dependency 'webmock'
  gem.add_development_dependency 'byebug'
  gem.add_development_dependency 'ruby-filemagic'
end
