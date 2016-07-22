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

  gem.add_dependency("ethon")
  gem.add_development_dependency 'rspec'
  gem.add_development_dependency 'dotenv'
  gem.add_development_dependency 'webmock'
  gem.add_development_dependency 'byebug'
  gem.add_development_dependency 'ruby-filemagic'
end
