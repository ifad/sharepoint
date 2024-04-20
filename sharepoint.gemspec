# frozen_string_literal: true

Gem::Specification.new do |gem|
  gem.name           = 'sharepoint'
  gem.version        = '0.1.0'
  gem.authors        = ['Antonio Delfin']
  gem.email          = ['a.delfin@ifad.org']
  gem.description    = 'Ruby client to consume Sharepoint services'
  gem.summary        = 'Ruby client to consume Sharepoint services'
  gem.homepage       = 'https://github.com/ifad/sharepoint'

  gem.files          = Dir.glob('{LICENSE,README.md,lib/**/*.rb}', File::FNM_DOTMATCH)
  gem.require_paths  = ['lib']

  gem.required_ruby_version = '>= 2.3'

  gem.add_dependency 'activesupport', '>= 4.0'
  gem.add_dependency 'ethon'

  gem.metadata['rubygems_mfa_required'] = 'true'
end
