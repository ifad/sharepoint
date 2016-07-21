require 'dotenv'
require 'pathname'
require 'byebug'
require 'vcr'
require 'filemagic/ext'

Dotenv.load('.env')

SPEC_BASE = Pathname.new(__FILE__).realpath.parent

$: << SPEC_BASE.parent + 'lib'
require 'sharepoint'

def fixture name
  SPEC_BASE + 'fixtures' + name
end

# Requires supporting ruby files with custom matchers and macros, etc,
# # in spec/support/ and its subdirectories.
Dir[File.join(SPEC_BASE, "support/**/*.rb")].each { |f| require f }


VCR.configure do |c|
  c.cassette_library_dir = (SPEC_BASE + 'fixtures' + 'cassettes').to_s
  c.hook_into :webmock
  c.configure_rspec_metadata!
end

RSpec::configure do |rspec|
  rspec.tty = true
  rspec.color = true
end
