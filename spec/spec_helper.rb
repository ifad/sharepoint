# frozen_string_literal: true

if ENV['RCOV'] || ENV['COVERAGE']
  require 'simplecov'

  SimpleCov.start do
    add_filter '/spec/'

    track_files 'lib/**/*.rb'
  end
end

require 'dotenv'
require 'pathname'
require 'byebug'
require 'filemagic/ext'

Dotenv.load('.env')

SPEC_BASE = Pathname.new(__FILE__).realpath.parent

$LOAD_PATH << ("#{SPEC_BASE.parent}lib")
require 'sharepoint'

def fixture(name)
  "#{SPEC_BASE}fixtures#{name}"
end

# Requires supporting ruby files with custom matchers and macros, etc,
# # in spec/support/ and its subdirectories.
Dir[File.join(SPEC_BASE, 'support/**/*.rb')].sort.each { |f| require f }

RSpec.configure do |rspec|
  rspec.tty = true
  rspec.color = true
  rspec.include Sharepoint::SpecHelpers
end
