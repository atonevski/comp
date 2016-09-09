# Ensure we require the local version and not one we might have installed already
require File.join([File.dirname(__FILE__),'lib','comp','version.rb'])
spec = Gem::Specification.new do |s| 
  s.name = 'comp'
  s.version = Comp::VERSION
  s.author = 'Your Name Here'
  s.email = 'your@email.address.com'
  s.homepage = 'http://your.website.com'
  s.platform = Gem::Platform::RUBY
  s.summary = 'A description of your project'
  s.files = `git ls-files`.split("
")
  s.require_paths << 'lib'
  s.has_rdoc = true
  s.extra_rdoc_files = ['README.rdoc','comp.rdoc']
  s.rdoc_options << '--title' << 'comp' << '--main' << 'README.rdoc' << '-ri'
  s.bindir = 'bin'
  s.executables << 'comp'
  s.add_development_dependency('rake')
  s.add_development_dependency('rdoc')
  s.add_development_dependency('aruba')
  s.add_runtime_dependency('gli','2.13.4')

  s.add_dependency 'activerecord'
  s.add_dependency 'sqlite3'
  s.add_dependency 'spreadsheet'
  s.add_dependency 'writeexcel'
  s.add_dependency 'axlsx'
end
