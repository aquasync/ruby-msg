$:.unshift File.dirname(__FILE__) + '/lib'
require 'mapi/version'

PKG_NAME = 'ruby-msg'
PKG_VERSION = Mapi::VERSION

Gem::Specification.new do |s|
	s.name = PKG_NAME
	s.version = PKG_VERSION
	s.summary = %q{Ruby Msg library.}
	s.description = %q{A library for reading and converting Outlook msg and pst files (mapi message stores).}
	s.authors = ['Charles Lowe']
	s.email = %q{aquasync@gmail.com}
	s.homepage = %q{http://code.google.com/p/ruby-msg}
	s.rubyforge_project = %q{ruby-msg}

	s.executables = ['mapitool']
	s.files  = ['README', 'COPYING', 'Rakefile', 'ChangeLog', 'ruby-msg.gemspec']
	s.files += Dir.glob('data/*.yaml')
	s.files += Dir.glob('lib/**/*.rb')
	s.files += Dir.glob('test/test_*.rb')
	s.files += Dir.glob('bin/*')

	s.has_rdoc = true
	s.extra_rdoc_files = ['README', 'ChangeLog']
	s.rdoc_options += [
		'--main', 'README',
		'--title', "#{PKG_NAME} documentation",
		'--tab-width', '2'
	]

	s.add_dependency 'ruby-ole', '>=1.2.8'
	s.add_dependency 'vpim', '>=0.360'
end

