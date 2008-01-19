require 'rake/rdoctask'
require 'rake/testtask'
require 'rake/packagetask'
require 'rake/gempackagetask'

require 'rbconfig'
require 'fileutils'

$:.unshift 'lib'

require 'mapi/msg'

PKG_NAME = 'ruby-msg'
PKG_VERSION = Mapi::Msg::VERSION

task :default => [:test]

Rake::TestTask.new(:test) do |t|
	t.test_files = FileList["test/test_*.rb"] - ['test/test_pst.rb']
	t.warning = false
	t.verbose = true
end

Rake::RDocTask.new do |t|
	t.rdoc_dir = 'doc'
	t.title    = "#{PKG_NAME} documentation"
	t.options += %w[--main README --line-numbers --inline-source --tab-width 2]
	t.rdoc_files.include 'lib/**/*.rb'
	t.rdoc_files.include 'README'
end

spec = Gem::Specification.new do |s|
	s.name = PKG_NAME
	s.version = PKG_VERSION
	s.summary = %q{Ruby Msg library.}
	s.description = %q{A library for reading Outlook msg files, and for converting them to RFC2822 emails.}
	s.authors = ["Charles Lowe"]
	s.email = %q{aquasync@gmail.com}
	s.homepage = %q{http://code.google.com/p/ruby-msg}
	#s.rubyforge_project = %q{ruby-msg}

	s.executables = ['msgtool']
	s.files  = Dir.glob('data/*.yaml') + ['Rakefile', 'README', 'FIXES']
	s.files += Dir.glob("lib/**/*.rb")
	s.files += Dir.glob("test/test_*.rb")
	s.files += Dir.glob("bin/*")
	
	s.has_rdoc = true
	s.rdoc_options += ['--main', 'README',
					   '--title', "#{PKG_NAME} documentation",
					   '--tab-width', '2']


	s.autorequire = 'msg'

	s.add_dependency 'ruby-ole', '>=1.2.3'
end

Rake::GemPackageTask.new(spec) do |p|
	p.gem_spec = spec
	p.need_tar = false #true
	p.need_zip = false
	p.package_dir = 'build'
end

