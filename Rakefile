require 'rake/rdoctask'
require 'rake/testtask'
require 'rake/packagetask'
require 'rake/gempackagetask'

require 'rbconfig'
require 'fileutils'

$: << './lib'
require 'msg.rb'

PKG_NAME = 'ruby-msg'
PKG_VERSION = Msg::VERSION

task :default => [:test]

Rake::TestTask.new(:test) do |t|
	t.test_files = FileList["test/test_*.rb"]
	t.warning = true
	t.verbose = true
end

=begin
Rake::PackageTask.new(PKG_NAME, PKG_VERSION) do |p|
	p.need_tar_gz = true
	p.package_dir = 'build'
	p.package_files.include("Rakefile", "README")
	p.package_files.include("contrib/*.c")
	p.package_files.include("test/test_*.rb", "test/*.doc", "lib/*.rb", "lib/ole/storage.rb")
end
=end

spec = Gem::Specification.new do |s|
	s.name = PKG_NAME
	s.version = PKG_VERSION
	s.summary = %q{Ruby Msg library.}
	s.description = %q{A library for reading Outlook msg files, and for converting them to RFC2822 emails.}
	s.authors = ["Charles Lowe"]
	s.email = %q{aquasync@gmail.com}
	s.homepage = %q{http://code.google.com/p/ruby-msg}
	#s.rubyforge_project = %q{ruby-msg}

	s.files = Dir.glob('data/*.yaml')
	exe = RUBY_PLATFORM['win'] ? '.exe' : ''
	# not great
	s.files += ['rtfdecompr' + exe, 'rtf2html' + exe]
	s.files += Dir.glob("lib/**/*.rb")
	s.files += Dir.glob("test/test_*.rb") + Dir.glob("test/*.doc")
	
	s.has_rdoc = true

	s.autorequire = 'msg'
end

Rake::GemPackageTask.new(spec) do |p|
	p.gem_spec = spec
	p.need_tar = true
	p.need_zip = true
	p.package_dir = 'build'
end

