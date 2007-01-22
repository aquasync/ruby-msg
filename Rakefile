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

Rake::PackageTask.new(PKG_NAME, PKG_VERSION) do |p|
	p.need_tar_gz = true
	p.package_dir = 'build'
	p.package_files.include("Rakefile")
	p.package_files.include("contrib/rtfdecompr.c")
	p.package_files.include("test/*.rb", "test/*.doc", "lib/msg.rb", "lib/ole/storage.rb")
end

# not the right way of doing it. doesn't show up in rake --tasks, and can't
# attach description
desc 'Install files'
task :install do
	dest = Config::CONFIG['sitelibdir'] + '/ole'
	Dir.mkdir dest rescue nil
	FileUtils.copy './lib/storage.rb', dest + '/storage.rb'
end

spec = Gem::Specification.new do |s|
	s.name = PKG_NAME
	s.version = PKG_VERSION
	s.summary = %q{Ruby Msg library.}
	s.description = %q{A library for reading Outlook msg files, and for converting them to RFC2822 emails.}
	s.authors = ["Charles Lowe"]
	s.email = %q{aquasync@gmail.com}
	s.homepage = %q{http://code.google.com/p/ruby-msg}
	#s.rubyforge_project = %q{rubyntlm}

	s.files = Dir.glob('data/*.yaml')
	exe = RUBY_PLATFORM['win'] ? '.exe' : ''
	s.files += ['rtfdecompr' + exe, 'rtf2html' + exe]
	s.files += Dir.glob("lib/**/*.rb")
	
	s.has_rdoc = false

	s.autorequire = 'msg'
end

Rake::GemPackageTask.new(spec) do |p|
	p.gem_spec = spec
	p.need_tar = true
	p.need_zip = true
	p.package_dir = 'build'
end

