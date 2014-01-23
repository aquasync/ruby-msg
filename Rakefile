require 'rubygems'
require 'rake/testtask'

require 'rbconfig'
require 'fileutils'

spec = eval File.read('ruby-msg.gemspec')

task :default => [:test]

Rake::TestTask.new do |t|
	t.test_files = FileList["test/test_*.rb"] - ['test/test_pst.rb']
	t.warning = false
	t.verbose = true
end

begin
	Rake::TestTask.new(:coverage) do |t|
		t.test_files = FileList["test/test_*.rb"] - ['test/test_pst.rb']
		t.warning = false
		t.verbose = true
		t.ruby_opts = ['-rsimplecov -e "SimpleCov.start; load(ARGV.shift)"']
	end
rescue LoadError
	# SimpleCov not available
end

begin
	require 'rdoc/task'
	RDoc::Task.new do |t|
		t.rdoc_dir = 'doc'
		t.rdoc_files.include 'lib/**/*.rb'
		t.rdoc_files.include 'README', 'ChangeLog'
		t.title    = "#{PKG_NAME} documentation"
		t.options += %w[--line-numbers --inline-source --tab-width 2]
		t.main	   = 'README'
	end
rescue LoadError
	# RDoc not available or too old (<2.4.2)
end

begin
	require 'rubygems/package_task'
	Gem::PackageTask.new(spec) do |t|
		t.need_tar = true
		t.need_zip = false
		t.package_dir = 'build'
	end
rescue LoadError
	# RubyGems too old (<1.3.2)
end

