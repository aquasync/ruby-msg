require 'rake/rdoctask'
require 'rake/testtask'
require 'rake/packagetask'
require 'rake/gempackagetask'

require 'rbconfig'
require 'fileutils'

require './lib/storage.rb'

PKG_NAME = 'ruby-msg'
PKG_VERSION = Ole::Storage::VERSION

task :default => [:test]

Rake::TestTask.new(:test) do |t|
  t.test_files = FileList["test/*.rb"]
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
task :install do
	dest = Config::CONFIG['sitelibdir'] + '/ole'
	Dir.mkdir dest rescue nil
	FileUtils.copy './lib/storage.rb', dest + '/storage.rb'
end

