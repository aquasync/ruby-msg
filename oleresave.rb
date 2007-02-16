#! /usr/bin/ruby

require 'lib/ole/storage'
require 'lib/ole/write_support'

srcfn, dstfn = ARGV

=begin
ole version:

src = Ole::Storage.load open(srcfn)
open dstfn, 'w' do |dst|
	src.save dst
end

=end

# new version:
Ole::Storage.open srcfn, 'r' do |src|
	Ole::Storage.open dstfn, 'w+' do |dst|
		Ole::Storage::Dirent.copy src.root, dst.root
	end
end

