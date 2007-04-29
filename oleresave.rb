#! /usr/bin/ruby

require 'lib/ole/storage'

srcfn, dstfn = ARGV

# new version:
Ole::Storage.open srcfn, 'r' do |src|
	Ole::Storage.open dstfn, 'w+' do |dst|
		Ole::Storage::Dirent.copy src.root, dst.root
	end
end

