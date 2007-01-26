#! /usr/bin/ruby -w

TEST_DIR = File.dirname __FILE__
$: << "#{TEST_DIR}/../lib"

require 'test/unit'
require 'ole/storage'

#
# = TODO
#
# These tests could be a lot more complete.
#
class TestStorage < Test::Unit::TestCase
	def setup
		@ole = Ole::Storage.load open("#{TEST_DIR}/test_word_6.doc", 'rb')
	end

	def teardown
		@ole.io.close
	end

	def test_header
		# should have further header tests, testing the validation etc.
		assert_equal 18, @ole.header.to_a.length
		assert_equal 117, @ole.header.dirent_start
		assert_equal 1, @ole.header.num_bat
		assert_equal 1, @ole.header.num_sbat
		assert_equal 0, @ole.header.num_mbat
	end

	def test_fat
		# the fat block has all the numbers from 5..118 bar 117
		bbat_table = [112] + ((5..118).to_a - [112, 117])
		assert_equal bbat_table, @ole.bbat.table.reject { |i| i >= (1 << 32) - 3 }, 'bbat'
		sbat_table = (1..43).to_a - [2, 3]
		assert_equal sbat_table, @ole.sbat.table.reject { |i| i >= (1 << 32) - 3 }, 'sbat'
	end

	def test_directories
		assert_equal 5, @ole.dirs.length, 'have all directories'
		# a more complicated one would be good for this
		assert_equal 4, @ole.root.children.length, 'properly nested directories'
	end

	def test_utf16_conversion
		assert_equal 'Root Entry', @ole.root.name
		assert_equal 'WordDocument', @ole.root.children[2].name
	end

	def test_data
		# test the ole storage type
		compobj = @ole.root["\001CompObj"]
		assert_equal 'Microsoft Word 6.0-Dokument', compobj.data[/^.{32}([^\x00]+)/m, 1]
		# i was actually not loading data correctly before, so carefully check everything here
		hashes = [-482597081, 285782478, 134862598, -863988921]
		assert_equal hashes, @ole.root.children.map { |child| child.data.hash }

		# test accessing data using a block, and a io object. not sure how i want the interface to
		# work, but i may want to restrict to a block so that the io object can be cleaned up, as it
		# will be seeking around the parent io when in use. basically, something like this should
		# work:
		assert_equal hashes.first, @ole.root.children.first.io.read.hash, 'io for small files'
		assert_equal hashes.last, @ole.root.children.last.io.read.hash, 'io for big files'
		# this should let me later do things like
		# io = Base64::IO.new(open('blah', 'r'))
		# io.read 16 # get base64 data
		# io = Base64::IO.new(open('blah', 'w'))
		# io.write 'abcdef==' # stream base64 data into binary file
		# thusly, if the Mime object is patched to allow a part to be an io object as well as a string.
		# then finally, I could write Mime.to_io, which provides an io object that wraps the mime message.
		# it could have an attachment like open('some file') etc etc etc...
		# blocks wouldn't work, so seeking thing hmm
	end
end

class TestRangesIO < Test::Unit::TestCase
	def setup
		# why not :) ?
		# repeats too
		ranges = [100..200, 0..10, 100..150]
		@io = RangesIO.new open("#{TEST_DIR}/test_storage.rb"), ranges, :close_parent => true
	end

	def teardown
		@io.close
	end

	def test_basic
		assert_equal 160, @io.size
		# this will map to the start of the file:
		@io.pos = 100
		assert_equal '#! /usr/bi', @io.read(10)
	end

	# should test range_and_offset specifically

	def test_reading
		# test selection of initial range, offset within that range
		pos = 100
		@io.seek pos
		# test advancing of pos properly, by...
		chunked = (0...10).map { @io.read 10 }.join
		# given the file is 160 long:
		assert_equal 60, chunked.length
		@io.seek pos
		# comparing with a flat read
		assert_equal chunked, @io.read(60)
	end
end

