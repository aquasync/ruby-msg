#! /usr/bin/ruby -w

TEST_DIR = File.dirname __FILE__
$: << "#{TEST_DIR}/../lib"

require 'test/unit'
require 'ole/storage'

# just for the new api segment
require 'ole/write_support'

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
		assert_equal 17,  @ole.header.to_a.length
		assert_equal 117, @ole.header.dirent_start
		assert_equal 1,   @ole.header.num_bat
		assert_equal 1,   @ole.header.num_sbat
		assert_equal 0,   @ole.header.num_mbat
	end

	def test_fat
		# the fat block has all the numbers from 5..118 bar 117
		bbat_table = [112] + ((5..118).to_a - [112, 117])
		assert_equal bbat_table, @ole.bbat.table.reject { |i| i >= (1 << 32) - 3 }, 'bbat'
		sbat_table = (1..43).to_a - [2, 3]
		assert_equal sbat_table, @ole.sbat.table.reject { |i| i >= (1 << 32) - 3 }, 'sbat'
	end

	def test_directories
		assert_equal 5, @ole.dirents.length, 'have all directories'
		# a more complicated one would be good for this
		assert_equal 4, @ole.root.children.length, 'properly nested directories'
	end

	def test_utf16_conversion
		assert_equal 'Root Entry', @ole.root.name
		assert_equal 'WordDocument', @ole.root.children[2].name
	end

	def test_data
		# test the ole storage type
		type = 'Microsoft Word 6.0-Dokument'
		assert_equal type, @ole.root["\001CompObj"].read[/^.{32}([^\x00]+)/m, 1]
		# i was actually not loading data correctly before, so carefully check everything here
		hashes = [-482597081, 285782478, 134862598, -863988921]
		assert_equal hashes, @ole.root.children.map { |child| child.read.hash }
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

