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
		compobj = @ole.root.children.find { |child| child.name == "\001CompObj" }
		assert_equal 'Microsoft Word 6.0-Dokument', compobj.data[/^.{32}([^\x00]+)/m, 1]
		# i was actually not loading data correctly before, so carefully check everything here
		hashes = [-482597081, 285782478, 134862598, -863988921]
		assert_equal hashes, @ole.root.children.map { |child| child.data.hash }
	end
end

