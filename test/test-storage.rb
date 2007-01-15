#! /usr/bin/ruby -w

require 'test/unit'

Dir.chdir File.dirname(__FILE__)
require './../lib/storage'

class TestStorage < Test::Unit::TestCase
	def setup
		@ole = Ole::Storage.new open('test-word-6.doc', 'rb')
	end

	def teardown
		@ole.io.close
	end

	def test_header
		# num_fat_blocks, root_start_block, unk2, unk3, dir_flag,
		# unk4, fat_next_block, num_extra_fat_blocks
		assert_equal [1, 117, 0, 4096, 2, 1, 4294967294, 0], @ole.header.to_a[2..-1]
	end

	def test_blocks
		# test getting the -1 block. blocks are -1 based, with that first one being the header block
		assert_equal @ole.blocks[-1][0...Ole::Storage::MAGIC.length], Ole::Storage::MAGIC, 'magic test'
		# other than that block, we have 119 in this file
		assert_equal 119, @ole.blocks.length, 'num blocks'
		# test loading of block data
		assert_equal Ole::Storage::Blocks::BLOCK_SIZE, @ole.blocks[0].length, 'block load'
		# 4 of the 5 directories are in the root_start_block
		assert_equal 4, @ole.blocks[117].dirs.length, 'dirs in block'
	end

	def test_fat
		# must confess i don't really understand the fat stuff
		# there is only one fat block in this file
		assert_equal [0], @ole.fat.reject { |i| i == (1 << 32) - 1 }, 'fat pointer table'
		# the fat block has all the numbers from 5..118 bar 117
		assert_equal((5..118).to_a - [117], @ole.blocks.get_fat_block(0).
			reject { |i| i >= (1 << 32) - 3 }.sort, 'fat block')
	end

	def test_directories
		assert_equal 5, @ole.dirs.length, 'have all directories'
		# a more complicated one would be good for this
		assert_equal 4, @ole.root.children.length, 'properly nested directories'
	end

	def test_utf16_conversion
		assert_equal 'Root Entry', @ole.root.name
	end
end

