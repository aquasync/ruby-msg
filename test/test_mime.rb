#! /usr/bin/ruby -w

TEST_DIR = File.dirname __FILE__
$: << "#{TEST_DIR}/../lib"

require 'test/unit'
require 'mime'

class TestMime < Test::Unit::TestCase
	# test out the way it partitions a message into parts
	def test_parsing_no_multipart
		mime = Mime.new "Header1: Value1\r\nHeader2: Value2\r\n\r\nBody text."
		assert_equal ['Value1'], mime.headers['Header1']
		assert_equal 'Body text.', mime.body
		assert_equal false, mime.multipart?
		assert_equal nil, mime.parts
		# we get round trip conversion. this is mostly fluke, as orderedhash hasn't been
		# added yet
		assert_equal "Header1: Value1\r\nHeader2: Value2\r\n\r\nBody text.", mime.to_s
	end
end

