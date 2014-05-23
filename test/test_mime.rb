#! /usr/bin/ruby -w

$: << File.dirname(__FILE__) + '/../lib'

require 'test/unit'
require 'mapi/mime'

class TestMime < Test::Unit::TestCase
	# test out the way it partitions a message into parts
	def test_parsing_no_multipart
		mime = Mapi::Mime.new "Header1: Value1\r\nHeader2: Value2\r\n\r\nBody text."
		assert_equal ['Value1'], mime.headers['Header1']
		assert_equal 'Body text.', mime.body
		assert_equal false, mime.multipart?
		assert_equal nil, mime.parts
		assert_equal "Header1: Value1\r\nHeader2: Value2\r\n\r\nBody text.", mime.to_s
	end
	
	def test_boundaries
		assert_match(/^----_=_NextPart_001_/, Mapi::Mime.make_boundary(1))
	end
end

