require 'minitest/autorun'

$:.unshift File.dirname(__FILE__) + '/../lib'
require 'mapi/types'

class TestMapiTypes < Minitest::Test
	include Mapi

	def test_constants
		assert_equal 3, Types::PT_LONG
	end

	def test_lookup
		assert_equal 'PT_LONG', Types::DATA[3].first
	end
end

