#! /usr/bin/ruby

TEST_DIR = File.dirname __FILE__
$: << "#{TEST_DIR}/../lib"

require 'test/unit'
require 'msg'

class TestMsg < Test::Unit::TestCase
	def test_blammo
		Msg.open "#{TEST_DIR}/test_Blammo.msg" do |msg|
			assert_equal '"TripleNickel" <TripleNickel@mapi32.net>', msg.from
			assert_equal 'BlammoBlammo', msg.subject
			assert_equal 0, msg.recipients.length
			assert_equal 0, msg.attachments.length
			# this is all properties
			assert_equal 66, msg.properties.raw.length
			# this is unique named properties
			assert_equal 48, msg.properties.to_h.length
			# get the named property keys
			keys = msg.properties.raw.keys.select { |key| String === key.code }
			assert_equal '55555555-5555-5555-c000-000000000046', keys[0].guid.format
			assert_equal 'Yippee555', msg.properties[keys[0]]
			assert_equal '66666666-6666-6666-c000-000000000046', keys[1].guid.format
			assert_equal 'Yippee666', msg.properties[keys[1]]
		end
	end
end

