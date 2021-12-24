#! /usr/bin/ruby -w

TEST_DIR = File.dirname __FILE__
$: << "#{TEST_DIR}/../lib"

require 'minitest/autorun'
require 'mapi/pst'

class TestPst < Minitest::Test
	def test
		load_pst "#{TEST_DIR}/pst/attachAndInline.pst"
		load_pst "#{TEST_DIR}/pst/msgInMsg.pst"
		load_pst "#{TEST_DIR}/pst/Outlook97-2002.pst"
		load_pst "#{TEST_DIR}/pst/Outlook2003.pst"
		load_pst "#{TEST_DIR}/pst/200 recipients.pst"
	end

	def load_pst filename
		count = 0
		open filename do |io|
			pst = Mapi::Pst.new io
			pst.to_a
		end
		printf("%d messages\n", count)
	end
end
