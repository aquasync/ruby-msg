#! /usr/bin/ruby -w

TEST_DIR = File.dirname __FILE__
$: << "#{TEST_DIR}/../lib"

require 'minitest/autorun'
require 'mapi/pst'

class TestPst < Minitest::Test
	def test_attachAndInline
		load_pst "#{TEST_DIR}/pst/attachAndInline.pst"
	end

	def test_msgInMsg
		load_pst "#{TEST_DIR}/pst/msgInMsg.pst"
	end

	def test_outlook97_2002
		load_pst "#{TEST_DIR}/pst/Outlook97-2002.pst"
	end

	def test_outlook2003
		load_pst "#{TEST_DIR}/pst/Outlook2003.pst"
	end

	def load_pst filename
		count = 0
		open filename do |io|
			pst = Mapi::Pst.new io
			pst.each do |message|
				count += 1
			end
		end
		printf("%d messages\n", count)
	end
end
