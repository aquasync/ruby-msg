require 'test/unit'

$:.unshift File.dirname(__FILE__) + '/../lib'
require 'mapi'
require 'mapi/convert'

class TestMapiPropertySet < Test::Unit::TestCase
	include Mapi

	def test_using_pseudo_properties
		# load some compressed rtf data
		data = File.read File.dirname(__FILE__) + '/test_rtf.data'
		store = {
			PropertySet::Key.new(0x0037) => 'Subject',
			PropertySet::Key.new(0x0c1e) => 'SMTP',
			PropertySet::Key.new(0x0c1f) => 'sender@email.com',
			PropertySet::Key.new(0x1009) => StringIO.new(data)
		}
		props = PropertySet.new store 
		msg = Message.new props
		def msg.attachments
			[]
		end
		def msg.recipients
			[]
		end
		# the ignoring of \r here should change. its actually not output consistently currently.
		assert_equal((<<-end), msg.to_mime.to_s.gsub(/NextPart[_0-9a-z\.]+/, 'NextPart_XXX').delete("\r"))
From: sender@email.com
Subject: Subject
Content-Type: multipart/alternative; boundary="----_=_NextPart_XXX"

This is a multi-part message in MIME format.

------_=_NextPart_XXX
Content-Type: text/plain


I will be out of the office starting  15.02.2007 and will not return until
27.02.2007.

I will respond to your message when I return. For urgent enquiries please
contact Motherine Jacson.



------_=_NextPart_XXX
Content-Type: text/html

<html>
<body>
<br>I will be out of the office starting  15.02.2007 and will not return until
<br>27.02.2007.
<br>
<br>I will respond to your message when I return. For urgent enquiries please
<br>contact Motherine Jacson.
<br>
<br></body>
</html>


------_=_NextPart_XXX--
		end
	end
end

