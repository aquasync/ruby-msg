require 'test/unit'

$:.unshift File.dirname(__FILE__) + '/../lib'
require 'mapi'
require 'mapi/convert'

class TestMapiPropertySet < Test::Unit::TestCase
	include Mapi

	def test_contact_from_property_hash
		make_key1 = proc { |id| PropertySet::Key.new id }
		make_key2 = proc { |id| PropertySet::Key.new id, PropertySet::PSETID_Address }
		store = {
			make_key1[0x001a] => 'IPM.Contact',
			make_key1[0x0037] => 'full name',
			make_key1[0x3a06] => 'given name',
			make_key1[0x3a08] => 'business telephone number',
			make_key1[0x3a11] => 'surname',
			make_key1[0x3a15] => 'postal address',
			make_key1[0x3a16] => 'company name',
			make_key1[0x3a17] => 'title',
			make_key1[0x3a18] => 'department name',
			make_key1[0x3a19] => 'office location',
			make_key2[0x8005] => 'file under',
			make_key2[0x801b] => 'business address',
			make_key2[0x802b] => 'web page',
			make_key2[0x8045] => 'business address street',
			make_key2[0x8046] => 'business address city',
			make_key2[0x8047] => 'business address state',
			make_key2[0x8048] => 'business address postal code',
			make_key2[0x8049] => 'business address country',
			make_key2[0x804a] => 'business address post office box',
			make_key2[0x8062] => 'im address',
			make_key2[0x8082] => 'SMTP',
			make_key2[0x8083] => 'email@address.com'
		}
		props = PropertySet.new store
		message = Message.new props
		assert_equal 'text/x-vcard', message.mime_type
		vcard = message.to_vcard
		assert_equal Vpim::Vcard, vcard.class
		assert_equal <<-'end', vcard.to_s
BEGIN:VCARD
VERSION:3.0
N:surname;given name;;;
FN:full name
ADR;TYPE=work:;;business address street;business address city\, business ad
 dress state;;;
X-EVOLUTION-FILE-AS:file under
EMAIL:email@address.com
ORG:company name
END:VCARD
		end
	end

	def test_contact_from_msg
		# load some msg contacts and convert them...
	end
end

