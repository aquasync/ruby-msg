require 'test/unit'

$:.unshift File.dirname(__FILE__) + '/../lib'
require 'mapi/property_set'

class TestMapiPropertySet < Test::Unit::TestCase
	include Mapi

	def test_constants
		assert_equal '00020328-0000-0000-c000-000000000046', PropertySet::PS_MAPI.format
	end

	def test_lookup
		guid = Ole::Types::Clsid.parse '00020328-0000-0000-c000-000000000046'
		assert_equal 'PS_MAPI', PropertySet::NAMES[guid]
	end

	def test_simple_key
		key = PropertySet::Key.new 0x0037
		assert_equal PropertySet::PS_MAPI, key.guid
		hash = {key => 'hash lookup'}
		assert_equal 'hash lookup', hash[PropertySet::Key.new(0x0037)]
		assert_equal '0x0037', key.inspect
		assert_equal :subject, key.to_sym
	end

	def test_complex_keys
		key = PropertySet::Key.new 'Keywords', PropertySet::PS_PUBLIC_STRINGS
		# note that the inspect string now uses symbolic guids
		assert_equal '#<Key PS_PUBLIC_STRINGS/"Keywords">', key.inspect
		# note that this isn't categories
		assert_equal :keywords, key.to_sym
		custom_guid = '00020328-0000-0000-c000-deadbeefcafe'
		key = PropertySet::Key.new 0x8000, Ole::Types::Clsid.parse(custom_guid)
		assert_equal "#<Key {#{custom_guid}}/0x8000>",  key.inspect
		key = PropertySet::Key.new 0x8005, PropertySet::PSETID_Address
		assert_equal 'file_under', key.to_s
	end

	def test_property_set_basics
		# the propertystore can be mocked with a hash:
		store = {
			PropertySet::Key.new(0x0037) => 'the subject',
			PropertySet::Key.new('Keywords', PropertySet::PS_PUBLIC_STRINGS) => ['some keywords'],
			PropertySet::Key.new(0x8888) => 'un-mapped value'
		}
		props = PropertySet.new store
		# can resolve subject
		assert_equal PropertySet::Key.new(0x0037), props.resolve('subject')
		# note that the way things are set up, you can't resolve body though. ie, only
		# existent (not all-known) properties resolve. maybe this should be changed. it'll
		# need to be, for <tt>props.body=</tt> to work as it should.
		assert_equal nil, props.resolve('body')
		assert_equal 'the subject', props.subject
		assert_equal ['some keywords'], props.keywords
		# other access methods
		assert_equal 'the subject', props['subject']
		assert_equal 'the subject', props[0x0037]
		assert_equal 'the subject', props[0x0037, PropertySet::PS_MAPI]
		# note that the store is accessible directly, as #raw currently (maybe i should rename)
		assert_equal store, props.raw
		# note that currently, props.each / props.to_h works with the symbolically
		# mapped properties, so the above un-mapped value won't be in the list:
		assert_equal({:subject => 'the subject', :keywords => ['some keywords']}, props.to_h)
		assert_equal [:keywords, :subject], props.keys.sort_by(&:to_s)
		assert_equal [['some keywords'], 'the subject'], props.values.sort_by(&:to_s)
	end

	# other things we could test - write support. duplicate key handling

	def test_pseudo_properties
		# load some compressed rtf data
		data = File.read File.dirname(__FILE__) + '/test_rtf.data'
		props = PropertySet.new PropertySet::Key.new(0x1009) => StringIO.new(data)
		# all these get generated from the rtf. still need tests for the way the priorities work
		# here, and also the html embedded in rtf stuff....
		assert_equal((<<-'end').chomp.gsub(/\n/, "\n\r"), props.body_rtf)
{\rtf1\ansi\ansicpg1252\fromtext \deff0{\fonttbl
{\f0\fswiss Arial;}
{\f1\fmodern Courier New;}
{\f2\fnil\fcharset2 Symbol;}
{\f3\fmodern\fcharset0 Courier New;}}
{\colortbl\red0\green0\blue0;\red0\green0\blue255;}
\uc1\pard\plain\deftab360 \f0\fs20 \par
I will be out of the office starting  15.02.2007 and will not return until\par
27.02.2007.\par
\par
I will respond to your message when I return. For urgent enquiries please\par
contact Motherine Jacson.\par
\par
}
		end
		assert_equal <<-'end', props.body_html
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
		end
		assert_equal <<-'end', props.body

I will be out of the office starting  15.02.2007 and will not return until
27.02.2007.

I will respond to your message when I return. For urgent enquiries please
contact Motherine Jacson.

		end
	end
end

