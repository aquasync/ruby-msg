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
		# note that the way things are set up, you can't resolve body though
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

	# other things we could test - write support. duplicate keys

	def test_pseudo_properties
		# writeme
	end
end

