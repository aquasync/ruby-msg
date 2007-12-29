require 'msg'

=begin

Mapi::PropertyStore
Mapi::Item
Mapi::Attachment
Mapi::Recipient

Pst::PropertyStore < Mapi::PropertyStore
Msg::PropertyStore < Mapi::PropertyStore

# simple hash like api. can be lazily loaded. however if you
# use something like #keys, or #values, then you'd probably
# just cache things. maybe this would be the raw interface, 
# which is wrapped by something which provides named access.
class Mapi::PropertyStore
	def initialize symbolic=true
		@symbolic = symbolic
	end

	def raw
		# gives access to the real property store object. the top level
		# wrapped one uses symbolic naming. 
		self.class.new false
	end

	def each symbolic=@symbolic
		# iterate through the key value pairs. keys are
		# Mapi::PropertyStore::Key instances, which allows for named
		# properties. values are of the appropriate class (which is
		# what?). 
	end

	def []
		# get a value based on a key
	end

	def to_h symbolic=@symbolic
	end
end

=end

module Mapi
	module Types
		#
		# Mapi property types, taken from http://msdn2.microsoft.com/en-us/library/bb147591.aspx.
		#
		# The fields are [mapi name, variant name, description]
		#
		# seen some synonyms here, like PT_I8 vs PT_LONG. seen stuff like PT_SRESTRICTION, not
		# sure what that is. look at `grep ' PT_' data/mapitags.yaml  | sort -u`
		# also, it has stuff like PT_MV_BINARY, where _MV_ probably means multi value, and is
		# likely just defined to | in 0x1000.
		#
		# Note that the last 2 are the only ones where the Mapi value differs from the Variant value
		# for the corresponding variant type. Odd. Also, the last 2 are currently commented out here
		# because of the clash.
		#
		# Note 2 - the strings here say VT_BSTR, but I don't have that defined in Ole::Types. Should
		# maybe change them to match. I've also seen reference to PT_TSTRING, which is defined as some
		# sort of get unicode first, and fallback to ansii or something.
		#
		DATA = {
			0x0001 => ['PT_NULL', 'VT_NULL', 'Null (no valid data)'],
			0x0002 => ['PT_SHORT', 'VT_I2', '2-byte integer (signed)'],
			0x0003 => ['PT_LONG', 'VT_I4', '4-byte integer (signed)'],
			0x0004 => ['PT_FLOAT', 'VT_R4', '4-byte real (floating point)'],
			0x0005 => ['PT_DOUBLE', 'VT_R8', '8-byte real (floating point)'],
			0x0006 => ['PT_CURRENCY', 'VT_CY', '8-byte integer (scaled by 10,000)'],
			0x000a => ['PT_ERROR', 'VT_ERROR', 'SCODE value; 32-bit unsigned integer'],
			0x000b => ['PT_BOOLEAN', 'VT_BOOL', 'Boolean'],
			0x000d => ['PT_OBJECT', 'VT_UNKNOWN', 'Data object'],
			0x001e => ['PT_STRING8', 'VT_BSTR', 'String'],
			0x001f => ['PT_UNICODE', 'VT_BSTR', 'String'],
			0x0040 => ['PT_SYSTIME', 'VT_DATE', '8-byte real (date in integer, time in fraction)'],
			#0x0102 => ['PT_BINARY', 'VT_BLOB', 'Binary (unknown format)'],
			#0x0102 => ['PT_CLSID', 'VT_CLSID', 'OLE GUID']
		}

		module Constants
			DATA.each { |num, (mapi_name, variant_name, desc)| const_set mapi_name, num }
		end

		include Constants
	end

	# FIXME this stuff below.
	# restruct msg project to share more code. eg, perhaps something like:
	# of course this is inverted at the moment, but msg will be changed transparently to pst
	# to fix that. 

	class PropertyStore < Msg::Properties
		TAGS = MAPITAGS
	end

	# IMessage essentially, but there's also stuff like IMAPIFolder etc. so, for this to form
	# basis for PST Item, it'd need to be more general.
	class Item < Msg
		# IAttach
		class Attachment < Msg::Attachment
		end


		class Recipient < Recipient
		end

		# +props+ should be a PropertyStore object.
		def initialize props
			@properties = props
			@mime = Mime.new props.transport_message_headers.to_s, true

			# hack
			@root = OpenStruct.new(:ole => OpenStruct.new(:dirents => [OpenStruct.new(:time => nil)]))
			populate_headers
		end
	end
end

