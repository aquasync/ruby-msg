require 'rubygems'
require 'yaml'
require 'ole/types'

module Mapi
	Log = Logger.new_with_callstack

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

	#
	# The Mapi::PropertySet class is used to wrap the lower level Msg or Pst property stores,
	# and provide a consistent and more friendly interface. It allows you to just say:
	#
	#   properties.subject
	#
	# instead of:
	#
	#   properites.raw[0x0037, PS_MAPI]
	#
	# The underlying store can be just a hash, or lazily loading directly from the file. A good
	# compromise is to cache all the available keys, and just return the values on demand, rather
	# than load up many possibly unwanted values.
	#
	class PropertySet
		# the property set guid constants
		# these guids are all defined with the macro DEFINE_OLEGUID in mapiguid.h.
		# see http://doc.ddart.net/msdn/header/include/mapiguid.h.html
		oleguid = proc do |prefix|
			Ole::Types::Clsid.parse "{#{prefix}-0000-0000-c000-000000000046}"
		end

		NAMES = {
			oleguid['00020328'] => 'PS_MAPI',
			oleguid['00020329'] => 'PS_PUBLIC_STRINGS',
			oleguid['00020380'] => 'PS_ROUTING_EMAIL_ADDRESSES',
			oleguid['00020381'] => 'PS_ROUTING_ADDRTYPE',
			oleguid['00020382'] => 'PS_ROUTING_DISPLAY_NAME',
			oleguid['00020383'] => 'PS_ROUTING_ENTRYID',
			oleguid['00020384'] => 'PS_ROUTING_SEARCH_KEY',
			# string properties in this namespace automatically get added to the internet headers
			oleguid['00020386'] => 'PS_INTERNET_HEADERS',
			# theres are bunch of outlook ones i think
			# http://blogs.msdn.com/stephen_griffin/archive/2006/05/10/outlook-2007-beta-documentation-notification-based-indexing-support.aspx
			# IPM.Appointment
			oleguid['00062002'] => 'PSETID_Appointment',
			# IPM.Task
			oleguid['00062003'] => 'PSETID_Task',
			# used for IPM.Contact
			oleguid['00062004'] => 'PSETID_Address',
			oleguid['00062008'] => 'PSETID_Common',
			# didn't find a source for this name. it is for IPM.StickyNote
			oleguid['0006200e'] => 'PSETID_Note',
			# for IPM.Activity. also called the journal?
			oleguid['0006200a'] => 'PSETID_Log',
		}

		module Constants
			NAMES.each { |guid, name| const_set name, guid }
		end

		include Constants

		# +Properties+ are accessed by <tt>Key</tt>s, which are coerced to this class.
		# Includes a bunch of methods (hash, ==, eql?) to allow it to work as a key in
		# a +Hash+.
		#
		# Also contains the code that maps keys to symbolic names.
		class Key
			include Constants

			attr_reader :code, :guid
			def initialize code, guid=PS_MAPI
				@code, @guid = code, guid
			end

			def to_sym
				# hmmm, for some stuff, like, eg, the message class specific range, sym-ification
				# of the key depends on knowing our message class. i don't want to store anything else
				# here though, so if that kind of thing is needed, it can be passed to this function.
				# worry about that when some examples arise.
				case code
				when Integer
					if guid == PS_MAPI # and < 0x8000 ?
						# the hash should be updated now that i've changed the process
						TAGS['%04x' % code].first[/_(.*)/, 1].downcase.to_sym rescue code
					else
						# handle other guids here, like mapping names to outlook properties, based on the
						# outlook object model.
						NAMED_MAP[self].to_sym rescue code
					end
				when String
					# return something like
					# note that named properties don't go through the map at the moment. so #categories
					# doesn't work yet
					code.downcase.to_sym
				end
			end
			
			def to_s
				to_sym.to_s
			end

			# FIXME implement these
			def transmittable?
				# etc, can go here too
			end

			# this stuff is to allow it to be a useful key
			def hash
				[code, guid].hash
			end

			def == other
				hash == other.hash
			end

			alias eql? :==

			def inspect
				# maybe the way to do this, would be to be able to register guids
				# in a global lookup, which are used by Clsid#inspect itself, to
				# provide symbolic names...
				guid_str = NAMES[guid] || "{#{guid.format}}"
				if Integer === code
					hex = '0x%04x' % code
					if guid == PS_MAPI
						# just display as plain hex number
						hex
					else
						"#<Key #{guid_str}/#{hex}>"
					end
				else
					# display full guid and code
					"#<Key #{guid_str}/#{code.inspect}>"
				end
			end
		end

		# duplicated here for now
		SUPPORT_DIR = File.dirname(__FILE__) + '/..'

		# data files that provide for the code to symbolic name mapping
		# guids in named_map are really constant references to the above
		TAGS = YAML.load_file "#{SUPPORT_DIR}/data/mapitags.yaml"
		NAMED_MAP = YAML.load_file("#{SUPPORT_DIR}/data/named_map.yaml").inject({}) do |hash, (key, value)|
			hash.update Key.new(key[0], const_get(key[1])) => value
		end

		attr_reader :raw
	
		# +raw+ should be an hash-like object that maps <tt>Key</tt>s to values. Should respond_to?
		# [], keys, values, each, and optionally []=, and delete.
		def initialize raw
			@raw = raw
		end

		# resolve +arg+ (could be key, code, string, or symbol), and possible +guid+ to a key.
		# returns nil on failure
		def resolve arg, guid=nil
			if guid;        Key.new arg, guid
			else
				case arg
				when Key;     arg
				when Integer; Key.new arg
				else          sym_to_key[arg.to_sym]
				end
			end
		end

		# this is the function that creates a symbol to key mapping. currently this works by making a
		# pass through the raw properties, but conceivably you could map symbols to keys using the
		# mapitags directly. problem with that would be that named properties wouldn't map automatically,
		# but maybe thats not too important.
		def sym_to_key
			return @sym_to_key if @sym_to_key
			@sym_to_key = {}
			raw.keys.each do |key|
				sym = key.to_sym
				Log.debug "couldn't find symbolic name for key #{key.inspect}" unless Symbol === sym
				if @sym_to_key[sym]
					Log.warn "duplicate key #{key.inspect}"
					# we give preference to PS_MAPI keys
					@sym_to_key[sym] = key if key.guid == PS_MAPI
				else
					# just assign
					@sym_to_key[sym] = key
				end
			end
			@sym_to_key
		end

		def keys
			sym_to_key.keys
		end
		
		def values
			keys.map { |key| raw[key] }
		end

		def [] arg, guid=nil
			raw[resolve(arg, guid)]
		end

		def []= arg, *args
			args.unshift nil if args.length == 1
			guid, value = args
			# FIXME this won't really work properly. it would need to go
			# to TAGS to resolve, as it often won't be there already...
			raw[resolve(arg, guid)] = value
		end

		def method_missing name, *args
			if name.to_s !~ /\=$/ and args.empty?
				self[name]
			elsif name.to_s =~ /(.*)\=$/ and args.length == 1
				self[$1] = args[0]
			else
				super
			end
		end

		def to_h
			sym_to_key.inject({}) { |hash, (sym, key)| hash.update sym => raw[key] }
		end
		
		# the other pseudo properties like body_html, body_rtf etc that are inferred will exist here.

		# -----
		
		# temporary pseudo tags
		
		# for providing rtf to plain text conversion. later, html to text too.
		def body
			return @body if defined?(@body)
			@body = (self[:body] rescue nil)
			@body = (::RTF::Converter.rtf2text body_rtf rescue nil) if !@body or @body.strip.empty?
			@body
		end

		# for providing rtf decompression
		def body_rtf
			return @body_rtf if defined?(@body_rtf)
			@body_rtf = (RTF.rtfdecompr rtf_compressed.read rescue nil)
		end

		# for providing rtf to html conversion
		def body_html
			return @body_html if defined?(@body_html)
			@body_html = (self[:body_html].read rescue nil)
			@body_html = (RTF.rtf2html body_rtf rescue nil) if !@body_html or @body_html.strip.empty?
			# last resort
			if !@body_html or @body_html.strip.empty?
				Log.warn 'creating html body from rtf'
				@body_html = (::RTF::Converter.rtf2text body_rtf, :html rescue nil)
			end
			@body_html
		end
	end

	# IMessage essentially, but there's also stuff like IMAPIFolder etc. so, for this to form
	# basis for PST Item, it'd need to be more general.
	class Item #< Msg
		# IAttach
		class Attachment #< Msg::Attachment
		end


		class Recipient #< Recipient
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

	class ObjectWithProperties
		attr_reader :properties
		alias props properties

		# +properties+ should be a PropertySet instance.
		def initialize properties
			@properties = properties
		end
	end

	# a general attachment class. is subclassed by Msg and Pst attachment classes
	class Attachment < ObjectWithProperties
		def filename
			props.attach_long_filename || props.attach_filename
		end

		def data
			@embedded_msg || @embedded_ole || props.attach_data
		end

		# with new stream work, its possible to not have the whole thing in memory at one time,
		# just to save an attachment
		#
		# a = msg.attachments.first
		# a.save open(File.basename(a.filename || 'attachment'), 'wb') 
		def save io
			raise "can only save binary data blobs, not ole dirs" if @embedded_ole
			data.each_read { |chunk| io << chunk }
		end

		def inspect
			"#<#{self.class.to_s[/\w+$/]}" +
				(filename ? " filename=#{filename.inspect}" : '') +
				(@embedded_ole ? " embedded_type=#{@embedded_ole.embedded_type.inspect}" : '') + ">"
		end
	end
	
	class Recipient < ObjectWithProperties
		# some kind of best effort guess for converting to standard mime style format.
		# there are some rules for encoding non 7bit stuff in mail headers. should obey
		# that here, as these strings could be unicode
		# email_address will be an EX:/ address (X.400?), unless external recipient. the
		# other two we try first.
		# consider using entry id for this too.
		def name
			name = props.transmittable_display_name || props.display_name
			# dequote
			name[/^'(.*)'/, 1] or name rescue nil
		end

		def email
			props.smtp_address || props.org_email_addr || props.email_address
		end

		RECIPIENT_TYPES = { 0 => :orig, 1 => :to, 2 => :cc, 3 => :bcc }
		def type
			RECIPIENT_TYPES[props.recipient_type]
		end

		def to_s
			if name = self.name and !name.empty? and email && name != email
				%{"#{name}" <#{email}>}
			else
				email || name
			end
		end

		def inspect
			"#<#{self.class.to_s[/\w+$/]}:#{self.to_s.inspect}>"
		end
	end

	# i refer to it as a message (as does mapi), although perhaps Item is better, as its a more general
	# concept than a message
	class Message < ObjectWithProperties
		# these 2 collections should be provided by our subclasses
		def attachments
			raise NotImplementedError
		end

		def recipients
			raise NotImplementedError
		end
	end
	
	# the #to_mime logic will be refactored out into mapi/to_mime.rb
	#   will probably ditch custom Mime class, and go with TMail. to_mime.rb
	#   will include that mapping, as well as vcard, etc etc. or, i'll have
	#   mapi/to_mime/message.rb mapi/to_mime/vcard.rb mapi/to_mime/ical.rb
	#   i will add minimal functionality to the base classes beyond that.
	# the msg parsing will be consolidated into mapi/msg.rb
	# the pst parsing will go into mapi/pst.rb
	# rtf related code will be in rtf.rb and mapi/rtf.rb
end

#message = Mapi::Msg.open 'test-swetlana_novikova.msg'
#message = Mapi::Msg.open ARGV.first #'test-swetlana_novikova.msg'
#p message.props
#p message.props.message_class
#puts message.convert

Mapi::Msg.open 'test_Blammo.msg' do |msg|
	puts msg.to_mime.to_tree
	puts msg.to_mime.to_s
end if $0 == __FILE__

__END__

pst = Mapi::Pst.open 'test.pst'
message2 = pst.find { |m| m.path =~ /^inbox\/test/i and m.subject == 'test' }

# what about conversion. i'm thinking something like:
# Mapi::PropertyStore#replace and/or #update
# and similarly, Message#replace. so you can do something like:

message3 = Mapi::Msg.open 'test2.msg', 'rb+' # creates empty msg file
message3.update message2 # recursively copy all properties
message3.close

# can define Message#to_msg as a shortcut to this. similarly, in the unlikely
# event that pst write support is ever implemented.

# needs to be cheap to create a message, not like the current situation. ie, fully lazily loading.
# on the other hand, the inspect strings will probably cause things to be loaded, but thats ok.
