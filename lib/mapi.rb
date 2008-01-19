require 'mapi/types'
require 'mapi/property_set'

module Mapi
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
