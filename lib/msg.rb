#! /usr/bin/ruby

$: << './lib'

require 'yaml'
require 'base64'
require 'logger'

require 'ole/storage'
require 'mime'
require 'support'

# turn binary guid into something displayable
module Ole
	class Storage
		def self.parse_guid s
			"{%08x-%04x-%04x-%02x%02x-#{'%02x' * 6}}" % s.unpack('L S S CC C6')
		end
	end
end

#
# Msg class
# ===============================================================================
# primary class interface to the vagaries of .msg files
#

class Msg
	VERSION = '1.2.9'

	Log = Logger.new STDERR
	Log.formatter = proc do |severity, time, progname, msg|
		# find where we were called from, in our code
		callstack = caller.dup
		callstack.shift while callstack.first =~ /\/logger\.rb:\d+:in/
		from = callstack.first.sub /:in `(.*?)'/, ":\\1"
		"[%s %s]\n%-7s%s\n" % [time.strftime('%H:%M:%S'), from, severity, msg.to_s]
	end
	# void logger
	# there should be something like Logger::VOID, as this wouldn't be uncommon.
	# or maybe you should just use STDERR, and set a level so that nothing prints anyway
	#.instance_eval do
	#	%w[warn debug info].each do |sym|
	#		define_method(sym) {}
	#	end
	#end

	attr_reader :ole, :attachments, :recipients, :headers, :properties
	alias props :properties

	def self.load io
		Msg.new Ole::Storage.load(io)
	end

	# +ole+ is an Ole::Storage object
	def initialize ole
		@ole = ole
		@root = @ole.root

		Log.warn "root name was #{@root.name.inspect}" unless @root.name == 'Root Entry'

		@attachments = []
		@recipients = []
		@properties = Properties.load @root

		# process the children which aren't properties
		@properties.unused.each do |child|
			if child.dir?
				case child.name
				# these first 2 will actually be of the form
				# 1\.0_#([0-9A-Z]{8}), where $1 is the 0 based index number in hex
				when /__attach_version1\.0_/
					@attachments << Attachment.new(child)
				when /__recip_version1\.0_/
					@recipients << Recipient.new(child)
				when /__nameid_version1\.0/
					# FIXME: ignore nameid quietly at the moment
				else ignore child
				end
			end
		end

		# if these headers exist at all, they can be helpful. we may however get a
		# application/ms-tnef mime root, which means there will be little other than
		# headers. we may get nothing.
		# and other times, when received from external, we get the full cigar, boundaries
		# etc and all.
		@mime = Mime.new props.transport_message_headers.to_s
		populate_headers
	end

	def headers
		@mime.headers
	end

	# copy data from msg properties storage to standard mime. headers
	def populate_headers
		# for all of this stuff, i'm assigning in utf8 strings.
		# thats ok i suppose, maybe i can say its the job of the mime class to handle that.
		# but a lot of the headers are overloaded in different ways. plain string, many strings
		# other stuff. what happens to a person who has a " in their name etc etc. encoded words
		# i suppose. but that then happens before assignment. and can't be automatically undone
		# until the header is decomposed into recipients.
		for type, recips in recipients.group_by { |r| r.type }
			# details of proper escaping and whatever etc are the job of recipient.to_s
			# don't know if this sort is really needed. header folding isn't our job
			# don't know why i bother, but if we can, we try to sort recipients by the numerical part
			# of the ole name.
			recips = (recips.sort_by { |r| r.obj.name[/\d{8}$/].hex } rescue recips)
			# are you supposed to use ; or , to separate?
			headers[type.to_s.sub(/^(.)/) { $1.upcase }] = [recips.join("; ")]
		end
		headers['Subject'] = [props.subject]

		# construct a From value
		# should this kind of thing only be done when headers don't exist already? maybe not. if its
		# sent, then modified and saved, the headers could be wrong?
		# hmmm. i just had an example where a mail is sent, from an internal user, but it has transport
		# headers, i think because one recipient was external. the only place the senders email address
		# exists is in the transport headers. so its maybe not good to overwrite from.
		# recipients however usually have smtp address available.
		# maybe we'll do it for all addresses that are smtp? (is that equivalent to 
		# sender_email_address !~ /^\//
		name, email = props.sender_name, props.sender_email_address
		if props.sender_addrtype == 'SMTP'
			headers['From'] = if name and email and name != email
				[%{"#{name}" <#{email}>}]
			else
				[email || name]
			end
		elsif !self.from
			# some messages were never sent, so that sender stuff isn't filled out. need to find another
			# way to get something
			# what about marking whether we thing the email was sent or not? or draft?
			# for partition into an eventual Inbox, Sent, Draft mbox set?
			if name or email
				Log.warn "* no smtp sender email address available (only X.400). creating fake one"
				# this is crap. though i've specially picked the logic so that it generates the correct
				# email addresses in my case.
				user = name.sub /(.*), (.*)/, "\\2.\\1"
				domain = email[%r{^/O=([^/]+)}i, 1].downcase + '.com'
				headers['From'] = [%{"#{name}" <#{user}@#{domain}>}]
			else
				Log.warn "* no sender email address available at all. FIXME"
			end
		# else we leave the transport message header version
		end

		# fill in a date value. by default, we won't mess with existing value hear
		if headers['Date'].empty?
			# we want to get a received date, as i understand it.
			# use this preference order, or pull the most recent?
			keys = %w[message_delivery_time client_submit_time last_modification_time creation_time]
			time = keys.each { |key| break time if time = props.send(key) }
			time = nil unless Date === time
			# can employ other methods for getting a time. heres one in a similar vein to msgconvert.pl,
			# ie taking the time from an ole object
			time ||= @ole.dirs.map { |dir| dir.time }.compact.sort.last

			# now convert and store
			# this is a little funky. not sure about time zone stuff either?
			# actually seems ok. maybe its always UTC and interpreted anyway. or can be timezoneless.
			# i have no timezone info anyway.
			# in gmail, i see stuff like 15 Jan 2007 00:48:19 -0000, and it displays as 11:48.
			# similarly, if I output 
			require 'time'
			headers['Date'] = [Time.iso8601(time.to_s).rfc2822] if time
		end

		if headers['Message-ID'].empty? and props.internet_message_id
			headers['Message-ID'] = [props.internet_message_id]
		end
		if headers['In-Reply-To'].empty? and props.in_reply_to_id
			headers['In-Reply-To'] = [props.in_reply_to_id]
		end
	end

	def ignore obj
		Log.warn "* ignoring #{obj.name} (#{obj.type.to_s})"
	end

=begin
usually message class is one of:
IPM.Contact (convert to .vcf)
IPM.Activity (this is from the journal)
IPM.Note (this is a mail -> .eml)
IPM.Appointment (from the calendar)
IPM.StickyNote (just a regular note. probably -> rtf)

FIXME: look at data/src/content_classes information
=end
	def type
		props.message_class[/IPM\.(.*)/, 1].downcase
	end

	# shortcuts to some things from the headers
	%w[From To Cc Bcc Subject].each do |key|
		define_method(key.downcase) { headers[key].join(' ') unless headers[key].empty? }
	end

	def inspect
		str = %w[from to cc bcc subject type].map do |key|
			send(key) and "#{key}=#{send(key).inspect}"
		end.compact.join(' ')
		"#<Msg #{str}>"
	end

	# --------
	# beginnings of conversion stuff

	def convert
		# 
		# for now, multiplex between returning a Mime object,
		# a Vpim::Vcard object,
		# a Vpim::Vcalendar object
		#
		# all of which should support a common serialization,
		# to save the result to a file.
		#
	end

	def body_to_mime
		# to create the body
		# should have some options about serializing rtf. and possibly options to check the rtf
		# for rtf2html conversion, stripping those html tags or other similar stuff. maybe want to
		# ignore it in the cases where it is generated from incoming html. but keep it if it was the
		# source for html and plaintext.
		if props.body_rtf or props.body_html
			# should plain come first?
			mime = Mime.new "Content-Type: multipart/alternative\r\n\r\n"
			mime.parts << Mime.new("Content-Type: text/plain\r\n\r\n" + props.body)
			mime.parts << Mime.new("Content-Type: text/html\r\n\r\n"  + props.body_html) if props.body_html
			#mime.parts << Mime.new("Content-Type: text/rtf\r\n\r\n"   + props.body_rtf)  if props.body_rtf
			# temporarily disabled the rtf. its just showing up as an attachment anyway.
			# for now, what i can do, is this:
			# (maybe i should just overload body_html do to this)
			if props.body_rtf and !props.body_html and props.body_rtf['htmltag']
				open('temp.html', 'w') { |f| f.write props.body_rtf }
				begin
					html = `./rtf2html < temp.html`
					mime.parts << Mime.new("Content-Type: text/html\r\n\r\n" + html.gsub(/(<\/html>).*\Z/mi, "\\1"))
				ensure
					File.unlink 'temp.html' rescue nil
				end
			end
			mime
		else
			# check no header case. content type? etc?. not sure if my Mime will accept
			Log.debug "taking that other path"
			# body can be nil, hence the to_s
			Mime.new "Content-Type: text/plain\r\n\r\n" + props.body.to_s
		end
	end

	def to_mime
		# intended to be used for IPM.note, which is the email type. can use it for others if desired,
		# YMMV
		Log.warn "to_mime used on a #{props.message_class}" unless props.message_class == 'IPM.Note'
		# we always have a body
		mime = body = body_to_mime

		# do we have attachments??
		unless attachments.empty?
			mime = Mime.new "Content-Type: multipart/mixed\r\n\r\n"
			mime.parts << body
			attachments.each { |attach| mime.parts << attach.to_mime }
		end

		# at this point, mime is either
		# - a single text/plain, consisting of the body
		# - a multipart/alternative, consiting of a few bodies
		# - a multipart/mixed, consisting of 1 of the above 2 types of bodies, and attachments.
		# we add this standard preamble if its multipart
		# FIXME preamble.replace, and body.replace both suck.
		# preamble= is doable. body= wasn't being done because body will get rewritten from parts
		# if multipart, and is only there readonly. can do that, or do a reparse...
		mime.preamble.replace "This is a multi-part message in MIME format.\r\n" if mime.multipart?

		# now that we have a root, we can mix in all our headers
		headers.each do |key, vals|
			# don't overwrite the content-type, encoding style stuff
			next unless mime.headers[key].empty?
			mime.headers[key] += vals
		end

		mime
	end

	def to_vcard
		require 'rubygems'
		require 'vpim/vcard'
		# a very incomplete mapping, but its a start...
		# can't find where to set a lot of stuff, like zipcode, jobtitle etc
		card = Vpim::Vcard::Maker.make2 do |m|
			# these are all standard mapi properties
			m.add_name do |n|
				n.given = props.given_name.to_s
				n.family = props.surname.to_s
				n.fullname = props.subject.to_s
			end

			# outlook seems to eschew the mapi properties this time,
			# like postal_address, street_address, home_address_city
			# so we use the named properties
			m.add_addr do |a|
				a.location = 'work'
				a.street = props.business_address_street.to_s
				# i think i can just assign the array
				a.locality = [props.business_address_city, props.business_address_state].compact.join ', '
				a.country = props.business_address_country.to_s
				a.postalcode = props.business_address_postal_code.to_s
			end

			# right type?
			m.birthday = props.birthday if props.birthday
			m.nickname = props.nickname.to_s

			# photo available?
			# FIXME finish, emails, telephones etc
		end
	end

	# after looking at other MAPI property stores, such as tnef / pst, i might want to separate
	# the parsing from other stuff here, but for now its ok.
	class Properties
		IDENTITY_PROC = proc { |a| a }
		ENCODINGS = {
			0x000d => 'Directory', # seems to be used when its going to be a directory instead of a file. eg nested ole. 3701 usually
			0x001f => Ole::Storage::UTF16_TO_UTF8, # unicode?
			0x001e => proc { |a| a[0..-2] }, # ascii?
			0x0102 => IDENTITY_PROC, # binary?
		}

		# these won't be strings for much longer.
		# maybe later, the Key#inspect could automatically show symbolic guid names if they
		# are part of this builtin list.
		# FIXME
		PS_MAPI =             '{not-really-sure-what-this-should-say}'
		PS_PUBLIC_STRINGS =   '{00020329-0000-0000-c000-000000000046}'
		# string properties in this namespace automatically get added to the internet headers
		PS_INTERNET_HEADERS = '{00020386-0000-0000-c000-000000000046}'
		# theres are bunch of outlook ones i think
		# http://blogs.msdn.com/stephen_griffin/archive/2006/05/10/outlook-2007-beta-documentation-notification-based-indexing-support.aspx
		# IPM.Appointment
		PSETID_Appointment =  '{00062002-0000-0000-c000-000000000046}'
		# IPM.Task
		PSETID_Task =         '{00062003-0000-0000-c000-000000000046}'
		# used for IPM.Contact
		PSETID_Address =      '{00062004-0000-0000-c000-000000000046}'
		PSETID_Common =       '{00062008-0000-0000-c000-000000000046}'
		# didn't find a source for this name. it is for IPM.StickyNote
		PSETID_Note =         '{0006200e-0000-0000-c000-000000000046}'
		# for IPM.Activity. also called the journal?
		PSETID_Log =          '{0006200a-0000-0000-c000-000000000046}'

		# Enumerable removed. not really needed

		# access the underlying raw properties by code. no implicit type conversion
		# by tag and whatever other higherlevel stuff i end up doing in this class.
		# unused is so that you can use this as a property store, and get the unused children
		attr_reader :raw, :unused, :nameid

		def initialize
			@raw = {}
			@unused = []
		end

		def self.load obj
			prop = Properties.new
			prop.load obj
			prop
		end

		def load obj
			# we need to do the nameid first, as it provides the map for later user defined properties
			children = obj.children.dup
			@nameid = if nameid_obj = children.find { |child| child.name == '__nameid_version1.0' }
				children.delete nameid_obj
				Properties.parse_nameid nameid_obj
			end
			# now parse the actual properties
			children.each do |child|
				if child.file?
					begin
						case child.name
						when /__properties_version1\.0/
							parse_properties child
						when /__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})(?:-([0-9A-F]{8}))?/
							parse_substg *($~[1..-1].map { |num| num.hex rescue nil } + [child])
						else raise "bad name for mapi property #{child.name.inspect}"
						end
					rescue
						Log.warn $!
						@unused << child
					end
				else @unused << child
				end
			end
		end

		def self.parse_nameid obj
			guids_obj = obj.children.find { |child| child.name == '__substg1.0_00020102' }
			props_obj = obj.children.find { |child| child.name == '__substg1.0_00030102' }
			names_obj = obj.children.find { |child| child.name == '__substg1.0_00040102' }
			remaining = obj.children.dup
			[guids_obj, props_obj, names_obj].each { |obj| remaining.delete obj }

			# parse guids
			# this is the guids for named properities (other than builtin ones)
			# i think PS_PUBLIC_STRINGS, and PS_MAPI are builtin.
			guids = [PS_PUBLIC_STRINGS] + guids_obj.data.scan(/.{16}/m).map do |str|
				Ole::Storage.parse_guid str
			end

			# parse names.
			# the string ids for named properties
			# they are no longer parsed, as they're referred to by offset not
			# index. they are simply sequentially packed, as a long, giving
			# the string length, then padding to 4 byte multiple, and repeat.

			# parse actual props.
			# not sure about any of this stuff really.
			# should flip a few bits in the real msg, to get a better understanding of how this works.
			props = props_obj.data.scan(/.{8}/m).map do |str|
				flags, offset = str[4..-1].unpack 'S2'
				# the property will be serialised as this pseudo property, mapping it to this named property
				pseudo_prop = 0x8000 + offset
				named = flags & 1 == 1
				prop = if named
					str_off = *str.unpack('L')
					data = names_obj.data
					len = *data[str_off, 4].unpack('L')
					Ole::Storage::UTF16_TO_UTF8[data[str_off + 4, len]]
				else
					a, b = str.unpack('S2')
					Log.debug "b not 0" if b != 0
					a
				end
				# a bit sus
				guid_off = flags >> 1
				# missing a few builtin PS_*
				Log.debug "guid off < 2 (#{guid_off})" if guid_off < 2
				guid = guids[guid_off - 2]
				[pseudo_prop, Key.new(prop, guid)]
			end

			Log.warn "* ignoring #{remaining.length} objects in nameid" unless remaining.empty?
			# this leaves a bunch of other unknown chunks of data with completely unknown meaning.
			# pp [:unknown, child.name, child.data.unpack('H*')[0].scan(/.{16}/m)]
			Hash[*props.flatten]
		end

		def parse_substg key, encoding, offset, obj
			if (encoding & 0x1000) != 0
				if !offset
					# there is typically one with no offset first, whose data is a series of numbers
					# equal to the lengths of all the sub parts. gives an implied array size i suppose.
					# maybe you can initialize the array at this time. the sizes are the same as all the
					# ole object sizes anyway, its to pre-allocate i suppose.
					#p obj.data.unpack('L*')
					# ignore this one
					return
				else
					# remove multivalue flag for individual pieces
					encoding &= ~0x1000
				end
			else
				Log.warn "offset specified for non-multivalue encoding #{obj.name}" if offset
				offset = nil
			end
			# offset is for multivalue encodings.
			unless encoder = ENCODINGS[encoding]
				Log.warn "unknown encoding #{encoding}"
				encoder = IDENTITY_PROC
			end
			add_property key, encoder[obj.data], offset
		end

		# i think this is fairly wrong
		def parse_properties obj
			data = obj.data
			# don't really understand this that well...
			pad = data.length % 16
			unless (pad == 0 || pad == 8) and data[0...pad] == "\000" * pad
				Log.warn "padding was not as expected #{pad} (#{data.length}) -> #{data[0...pad].inspect}"
			end
			data[pad..-1].scan(/.{16}/m).each do |data|
				property, encoding = ('%08x' % data.unpack('L')).scan /.{4}/
				key = property.hex
				# doesn't make any sense to me. probably because its a serialization of some internal
				# outlook structure...
				next if property == '0000'
				case encoding
				when '0102', '001e', '001f', '101e', '101f'
					# ignore on purpose. not sure what its for
					# multivalue versions ignored also
				when '0003' # long
					# don't know what all the other data is for
					add_property key, *data[8, 4].unpack('L')
				when '000b' # boolean
					# again, heaps more data than needed. and its not always 0 or 1.
					# they are in fact quite big numbers. this is wrong.
#					p [property, data[4..-1].unpack('H*')[0]]
					add_property key, data[8, 4].unpack('L')[0] != 0
				when '0040' # systime
					# seems to work:
					add_property key, Ole::Storage::OleDir.parse_time(*data[8..-1].unpack('L*'))
				else
					Log.warn "ignoring data in __properties section, encoding: #{encoding}"
					Log << data.unpack('H*').inspect + "\n"
				end
			end
		end

		def add_property key, value, pos=nil
			# map keys in the named property range through nameid
			if Integer === key and key >= 0x8000
				if real_key = @nameid[key]
					key = real_key
				else
					Log.warn "property in named range not in nameid #{key.inspect}"
					key = Key.new key
				end
			else
				key = Key.new key
			end
			if pos
				@raw[key] ||= []
				Log.warn "duplicate property" unless Array === @raw[key]
				# ^ this is actually a trickier problem. the issue is more that they must all be of
				# the same type.
				@raw[key][pos] = value
			else
				# take the last.
				Log.warn "duplicate property #{key.inspect}" if @raw[key]
				@raw[key] = value
			end
		end

		# resolve an arg (could be key, code, string, or symbol), and possible guid to a key
		def resolve arg, guid=nil
			if guid;        Key.new arg, guid
			else
				case arg
				when Key;     arg
				when Integer; Key.new arg
				else          sym_to_key[arg.to_sym]
				end
			end or raise "unable to resolve key from #{[arg, guid].inspect}"
		end

		# just so i can get an easy unique list of missing ones
		@@quiet_property = {}

		def sym_to_key
			# create a map for converting symbols to keys. cache it
			unless @sym_to_key
				@sym_to_key = {}
				@raw.each do |key, value|
					sym = key.to_sym
					# used to use @@quiet_property to only ignore once
					Log.info "couldn't find symbolic name for key #{key.inspect}" unless Symbol === sym
					if @sym_to_key[sym]
						Log.warn "duplicate key #{key.inspect}"
						# we give preference to PS_MAPI keys
						@sym_to_key[sym] = key if key.guid == PS_MAPI
					else
						# just assign
						@sym_to_key[sym] = key
					end
				end
			end
			@sym_to_key
		end

		# accessors

		def [] arg, guid=nil
			@raw[resolve(arg, guid)] rescue nil
		end

		# need to rewrite the above so that i can leverage the logic for this version. maybe have
		# a resolve, and a create_sym_to_key_mapping function
		#def []= arg, guid=nil, value
		#	@raw[resolve(arg, guid)] = value
		#end

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
			hash = {}
			sym_to_key.each { |sym, key| hash[sym] = self[key] if Symbol === sym }
			hash
		end

		def inspect
			'#<Properties ' + to_h.map do |k, v|
				v = v.inspect
				"#{k}=#{v.length > 32 ? v[0..29] + '..."' : v}"
			end.join(' ') + '>'
		end

		# -----
		
		# temporary pseudo tag.
		def body_rtf
			return nil unless rtf_compressed
			return @body_rtf if @body_rtf
			open('temp.rtf', 'wb') { |f| f.write rtf_compressed }
			begin
				@body_rtf = `./rtfdecompr temp.rtf`
			ensure
				File.unlink 'temp.rtf'
			end
		end

		# ------
		# key class for accessing properties

		class Key
			attr_reader :code, :guid
			def initialize code, guid=PS_MAPI
				@code, @guid = code, guid
			end

			def to_sym
				# try to make a nice name out of ourselves. is this going to intern
				# too much stuff?
				# hmmm, for some stuff, like, eg, the message class specific range, sym-ification
				# of the key depends on knowing our message class. i don't want to store anything else
				# here though, so if that kind of thing is needed, it can be passed to this function.
				# worry about that when some examples arise.
				case code
				when Integer
					if guid == PS_MAPI # and < 0x8000 ?
						# the hash should be updated now that i've changed the process
						MAPITAGS['%04x' % code].first[/_(.*)/, 1].downcase.to_sym rescue code
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
				if Integer === code
					hex = '0x%04x' % code
					if guid == PS_MAPI
						# just display as plain hex number
						hex
					else
						"#<Key #{guid}/#{hex}>"
					end
				else
					# display full guid and code
					"#<Key #{guid}/#{code.inspect}>"
				end
			end
		end

		# YUCK moved here because we need Key
		# data files that provide for the code to symbolic name mapping
		# guids in named_map are really constant references to the above
		MAPITAGS = open('data/mapitags.yaml') { |file| YAML.load file }
		NAMED_MAP = Hash[*open('data/named_map.yaml') { |file| YAML.load file }.map do |key, value|
			[Key.new(key[0], const_get(key[1])), value]
		end.flatten]
	end

	class Attachment
		attr_reader :obj, :properties
		alias props :properties

		def initialize obj
			@obj = obj
			@properties = Properties.load @obj
			@properties.unused.each do |child|
					# FIXME warn
# FIXME: this is out of scope so doesn't warn anymore
#				else ignore child
#				if property == '3701' # data property means nested msg
#					puts "* ignoring nested msg."
			end
		end

		def filename
			props.attach_long_filename || props.attach_filename
		end

		def data
			props.attach_data
		end

		alias to_s :data

		def to_mime
			# TODO: smarter mime typing.
			mimetype = props.attach_mime_tag || 'application/octet-stream'
			mime = Mime.new "Content-Type: #{mimetype}\r\n\r\n"
			mime.headers['Content-Disposition'] = [%{attachment; filename="#{filename}"}]
			mime.headers['Content-Transfer-Encoding'] = ['base64']
			# data.to_s for now. data was nil for some reason.
			# perhaps it was a data object not correctly handled?
			mime.body.replace Base64.encode64(data.to_s).gsub(/\n/, "\r\n")
			mime
		end

		def inspect
			"#<#{self.class.to_s[/\w+$/]} filename=#{filename.inspect}>"
		end
	end

	#
	# Recipient class
	# ----------------------------------------------------------------------------
	# serves as a container for the recip directories in the .msg.
	# has things like office_location, business_telephone_number, but i don't
	# think enough to make a vCard out of??
	#
	class Recipient
		attr_reader :obj, :properties
		alias props :properties

		def initialize obj
			@obj = obj
			@properties = Properties.load @obj
			@properties.unused.each do |child|
				# FIXME warn
			end
		end

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
end

if $0 == __FILE__
	quiet = if ARGV[0] == '-q'
		ARGV.shift
		true
	end
	# just shut up and convert a message to eml
	Msg::Log.level = Logger::WARN
	Msg::Log.level = Logger::FATAL if quiet
	msg = Msg.load open(ARGV[0])
	puts msg.to_mime.to_s
end

