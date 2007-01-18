#! /usr/bin/ruby

$: << './lib'

require 'yaml'
require 'base64'
require 'ole/storage'
require 'support'

#
# Mime class
# -------------------------------------------------------------------------------
# basic mime class for really basic and probably non-standard parsing and
# construction of MIME messages. intended for assistance in parsing the
# transport_message_headers provided in .msg files, and as the container that
# can serialize itself for final conversion to standard format.
#

# some of this stuff seems to duplicate a bit of the work in net/http.rb's HTTPHeader, but I
# don't know if the overlap is sufficient. I don't want to lower case things, just for starters,
# meaning i can use this as read/write mime representation for constructing output, and not
# get screwed up case.

class Mime
	attr_reader :headers, :body, :parts, :content_type, :preamble, :epilogue

	def initialize str
		headers, @body = $~[1..-1] if str[/(.*?\r?\n)(?:\r?\n(.*))?\Z/m]

		@headers = Hash.new { |hash, key| hash[key] = [] }
		@body ||= ''
		headers.to_s.scan(/^\S+:\s*.*(?:\n\t.*)*/).each do |header|
			@headers[header[/(\S+):/, 1]] << header[/\S+:\s*(.*)/m, 1].gsub(/\s+/m, ' ').strip # this is kind of wrong
		end

		# don't have to have content type i suppose
		@content_type, attrs = nil, {}
		if content_type = @headers['Content-Type'][0]
			@content_type, attrs = Mime.split_header content_type
		end

		if multipart?
			if body.empty?
				@preamble = ''
				@epilogue = ''
				@parts = []
			else
				# we need to split the message at the boundary
				boundary = attrs['boundary'] or raise "no boundary for multipart message"

				# splitting the body:
				parts = body.split /--#{Regexp.quote boundary}/m
				unless parts[-1] =~ /^--/; warn "bad multipart boundary (missing trailing --)"
				else parts[-1][0..1] = ''
				end
				parts.each_with_index do |part, i|
					part =~ /^(\r?\n)?(.*?)(\r?\n)?\Z/m
					part.replace $2
					warn "bad multipart boundary" if (1...parts.length-1) === i and !($1 && $3)
				end
				@preamble = parts.shift
				@epilogue = parts.pop
				@parts = parts.map { |part| Mime.new part }
			end
		end
	end

	def multipart?
		@content_type and @content_type[/^multipart/]
	end

	def inspect
		# add some extra here.
		"#<Mime content_type=#{@content_type.inspect}>"
	end

	def to_tree
		if multipart?
			str = "- #{inspect}\n"
			parts.each_with_index do |part, i|
				last = i == parts.length - 1
				part.to_tree.split(/\n/).each_with_index do |line, j|
					str << "  #{last ? (j == 0 ? "\\" : ' ') : '|'}" + line + "\n"
				end
			end
			str
		else
			"- #{inspect}\n"
		end
	end

	def to_s opts={}
		opts = {:boundary_counter => 0}.merge opts
		if multipart?
			boundary = Mime.make_boundary opts[:boundary_counter] += 1, self
			@body = [preamble, parts.map { |part| "\r\n" + part.to_s(opts) + "\r\n" }, "--\r\n" + epilogue].
				flatten.join("\r\n--" + boundary)
			content_type, attrs = Mime.split_header @headers['Content-Type'][0]
			attrs['boundary'] = boundary
			@headers['Content-Type'] = [([content_type] + attrs.map { |key, val| %{#{key}="#{val}"} }).join('; ')]
		end

		str = ''
		@headers.each do |key, vals|
			vals.each { |val| str << "#{key}: #{val}\r\n" }
		end
		str << "\r\n" + @body
	end

	def self.split_header header
		# FIXME: haven't read standard. not sure what its supposed to do with " in the name, or if other
		# escapes are allowed. can't test on windows as " isn't allowed anyway. can be fixed with more
		# accurate parser later.
		# maybe move to some sort of Header class. but not all headers should be of it i suppose.
		# at least add a join_header then, taking name and {}. for use in Mime#to_s (for boundary
		# rewrite), and Attachment#to_mime, among others...
		attrs = {}
		header.scan(/;\s*([^\s=]+)\s*=\s*("[^"]*"|[^\s;]*)\s*/m).each do |key, value|
			if attrs[key]; warn "ignoring duplicate header attribute #{key.inspect}"
			else attrs[key] = value[/^"/] ? value[1..-2] : value
			end
		end

		[header[/^[^;]+/].strip, attrs]
	end

	# +i+ is some value that should be unique for all multipart boundaries for a given message
	def self.make_boundary i, extra_obj = Mime
		"----_=_NextPart_#{'%03d' % i}_#{'%08x' % extra_obj.object_id}.#{'%08x' % Time.now}"
	end
end

#
# Msg class
# ===============================================================================
# primary class interface to the vagaries of .msg files
#

class Msg
	attr_reader :ole, :attachments, :recipients, :headers, :properties
	alias props :properties

	def self.load io
		Msg.new Ole::Storage.load(io)
	end

	# +ole+ is an Ole::Storage object
	def initialize ole
		@ole = ole
		@root = @ole.root

		warn "root name was #{@root.name.inspect}" unless @root.name == 'Root Entry'

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
				warn "* no smtp sender email address available (only X.400). creating fake one"
				# this is crap. though i've specially picked the logic so that it generates the correct
				# email addresses in my case.
				user = name.sub /(.*), (.*)/, "\\2.\\1"
				domain = email[%r{^/O=([^/]+)}i, 1].downcase + '.com'
				headers['From'] = [%{"#{name}" <#{user}@#{domain}>}]
			else
				warn "* no sender email address available at all. FIXME"
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
		warn "* ignoring #{obj.name} (#{obj.type.to_s})"
	end

=begin
usually message class is one of:
IPM.Contact (convert to .vcf)
IPM.Activity (this is from the journal)
IPM.Note (this is a mail -> .eml)
IPM.Appointment (from the calendar)
IPM.StickyNote (just a regular note. probably -> rtf)

FIXME: look at data/content_classes information
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
			warn "taking that other path"
			# body can be nil, hence the to_s
			Mime.new "Content-Type: text/plain\r\n\r\n" + props.body.to_s
		end
	end

	def to_mime
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

	# after looking at other MAPI property stores, such as tnef / pst, i might want to separate
	# the parsing from other stuff here, but for now its ok.
	class Properties
		include Enumerable

		IDENTITY_PROC = proc { |a| a }
		ENCODINGS = {
			'000d' => 'Directory', # seems to be used when its going to be a directory instead of a file. eg nested ole. 3701 usually
			'001f' => Ole::Storage::UTF16_TO_UTF8, # unicode?
			'001e' => IDENTITY_PROC, # ascii?
			'0102' => IDENTITY_PROC, # binary?
		}

		# seems a bit ugly
		MAPITAGS = open('data/mapitags.yaml') { |file| YAML.load file }
		MAPITAGS_BY_NAME = MAPITAGS.to_a.inject({}) do |hash, pair|
			hash[pair[1][0][/^PR_(.*)/, 1].downcase] = pair[0]
			hash
		end

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
				parse_nameid nameid_obj
			end
			children.each do |child|
				if child.file?
					# treat everything else as a property.
					begin
						self << child
					rescue
						warn $!
						@unused << child
					end
				else
					@unused << child
				end
			end
		end

		# move this to the property class...
		PS_PUBLIC_STRINGS = '{00020329-0000-0000-c000-000000000046}'
		def parse_nameid obj
			guids_obj = obj.children.find { |child| child.name == '__substg1.0_00020102' }
			props_obj = obj.children.find { |child| child.name == '__substg1.0_00030102' }
			names_obj = obj.children.find { |child| child.name == '__substg1.0_00040102' }
			remaining = obj.children.dup
			[guids_obj, props_obj, names_obj].each { |obj| remaining.delete obj }

			# parse guids
			# this is the guids for named properities (other than builtin ones)
			# i think PS_PUBLIC_STRINGS, and PS_MAPI are builtin.
			guids = [PS_PUBLIC_STRINGS] + guids_obj.data.scan(/.{16}/m).map { |str| parse_guid str }

			# parse names.
			# this is the string ids for named properties
			# this isn't used, as they are referred to by offset, not index.
			names = []
			data, i = names_obj.data, 0
			while i < data.length
				len = *data[i, 4].unpack('L')
				names << Ole::Storage::UTF16_TO_UTF8[data[i += 4, len]]
				# skip text, with padding to multiple of 4
				i += (len + 3) & ~3
			end

			# parse actual props.
			# not sure about any of this stuff really.
			# should flip a few bits in the real msg, to get a better understanding of how this works.
			props = props_obj.data.scan(/.{8}/m).map do |str|
				flags, offset = str[4..-1].unpack 'S2'
				# the property will be serialised as this pseudo property, mapping it to this named property
				pseudo_prop = '%04x' % ("8000".hex + offset)
				named = flags & 1 == 1
				prop = if named
					str_off = *str.unpack('L')
					data = names_obj.data
					len = *data[str_off, 4].unpack('L')
					Ole::Storage::UTF16_TO_UTF8[data[str_off + 4, len]]
				else
					a, b = str.unpack('S2')
					warn "b not 0" if b != 0
					'%04x' % a
				end
				# a bit sus
				guid_off = flags >> 1
				# missing a few builtin PS_*
				warn "guid off < 2 (#{guid_off})" if guid_off < 2
				guid = guids[guid_off - 2]
				[pseudo_prop, guid, named, prop]
			end

			warn "* ignoring #{remaining.length} objects in nameid" unless remaining.empty?
			# this leaves a bunch of other unknown chunks of data with completely unknown meaning.
			# pp [:unknown, child.name, child.data.unpack('H*')[0].scan(/.{16}/m)]
			props
		end

		# just so i can get an easy unique list of missing ones
		@@quiet_property = {}

		def << obj
			# i got one like this: `__substg1.0_800B101F-00000000', so anchor removed from end
			# as expected, i then got duplicate property warnings. so that must have been what its
			# for. i should maybe look into user defined properties. maybe things occur multiple
			# times if user defined. no, that doesn't really explain it. not sure what the meaning
			# of multiple properties should be.... hmmm. maybe its an ole thing. maybe the first
			# one was "trash".
			# hmmm, think i might know what the other stuff is more - multivalue array position
			case obj.name
			when /__properties_version1\.0/
				data = obj.data
				# don't really understand this that well...
				pad = data.length % 16
				unless (pad == 0 || pad == 8) and data[0...pad] == "\000" * pad
					warn "padding was not as expected #{pad} (#{data.length}) -> #{data[0...pad].inspect}"
				end
				data[pad..-1].scan(/.{16}/m).each do |data|
					property, encoding = ('%08x' % data.unpack('L')).scan /.{4}/
					# doesn't make any sense to me. probably because its a serialization of some internal
					# outlook structure...
					next if property == '0000'
					case encoding
					when '0102', '001e', '001f'
						# ignore on purpose. not sure what its for
					when '0003' # long
						# don't know what all the other data is for
						add_property property, *data[8, 4].unpack('L')
					when '000b' # boolean
						# again, heaps more data than needed. and its not always 0 or 1.
						# they are in fact quite big numbers. this is wrong.
#						p [property, data[4..-1].unpack('H*')[0]]
						add_property property, data[8, 4].unpack('L')[0] != 0
					when '0040' # systime
						# seems to work:
						add_property property, Ole::Storage::OleDir.parse_time(*data[8..-1].unpack('L*'))
					else
						warn "ignoring data in __properties section, encoding: #{encoding}"
					end
				end
			when /^__substg1\.0_([0-9A-F]{4})([0-9A-F]{4})(?:-([0-9A-F]{8}))?/
				property, encoding, offset = $~[1..-1]
				if (encoding.hex & 0x1000) != 0
					if !offset
						# there is typically one with no offset first, whose data is a series of numbers
						# equal to the lengths of all the sub parts. gives an implied array size i suppose.
						# maybe you can initialize the array at this time. the sizes are the same as all the
						# ole object sizes anyway, its to pre-allocate i suppose.
						#p obj.data.unpack('L*')
						# ignore this one
						return self
					else
						offset = offset.to_s.hex
						# the encoding of the individual pieces is like this:
						encoding = '%04x' % (encoding.hex & ~0x1000)
					end
				else
					warn "offset specified for non-multivalue encoding #{obj.name}" if offset
					offset = nil
				end
				# offset is for multivalue encodings.
				unless encoder = ENCODINGS[encoding.downcase]
					warn "unknown encoding #{encoding}"
					encoder = IDENTITY_PROC
				end
				add_property property.downcase, encoder[obj.data], offset
			else
				raise "bad name for mapi property #{obj.name.inspect}"
			end
			self
		end

		def add_property property, data, pos=nil
			# fix the property
			if property.hex >= '8000'.hex
				# in the named range
				if realprop = (@nameid || []).assoc(property)
					property = realprop.values_at(1, 3).join
				else
					warn "property in named range not in nameid #{property.inspect}"
				end
			end
			unless MAPITAGS[property]
				# (string) named properties shouldn't need an associated name.
				warn "no name associated to #{property}" unless @@quiet_property[property]
				@@quiet_property[property] = true
			end
			if pos
				@raw[property] ||= []
				warn "duplicate property" unless Array === @raw[property]
				# ^ this is actually a trickier problem. the issue is more that they must all be of
				# the same type.
				@raw[property][pos] = data
			else
				# take the last.
				warn "duplicate property #{property}" if @raw[property]
				@raw[property] = data
			end
		end

		def [] name
			raise "unknown tag name #{name.inspect}" unless key = MAPITAGS_BY_NAME[name.to_s]
			@raw[key]
		end

		def []= name, value
			raise "unknown tag name #{name.inspect}" unless key = MAPITAGS_BY_NAME[name.to_s]
			@raw[key] = value
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

		def keys
			@raw.keys.map { |key| name = MAPITAGS[key] and name[0][/^PR_(.*)/, 1].downcase }.compact
		end

		def each
			keys.each { |key| yield key, self[key] }
		end

		def values
			keys.map { |key| self[key] }
		end

		def to_h
			hash = {}
			each { |key, value| hash[key] = value }
			hash
		end

		def inspect
			'#<Properties ' + map do |k, v|
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
			# it kind of sucks, having all this stuff in memory. Ole and Mime classes should support
			# streams. for ole, if data length > 1024, it could automatically just become a
			# promise, wrapping a stream. most methods wouold just force the load to mem and then
			# it would delegate to the string. however, #to_io could return underlying io object.
			# unfortunately base64 builtin doesn't work on streams though, or provide mechanism
			# that i can see. but one can of course pass chunks of appropriate size to it. ideally,
			# attachments would just be streamed from the msg file, through appropriate encoding,
			# and to output file. so mime#to_s will be replaced with a call to something like
			# mime#serialize, using a StringIO etc. this then allows:
			# msg = Msg.load 'input.msg'
			# open('output.eml', 'w') { |file| msg.to_mime.serialize file }
			# the existence of promises etc will mean input.msg is actually still open though.s o
			# i'll then need either
			# msg.close
			# or block form. a finalizer can also release the file. 
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
	# serves as a container for the recip directories in the .msg. not sure of the
	# usefullness. has things like office_location, business_telephone_number, but
	# not enough for vcard. and doesn't have a normal email address anywhere, just
	# that exchange crap EX:/O=.../OU=.../CN=...
	# should look what msgconvert.pl was doing with them.
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
	msg = Msg.load open(ARGV[0])
	puts msg.to_mime.to_s
end

=begin


(28 bytes binary data and msg.properties.sender_email_address and "\000")
entryids are a strange format. for internal or from exchange whatever, they have that
EX:/O=XXXXXX/...
otherwise, they may have SMTP in them.
  such as msg.properties.sent_representing_search_key
	  == "SMTP:SOMEGUY@XXX.COM\000"
	but Ole::UTF16_TO_UTF8[msg2.properties.sender_entryid[/.*\000\000(.+)\000/, 1][0..-2]]
	  == "SomeGuy@XXX.COM"
	for external people, entry ids have displayname and address.
=end

# turn binary guid into something displayable
def parse_guid s
	"{%08x-%04x-%04x-%02x%02x-#{'%02x' * 6}}" % s.unpack('L S S CC C6')
end

