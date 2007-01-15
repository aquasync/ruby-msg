#! /usr/bin/ruby

require 'yaml'
require 'base64'
require 'ostruct'
require 'iconv'
require 'ole/storage'

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
# Ole class
# ------------------------------------------------------------------------------
# simple class that is intended to abstract away the details of the access to
# the ole objects. currently wraps a simple perl dumper based on
# OLE::Storage_Lite, with base64 encoding to get around some YAML incompatibilities
#

class Ole < OpenStruct
	UTF16_TO_UTF8 = Iconv.new('utf-8', 'utf-16le').method :iconv

	TYPE_MAP = {
		1 => :dir,
		2 => :file,
		5 => :root
	}

	def self.load filename
		data = File.popen"./ole2yaml.pl '#{filename.gsub /'/, %{'"'"'}}'" do |pipe|
			YAML.load(pipe)
		end
		Ole.decode data
	end
	
	def self.decode data
		data['name'] = UTF16_TO_UTF8[data['name']]
		data['data'] = Base64.decode64(data['data']) if data['data']
		data['kind'] = TYPE_MAP[data['kind']] or raise "unknown ole type #{data['kind']}"

		if data['children']
			data['children'].map! { |child| Ole.decode child }
		end

		Ole.new data
	end

	def file?
		kind == :file
	end

	def inspect
		"#<Ole @name=#{name.inspect}>"
	end
end

#
# Msg class
# ===============================================================================
# primary class interface to the vagaries of .msg files
#

class Msg
	class Properties
		include Enumerable

		IDENTITY_PROC = proc { |a| a }

		ENCODINGS = {
			'000d' => 'Directory', # seems to be used when its going to be a directory instead of a file. eg nested ole. 3701 usually
			'001f' => Ole::UTF16_TO_UTF8, # unicode?
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
		attr_reader :raw

		def initialize
			@raw = {}
		end

		# just so i can get an easy unique list of missing ones
		@@quiet_property = {}

		def << obj
			# i got one like this: `__substg1.0_800B101F-00000000', so anchor removed from end
			# as expected, i then got duplicate property warnings. so that must have been what its
			# for. i should maybe look into user defined properties. maybe things occur multiple
			# times if user defined. no, that doesn't really explain it. not sure what the meaning
			# of multiple properties should be.... hmmm. maybe its an ole thing. maybe the first
			# one was "trash"
			case obj.name
			when /__properties_version1\.0/
				data = obj.data
				# don't really understand this that well...
				pad = data.length % 16
				unless (pad == 0 || pad == 8) and data[0...pad] == "\000" * pad
					warn "padding was not as expected #{pad} (#{data.length}) -> #{data[0...pad].inspect}"
				end
				# not really sure about using regexps to break binary data into byte chunks
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
						# they are in fact quite big numbers.
#						p [property, data[4..-1].unpack('H*')[0]]
						add_property property, data[8, 4].unpack('L')[0] != 0
					when '0040' # systime
						p [property, data[4..-1].unpack('H*')[0]]
					else
						warn "ignoring data in __properties section, encoding: #{encoding}"
					end
				end
			when /^__substg1\.0_(....)(....)/
				property, encoding = $~[1..-1]
				unless encoder = ENCODINGS[encoding.downcase]
					warn "unknown encoding #{encoding}"
					encoder = IDENTITY_PROC
				end
				add_property property.downcase, encoder[obj.data]
			else
				raise "bad name for mapi property #{obj.name.inspect}"
			end
			self
		end

		def add_property property, data
			if @raw[property]
				warn "duplicate property #{property}"
				# take the last.
			end
			unless MAPITAGS[property]
				warn "unknown property #{property}" unless @@quiet_property[property]
				@@quiet_property[property] = true
			end
			@raw[property] = data
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
			@properties = Properties.new

			@obj.children.each do |child|
				if child.file?
					begin   @properties << child
					rescue; warn $!
					end
				else
					# FIXME warn
# FIXME: this is out of scope so doesn't warn anymore
#				else ignore child
#				if property == '3701' # data property means nested msg
#					puts "* ignoring nested msg."
				end
			end
		end

		def filename
			props.attach_long_filename || props.attach_filename
		end

		def data
			props.attach_data
		end

		alias to_s :data

		def inspect
			"#<Attachment filename=#{filename.inspect}>"
		end

		def to_mime
			# TODO: smarter mime typing.
			# it kind of sucks, having all this stuff in memory. ole and mime should support
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
			mime.body.replace Base64.encode64(data).gsub(/\n/, "\r\n")
			mime
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
			@properties = Properties.new

			@obj.children.each do |child|
				if child.file?
					# treat everything else as a property.
					begin   @properties << child
					rescue; warn $!
					end
				else
					# FIXME warn
				end
			end
		end

		# some kind of best effort guess for converting to standard mime style format.
		# there are some rules for encoding non 7bit stuff in mail headers. should obey
		# that here, as these strings could be unicode
		# email_address will be and EX:/ exchange address, unless external recipient. the
		# other two we try first.
		# consider using entry id for this too.
		def name
			name = props.transmittable_display_name || props.display_name
			name[/^'(.*)'/, 1] or name
		end

		def email
			props.smtp_address || props.org_email_addr || props.email_address
		end

		def kind
			{ 0 => :orig, 1 => :to, 2 => :cc, 3 => :bcc }[props.recipient_type]
		end

		def to_s
			name && !name.empty? && email ? %{"#{name}" <#{email}>} : (email || name)
		end

		def inspect
			"#<Recipient:#{self.to_s.inspect}>"
		end
	end

	attr_reader :obj, :attachments, :recipients, :headers
	alias props :properties

	def self.load filename
		Msg.new Ole.load(filename)
	end

	# +obj+ is the ole root object
	def initialize obj
		@obj = obj

		@attachments = []
		@recipients = []
		@properties = Properties.new

		warn "root name was #{@obj.name.inspect}" unless @obj.name == 'Root Entry'

		# process the direct children of the root.
		@obj.children.each do |child|
			if child.file?
				# treat everything else as a property.
				begin   @properties << child
				rescue; warn $!
				end
			else
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
		for kind, recips in recipients.group_by { |r| r.kind }
			# details of proper escaping and whatever etc are the job of recipient.to_s
			# don't know if this sort is really needed. header folding isn't our job
			headers[kind.to_s.sub(/^(.)/) { $1.upcase }] =
				[recips.sort_by { |r| r.obj.name[/\d{8}$/].hex }.join(" ")]
		end
		headers['Subject'] = [props.subject]
	end

	# the trick now, is a to_s, or to_mime here, that converts .msg to something normal.
	# that is really the aim.

	def ignore obj
		warn "* ignoring #{obj.name} (#{obj.kind.to_s})"
	end

=begin
usuall message class is one of:
IPM.Contact (convert to .vcf)
IPM.Activity (this is from the journal)
IPM.Note (this is a mail -> .eml)
IPM.Appointment (from the calendar)
IPM.StickyNote (just a regular note. probably -> rtf)

FIXME: look at data/content_classes information
=end
	def kind
		props.message_class[/IPM\.(.*)/, 1].downcase
	end

	def inspect
		# the gsubs are just because of the inspecting.
		str = %w[From To Bcc Cc].map do |kind|
			next if headers[kind].empty?
			kind.downcase + '=' + headers[kind].join(' ').inspect.gsub(/\\"/, "'")
		end.compact.join(' ')
		to   = headers['To'].join(' ').inspect.gsub(/\\"/, "'")
		"#<Msg subject=#{props.subject.inspect} #{str} kind=#{kind.inspect}>"
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
					File.unlink 'temp.html'
				end
			end
			mime
		else
			# check no header case. content type? etc?. not sure if my Mime will accept
			warn "taking that other path"
			Mime.new "Content-Type: text/plain\r\n\r\n" + props.body
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
end


if $0 == __FILE__
	msg = Msg.new ARGV[0]
	p msg
end

=begin
# need to work on reducing this. i think latest platform sdk, and
# some tnef headers will reduce this list further.

# somehow webdav related. not very interesting
newer formula, due to escaping:
	CGI.escape(msg.properties.subject.strip).tr('+', ' ') + '.EML' == msg.properties.url_comp_name
rs[:cc].sort_by { |r| r.obj.name[/\d{8}$/].hex }
suppose i should stop trying to figure these out:
"...The range for named property identifiers is between 0x8000 and 0xFFFE."
oh.

(mapidefs.h)
174 /* Types of message receivers */
175 #ifndef MAPI_ORIG
176 #define MAPI_ORIG      0          /* The original author */
177 #define MAPI_TO        1          /* The primary message receiver */
178 #define MAPI_CC        2          /* A carbon copy receiver */
179 #define MAPI_BCC       3          /* A blind carbon copy receiver */
180 #define MAPI_P1        0x10000000 /* A message resend */
181 #define MAPI_SUBMITTED 0x80000000 /* This message has already been sent */
182 #endif

when i start handling recipients as well, i get lots more
unknown property 5ff6
 * appears to be equal to display_name, and transmitable_display_name
unknown property 5ff7
 * recipient.properties.raw['5ff7'].upcase == recipient.properties.entryid
   is equivalent, but not all uppercase.
	 everything else is upper though. maybe a displayname kind of thing.

get the property ranges figured out, so i can group
properties by type (transmitable or not, etc).

TODO
- fill out some of the tag holes. mostly done
- integrate rtf decompression, as pseudo tag. mostly done
- preliminary mime conversion. mostly done. currently relies on transport_message_headers.
  need to also support internal mail, where we write all those headers ourselves instead,
	from the msg.props.
	like
	mime.headers['Subject'] = [msg.props.subject]
	...
	etc, for heaps of things, like Thread-ID, Message-ID, basics like From, To, CC etc etc.
	a more advanced possibility is some kind of smart merge, taking information superset, but
	i think i'll keep it simple for now.
- figure out some things, like entryids, and the properties directories,
  and the ntsecurity data 0e27.
- conversion of inline images.
  Content-Location, cid: urls etc etc.
	what would be cool, is conversion of outlooks text/rtf's only real "feature" over
	text/html - convert inline attachments to be <a href> links, using cid: urls to the
	actual content data, and using an <img with cid: url to a converted image from the
	attach_rendering property (image data), along with the text itself. (although i think
	the rendering may actually include the text ??. that would explain why its always clipped.
- entryid format cf. entry_id.h. another serialized structure.
  entryids are for the addressbook connection. EMS (exchange message something), AB
	address book. MUIDEMSAB. makes sense.

# ways to get the sender of a mail.
# if external, you can do this (different for internal).
name, protocol, email = Ole::UTF16_TO_UTF8[msg.props.sender_entryid[24..-1]].split(/\x00/)
# here is an excerpt from a yaml dump.
# need to consider how to get it. also when its a draft, and other stuff.
creator_name: GeorgiaL@ag.nsw.gov.au
sent_representing_name: GeorgiaL@ag.nsw.gov.au
last_modifier_name: GeorgiaL@ag.nsw.gov.au
sender_email_address: GeorgiaL@ag.nsw.gov.au
sent_representing_email_address: GeorgiaL@ag.nsw.gov.au
sender_name: GeorgiaL@ag.nsw.gov.au
=> nil



http://peach.ease.lsoft.com/scripts/wa.exe?A2=ind0207&L=mapi-l&P=24515
"In case anybody is interested, Exchange stores PR_NT_SECURITY_DESCRIPTOR as a header plus the regular self=-relative SECURITY_DESCRIPTOR structure. The first two bytes of the header (WORD) is the length of the header; Read these two bytes to find out how many bytes you must skip to get to the real data. Many thanks to Stephen Griffin for this info."
using outlook spy gives an actual dump. for example:
<<
Control: SE_DACL_AUTO_INHERITED | SE_DACL_DEFAULTED | SE_DACL_PRESENT | SE_GROUP_DEFAULTED | SE_OWNER_DEFAULTED | SE_SACL_AUTO_INHERITED | SE_SACL_DEFAULTED | SE_SELF_RELATIVE
Owner: 
  SID: S-1-5-21-1004336348-602609370-725345543-44726
  Name: lowecha
  DomainName: XXX
Group: 
  SID: S-1-5-21-1004336348-602609370-725345543-513
  Name: Domain Users
  DomainName: XXX
Dacl: 
    Header:
      AceType: ACCESS_DENIED_ACE_TYPE
      AceFlags: INHERITED_ACE
    Mask: fsdrightReadBody (fsdrightListContents) | fsdrightWriteBody (fsdrightCreateItem) | fsdrightAppendMsg (fsdrightCreateContainer) | fsdrightReadProperty | fsdrightWriteProperty | fsdrightExecute | fsdrightReadAttributes | fsdrightWriteAttributes | fsdrightWriteOwnProperty | fsdrightDeleteOwnItem | fsdrightViewItem | fsdrightWriteSD | fsdrightDelete | fsdrightWriteOwner | fsdrightReadControl | fsdrightSynchronize
    Sid:
      SID: S-1-5-7
      Name: ANONYMOUS LOGON
      DomainName: NT AUTHORITY
>>
Not something i care about at the moment.

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

def parse_nameid obj
	obj.children.each do |child|
		case child.name
		when /_0002/
			# this is the guids for named properities (other than builtin ones)
			# i think PS_PUBLIC_STRINGS, and PS_MAPI are builtin.
			guids = child.data.scan(/.{16}/m).map { |str| parse_guid str }
			pp [:guids, child.name, guids]
		when /_0004/
			# this is the string ids for named properties
			names = []
			data, i = child.data, 0
			while i < data.length
				len = *data[i, 4].unpack('L')
				names << Ole::UTF16_TO_UTF8[data[i += 4, len]]
				# skip text, with padding to multiple of 4
				i += (len + 3) & ~3
			end
			pp [:names, child.name, names]
		else
				pp [:unknown, child.name, child.data.unpack('H*')[0].scan(/.{16}/m)]
		end
	end
	nil
end

=begin
basically, i will just pass them through the same code path. because they are typed tags,
i will probably make the raw versions converted to appropriate integer values.
then the higher level versions will map enums to symbols. eg

r.props.recipient_type == :to
then,
r2 = recipients.group_by { |r| r.props.recipient_type }

r2[:to]
r2[:from]
r2[:bcc]
... etc

=end

