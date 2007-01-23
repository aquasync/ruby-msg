#! /usr/bin/ruby

$: << File.dirname(__FILE__)

require 'yaml'
require 'base64'

require 'support'
require 'ole/storage'
require 'msg/properties'
require 'mime'

module Ole
	class Storage
		# turn a binary guid into something displayable
		def self.parse_guid s
			"{%08x-%04x-%04x-%02x%02x-#{'%02x' * 6}}" % s.unpack('L S S CC C6')
		end
	end
end

#
# = Introduction
#
# Primary class interface to the vagaries of .msg files.
#
# The core of the work is done by the <tt>Msg::Properties</tt> class.
#

class Msg
	VERSION = '1.2.12'
	# we look here for the yaml files in data/, and the exe files for support
	# decoding at the moment.
	SUPPORT_DIR = File.dirname(__FILE__) + '/..'

	Log = Logger.new_with_callstack

	attr_reader :root, :attachments, :recipients, :headers, :properties
	alias props :properties

	def self.load io
		Msg.new Ole::Storage.load(io).root
	end

	# +root+ is an Ole::Storage::OleDir object
	def initialize root
		@root = root
		@attachments = []
		@recipients = []
		@properties = Properties.load @root

		# process the children which aren't properties
		@properties.unused.each do |child|
			if child.dir?
				case child.name
				# these first 2 will actually be of the form
				# 1\.0_#([0-9A-Z]{8}), where $1 is the 0 based index number in hex
				# should i parse that and use it as an index?
				when /__attach_version1\.0_/
					attach = Attachment.new(child)
					@attachments << attach if attach.valid?
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
				domain = (email[%r{^/O=([^/]+)}i, 1].downcase + '.com' rescue email)
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

	# redundant?
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
			# its actually possible for plain body to be empty, but the others not.
			# if i can get an html version, then maybe a callout to lynx can be made...
			mime.parts << Mime.new("Content-Type: text/plain\r\n\r\n" + props.body) if props.body
			mime.parts << Mime.new("Content-Type: text/html\r\n\r\n"  + props.body_html) if props.body_html
			#mime.parts << Mime.new("Content-Type: text/rtf\r\n\r\n"   + props.body_rtf)  if props.body_rtf
			# temporarily disabled the rtf. its just showing up as an attachment anyway.
			# for now, what i can do, is this:
			# (maybe i should just overload body_html do to this)
			# its thus currently possible to get no body at all if the only body is rtf. that is not
			# really acceptable
			if props.body_rtf and !props.body_html and props.body_rtf['htmltag']
				open('temp.html', 'w') { |f| f.write props.body_rtf }
				begin
					html = `#{SUPPORT_DIR}/rtf2html < temp.html`
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

	class Attachment
		attr_reader :obj, :properties
		alias props :properties

		def initialize obj
			@obj = obj
			@properties = Properties.load @obj
			@embedded_ole = nil
			@embedded_msg = nil

			@properties.unused.each do |child|
				# this is fairly messy stuff.
				if child.dir? and child.name =~ Properties::SUBSTG_RX and
					 $1 == '3701' and $2.downcase == '000d'
					@embedded_ole = child
					class << @embedded_ole
						def compobj
							return nil unless compobj = children.find { |child| child.name == "\001CompObj" }
							compobj.data[/^.{32}([^\x00]+)/m, 1]
						end

						def embedded_type
							temp = compobj and return temp
							# try to guess more
							if children.select { |child| child.name =~ /__(substg|properties|recip|attach|nameid)/ }.length > 2
								return 'Microsoft Office Outlook Message'
							end
							nil
						end
					end
					if @embedded_ole.embedded_type == 'Microsoft Office Outlook Message'
						@embedded_msg = Msg.new @embedded_ole
					end
				end
				# FIXME warn
			end
		end

		def valid?
			# something i started to notice when handling embedded ole object attachments is
			# the particularly strange case where they're are empty attachments
			props.raw.keys.length > 0
		end

		def filename
			props.attach_long_filename || props.attach_filename
		end

		def data
			@embedded_msg || @embedded_ole || props.attach_data
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
			"#<#{self.class.to_s[/\w+$/]}" +
				(filename ? " filename=#{filename.inspect}" : '') +
				(@embedded_ole ? " embedded_type=#{@embedded_ole.embedded_type.inspect}" : '') + ">"
		end
	end

	#
	# +Recipient+ serves as a container for the +recip+ directories in the .msg.
	# It has things like office_location, business_telephone_number, but I don't
	# think enough to make a vCard out of?
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

