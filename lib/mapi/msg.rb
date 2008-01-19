require 'rubygems'
require 'ole/storage'
require 'mapi'
require 'mapi/rtf'
require 'mapi/msg/properties'
require 'mime'

module Mapi
	#
	# = Introduction
	#
	# Primary class interface to the vagaries of .msg files.
	#
	# The core of the work is done by the <tt>Msg::PropertyStore</tt> class.
	#
	class Msg < Message
		VERSION = '1.4.0'

		# these 2 will actually be of the form
		# 1\.0_#([0-9A-Z]{8}), where $1 is the 0 based index number in hex
		# should i parse that and use it as an index, or just return in
		# file order? probably should use it later...
		ATTACH_RX = /^__attach_version1\.0_.*/
		RECIP_RX = /^__recip_version1\.0_.*/
		VALID_RX = /#{PropertyStore::VALID_RX}|#{ATTACH_RX}|#{RECIP_RX}/

		attr_reader :root
		attr_accessor :close_parent

		# Alternate constructor, to create an +Msg+ directly from +arg+ and +mode+, passed
		# directly to Ole::Storage (ie either filename or seekable IO object).
		def self.open arg, mode=nil
			msg = new Ole::Storage.open(arg, mode).root
			# we will close the ole when we are #closed
			msg.close_parent = true
			if block_given?
				begin   yield msg
				ensure; msg.close
				end
			else msg
			end
		end

		# Create an Msg from +root+, an <tt>Ole::Storage::Dirent</tt> object
		def initialize root
			@root = root
			@close_parent = false
			super PropertySet.new(PropertyStore.load(@root).raw)
			Msg.warn_unknown @root
		end

		def self.warn_unknown obj
			# bit of validation. not important if there is extra stuff, though would be
			# interested to know what it is. doesn't check dir/file stuff.
			unknown = obj.children.reject { |child| child.name =~ VALID_RX }
			Log.warn "skipped #{unknown.length} unknown msg object(s)" unless unknown.empty?
		end

		def close
			@root.ole.close if @close_parent
		end

		def attachments
			@attachments ||= @root.children.
				select { |child| child.dir? and child.name =~ ATTACH_RX }.
				map { |child| Attachment.new child }.
				select { |attach| attach.valid? }
		end

		def recipients
			@recipients ||= @root.children.
				select { |child| child.dir? and child.name =~ RECIP_RX }.
				map { |child| Recipient.new child }
		end

		def inspect
			#str = %w[from to cc bcc subject type].map do |key|
			#	send(key) and "#{key}=#{send(key).inspect}"
			#end.compact.join(' ')
			"#<Msg ...>" ##{str}>"
		end

		class Attachment < Mapi::Attachment
			attr_reader :obj, :properties
			alias props :properties

			def initialize obj
				@obj = obj
				@properties = Properties.load @obj
				@embedded_ole = nil
				@embedded_msg = nil

				super PropertySet.new(PropertyStore.load(@obj).raw)
				Msg.warn_unknown @obj

				@obj.children.each do |child|
					# temp hack. PropertyStore doesn't do directory properties atm - FIXME
					if child.dir? and child.name =~ Properties::SUBSTG_RX and
						 $1 == '3701' and $2.downcase == '000d'
						@embedded_ole = child
						class << @embedded_ole
							def compobj
								return nil unless compobj = self["\001CompObj"]
								compobj.read[/^.{32}([^\x00]+)/m, 1]
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
				end
			end

			def valid?
				# something i started to notice when handling embedded ole object attachments is
				# the particularly strange case where there are empty attachments
				not props.raw.keys.empty?
			end
		end

		#
		# +Recipient+ serves as a container for the +recip+ directories in the .msg.
		# It has things like office_location, business_telephone_number, but I don't
		# think enough to make a vCard out of?
		#
		class Recipient < Mapi::Recipient
			attr_reader :obj, :properties
			alias props :properties

			def initialize obj
				@obj = obj
				super PropertySet.new(PropertyStore.load(@obj).raw)
				Msg.warn_unknown @obj
			end
		end
	end
end

