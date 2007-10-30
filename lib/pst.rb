#! /usr/bin/ruby

#
# = Introduction
#
# This file is mostly an attempt to port libpst to ruby, and simplify it in the process. It
# will leverage much of the existing MAPI => MIME conversion developed for Msg files, and as
# such is purely concerned with the file structure details.
#
# Already it works better for me on my outlook 97 psts, though a goal of the project is to
# support version 2003 psts also.
#
# = TODO
# 
# 1. xattribs
# 1.5 io streams like ruby-msg, with set of encoding / decoding functions.
# 1.75 fix the usage of Mapi:: stuff.
# 2. better tests
# 3. use tests to isolate any more of the encrypt vs non-encrypt issues, try and understand the & 1
#    & 2 stuff etc.
# 4. refactor index load
# 5. cleanup general. try to rationalise the code.
# 6. eml extraction - compare accuracy.
# 6.5 msg serialization.
# 7. outlook 2003
#

# restruct msg project to share more code. eg, perhaps something like:
# of course this is inverted at the moment, but msg will be changed transparently to pst
# to fix that. 

require 'rubygems'
require 'msg'
require 'enumerator'
require 'ostruct'

module Mapi
	class PropertyStore < Msg::Properties
		TAGS = MAPITAGS
	end

	class Item < Msg
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

=begin

Mapi::Pst.open('myfile.pst') do |pst|
	message = pst.find { |item| item.type == :message }
	message.class # Mapi::Pst::Item
	message.props # Mapi::Pst::PropertyStorage
	message.storage # => :pst
	p message.attachments.first
	message.to_mime
end

Mapi::Msg.open('myfile.msg') do |pst|
	message = pst.find { |item| item.type == :message }
	message.class # Mapi::Msg::Item
	message.storage # => :msg
	p message.recipients.first
	message.to_mime
end

# what about saving the pst to msg files:

Mapi::Pst.open 'myfile.pst' do |pst|
	pst.each_with_index do |item, i|
		# this will be the first cut of msg serialization. serialize an arbitrary
		# Mapi::Item, whether from a pst or not, to an '.msg' file. independent of propertystore
		# class etc.
		item.to_msg "item_#{i}.msg"
	end
end

where Mapi::Pst and ::Msg are just some sort of generalised property store style backends, and the
majority of the message code is independent of it. will make write access easier too. this way the
various Mapi => standards (rfc2822, vcard, etc etc) stuff can be leveraged across the whole thing.
whether you get a message by Msg.open('filename.msg') or
Pst.open('filename.pst').find { |msg| msg.subject =~ /blah blah/ }, either way you get a message
object that can be manipulated the same way. 

mapi property types (from http://msdn2.microsoft.com/en-us/library/bb147591.aspx)

value     mapi name       variant name    description
-------------------------------------------------------------------------------
0x0001    PT_NULL         VT_NULL         Null (no valid data)
0x0002    PT_SHORT        VT_I2           2-byte integer (signed)
0x0003    PT_LONG         VT_I4           4-byte integer (signed)
0x0004    PT_FLOAT        VT_R4           4-byte real (floating point)
0x0005    PT_DOUBLE       VT_R8           8-byte real (floating point)
0x0006    PT_CURRENCY     VT_CY           8-byte integer (scaled by 10, 000)
0x000A    PT_ERROR        VT_ERROR        SCODE value; 32-bit unsigned integer
0x000B    PT_BOOLEAN      VT_BOOL         Boolean
0x000D    PT_OBJECT       VT_UNKNOWN      Data object
0x001E/001F    PT_STRING8 VT_BSTR         String
0x0040    PT_SYSTIME      VT_DATE         8-byte real (date in integer, time in fraction)
0x0102    PT_BINARY       VT_BLOB         Binary (unknown format)
0x0102    PT_CLSID        VT_CLSID        OLE GUID

all the values except for the last 2, are the same value as the variant's value it seems.

ole variant types (from http://www.marin.clara.net/COM/variant_type_definitions.htm)

value		variant name
-------------------------------------------------------------------------------
0x0000  VT_EMPTY
0x0001  VT_NULL
0x0002  VT_I2
0x0003  VT_I4
0x0004  VT_R4
0x0005  VT_R8
0x0006  VT_CY
0x0007  VT_DATE
0x0008  VT_BSTR
0x0009  VT_DISPATCH
0x000a  VT_ERROR
0x000b  VT_BOOL
0x000c  VT_VARIANT
0x000d  VT_UNKNOWN
0x000e  VT_DECIMAL
0x0010  VT_I1
0x0011  VT_UI1
0x0012  VT_UI2
0x0013  VT_UI4
0x0014  VT_I8
0x0015  VT_UI8
0x0016  VT_INT
0x0017  VT_UINT
0x0018  VT_VOID
0x0019  VT_HRESULT
0x001a  VT_PTR
0x001b  VT_SAFEARRAY
0x001c  VT_CARRAY
0x001d  VT_USERDEFINED
0x001e  VT_LPSTR
0x001f  VT_LPWSTR
0x0040  VT_FILETIME
0x0041  VT_BLOB
0x0042  VT_STREAM
0x0043  VT_STORAGE
0x0044  VT_STREAMED_OBJECT
0x0045  VT_STORED_OBJECT
0x0046  VT_BLOB_OBJECT
0x0047  VT_CF
0x0048  VT_CLSID
0x0fff  VT_ILLEGALMASKED
0x0fff  VT_TYPEMASK
0x1000  VT_VECTOR
0x2000  VT_ARRAY
0x4000  VT_BYREF
0x8000  VT_RESERVED
0xffff  VT_ILLEGAL

=end

class Pst
	VERSION = '0.5.0'

	class FormatError < StandardError
	end

	# also used in ole/storage, and msg/mime. time to refactor
	module ToTree
		def to_tree
			# we want to call this only once
			children = self.children

			if children.empty?; "- #{inspect}\n"
			else
				str = "- #{inspect}\n"
				children.each_with_index do |child, i|
					last = i == children.length - 1
					child.to_tree.split(/\n/).each_with_index do |line, j|
						str << "  #{last ? (j == 0 ? "\\" : ' ') : '|'}#{line}\n"
					end
				end
				str
			end
		end
	end

	#
	# this is the header and encryption encapsulation code
	# ----------------------------------------------------------------------------
	#

	# class which encapsulates the pst header
	class Header
		SIZE = 512
		MAGIC = 0x2142444e

		# these are the constants defined in libpst.c, that
		# are referenced in pst_open()
		INDEX_TYPE_OFFSET = 0x0A
		FILE_SIZE_POINTER = 0xA8
		FILE_SIZE_POINTER_64 = 0xB8
		SECOND_POINTER = 0xBC
		INDEX_POINTER = 0xC4
		SECOND_POINTER_64 = 0xE0
		INDEX_POINTER_64 = 0xF0
		ENC_OFFSET = 0x1CD

		attr_reader :magic, :index_type, :encrypt_type, :size
		attr_reader :index1_count, :index1, :index2_count, :index2
		attr_reader :version
		def initialize data
			@magic = data.unpack('N')[0]
			@index_type = data[INDEX_TYPE_OFFSET]
			@version = @index_type == 0xe ? 1997 : 2003

			if version_2003?
				# don't know?
				@encrypt_type = 0

				@index2_count, @index2 = data[SECOND_POINTER_64 - 4, 8].unpack('V2')
				@index1_count, @index1 = data[INDEX_POINTER_64  - 4, 8].unpack('V2')

				@size = data[FILE_SIZE_POINTER_64, 4].unpack('V')[0]
			else
				@encrypt_type = data[ENC_OFFSET]

				@index2_count, @index2 = data[SECOND_POINTER - 4, 8].unpack('V2')
				@index1_count, @index1 = data[INDEX_POINTER  - 4, 8].unpack('V2')

				@size = data[FILE_SIZE_POINTER, 4].unpack('V')[0]
			end

			validate!
		end

		def version_2003?
			version == 2003
		end

		def encrypted?
			encrypt_type != 0
		end

		def validate!
			raise FormatError, "bad signature on pst file (#{'0x%x' % magic})" unless magic == MAGIC
			raise FormatError, "only index types 0xe and 0x17 are handled (#{'0x%x' % index_type})" unless [0x0e, 0x17].include?(index_type)
			raise FormatError, "only encrytion types 0 and 1 are handled (#{encrypt_type.inspect})" unless [0, 1].include?(encrypt_type)
		end
	end

	# compressible encryption! :D
	#
	# simple substitution. see libpst.c
	# maybe test switch to using a String#tr!
	class CompressibleEncryption
		DECRYPT_TABLE = [
			0x47, 0xf1, 0xb4, 0xe6, 0x0b, 0x6a, 0x72, 0x48,
			0x85, 0x4e, 0x9e, 0xeb, 0xe2, 0xf8, 0x94, 0x53, # 0x0f
			0xe0, 0xbb, 0xa0, 0x02, 0xe8, 0x5a, 0x09, 0xab,
			0xdb, 0xe3, 0xba, 0xc6, 0x7c, 0xc3, 0x10, 0xdd, # 0x1f
			0x39, 0x05, 0x96, 0x30, 0xf5, 0x37, 0x60, 0x82,
			0x8c, 0xc9, 0x13, 0x4a, 0x6b, 0x1d, 0xf3, 0xfb, # 0x2f
			0x8f, 0x26, 0x97, 0xca, 0x91, 0x17, 0x01, 0xc4,
			0x32, 0x2d, 0x6e, 0x31, 0x95, 0xff, 0xd9, 0x23, # 0x3f
			0xd1, 0x00, 0x5e, 0x79, 0xdc, 0x44, 0x3b, 0x1a,
			0x28, 0xc5, 0x61, 0x57, 0x20, 0x90, 0x3d, 0x83, # 0x4f
			0xb9, 0x43, 0xbe, 0x67, 0xd2, 0x46, 0x42, 0x76,
			0xc0, 0x6d, 0x5b, 0x7e, 0xb2, 0x0f, 0x16, 0x29, # 0x5f
			0x3c, 0xa9, 0x03, 0x54, 0x0d, 0xda, 0x5d, 0xdf,
			0xf6, 0xb7, 0xc7, 0x62, 0xcd, 0x8d, 0x06, 0xd3, # 0x6f
			0x69, 0x5c, 0x86, 0xd6, 0x14, 0xf7, 0xa5, 0x66,
			0x75, 0xac, 0xb1, 0xe9, 0x45, 0x21, 0x70, 0x0c, # 0x7f
			0x87, 0x9f, 0x74, 0xa4, 0x22, 0x4c, 0x6f, 0xbf,
			0x1f, 0x56, 0xaa, 0x2e, 0xb3, 0x78, 0x33, 0x50, # 0x8f
			0xb0, 0xa3, 0x92, 0xbc, 0xcf, 0x19, 0x1c, 0xa7,
			0x63, 0xcb, 0x1e, 0x4d, 0x3e, 0x4b, 0x1b, 0x9b, # 0x9f
			0x4f, 0xe7, 0xf0, 0xee, 0xad, 0x3a, 0xb5, 0x59,
			0x04, 0xea, 0x40, 0x55, 0x25, 0x51, 0xe5, 0x7a, # 0xaf
			0x89, 0x38, 0x68, 0x52, 0x7b, 0xfc, 0x27, 0xae,
			0xd7, 0xbd, 0xfa, 0x07, 0xf4, 0xcc, 0x8e, 0x5f, # 0xbf
			0xef, 0x35, 0x9c, 0x84, 0x2b, 0x15, 0xd5, 0x77,
			0x34, 0x49, 0xb6, 0x12, 0x0a, 0x7f, 0x71, 0x88, # 0xcf
			0xfd, 0x9d, 0x18, 0x41, 0x7d, 0x93, 0xd8, 0x58,
			0x2c, 0xce, 0xfe, 0x24, 0xaf, 0xde, 0xb8, 0x36, # 0xdf
			0xc8, 0xa1, 0x80, 0xa6, 0x99, 0x98, 0xa8, 0x2f,
			0x0e, 0x81, 0x65, 0x73, 0xe4, 0xc2, 0xa2, 0x8a, # 0xef
			0xd4, 0xe1, 0x11, 0xd0, 0x08, 0x8b, 0x2a, 0xf2,
			0xed, 0x9a, 0x64, 0x3f, 0xc1, 0x6c, 0xf9, 0xec  # 0xff
		]

		ENCRYPT_TABLE = [nil] * 256
		DECRYPT_TABLE.each_with_index { |i, j| ENCRYPT_TABLE[i] = j }

		def self.decrypt_alt encrypted
			decrypted = ''
			encrypted.length.times { |i| decrypted << DECRYPT_TABLE[encrypted[i]] }
			decrypted
		end

		def self.encrypt_alt decrypted
			encrypted = ''
			decrypted.length.times { |i| encrypted << ENCRYPT_TABLE[decrypted[i]] }
			encrypted
		end

		# an alternate implementation that is possibly faster....
		DECRYPT_STR, ENCRYPT_STR = [DECRYPT_TABLE, (0...256)].map do |values|
			values.map { |i| i.chr }.join.gsub(/([\^\-\\])/, "\\\\\\1")
		end

		def self.decrypt encrypted
			encrypted.tr ENCRYPT_STR, DECRYPT_STR
		end

		def self.encrypt decrypted
			decrypted.tr DECRYPT_STR, ENCRYPT_STR
		end
	end

	class RangesIOEncryptable < RangesIO
		def initialize io, ranges, opts={}
			@decrypt = !!opts[:decrypt]
			super
		end

		def encrypted?
			@decrypt
		end

		def read limit=nil
			buf = super
			buf = CompressibleEncryption.decrypt(buf) if encrypted?
			buf
		end
	end

	attr_reader :io, :header, :idx, :desc, :special_folder_ids

	# corresponds to
	# * pst_open
	# * pst_load_index
	def initialize io
		@io = io
		io.pos = 0
		@header = Header.new io.read(Header::SIZE)

		# would prefer this to be in Header#validate, but it doesn't have the io size.
		# should downgrade this to just be a warning...
		raise FormatError, "header size field invalid (#{header.size} != #{io.size}}" unless header.size == io.size

		load_idx
		load_desc
		load_xattrib

		@special_folder_ids = {}
	end

	def encrypted?
		@header.encrypted?
	end

	#
	# this is the index and desc record loading code
	# ----------------------------------------------------------------------------
	#

	# more constants from libpst.c
	# these relate to the index block
	ITEM_COUNT_OFFSET = 0x1f0 # count byte
	LEVEL_INDICATOR_OFFSET = 0x1f3 # node or leaf
	BACKLINK_OFFSET = 0x1f8 # backlink u1 value

	# these 3 classes are used to hold various file records

	# i think the max size is 8192. i'm not sure how bigger
	# items are serialized. somehow, they are serialized as a table that lists a
	# bunch of ids. 
	#
	# pst_index
	class Index < Struct.new(:id, :offset, :size, :u1)
		UNPACK_STR = 'VVvv'
		SIZE = 12
		BLOCK_SIZE = 512 # index blocks was 516 but bogus
		COUNT_MAX = 41 # max active items (ITEM_COUNT_OFFSET / Index::SIZE = 41)

		attr_accessor :pst
		def initialize data
			super(*data.unpack(UNPACK_STR))
		end

		def read decrypt=true
			# don't decrypt odd ids??. wait, that'd be & 1. this is weirder.
			# my tests fail without this line, so whatever its doing is important.
			decrypt = false if (id & 2) != 0
			pst.pst_read_block_size offset, size, decrypt
		end

		# show all numbers in hex
		def inspect
			super.gsub(/=(\d+)/) { '=0x%x' % $1.to_i }
		end
	end

	# _pst_table_ptr_struct
	class TablePtr < Struct.new(:start, :u1, :offset)
		UNPACK_STR = 'V3'
		SIZE = 12

		def initialize data
			data = data.unpack(UNPACK_STR) if String === data
			super(*data)
		end
	end

	# pst_desc
	class Desc < Struct.new(:desc_id, :idx_id, :idx2_id, :parent_desc_id)
		UNPACK_STR = 'V4'
		SIZE = 16
		BLOCK_SIZE = 512 # descriptor blocks was 520 but bogus
		COUNT_MAX = 31 # max active desc records (ITEM_COUNT_OFFSET / Desc::SIZE = 31)

		include ToTree

		attr_accessor :pst
		attr_reader :children
		def initialize data
			super(*data.unpack(UNPACK_STR))
			@children = []
		end

		def desc
			pst.idx_from_id idx_id
		end

		def list_index
			pst.idx_from_id idx2_id
		end

		# show all numbers in hex
		def inspect
			super.gsub(/=(\d+)/) { '=0x%x' % $1.to_i }
		end
	end

	# corresponds to
	# * _pst_build_id_ptr
	def load_idx
		@idx = []
		@idx_offsets = []
		load_idx_rec header.index1, header.index1_count, 0

		# we'll typically be accessing by id, so create a hash as a lookup cache
		@idx_from_id = {}
 		@idx.each do |idx|
			warn "there are duplicate idx records with id #{idx.id}" if @idx_from_id[idx.id]
			@idx_from_id[idx.id] = idx
		end
	end

	# load the flat idx table, which maps ids to file ranges. this is the recursive helper
	#
	# corresponds to
	# * _pst_build_id_ptr
	def load_idx_rec offset, linku1, start_val
		@idx_offsets << offset

		#_pst_read_block_size(pf, offset, BLOCK_SIZE, &buf, 0, 0) < BLOCK_SIZE)
		buf = pst_read_block_size offset, Index::BLOCK_SIZE, false, false 

		item_count = buf[ITEM_COUNT_OFFSET]
		raise "have too many active items in index (#{item_count})" if item_count > Index::COUNT_MAX

		idx = Index.new buf[BACKLINK_OFFSET, Index::SIZE]
		raise 'blah 1' unless idx.id == linku1

		if buf[LEVEL_INDICATOR_OFFSET] == 0
			# leaf pointers
			# split the data into item_count index objects
			buf[0, Index::SIZE * item_count].scan(/.{#{Index::SIZE}}/mo).each_with_index do |data, i|
				idx = Index.new data
				# first entry
				raise 'blah 3' if i == 0 and start_val != 0 and idx.id != start_val
				idx.pst = self
				# this shouldn't really happen i'd imagine
				break if idx.id == 0
				@idx << idx
			end
		else
			# node pointers
			# split the data into item_count table pointers
			buf[0, TablePtr::SIZE * item_count].scan(/.{#{TablePtr::SIZE}}/mo).each_with_index do |data, i|
				table = TablePtr.new data
				# for the first value, we expect the start to be equal
				raise 'blah 3' if i == 0 and start_val != 0 and table.start != start_val
				# this shouldn't really happen i'd imagine
				break if table.start == 0
				load_idx_rec table.offset, table.u1, table.start
			end
		end
	end

	# most access to idx objects will use this function
	#
	# corresponds to
	# * _pst_getID
	def idx_from_id id
		@idx_from_id[id]
	end

	# corresponds to
	# * _pst_build_desc_ptr
	# * record_descriptor
	def load_desc
		@desc = []
		@desc_offsets = []
		load_desc_rec header.index2, header.index2_count, 0x21

		# first create a lookup cache
		@desc_from_id = {}
 		@desc.each do |desc|
			desc.pst = self
			warn "there are duplicate desc records with id #{desc.desc_id}" if @desc_from_id[desc.desc_id]
			@desc_from_id[desc.desc_id] = desc
		end

		# now turn the flat list of loaded desc records into a tree

		# well, they have no parent, so they're more like, the toplevel descs.
		@orphans = []
		# now assign each node to the parents child array, putting the orphans in the above
		@desc.each do |desc|
			parent = @desc_from_id[desc.parent_desc_id]
			# note, besides this, its possible to create other circular structures.
			if parent == desc
				# this actually happens usually, for the root_item it appears.
				#warn "desc record's parent is itself (#{desc.inspect})"
			# maybe add some more checks in here for circular structures
			elsif parent
				parent.children << desc
				next
			end
			@orphans << desc
		end

		# maybe change this to some sort of sane-ness check. orphans are expected
#		warn "have #{@orphans.length} orphan desc record(s)." unless @orphans.empty?
	end

	# load the flat list of desc records recursively
	#
	# corresponds to
	# * _pst_build_desc_ptr
	# * record_descriptor
	def load_desc_rec offset, linku1, start_val
		@desc_offsets << offset
		
		buf = pst_read_block_size offset, Desc::BLOCK_SIZE, false, false
		item_count = buf[ITEM_COUNT_OFFSET]

		# not real desc
		desc = Desc.new buf[BACKLINK_OFFSET, 4]
		raise 'blah 1' unless desc.desc_id == linku1

		if buf[LEVEL_INDICATOR_OFFSET] == 0
			# leaf pointers
			raise "have too many active items in index (#{item_count})" if item_count > Desc::COUNT_MAX
			# split the data into item_count desc objects
			buf[0, Desc::SIZE * item_count].scan(/.{#{Desc::SIZE}}/mo).each_with_index do |data, i|
				desc = Desc.new data
				# first entry
				raise 'blah 3' if i == 0 and start_val != 0 and desc.desc_id != start_val
				# this shouldn't really happen i'd imagine
				break if desc.desc_id == 0
				@desc << desc
			end
		else
			# node pointers
			raise "have too many active items in index (#{item_count})" if item_count > Index::COUNT_MAX
			# split the data into item_count table pointers
			buf[0, TablePtr::SIZE * item_count].scan(/.{#{TablePtr::SIZE}}/mo).each_with_index do |data, i|
				table = TablePtr.new data
				# for the first value, we expect the start to be equal
				raise 'blah 3' if i == 0 and start_val != -1 and table.start != start_val
				# this shouldn't really happen i'd imagine
				break if table.start == 0
				load_desc_rec table.offset, table.u1, table.start
			end
		end
	end

	# as for idx
	# 
	# corresponds to:
	# * _pst_getDptr
	def desc_from_id id
		@desc_from_id[id]
	end

	# corresponds to
	# * pst_load_extended_attributes
	def load_xattrib
		unless desc = desc_from_id(0x61)
			warn "no extended attributes desc record found"
			return
		end
		unless desc.desc
			warn "no desc idx for extended attributes"
			return
		end
		if desc.list_index
		end
		#warn "skipping loading xattribs"
		# FIXME implement loading xattribs
	end

	# corresponds to:
	# * _pst_read_block_size
	# * _pst_read_block ??
	# * _pst_ff_getIDblock_dec ??
	# * _pst_ff_getIDblock ??
	def pst_read_block_size offset, size, decrypt=true, is_index=false
		io.seek offset
		buf = io.read size
		warn "tried to read #{size} bytes but only got #{buf.length}" if buf.length != size

		if is_index and buf[0] == 0x01 and buf[1] == 0x00
			warn "not doing weird index stuff..."
			# FIXME what was going on here. and what is the real difference between this
			# function, and all the _ff_read block functions. and what's the deal with the & 1 and & 2 stuff.
		end

		encrypted? && decrypt ? CompressibleEncryption.decrypt(buf) : buf
	end

	#
	# id2 
	# ----------------------------------------------------------------------------
	#

	class ID2Assoc < Struct.new(:id2, :id, :table2)
		UNPACK_STR = 'V3'
		SIZE = 12

		def initialize data
			data = data.unpack(UNPACK_STR) if String === data
			super(*data)
		end
	end

	def load_idx2 idx
		load_idx2_rec idx
	end

	# corresponds to
	# * _pst_build_id2
	def load_idx2_rec idx
		buf = pst_read_block_size idx.offset, idx.size, false, false
		type, count = buf.unpack 'v2'
		unless type == 0x0002
			raise 'unknown id2 type 0x%04x' % type
			#return
		end
		id2 = []
		count.times do |i|
			assoc = ID2Assoc.new buf[4 + ID2Assoc::SIZE * i, ID2Assoc::SIZE]
			id2 << assoc
			if assoc.table2 != 0
				id2 += load_idx2_rec idx_from_id(assoc.table2)
			end
		end
		id2
	end

	class RangesIOIdxChain < RangesIOEncryptable
		def initialize pst, idx_head
			@idxs = pst.id2_block_idx_chain idx_head
			# whether or not a given idx needs encrypting
			decrypts = @idxs.map do |idx|
				decrypt = (idx.id & 2) != 0 ? false : pst.encrypted?
			end.uniq
			raise NotImplementedError, 'partial encryption in RangesIOID2' if decrypts.length > 1
			decrypt = decrypts.first
			# convert idxs to ranges
			ranges = @idxs.map { |idx| [idx.offset, idx.size] }
			super pst.io, ranges, :decrypt => decrypt
		end
	end

	class RangesIOID2 < RangesIOIdxChain
		def self.new pst, id2, id2_head
			RangesIOIdxChain.new pst, pst.idx_from_id(pst.pst_getID2(id2_head, id2))
		end
	end

	# corresponds to:
	# * _pst_ff_getID2block
	# * _pst_ff_getID2data
	# * _pst_ff_compile_ID
	def id2_block_idx_chain idx
		if (idx.id & 0x2) == 0
			[idx]
		else
			buf = idx.read
			type, fdepth, count = buf[0, 4].unpack 'CCv'
			unless type == 1 # libpst.c:3958
				warn 'Error in idx_chain - %p, %p, %p - attempting to ignore' % [type, fdepth, count]
				return [idx]
			end
			# there are 4 unaccounted for bytes here, 4...8
			ids = buf[8, count * 4].unpack('V*')
			if fdepth == 1
				ids.map { |id| idx_from_id id }
			else
				ids.map { |id| pst_getID2block_idxs_rec id }.flatten
			end
		end
	end

	# id2 is the array of id2assoc. id is the id2 value we're looking for
	#
	# corresponds to:
	# * _pst_getID2
	def pst_getID2 id2, id
		id2 = id2.find { |x| x.id2 == id }
		id2 and id2.id
	end

	#
	# main block parsing code. gets raw properties
	# ----------------------------------------------------------------------------
	#

	# the job of this class, is to take a desc record, and be able to enumerate through the
	# mapi properties of the associated thing.
	#
	# corresponds to
	# * _pst_parse_block
	# * _pst_process (in some ways. although perhaps thats more the Item::Properties#add_property)
	class BlockParser
		TYPES = {
			0xbcec => 1,
			0x7cec => 2,
			#0x0101 => 3
		}

		PR_SUBJECT = Msg::Properties::MAPITAGS.find { |num, (name, type)| name == 'PR_SUBJECT' }.first.hex
		PR_BODY_HTML = Msg::Properties::MAPITAGS.find { |num, (name, type)| name == 'PR_BODY_HTML' }.first.hex

		# this stuff could maybe be moved to Ole::Types? or leverage it somehow?
		IMMEDIATE_TYPES = [0x0002, 0x0003, 0x000b]
		INDIRECT_TYPES = [
			0x0005, 0x000d, 0x0014, 0x001e, 0x0040, 0x0048, 0x0102, 0x1003, 0x01014, 0x101e, 0x1102
		]

		# the attachment and recipient arrays appear to be always stored with these fixed
		# id2 values. seems strange. are there other extra streams? can find out by making higher
		# level IO wrapper, which has the id2 value, and doing the diff of available id2 values versus
		# used id2 values in properties of an item.
		ID2_ATTACHMENTS = 0x671
		ID2_RECIPIENTS = 0x692

		attr_reader :desc, :data
		def initialize desc
			raise FormatError, "unable to get associated index record for #{desc.inspect}" unless desc.desc
			@desc = desc
			#@data = desc.desc.read
			if Pst::Index === desc.desc
				@data = RangesIOIdxChain.new(desc.pst, desc.desc).read
			else
				# fake desc
				@data = desc.desc.read
			end
			load_header
		end

		# a given desc record may or may not have associated idx2 data. we lazily load it here, so it will never
		# actually be requested unless get_data_indirect actually needs to use it.
		def idx2
			return @idx2 if @idx2
			raise FormatError, 'idx2 requested but no idx2 available' unless desc.list_index
			# should check this can't return nil
			@idx2 = desc.pst.load_idx2 desc.list_index
		end

		def load_header
			@index_offset, type, @offset1 = data.unpack 'vvV'
=begin
			# not a real type. the 0x0101 is for a nested idx.
			if !TYPES[type] and @index_offset == 0x0101
				# libpst doesn't really handle this case properly, it just reads and discards the data.
				warn "parsing block type 3!!!! offset1 = 0x%04x" % @offset1 # <- want to see if this value is constant.
				# offset1 values: 0x3a5c
				# apparently, type will actually be a count of records. check that
				# hmm, it appears that it is used for RawPropertyStoreTable, where the count is high. eg, for
				# the recipients, i had a message with 41 recipients, and it triggered this. perhaps any time when
				# the block is going to be too big in some fashion.
				expect = (data.length - 8) / 4
				warn "expected type field to be #{expect}, but got #{type}" if type != expect
				@data = data[8, 4 * type].unpack('V*').map { |id| desc.pst.idx_from_id(id).read }.join
				$data = @data
				$obj = self
				# this doesn't seem to work quite right yet
				return load_header
			end
=end
			raise FormatError, 'unknown block type signature 0x%04x' % type unless TYPES[type]
			@type = TYPES[type]
		end

		# based on the value of offset, return either some data from buf, or some data from the
		# id2 chain id2, where offset is some key into a lookup table that is stored as the id2
		# chain. i think i may need to create a BlockParser class that wraps up all this mess.
		#
		# corresponds to:
		# * _pst_getBlockOffsetPointer
		# * _pst_getBlockOffset
		def get_data_indirect offset
			if offset == 0
				nil
			elsif (offset & 0xf) == 0xf
				#warn "id2 value unhandled"
				#warn 'called into id2 part of getBlockOffsetPointer but id2 undefined' unless idx2
				#desc.pst.pst_getID2block offset, idx2
				#nil
				RangesIOID2.new(desc.pst, offset, idx2).read
			else
				low, high = offset & 0xf, offset >> 4
				raise FormatError if @index_offset == 0 or low != 0 or @index_offset + 2 + high + 4 > data.length
				from, to = data[@index_offset + 2 + high, 4].unpack 'v2'
				raise FormatError if from > to
				data[from...to]
			end
		end

		def handle_indirect_values key, type, value
			case type
			when 0x000b
				value = value != 0
			when *IMMEDIATE_TYPES # not including 0x000b which we just did
				# the value is actually the value....
			when *INDIRECT_TYPES
				# the value is a pointer
				#p '0x%04x %04x' % [type, ref_type]
				#begin
				if String === value # ie, value size > 4 above
					value = StringIO.new value
				else
					value = get_data_indirect_io(value)
				end
				value = value.read if value and type == 0x001e or type == 0x001f
				#rescue 
				#	puts $!
				#	value = :novalue
				#end
				# special subject handling
				if key == PR_BODY_HTML and value
					# to keep the msg code happy, which thinks body_html will be an io
					value = StringIO.new value
				end
				if key == PR_SUBJECT and value
					ignore, offset = value.unpack 'C2'
					offset = (offset == 1 ? nil : offset - 3)
					value = value[2..-1]
=begin
					index = value =~ /^[A-Z]*:/ ? $~[0].length - 1 : nil
					unless ignore == 1 and offset == index
						warn 'something wrong with subject hack' 
						$x = [ignore, offset, value]
						require 'irb'
						IRB.start
						exit
					end
=end
=begin
new idea:

making sense of the \001\00[156] i've seen prefixing subject. i think its to do with the placement
of the ':', or the ' '. And perhaps an optimization to do with thread topic, and ignoring the prefixes
added by mailers. thread topic is equal to subject with all that crap removed.

can test by creating some mails with bizarre subjects.

subject="\001\005RE: blah blah"
subject="\001\001blah blah"
subject="\001\032Out of Office AutoReply: blah blah"
subject="\001\020Undeliverable: blah blah"

looks like it

=end

					# now what i think, is that perhaps, value[offset..-1] ...
					# or something like that should be stored as a special tag. ie, do a double yield
					# for this case. probably PR_CONVERSATION_TOPIC, in which case i'd write instead:
					# yield [PR_SUBJECT, ref_type, value]
					# yield [PR_CONVERSATION_TOPIC, ref_type, value[offset..-1]
					# next # to skip the yield.
				end

				# special handling for embedded objects
				# used for attach_data for attached messages. in which case attach_method should == 5,
				# for embedded object.
				if type == 0x000d and value
					value = value.read if value.respond_to?(:read)
					id2, unknown = value.unpack 'V2'
					io = RangesIOID2.new desc.pst, id2, idx2

					# hacky
					desc2 = OpenStruct.new(:desc => io, :pst => desc.pst, :list_index => desc.list_index, :children => [])
					# put nil instead of desc.list_index, otherwise the attachment is attached to itself ad infinitum.
					# should try and fix that FIXME
					value = Item.new desc2, RawPropertyStore.new(desc2).to_a
					desc2.list_index = nil
				end
			else
				# unhandled
				raise FormatError, 'unsupported ref_type %04x' % ref_type
			end
			[key, type, value]
		end

		def get_data_indirect_io offset
			if offset == 0
				nil
			elsif (offset & 0xf) == 0xf
				#warn "id2 value unhandled"
				#warn 'called into id2 part of getBlockOffsetPointer but id2 undefined' unless idx2
				#desc.pst.pst_getID2block offset, idx2
				#nil
				RangesIOID2.new desc.pst, offset, idx2
			else
				low, high = offset & 0xf, offset >> 4
				raise FormatError if @index_offset == 0 or low != 0 or @index_offset + 2 + high + 4 > data.length
				from, to = data[@index_offset + 2 + high, 4].unpack 'v2'
				raise FormatError if from > to
				StringIO.new data[from...to]
			end
		end
	end

=begin
Two items that currently break attachments and recipients

* recipients:

	affects: ["0x200764", "0x2011c4", "0x201b24", "0x201b44", "0x201ba4", "0x201c24", "0x201cc4", "0x202504"]

* attachments:

	this is just from the 0x000d object
	fixed now

=end

	# RawPropertyStore is used to iterate through the properties of an item, or the auxiliary
	# data for an attachment. its just a parser for the way the properties are serialized, when the
	# properties don't have to conform to a column structure.
	class RawPropertyStore < BlockParser
		include Enumerable

		attr_reader :length
		def initialize desc
			super
			raise FormatError, "expected type 1 - got #{@type}" unless @type == 1

			# the way that offset works, data1 may be a subset of buf, or something from id2. if its from buf,
			# it will be offset based on index_offset and offset. so it could be some random chunk of data anywhere
			# in the thing.
			header_data = get_data_indirect @offset1
			raise FormatError if header_data.length < 8
			signature, offset2 = header_data.unpack 'V2'
			raise FormatError, 'unhandled block signature 0x%08x' % type if signature != 0x000602b5
			# this is actually a big chunk of tag tuples.
			@index_data = get_data_indirect offset2
			@length = @index_data.length / 8
		end

		# iterate through the property tuples
		def each
			length.times do |i|
				key, type, value = handle_indirect_values(*@index_data[8 * i, 8].unpack('vvV'))
				yield key, type, value
			end
		end
	end

	# RawPropertyStoreTable is kind of like a database table.
	# it has a fixed set of columns.
	# #[] is kind of like getting a row from the table.
	# those rows are currently encapsulated by Row, which has #each like
	# RawPropertyStore.
	# only used for the recipients array, and the attachments array. completely lazy, doesn't
	# load any of the properties upon creation. 
	class RawPropertyStoreTable < BlockParser
		include Enumerable

		attr_reader :length, :index_data, :data2, :data3
		def initialize desc
			super
			raise FormatError, "expected type 2 - got #{@type}" unless @type == 2

			header_data = get_data_indirect @offset1
			# seven_c_blk
			# often: u1 == u2 and u3 == u2 + 2, then rec_size == u3 + 4. wtf
			seven_c, @num_list, u1, u2, u3, @rec_size, b_five_offset,
				ind2_offset, u7, u8 = header_data[0, 22].unpack('CCv4V2v2')
			raise FormatError unless seven_c == 0x7c
			@index_data = header_data[22..-1]
			warn 'Something looks wrong' unless @num_list == (@index_data.length / 8)

			header_data2 = get_data_indirect b_five_offset
			raise FormatError if header_data2.length < 8
			signature, offset2 = header_data2.unpack 'V2'
			raise FormatError, 'unhandled block signature 0x%08x' % type if signature != 0x000204b5

			# this holds all the row data
			@data3 = get_data_indirect ind2_offset

			# there must be something to the data in data2. i think data2 is the array of objects essentially.
			# currently its only used to imply a length
			@data2 = get_data_indirect offset2
			if data2
				@length = (data2.length / 6.0).ceil
			else
				# hmmm, actually, we can still figure it out:
				@length = @data3.length / @rec_size
			end
		end

		def [] idx
			Row.new self, idx * @rec_size
		end

		def each
			length.times { |i| yield self[i] }
		end

		class Row
			include Enumerable

			def initialize array_parser, x
				@array_parser, @x = array_parser, x
			end

			# iterate through the property tuples
			def each
				(@array_parser.index_data.length / 8).times do |i|
					ref_type, type, ind2_off, size, slot = @array_parser.index_data[8 * i, 8].unpack 'v3CC'
					# check this rescue too
					value = @array_parser.data3[@x + ind2_off, size]
#					if INDIRECT_TYPES.include? ref_type
					if size <= 4
						value = value.unpack('V')[0]
					end
					#p ['0x%04x' % ref_type, '0x%04x' % type, (Msg::Properties::MAPITAGS['%04x' % type].first[/^.._(.*)/, 1].downcase rescue nil),
					#		value_orig, value, (get_data_indirect(value_orig.unpack('V')[0]) rescue nil), size, ind2_off, slot]
					key, type, value = @array_parser.handle_indirect_values type, ref_type, value
					yield key, type, value
				end
			end
		end
	end

	class AttachmentTable < BlockParser
		# a "fake" MAPI property name for this constant. if you get a mapi property with
		# this value, it is the id2 value to use to get attachment data.
		PR_ATTACHMENT_ID2 = 0x67f2

		attr_reader :desc, :table
		def initialize desc
			@desc = desc
			# no super, we only actually want BlockParser2#idx2
			@table = nil
			return unless desc.list_index
			return unless id = desc.pst.pst_getID2(idx2, ID2_ATTACHMENTS)
			return unless idx = desc.pst.idx_from_id(id)
			# FIXME make a fake desc.
			@desc2 = OpenStruct.new :desc => idx, :pst => desc.pst, :list_index => desc.list_index
			@table = RawPropertyStoreTable.new @desc2
		end

		def to_a
			return [] if !table
			table.map do |attachment|
				attachment = attachment.to_a
				# potentially merge with yet more properties
				if attachment_id2 = attachment.assoc(PR_ATTACHMENT_ID2)
					idx = desc.pst.idx_from_id desc.pst.pst_getID2(idx2, attachment_id2.last)
					@desc2.desc = idx
					RawPropertyStore.new(@desc2).each do |a, b, c|
						record = attachment.assoc a
						attachment << record = [] unless record
						record.replace [a, b, c]
					end
				end
				attachment
			end
		end
	end

	# there is no equivalent to this in libpst. ID2_RECIPIENTS was just guessed given the above
	# AttachmentTable.
	class RecipientTable < BlockParser
		attr_reader :desc, :table
		def initialize desc
			@desc = desc
			# no super, we only actually want BlockParser2#idx2
			@table = nil
			return unless desc.list_index
			return unless id = desc.pst.pst_getID2(idx2, ID2_RECIPIENTS)
			return unless idx = desc.pst.idx_from_id(id)
			# FIXME make a fake desc.
			desc2 = OpenStruct.new :desc => idx, :pst => desc.pst, :list_index => desc.list_index
			@table = RawPropertyStoreTable.new desc2
		end

		def to_a
			return [] if !table
			table.map { |x| x.to_a }
		end
	end

	#
	# higher level item code. wraps up the raw properties above, and gives nice
	# objects to work with. handles item relationships too.
	# ----------------------------------------------------------------------------
	#

	class Item < Mapi::Item
		class PropertyStore < Mapi::PropertyStore
			def add_property key, type, value
				super key, value if key < 0x8000
			end
		end

		class Attachment < Mapi::Item::Attachment
			def initialize list
				@properties = PropertyStore.new
				list.each { |a, b, c| @properties.add_property a, b, c }

				@embedded_msg = props.attach_data if Item === props.attach_data
			end
		end

		class Recipient < Mapi::Item::Recipient
			def initialize list
				@properties = PropertyStore.new
				list.each { |a, b, c| @properties.add_property a, b, c }
			end
		end

		class EntryID < Struct.new(:u1, :entry_id, :id)
			UNPACK_STR = 'VA16V'

			def initialize data
				data = data.unpack(UNPACK_STR) if String === data
				super(*data)
			end
		end

		include ToTree

		attr_reader :properties
		attr_accessor :type
		def initialize desc, list, type=nil
			@desc = desc
			@properties = PropertyStore.new
			list.each { |a, b, c| @properties.add_property a, b, c }

			# this is kind of weird, but the ids of the special folders are stored in a hash
			# when the root item is loaded
			if ipm_wastebasket_entryid
				desc.pst.special_folder_ids[ipm_wastebasket_entryid] = :wastebasket
			end

			if finder_entryid
				desc.pst.special_folder_ids[finder_entryid] = :finder
			end

			# and then here, those are used, along with a crappy heuristic to determine if we are an
			# item
=begin
i think the low bits of the desc_id can give some info on the type.

it seems that 0x4 is for regular messages (and maybe contacts etc)
0x2 is for folders, and 0x8 is for special things like rules etc, that aren't visible.
=end
			unless type
				type = props.valid_folder_mask || ipm_subtree_entryid || props.content_count || props.subfolders ? :folder : :item
				if type == :folder
					type = desc.pst.special_folder_ids[desc.desc_id] || type
				end
			end

			@type = type

			super props
		end

		alias props :properties

		def children
			id = ipm_subtree_entryid
			if id
				root = @desc.pst.desc_from_id id
				raise "couldn't find root" unless root
				raise 'both kinds of children' unless @desc.children.empty?
				children = root.children
				# lets look up the other ids we have.
				# typically the wastebasket one "deleted items" is in the children already, but
				# the search folder isn't.
				extras = [ipm_wastebasket_entryid, finder_entryid].compact.map do |id|
					root = @desc.pst.desc_from_id id
					warn "couldn't find root for id #{id}" unless root
					root
				end.compact
				# i do this instead of union, so as not to mess with the order of the
				# existing children.
				children += (extras - children)
				children
			else
				@desc.children
			end.map do |desc|
				@desc.pst.pst_parse_item desc
			end
		end

		# these are still around because they do different stuff

		# Top of Personal Folder Record
		def ipm_subtree_entryid
			@ipm_subtree_entryid ||= EntryID.new(props.ipm_subtree_entryid.read).id rescue nil
		end

		# Deleted Items Folder Record
		def ipm_wastebasket_entryid
			@ipm_wastebasket_entryid ||= EntryID.new(props.ipm_wastebasket_entryid.read).id rescue nil
		end

		# Search Root Record
		def finder_entryid
			@finder_entryid ||= EntryID.new(props.finder_entryid.read).id rescue nil
		end

		# all these have been replaced with the method_missing below
=begin
		# States which folders are valid for this message store 
		#def valid_folder_mask
		#	props[0x35df]
		#end

		# Number of emails stored in a folder
		def content_count
			props[0x3602] 
		end

		# Has children
		def subfolders
			props[0x360a]
		end
=end

		# i think i will change these, so they can inherit the lazyness from RawPropertyStoreTable.
		# so if you want the last attachment, you can get it without creating the others perhaps.
		# it just has to handle the no table at all case a bit more gracefully.

		def attachments
			@attachments ||= AttachmentTable.new(@desc).to_a.map { |list| Attachment.new list }
		end

		def recipients
			[]
			#@recipients ||= RecipientTable.new(@desc).to_a.map { |list| Recipient.new list }
		end

		def each_recursive(&block)
			p :self => self
			children.each do |child|
				p :child => child
				block[child]
				child.each_recursive(&block)
			end
		end

		def inspect
			attrs = %w[display_name subject sender_name subfolders]
#			attrs = %w[display_name valid_folder_mask ipm_wastebasket_entryid finder_entryid content_count subfolders]
			str = attrs.map { |a| b = props.send a; " #{a}=#{b.inspect}" if b }.compact * ','

			type_s = type == :item ? 'Item' : type == :folder ? 'Folder' : type.to_s.capitalize + 'Folder'
			str2 = 'desc_id=0x%x' % @desc.desc_id

			!str.empty? ? "#<Pst::#{type_s} #{str2}#{str}>" : "#<Pst::#{type_s} #{str2} props=#{props.inspect}>" #\n" + props.transport_message_headers + ">"
		end
	end

	# corresponds to
	# * _pst_parse_item
	def pst_parse_item desc
		Item.new desc, RawPropertyStore.new(desc).to_a
	end

	#
	# other random code
	# ----------------------------------------------------------------------------
	#

	def dump_debug_info
		puts "* pst header"
		p header

=begin
Looking at the output of this, for blank-o1997.pst, i see this part:
...
- (26624,516) desc block data (overlap of 4 bytes)
- (27136,516) desc block data (gap of 508 bytes)
- (28160,516) desc block data (gap of 2620 bytes)
...

which confirms my belief that the block size for idx and desc is more likely 512
=end
		if 0 + 0 == 1
			puts '* file range usage'
			file_ranges =
				# these 3 things, should account for most of the data in the file.
				[[0, Header::SIZE, 'pst file header']] +
				@idx_offsets.map { |offset| [offset, Index::BLOCK_SIZE, 'idx block data'] } +
				@desc_offsets.map { |offset| [offset, Desc::BLOCK_SIZE, 'desc block data'] } +
				@idx.map { |idx| [idx.offset, idx.size, 'data for idx id=0x%x' % idx.id] }
			(file_ranges.sort_by { |idx| idx.first } + [nil]).to_enum(:each_cons, 2).each do |(offset, size, name), next_record|
				# i think there is a padding of the size out to 64 bytes
				# which is equivalent to padding out the final offset, because i think the offset is 
				# similarly oriented
				pad_amount = 64
				warn 'i am wrong about the offset padding' if offset % pad_amount != 0
				# so, assuming i'm not wrong about that, then we can calculate how much padding is needed.
				pad = pad_amount - (size % pad_amount)
				pad = 0 if pad == pad_amount
				gap = next_record ? next_record.first - (offset + size + pad) : 0
				extra = case gap <=> 0
					when -1; ["overlap of #{gap.abs} bytes)"]
					when  0; []
					when +1; ["gap of #{gap} bytes"]
				end
				# how about we check that padding
				@io.pos = offset + size
				pad_bytes = @io.read(pad)
				extra += ["padding not all zero"] unless pad_bytes == 0.chr * pad
				puts "- #{offset}:#{size}+#{pad} #{name.inspect}" + (extra.empty? ? '' : ' [' + extra * ', ' + ']')
			end
		end

		# i think the idea of the idx, and indeed the idx2, is just to be able to
		# refer to data indirectly, which means it can get moved around, and you just update
		# the idx table. it is simply a list of file offsets and sizes.
		# not sure i get how id2 plays into it though....
		# the sizes seem to be all even. is that a co-incidence? and the ids are all even. that
		# seems to be related to something else (see the (id & 2) == 1 stuff)
		puts '* idx entries'
		@idx.each { |idx| puts "- #{idx.inspect}" }

		# if you look at the desc tree, you notice a few things:
		# 1. there is a desc that seems to be the parent of all the folders, messages etc.
		#    it is the one whose parent is itself.
		#    one of its children is referenced as the subtree_entryid of the first desc item,
		#    the root.
		# 2. typically only 2 types of desc records have idx2_id != 0. messages themselves,
		#    and the desc with id = 0x61 - the xattrib container. everything else uses the
		#    regular ids to find its data. i think it should be reframed as small blocks and
		#    big blocks, but i'll look into it more.
		#
		# idx_id and idx2_id are for getting to the data. desc_id and parent_desc_id just define
		# the parent <-> child relationship, and the desc_ids are how the items are referred to in
		# entryids.
		# note that these aren't unique! eg for 0, 4 etc. i expect these'd never change, as the ids
		# are stored in entryids. whereas the idx and idx2 could be a bit more volatile.
		puts '* desc tree'
		# make a dummy root hold everything just for convenience
		root = Desc.new ''
		def root.inspect; "#<Pst::Root>"; end
		root.children.replace @orphans
		puts root.to_tree.gsub(/, (parent_desc_id|idx2_id)=0x0(?!\d)/, '')

		# this is fairly easy to understand, its just an attempt to display the pst items in a tree form
		# which resembles what you'd see in outlook.
		puts '* item tree'
		puts root_item.to_tree
	end

	def root_desc
		@desc.first
	end

	def root_item
		item = pst_parse_item root_desc
		item.type = :root
		item
	end

	def root
		root_item
	end

	# depth first search of all items
	include Enumerable

	def each(&block)
		root = self.root
		block[root]
		root.each_recursive(&block)
	end

	def name
		@name ||= root_item.props.display_name
	end
	
	def inspect
		"#<Pst name=#{name.inspect} io=#{io.inspect}>"
	end
end

if $0 == __FILE__
	require 'test/unit'

	class TestPst2 < Test::Unit::TestCase
		attr_reader :pst
		def setup
			@io = open 'test2-o1997.pst', 'rb'
			@pst = Pst.new @io
		end

		def teardown
			@io.close
		end

		def test_recipients
			message = pst.root_item.children.select { |child| child.type == :item }.first
			recipients = message.recipients
			assert_equal 3, recipients.length
		end

		# test a particular message
		def test_message
			folder = pst.root_item.children.select { |child| child.type == :folder }.first
			children = folder.children
			assert_equal 328, children.length
			sub_folder = folder.children.select { |child| child.type == :folder }.first
			message = sub_folder.children.last
			attachment = message.attachments.first
			# FIXME look into why these are different. which one is correct
			assert_equal 4086955, attachment.props.attach_size
			assert_equal 4086712, attachment.props.attach_data.size
			# the file is spread out in 500 chunks across the pst.
			# some of those chunks are possibly contiguous though. TODO - implement
			# optional merging of consecutive ranges in RangesIO.
			assert_equal 500, attachment.props.attach_data.ranges.length
		end

		def test_attached_message
			message = pst.pst_parse_item pst.desc_from_id(0x205524)
			assert_equal 1, message.attachments.length
			attachment = message.attachments.first
			assert_equal 5, attachment.props.attach_method # embedded object
			assert_equal Pst::Item, attachment.props.attach_data.class
			assert_equal 'message/rfc822', attachment.props.attach_mime_tag
			message2 = attachment.props.attach_data
			# this should generate an attached email message
			assert_equal <<-'end', message.to_mime.to_tree
- #<Mime content_type="multipart/mixed">
  |- #<Mime content_type="text/plain">
  \- #<Mime content_type="message/rfc822">
			end
		end
	end

	class TestPst3 < Test::Unit::TestCase
		attr_reader :pst
		def setup
			@io = open 'test3-o1997.pst', 'rb'
			@pst = Pst.new @io
		end

		def teardown
			@io.close
		end

		def test_pst_structure
			assert_equal 'Personal Folders', pst.name
			assert_equal 47, pst.idx.length
			assert_equal 34, pst.desc.length
			assert_equal 3, pst.root_item.children.length # trash, message, search
		end

		def test_message_properties
			message = pst.root_item.children[1]
			expected_properties = {
				# bools
				:alternate_recipient_allowed => true,
				:read_receipt_requested => false,
				:delete_after_submit => false,
				:originator_delivery_report_requested => false,

				# numbers
				:priority => 0,
				:sensitivity => 0,
				:importance => 1,
				:internet_cpid => 20127,
				:action => 4294967295,
				:message_size => 39974,
				:message_flags => 25,
				:profile_offline_store_path => 746,

				# strings
				:conversation_topic => 'draft message test with attachment',
				:subject => 'draft message test with attachment',
				:last_modifier_name => 'Charles Lowe',
				:message_class => 'IPM.Note',
				:display_to => 'nobody in particular',
				:body => "012346789\r\n" * 1000 + "\r\n \r\n",
				:body_html => "<!DOCTYPE HTML PUBLIC \"-//W3C//DTD HTML 4.0 Transitional//EN\">\r\n" \
											"<HTML><HEAD>\r\n<META http-equiv=Content-Type content=\"text/html; charset=us-ascii\">\r\n" \
											"<META content=\"MSHTML 6.00.2900.3199\" name=GENERATOR></HEAD>\r\n" \
											"<BODY><FONT face=Arial \r\nsize=2>" + '012346789<BR>' * 1000 +
											"</FONT></FONT>\r\n<DIV><FONT face=Arial size=2></FONT>&nbsp;</DIV></BODY></HTML>\r\n",

				# other
				# are these entry ids?
				:predecessor_change_list => "\024R\237\307\010\203\357\331J\265\231[\006s\254\373U\000\000x}",
				:change_key => "R\237\307\010\203\357\331J\265\231[\006s\254\373U\000\000x}",
				:search_key=>"PJ\204O\244\225\374N\224\376\3568\347\325\037\230",
				:sentmail_entryid => "\000\000\000\000\330\t\305\234!\302\vD\263\351\223[\217d;.\001\000\357\372\346\t\270fcL\274X\241\201\322\367\000\306\000\000\001!F\301\000\000",
				# this will get parsed as a time later.
				:last_modification_time => "\000\214ahX\024\310\001",
				:creation_time=>"\000\214ahX\024\310\001",
				:message_delivery_time => "\300\344J\247X\024\310\001",
			}

			# check hashes have the same keys
			assert_equal expected_properties.keys.sort_by { |k| k.to_s }, message.props.to_h.keys.sort_by { |k| k.to_s }
			# assert each component is equal. this is for easier to read failures.
			expected_properties.each do |key, value|
				got = message.props[key]
				got = got.read if got.respond_to? :read
				assert_equal value, got, "#{key} message property"
			end

			assert_equal 1, message.attachments.length
			attachment = message.attachments.first
			expected_properties = {
				# these are probably bugs. shouldn't have nils.
				:attach_content_location => nil,
				:attach_content_id => nil,
				:attach_long_pathname => nil,
				:attach_tag => nil,
				:attach_additional_info => nil,
				:attach_mime_tag => nil,
				:attach_pathname => nil,

				# bools
				:attachment_hidden => false,

				# numbers
				:attach_method => 1,
				:attach_flags => 0,
				:attachment_linkid => 0,
				:attach_size => 14589,
				:attachment_flags => 0,
				:rendering_position => 4294967295,

				# strings
				:attach_extension => '.txt',
				:display_name => 'x.txt',
				:attach_long_filename => 'x.txt',
				:attach_filename => 'x.txt',

				# should be changed to not be a string.
				:attach_data => "012346789\r\n" * 1000,
				:attach_rendering => /.*/, # just ignore this. 
				:attach_encoding => nil,
				:creation_time => "\034s\345zX\024\310\001",
				:exception_endtime => "\000@\335\243WE\263\f",
				:exception_starttime => "\000@\335\243WE\263\f",
				:last_modification_time => "\034s\345zX\024\310\001",
			}

			# check hashes have the same keys
			assert_equal expected_properties.keys.sort_by { |k| k.to_s }, attachment.props.to_h.keys.sort_by { |k| k.to_s }
			# assert each component is equal. this is for easier to read failures.
			expected_properties.each do |key, value|
				got = attachment.props[key]
				got = got.read if got.respond_to? :read
				send "assert_#{Regexp === value ? :match : :equal}", value, got, "#{key} attachment property" 
			end

			assert_equal 1, message.recipients.length
			recipient = message.recipients.first
			expected_properties = {
				:"7bit_display_name" => nil,

				:responsibility => true,
				:send_rich_info => true,

				:object_type => 6,
				:send_internet_encoding => 0,
				:display_type => 0,
				:recipient_type => 1,

				:display_name => 'nobody in particular',
				:email_address => 'nobody@nobody.com',
				:addrtype => 'SMTP',
				:search_key => "SMTP:NOBODY@NOBODY.COM\000",

				# email address is encoded in here too
				:entryid => "\000\000\000\000\201+\037\244\276\243\020\031\235n\000\335\001\017T\002\000\000\001\220n\000o\000b\000o\000d\000y\000 \000i\000n\000 \000p\000a\000r\000t\000i\000c\000u\000l\000a\000r\000\000\000S\000M\000T\000P\000\000\000n\000o\000b\000o\000d\000y\000@\000n\000o\000b\000o\000d\000y\000.\000c\000o\000m\000\000\000",
				:record_key => "\000\000\000\000\201+\037\244\276\243\020\031\235n\000\335\001\017T\002\000\000\001\220n\000o\000b\000o\000d\000y\000 \000i\000n\000 \000p\000a\000r\000t\000i\000c\000u\000l\000a\000r\000\000\000S\000M\000T\000P\000\000\000n\000o\000b\000o\000d\000y\000@\000n\000o\000b\000o\000d\000y\000.\000c\000o\000m\000\000\000",
			}

			# check hashes have the same keys
			assert_equal expected_properties.keys.sort_by { |k| k.to_s }, recipient.props.to_h.keys.sort_by { |k| k.to_s }
			# assert each component is equal. this is for easier to read failures.
			expected_properties.each do |key, value|
				got = recipient.props[key]
				got = got.read if got.respond_to? :read
				assert_equal value, got, "#{key} recipient property" 
			end
		end
	end
end

=begin

* test3-o1997.pst. idx records whose data contains '01234':

[#<struct Pst::Index id=240, offset=34112, size=8180, u1=2>,  <- attachment part 1
 #<struct Pst::Index id=248, offset=85568, size=2820, u1=2>,  <- attachment part 2
 #<struct Pst::Index id=296, offset=92032, size=8180, u1=2>,  <- body_html part 1
 #<struct Pst::Index id=304, offset=100224, size=5142, u1=2>, <- body_html part 2
 #<struct Pst::Index id=328, offset=105408, size=8180, u1=2>, <- body part 1
 #<struct Pst::Index id=336, offset=113600, size=2825, u1=2>] <- body part 2

# there is no Q that is always little endian, so use 2 V each
# 24 chunks
blocks.map do |s|
	id, size, offset = s.unpack('V6').to_enum(:each_slice, 2).map { |low, high| low + (high << 32) }
end
^ 64 bit values.

=end


