#! /usr/bin/ruby

require 'rubygems'
require 'msg'
require 'enumerator'

class Pst
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

	class Header
		SIZE = 512
		MAGIC = 0x2142444e

		# these are the constants defined in libpst.c, that
		# are referenced in pst_open()
		INDEX_TYPE_OFFSET = 0x0A
		ENC_OFFSET = 0x1CD
		FILE_SIZE_POINTER = 0xA8
		FILE_SIZE_POINTER_64 = 0xB8
		INDEX_POINTER = 0xC4
		INDEX_POINTER_64 = 0xF0
		SECOND_POINTER = 0xBC
		SECOND_POINTER_64 = 0xE0

		attr_reader :magic, :index_type, :encrypt_type, :size
		attr_reader :index1_count, :index1, :index2_count, :index2
		def initialize data
			@magic = data.unpack('N')[0]
			@index_type = data[INDEX_TYPE_OFFSET]
			@encrypt_type = data[ENC_OFFSET]

			@index2_count, @index2 = data[SECOND_POINTER - 4, 8].unpack('V2')
			@index1_count, @index1 = data[INDEX_POINTER  - 4, 8].unpack('V2')

			@size = data[FILE_SIZE_POINTER, 4].unpack('V')[0]

			validate!
		end

		def validate!
			p self
			raise "bad signature on pst file (#{'0x%x' % @magic})" unless @magic == MAGIC
			raise "only index type 0xe is handled (#{'0x%x' % @index_type})" unless @index_type == 0x0e
			raise "only encrytion types 0 and 1 are handled (#{@encrypt_type.inspect})" unless [0, 1].include?(@encrypt_type)
		end
	end

	# compressible encryption! :D
	#
	# simple substitution. see libpst.c
	# maybe switch to using a .tr!
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

		def self.decrypt encrypted
			decrypted = ''
			encrypted.length.times { |i| decrypted << DECRYPT_TABLE[encrypted[i]] }
			decrypted
		end

		def self.encrypt decrypted
			encrypted = ''
			decrypted.length.times { |i| encrypted << ENCRYPT_TABLE[decrypted[i]] }
			encrypted
		end

		# an alternate implementation that is possibly faster....
		DECRYPT_STR, ENCRYPT_STR = [DECRYPT_TABLE, (0...256)].map do |values|
			values.map { |i| i.chr }.join.gsub(/([\^\-\\])/, "\\\\\\1")
		end

		def self.decrypt encrypted
			a = encrypted.tr ENCRYPT_STR, DECRYPT_STR
			#if a != decrypt2(encrypted)
			#	require 'irb'
			#	$a = encrypted
			#	IRB.start
			#	exit
			#end
			a
		end

		def self.encrypt2 decrypted
			decrypted.tr DECRYPT_STR, ENCRYPT_STR
		end
	end

	# more constants from libpst.c
	# these relate to the index block
	BLOCK_SIZE = 516 # index blocks
	DESC_BLOCK_SIZE = 516 # descriptor blocks was 520 but bogus
	ITEM_COUNT_OFFSET = 0x1f0 # count byte
	LEVEL_INDICATOR_OFFSET = 0x1f3 # node or leaf
	BACKLINK_OFFSET = 0x1f8 # backlink u1 value
	# these constants are in the classes
	#ITEM_SIZE = 12
	#DESC_SIZE = 16
	# i think these may simply be implied by the size of the blocks and the size
	# of the structures packed within. check it out
	# 516 / Index::SIZE == 43. and theres the header bit. but, given that that header
	# bit starts at 0x1f0, then 0x1f0 / 12 is really the max, which is 41. and, if we
	# assume that the record there is 8, not 12 bytes, then block size is 512, which
	# makes more sense. ie, that the header goes from 0x1f0 -> 512, which is a 16 byte
	# header record.
	# similarly, for the Desc - 0x1f0 / 16 -> 31. so theres nothing unexpected about
	# this. 
	# so its seems these are simply tightly packed, nested tree thingies.
	INDEX_COUNT_MAX = 41 # max active items
	DESC_COUNT_MAX = 31 # max active items

	attr_reader :io, :header, :idx, :desc, :special_folder_ids
	def initialize io
		@io = io
		io.pos = 0
		@header = Header.new io.read(Header::SIZE)

		load_idx
		load_desc
		load_xattrib

		@special_folder_ids = {}
	end

	def encrypted?
		@header.encrypt_type != 0
	end

	#
	# this is the index and desc record loading code
	# ----------------------------------------------------------------------------
	#

	# these 3 classes are used to hold various file records

	# i think the max size is 8192. i'm not sure how bigger
	# items are serialized. somehow, they are serialized as a table that lists a
	# bunch of ids. 
	class Index < Struct.new(:id, :offset, :size, :u1)
		UNPACK_STR = 'VVvv'
		SIZE = 12

		attr_accessor :pst
		def initialize data
			super(*data.unpack(UNPACK_STR))
		end

		def read decrypt=true
			# don't decrypt odd ids??. wait, that'd be & 1. this is weirder.
			decrypt = false if (id & 2) != 0
			pst.pst_read_block_size offset, size, decrypt
		end

		# show all numbers in hex
		def inspect
			super.gsub(/=(\d+)/) { '=0x%x' % $1.to_i }
		end
	end

	class TablePtr < Struct.new(:start, :u1, :offset)
		UNPACK_STR = 'V3'
		SIZE = 12

		def initialize data
			data = data.unpack(UNPACK_STR) if String === data
			super(*data)
		end
	end

	class Desc < Struct.new(:desc_id, :idx_id, :idx2_id, :parent_desc_id)
		UNPACK_STR = 'V4'
		SIZE = 16

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

	def load_idx
		@idx = []
		@idx_offsets = []
		load_idx_rec header.index1, 0, header.index1_count, 0, 1 << 31

		# we'll typically be accessing by id, so create a hash as a lookup cache
		@idx_from_id = {}
 		@idx.each do |idx|
			warn "there are duplicate idx records with id #{idx.id}" if @idx_from_id[idx.id]
			@idx_from_id[idx.id] = idx
		end
	end

	# most access to idx objects will use this function
	def idx_from_id id
		@idx_from_id[id]
	end

	# load the flat idx table, which maps ids to file ranges. this is the recursive helper
	def load_idx_rec offset, depth, linku1, start_val, end_val
		if end_val <= start_val
			warn "end <= start"
			return -1
		end
		@idx_offsets << offset

		#_pst_read_block_size(pf, offset, BLOCK_SIZE, &buf, 0, 0) < BLOCK_SIZE)
		buf = pst_read_block_size offset, BLOCK_SIZE, false, false 

		item_count = buf[ITEM_COUNT_OFFSET]
		raise "have too many active items in index (#{item_count})" if item_count > INDEX_COUNT_MAX

		idx = Index.new buf[BACKLINK_OFFSET, Index::SIZE]
		#p idx
		raise 'blah 1' unless idx.id == linku1

		if buf[LEVEL_INDICATOR_OFFSET] == 0
			# leaf pointers
			last = start_val
			item_count.times do |i|
				idx = Index.new buf[Index::SIZE * i, Index::SIZE]
				idx.pst = self
				#p idx
				break if idx.id == 0
				raise 'blah 2' unless (last...end_val) === idx.id
				last = idx.id
				# first entry
				raise 'blah 3' if i == 0 and start_val != 0 and idx.id != start_val
				@idx << idx
			end
		else
			# node pointers
			last = start_val
			item_count.times do |i|
				table = TablePtr.new buf[TablePtr::SIZE * i, TablePtr::SIZE]
				break if table.start == 0
				table2 = if i == item_count - 1
					TablePtr.new [end_val, nil, nil]
				else
					TablePtr.new buf[TablePtr::SIZE * (i + 1), TablePtr::SIZE]
				end
				raise 'blah 2' unless (last...end_val) === table.start
				last = table.start
				# first entry
				raise 'blah 3' if i == 0 and start_val != 0 and table.start != start_val
				load_idx_rec table.offset, depth + 1, table.u1, table.start, table2.start
			end
		end
	end

	def load_desc
		@desc = []
		@desc_offsets = []
		load_desc_rec header.index2, 0, header.index2_count, 0x21, 1 << 31

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
				warn "desc record's parent is itself (#{desc.inspect})"
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

	# as for idx
	def desc_from_id id
		@desc_from_id[id]
	end

	# load the flat list of desc records recursively
	def load_desc_rec offset, depth, linku1, start_val, end_val
		@desc_offsets << offset
		
		buf = pst_read_block_size offset, DESC_BLOCK_SIZE, false, false
		item_count = buf[ITEM_COUNT_OFFSET]

		# not real desc
		desc = Desc.new buf[BACKLINK_OFFSET, 4]
		raise 'blah 1' unless desc.desc_id == linku1

		if buf[LEVEL_INDICATOR_OFFSET] == 0
			# leaf pointers
			raise "have too many active items in index (#{item_count})" if item_count > DESC_COUNT_MAX
			last = start_val
			item_count.times do |i|
				desc = Desc.new buf[Desc::SIZE * i, Desc::SIZE]
				break if desc.desc_id == 0
				# this same thing is seen in the code all over the place. ensures that the series
				# is strictly increasing, and never greater than end_val.
				raise 'blah 2' unless (last...end_val) === desc.desc_id
				last = desc.desc_id
				# first entry
				raise 'blah 3' if i == 0 and start_val != 0 and desc.desc_id != start_val
				@desc << desc
			end
		else
			# node pointers
			raise "have too many active items in index (#{item_count})" if item_count > INDEX_COUNT_MAX
			last = start_val
			item_count.times do |i|
				table = TablePtr.new buf[TablePtr::SIZE * i, TablePtr::SIZE]
				break if table.start == 0
				table2 = if i == item_count - 1
					TablePtr.new [end_val, nil, nil]
				else
					TablePtr.new buf[TablePtr::SIZE * (i + 1), TablePtr::SIZE]
				end
				raise 'blah 2' unless (last...end_val) === table.start
				last = table.start
				# first entry
				raise 'blah 3' if i == 0 and start_val != -1 and table.start != start_val
				load_desc_rec table.offset, depth + 1, table.u1, table.start, table2.start
			end
		end
	end

	# ----------------------------------------------------------------------------

	class ID2Assoc < Struct.new(:id2, :id, :table2)
		UNPACK_STR = 'V3'
		SIZE = 12

		def initialize data
			data = data.unpack(UNPACK_STR) if String === data
			super(*data)
		end
	end

	def pst_builidx_id2 idx, idx2=0
		buf = pst_read_block_size idx.offset, idx.size, false, false
		type, count = buf.unpack 'v2'
		unless type == 0x0002
			warn 'unknown id2 type 0x%04x' % type
			return
		end
		id2 = []
		count.times do |i|
			assoc = ID2Assoc.new buf[4 + ID2Assoc::SIZE * i, ID2Assoc::SIZE]
			id2 << assoc
			if assoc.table2 != 0
				# FIXME implement recursive id2 table loading
				#warn "deeper id2 tables not handled."
			end
		end
		id2
	end

	def dump_debug_info
		puts "* pst header"
		p header

		# these 3 things, should account for most of the data in the file.
=begin
Looking at the output of this, for blank-o1997.pst, i see this part:
...
- (26624,516) desc block data (overlap of 4 bytes)
- (27136,516) desc block data (gap of 508 bytes)
- (28160,516) desc block data (gap of 2620 bytes)
...

which confirms my belief that the block size for idx and desc is more likely 512
=end
		puts '* file range usage'
		file_ranges =
			[[0, Header::SIZE, 'pst file header']] +
			@idx_offsets.map { |offset| [offset, BLOCK_SIZE - 4, 'idx block data'] } +
			@desc_offsets.map { |offset| [offset, DESC_BLOCK_SIZE - 4, 'desc block data'] } +
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
				when 0;  []
				when +1; ["gap of #{gap} bytes"]
			end
			# how about we check that padding
			@io.pos = offset + size
			pad_bytes = @io.read(pad)
			extra += ["padding not all zero"] unless pad_bytes == 0.chr * pad
			puts "- #{offset}:#{size}+#{pad} #{name.inspect}" + (extra.empty? ? '' : ' [' + extra * ', ' + ']')
		end
		#return
		#% idx.idsort_by { |idx| idx.offset }.each { |idx| puts "- #{idx.inspect}" }

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
		warn "skipping loading xattribs"
		# FIXME implement loading xattribs
	end

	def pst_read_block_size offset, size, decrypt=true, is_index=false
		old_pos = io.pos
		begin
			io.pos = offset
			buf = io.read size
			warn "tried to read #{size} bytes but only got #{buf.length}" if buf.length != size

			if is_index and buf[0] == 0x01 and buf[1] == 0x00
				warn "not doing weird index stuff..."
				# FIXME what was going on here. and what is the real difference between this
				# function, and all the _ff_read block functions. and what's the deal with the & 1 and & 2 stuff.
			end

			buf = CompressibleEncryption.decrypt buf if decrypt and encrypted?

			buf
		ensure
			io.pos = old_pos
		end
	end

	class Item
		class Properties < Msg::Properties
			def initialize list
				super()
				list.each { |type, ref_type, value| add_property type, value if type < 0x8000 }
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
			@properties = Properties.new list

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
			unless type
				type = valid_folder_mask || ipm_subtree_entryid || content_count || subfolders ? :folder : :item
				if type == :folder
					type = desc.pst.special_folder_ids[desc.desc_id] || type
				end
			end

			@type = type
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
				#p desc
				item = @desc.pst.pst_parse_item desc
				#p item
				item
			end
		end

		# these are still around because they do different stuff

		# Top of Personal Folder Record
		def ipm_subtree_entryid
			EntryID.new(props.ipm_subtree_entryid).id rescue nil
		end

		# Deleted Items Folder Record
		def ipm_wastebasket_entryid
			EntryID.new(props.ipm_wastebasket_entryid).id rescue nil
		end

		# Search Root Record
		def finder_entryid
			EntryID.new(props.finder_entryid).id rescue nil
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

		def method_missing name
			props.send name
		end

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

		def inspect
			attrs = %w[display_name subject sender_name subfolders]
#			attrs = %w[display_name valid_folder_mask ipm_wastebasket_entryid finder_entryid content_count subfolders]
			str = attrs.map { |a| b = send a; " #{a}=#{b.inspect}" if b }.compact * ','

			type_s = type == :item ? 'Item' : type == :folder ? 'Folder' : type.to_s.capitalize + 'Folder'
			str2 = 'desc_id=0x%x' % @desc.desc_id

			!str.empty? ? "#<Pst::#{type_s} #{str2}#{str}>" : "#<Pst::#{type_s} #{str2} props=#{props.inspect}>" #\n" + props.transport_message_headers + ">"
		end
	end

	class RecursiveStub
		include ToTree

		def children
			[]
		end

		def inspect
			"#<Pst::Item::RecursiveStub ...>"
		end
	end

	def pst_parse_item desc
		#puts "> parsing desc #{desc.desc_id}"

		@parsed ||= {}
		if @parsed[desc.desc_id]
			# must be recursive
			return RecursiveStub.new
		end
		@parsed[desc.desc_id] = true
	
		unless desc.desc
			warn "unable to parse #{desc.inspect}"
			return nil
		end
		id2 = if desc.list_index
			pst_builidx_id2 desc.list_index
		else
			# i don't think this is a weird case
			#warn "no list_index"
			nil
		end
		item = Item.new desc, pst_parse_block(desc.idx_id, id2)
	end

	def pst_parse_block block_id, id2=nil
		buf = idx_from_id(block_id).read

		index_offset, type, offset = buf.unpack 'vvV'
		#p [:block_header, index_offset, type, offset]

		case type
		when 0xbcec
			#p 
			block_type = 1
			#pst_getBlockOffsetPointer i2_head, buf, read_size, ind_ptr, block_hdr.offset, &block_offset1
			# the way that offset works, data1 may be a subset of buf, or something from id2. if its from buf,
			# it will be offset based on index_offset and offset. so it could be some random chunk of data anywhere
			# in the thing. 
			data1 = pst_getBlockOffsetPointer id2, buf, index_offset, offset
			warn 'not good' if data1.length < 8
			type, ref_type, value = data1.unpack 'vvV'
			raise "unhandled table rec type 0x%04x" % type if type != 0x02b5
			# this is actually a big chunk of tag tuples.
			data2 = pst_getBlockOffsetPointer id2, buf, index_offset, value
			num_list = data2.length / 8
			num_recs = 1
		when 0x7cec
			# seen this only in the parsing of an attachment record so far.
			raise 2
		when 0x0101
			raise 3
		else
			raise "unknown block type 0x%04x" % type
		end

		immediate_types = [0x0002, 0x0003, 0x000b]
		indirect_types = [0x0005, 0x000d, 0x0014, 0x001e, 0x0040, 0x0048, 0x0102, 0x1003, 0x01014, 0x101e, 0x1102]
		x = []

		num_recs.times do |rec|
			num_list.times do |list|
				if block_type == 1
					type, ref_type, value = data2[8 * list, 8].unpack 'vvV'
					if immediate_types.include? ref_type
						# the value is actually the value....
						# do some coercion
						case ref_type
						when 0x000b
							value = value != 0
						end
					elsif indirect_types.include? ref_type
						# the value is a pointer
						value = pst_getBlockOffsetPointer id2, buf, index_offset, value
					else
						# unhandled
						raise "unsupported ref_type %04x" % ref_type
					end
					x << [type, ref_type, value]
				end
			end
		end

		x
	end

	# the job of this class, is to take a desc record, and be able to enumerate through the
	# mapi properties of the associated thing.
	class BlockParser
		attr_reader :desc, :idx2, :data
		def initialize desc
			raise "unable to get associated index record for #{desc.inspect}" unless desc.desc
			@desc = desc
			# a given desc record may or may not have associated idx2 data.
			@idx2 = desc.pst.pst_builidx_id2 desc.list_index if desc.list_index
		end

		def load data
			@data = data


		item = Item.new desc, pst_parse_block(desc.idx_id, id2)
	# based on the value of offset, return either some data from buf, or some data from the
	# id2 chain id2, where offset is some key into a lookup table that is stored as the id2
	# chain. i think i may need to create a BlockParser class that wraps up all this mess.
	def get_data_indirect id2, buf, i_offset, offset
		if offset == 0
			nil
		elsif (offset & 0xf) == 0xf
			#warn "id2 value unhandled"
			warn 'called into id2 part of getBlockOffsetPointer but id2 undefined' unless id2
			pst_getID2block offset, id2
			#nil
		else
			low, high = offset & 0xf, offset >> 4
			raise 'blah' if i_offset == 0 or low != 0 or i_offset + 2 + high + 4 > buf.length
			from, to = buf[i_offset + 2 + high, 4].unpack 'v2'
			raise 'blah 2' if from > to
			buf[from...to]
		end
	end

	def pst_getBlockOffsetPointer id2, buf, i_offset, offset
		if offset == 0
			nil
		elsif (offset & 0xf) == 0xf
			#warn "id2 value unhandled"
			warn 'called into id2 part of getBlockOffsetPointer but id2 undefined' unless id2
			pst_getID2block offset, id2
			#nil
		else
			low, high = offset & 0xf, offset >> 4
			raise 'blah' if i_offset == 0 or low != 0 or i_offset + 2 + high + 4 > buf.length
			from, to = buf[i_offset + 2 + high, 4].unpack 'v2'
			raise 'blah 2' if from > to
			buf[from...to]
		end
	end

	def pst_getID2block id2, id2_head
		idx = idx_from_id pst_getID2(id2_head, id2)
		if (idx.id & 0x2) == 0
			idx.read
		else
			#warn 'weird compile id thing not handled yet'
			return nil
		end
	end

	# id2 is the array of id2assoc. id is the id2 value we're looking for
	def pst_getID2 id2, id
		id2.find { |x| x.id2 == id }.id rescue nil
	end

	def inspect
		"#<Pst name=#{(@name ||= root_item.display_name).inspect} io=#{io.inspect}>"
	end

	# this searches for occurences of 0xbcec, sees if it is part of the header signature of a block containing
	# mapi string tags, and then collates all that it finds. this is in order to dump a series of index block
	# offsets, which will help me find out how indexes are stored in outlook 2003 files. note that it works on
	# outlook 97 files too. could be part of a kind of recovery algo for block searching...
	def self.temp_hack buf, index_offset, offset
		if offset == 0
			#nil
			raise 'ignore'
		elsif (offset & 0xf) == 0xf
			raise 'id2 not handled'
		else
			low, high = offset & 0xf, offset >> 4
			# seems like the bottom 4 bits are either all set, or all not set.
			raise 'blah' if index_offset == 0 or low != 0 or i_offset + 2 + high + 4 > buf.length
			from, to = buf[index_offset + 2 + high, 4].unpack 'v2'
			raise 'blah 2' if from > to
			buf[from...to]
		end
	end

	# like String#index, but gives all matches
	def self.string_indexes string, substring
		offsets = []
		string.scan(/#{Regexp.quote substring}/m) { offsets << $~.begin(0) }
		offsets
	end

	def self.string_search filename
		# we hold the entire file in memory
		file_data = File.read filename
		orig_file_data = file_data
		# now search for bcec
		bcec_offsets = string_indexes file_data, [0xbcec].pack('v')
		if bcec_offsets.length < 3
			# lets try it again decrypted
			file_data = CompressibleEncryption.decrypt file_data
			bcec_offsets = string_indexes file_data, [0xbcec].pack('v')
		end
		# now go through those parsing them as blocks
		strings = []
		bcec_offsets.each do |bcec_offset|
			block_offset = bcec_offset - 2
			# 8192 is a reasonably large block size. also was the biggest i ever saw in '97 files
			block_data = file_data[block_offset, 8192]
			index_offset, type, offset = block_data.unpack 'vvV'
			signature, offset = temp_hack(block_data, index_offset, offset).unpack('VV') rescue next
			next unless signature == 0x000602b5
			next if offset > 8192
			index_data = temp_hack(block_data, index_offset, offset) rescue next
			index_data.scan(/.{8}/om).map do |chunk|
				tag, ref_type, value = chunk.unpack 'vvV'
				next unless ref_type == 0x001e or ref_type == 0x001f
				string = temp_hack(block_data, index_offset, value) rescue next
				string = Ole::Types::FROM_UTF16.iconv(string) if ref_type == 0x001f
				strings << [block_offset, '0x%04x' % tag, string] unless string.empty?
			end
		end
		require 'pp'
		offsets = strings.transpose[0].uniq.map do |offset|
			search = [offset].pack('v')
			indexes = string_indexes(orig_file_data, search)
			indexes = [] if indexes.length > 10
			[offset, indexes]
		end
		require 'narray'
		na = NArray[*offsets.transpose[1].flatten]
		mask = (na > (na.mean + na.stddev * 1.5)) | (na < (na.mean - na.stddev * 1.5))
		#p na.to_a
		na = na[mask.not]
		#p na.to_a
		require 'yaml'
		y :mean => na.mean, :min => na.min, :max => na.max, :count => na.length, :median => na.median, :stddev => na.stddev
		pp offsets
		strings
	end

	# trying to figure out outlook2003 pst file format
	def self.o2003_indexdump filename
		# we hold the entire file in memory
		file_data = File.read filename
		# how to find the 19456 magic number??
		#p file_data[19456 + 0x1f0, 16].unpack('c*')
		indexes = []
		file_data[19456, 512].scan(/.{24}/m).each do |index_data|
			# the size of the index records has been doubled in outlook2003 file format.
			x = index_data.unpack('V2V2VV')
			indexes << x
			# i think the size is bogus
			id1, id2, of1, of2, size, unk = x
			next if id1 == 0
			puts '* 0x%08x%08x 0x%08x%08x 0x%08x 0x%08x' % [id2, id1, of2, of1, size, unk]
			# this shows the corresponding block information. see magic signatures like bcec, and 7cec
			puts '  (0x%04x 0x%04x 0x%08x)' % file_data[of1, size].unpack('vvV')
		end
		# if the idx records are now 64 bit, we should be able to find the desc tree pretty easily.
		# first, we'll try to find which one has the top of folder string name
		text = Ole::Types::TO_UTF16.iconv 'Name of the Personal Folder'
		rx = /#{Regexp.quote text}/i
		indexes.each do |id1, id2, of1, of2, size, unk|
			next unless file_data[of1, 256] =~ rx
			p [id1, of1, size]
		end
		nil
	end
end

if $0 == __FILE__ #or true
	open 'test-o1997.pst' do |f|
#	open 'blank-o1997.pst' do |f|
#	open 'blank-o2003.pst' do |f|
		pst = Pst.new(f)
		pst.dump_debug_info
	end
end

