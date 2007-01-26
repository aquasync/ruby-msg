#! /usr/bin/ruby -w

require 'iconv'
require 'date'
require 'support'

#
# = Introduction
#
# +RangesIO+ is a basic class for wrapping another IO object allowing you to arbitrarily reorder
# slices of the input file by providing a list of ranges. Intended as an initial measure to curb
# inefficiencies in the OleDir#data method just reading all of a file's data in one hit, with
# no method to stream it.
# 
# This class will encapuslate the ranges (corresponding to big or small blocks) of any ole file
# and thus allow reading/writing directly to the source bytes, in a streamed fashion (so just
# getting 16 bytes doesn't read the whole thing).
#
# In the simplest case it can be used with a single range to provide a limited io to a section of
# a file.
#
# = Limitations
#
# * Writing code not written yet. easy enough
# * May be useful to have a facility for writing more than initially allocated ranges? or provide
#   a way for a provider of ranges to catch that and transparently provide a new sink io, that
#   can be read from at serialization time... Or allocate blocks straight away and provide ranges?
# * No buffering. by design at the moment. Intended for large reads
#
class RangesIO
	attr_reader :io, :ranges, :size, :pos
	# +io+ is the parent io object that we are wrapping.
	# 
	# +ranges+ are byte offsets, either
	# 1. an array of ranges [1..2, 4..5, 6..8] or
	# 2. an array of arrays, where the second is length [[1, 1], [4, 1], [6, 2]] for the above
	#    (think the way String indexing works)
	# The +ranges+ provide sequential slices of the file that will be read. they can overlap.
	def initialize io, ranges, opts={}
		@opts = {:close_parent => false}.merge opts
		@io = io
		# convert ranges to arrays. check for negative ranges?
		@ranges = ranges.map { |r| Range === r ? [r.begin, r.end - r.begin] : r }
		# calculate size
		@size = @ranges.inject(0) { |total, (pos, len)| total + len }
		# initial position in the file
		@pos = 0
	end

	def seek pos, whence=IO::SEEK_SET
		# just a simple pos calculation. invalidate buffers if we had them
		@pos = pos
		# FIXME
	end
	alias pos= :seek

	def close
		@io.close if @opts[:close_parent]
	end

	def range_and_offset pos
		off = nil
		r = ranges.inject(0) do |total, r|
			to = total + r[1]
			if pos <= to
				off = pos - total
				break r
			end
			to
		end
		# should be impossible for any valid pos, (0...size) === pos
		raise "unable to find range for pos #{pos.inspect}" unless off
		return r, off
	end

	def eof?
		@pos == @size
	end

	# read bytes from file, to a maximum of +limit+, or all available if unspecified.
	def read limit=nil
		data = ''
		limit ||= size
		# special case eof
		return data if eof?
		r, off = range_and_offset @pos
		i = ranges.index r
		# this may be conceptually nice (create sub-range starting where we are), but
		# for a large range array its pretty wasteful. even the previous way was. but
		# i'm not trying to optimize this atm. it may even go to c later if necessary.
		([[r[0] + off, r[1] - off]] + ranges[i+1..-1]).each do |pos, len|
			@io.seek pos
			if limit < len
				# FIXME this += isn't correct if there is a read error
				# or something.
				@pos += limit
				break data << @io.read(limit) 
			end
			data << @io.read(len)
			@pos += len
			limit -= len
		end
		data
	end

	# this will be generalised to a module later
	def each_read blocksize=4096
		yield read(blocksize) until eof?
	end

	# write should look fairly similar to the above.
	
	def inspect
		# the rescue is for empty files
		pos, len = *(range_and_offset(@pos)[0] rescue [nil, nil])
		range_str = pos ? "#{pos}..#{pos+len}" : 'nil'
		"#<RangesIO io=#{io.inspect} size=#@size pos=#@pos "\
			"current_range=#{range_str}>"
	end
end

module Ole # :nodoc:
	Log = Logger.new_with_callstack

	# 
	# = Introduction
	#
	# <tt>Ole::Storage</tt> is a simple class intended to abstract away details of the
	# access to OLE2 structured storage files, such as those produced by
	# Microsoft Office, eg *.doc, *.msg etc.
	#
	# Initially based on chicago's libole, source available at
	# http://prdownloads.sf.net/chicago/ole.tgz
	# Later augmented with some corrections by inspecting pole, and (purely
	# for header definitions) gsf.
	#
	# = Usage
	#
	# Usage should be fairly straight forward:
	#
	#   # get the parent ole storage object
	#   ole = Ole::Storage.load open('myfile.msg')
	#   # => #<Ole::Storage io=#<File:myfile.msg> root=#<OleDir:"Root Entry"
	#          size=2816>>
	#   # get the top level root object and output a tree structure for
	#   # debugging
	#   puts ole.root.to_tree
	#   # =>
	#   - #<OleDir:"Root Entry" size=3840 time="2006-11-03T00:52:53Z">
	#     |- #<OleDir:"__nameid_version1.0" size=0 time="2006-11-03T00:52:53Z">
	#     |  |- #<OleDir:"__substg1.0_00020102" size=16 data="CCAGAAAAAADAAA...">
	#     ...
	#     |- #<OleDir:"__substg1.0_8002001E" size=4 data="MTEuMA==">
	#     |- #<OleDir:"__properties_version1.0" size=800 data="AAAAAAAAAAABAA...">
	#     \- #<OleDir:"__recip_version1.0_#00000000" size=0 time="2006-11-03T00:52:53Z">
	#        |- #<OleDir:"__substg1.0_0FF60102" size=4 data="AAAAAA==">
	#   	 ...
	#
	# = TODO
	#
	# 1. Some sort of streamed access to data, for scalability. (initial attempt done).
	# 2. Other accessors for +OleDir+'s, such as #each, and <tt>#[]</tt> taking index
	#    and a relative string path. (partially done)
	#    Maybe consider using the '/' operator, like, eg hpricot:
	#      blah = ole.root/'__nameid_version1.0'/'__substg1.0_00020102'
	#    good solution if '/' is a legal character, excluding ole.root['__nameid.../...']
	#    but should also be a single operator version, such that a full path is a single object.
	# 3. Create/Update capability.
	#

	class Storage
		VERSION = '1.0.10'
		# All +OleDir+ names are in UTF16, which we convert
		UTF16_TO_UTF8 = Iconv.new('utf-8', 'utf-16le').method :iconv

		# The top of the ole tree structure
		attr_reader :root
		# The tree structure in its original flattened form
		attr_reader :dirs
		# The underlying io object to/from which the ole object is serialized
		attr_reader :io
		# Low level internals, not generally useful
		attr_reader :header, :bbat, :sbat, :sb_blocks

		# Note that creation of new ole objects not properly supported as yet
		def initialize
		end

		# A short cut
		def self.load io
			ole = Storage.new
			ole.load io
			ole
		end

		# Load an ole document.
		# +io+ needs to be seekable.
		def load io
			# we always read 512 for the header block. if the block size ends up being different,
			# what happens to the 109 fat entries. are there more/less entries?
			@io = io
			@io.seek 0
			header_block = @io.read 512
			@header = Header.load header_block

			bbat_chain_data =
				header_block[Header::SIZE..-1] +
				read_big_blocks((0...@header.num_mbat).map { |i| i + @header.mbat_start })
			@bbat = AllocationTable.load self, bbat_chain_data.unpack('L*')[0, @header.num_bat]
			@sbat = AllocationTable.load self, @bbat.chain(@header.sbat_start)

			# get block chain for directories and load them
			@dirs = read_big_blocks(@bbat.chain(@header.dirent_start)).scan(/.{#{OleDir::SIZE}}/mo).
			# semantics mayn't be quite right. used to cut at first dir where dir.type == 0
				map { |str| OleDir.load str }.reject { |dir| dir.type_id == 0 }
			@dirs.each { |dir| dir.ole = self }
			#p @dirs

			# now reorder from flat into a tree
			# links are stored in some kind of balanced binary tree
			# check that everything is visited at least, and at most once
			# similarly with the blocks of the file.
			class << @dirs
				def to_tree idx=0
					return [] if idx == (1 << 32) - 1
					dir = self[idx]
					dir.children = to_tree dir.child
					raise "directory #{dir.inspect} used twice" if dir.idx
					dir.idx = idx
					to_tree(dir.prev) + [dir] + to_tree(dir.next)
				end
			end

			@root = @dirs.to_tree.first
			Log.warn "root name was #{@root.name.inspect}" unless @root.name == 'Root Entry'
			@sb_blocks = @bbat.chain @root.first_block

			unused = @dirs.reject { |dir| dir.idx }.length
			Log.warn "* #{unused} unused directories" if unused > 0
		end

		# Read a chain (an array given by +blocks+) of big blocks, optionally
		# truncating to +size+.
		# Big blocks are of size Ole::Storage::Header#b_size, and are stored
		# linearly.
		def read_big_blocks blocks, size=nil
			block_size = 1 << @header.b_shift
			data = ''
			blocks.each do |block|
				@io.seek block_size * (block + 1)
				data << @io.read(block_size)
			end
			data = data[0, size] if size and size < data.length
			data
		end

		# Read a chain (an array given by +blocks+) of small blocks, optionally
		# truncating to +size+.
		# Small blocks are of size Ole::Storage::Header#s_size, and are stored
		# as a single file, serialized using big blocks. Single blocks are
		# mapped to big blocks using Ole::Storage#sb_blocks
		# 
		# pretty much deprecated. should be refactored, with the above, into a
		# +small_blocks_to_ranges+ type function, that then gets passed to the RangesIO for reading
		# and writing.
		def read_small_blocks blocks, size=nil
			data = ''
			blocks.each do |block|
				# this tries to efficiently map a small block file to its position in the parent file.
				idx, pos = (block * (1 << @header.s_shift)).divmod 1 << @header.b_shift
				pos += (1 << @header.b_shift) * (@sb_blocks[idx] + 1)
				@io.seek pos
				data << @io.read(1 << @header.s_shift)
			end
			data = data[0, size] if size and size < data.length
			data
		end

		def inspect
			"#<#{self.class} io=#{@io.inspect} root=#{@root.inspect}>"
		end

		# A class which wraps the ole header
		class Header
			MEMBERS = [
				:magic, :clsid, :minor_ver, :major_ver, :byte_order, :b_shift, :s_shift,
				:reserved, :csectdir, :num_bat, :dirent_start, :transacting_signature, :threshold,
				:sbat_start, :num_sbat, :mbat_start, :num_mbat
			]
			PACK = 'a8 a16 S2 a2 S2 a6 L3 a4 L6'
			SIZE = 0x4c
			MAGIC = "\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # expected value of Header#magic

			# 2 basic initializations, from scratch, or from a data string.
			# from scratch will be geared towards creating a new ole object
			def initialize
				@values = []
			end

			def self.load str
				h = Header.new
				h.to_a.replace str.unpack(PACK)
				h.validate!
				h
			end

			def validate!
				raise "OLE2 signature is invalid" unless magic == MAGIC
				if num_bat == 0 or # is that valid for a completely empty file?
					 # not sure about this one. basically to do max possible bat given size of mbat
					 num_bat > 109 && num_bat > 109 + num_mbat * (1 << b_shift - 2) or
					 # shouldn't need to use the mbat as there is enough space in the header block
					 num_bat < 109 && num_mbat != 0 or
					 # given the size of the header is 76, if b_shift <= 6, blocks address the header.
					 s_shift > b_shift or b_shift <= 6 or b_shift >= 31 or
					 # we only handle little endian
					 byte_order != "\xfe\xff"
					raise "not valid OLE2 structured storage file"
				end
				# relaxed this, due to test-msg/qwerty_[1-3]*.msg they all had
				# 3 for this value. 
				# transacting_signature != "\x00" * 4 or
				if threshold != 4096 or
					 reserved != "\x00" * 6
					Log.warn "may not be a valid OLE2 structured storage file"
				end
				true
			end

			def b_size
				@b_size ||= 1 << b_shift
			end

			def s_size
				@s_size ||= 1 << s_shift
			end

			MEMBERS.each_with_index do |sym, i|
				define_method(sym) { @values[i] }
				define_method(sym.to_s + '=') { |val| @values[i] = val }
			end

			def to_a
				@values
			end

			def inspect
				"#<#{self.class} " +
					MEMBERS.zip(@values).map { |k, v| "#{k}=#{v.inspect}" }.join(" ") +
					">"
			end
		end

		#
		# +AllocationTable+'s hold the chains corresponding to files. Given
		# an initial index, <tt>AllocationTable#chain</tt> follows the chain, returning
		# the blocks that make up that file.
		#
		# There are 2 allocation tables, the bbat, and sbat, for big and small
		# blocks respectively. The block chain should be loaded using either
		# <tt>Storage#read_big_blocks</tt> or <tt>Storage#read_small_blocks</tt>
		# as appropriate.
		#
		# Whether or not big or small blocks are used for a file depends on
		# whether its size is over the <tt>Header#threshold</tt> level.
		#
		# An <tt>Ole::Storage</tt> document is serialized as a series of directory objects,
		# which are stored in blocks throughout the file. The blocks are either
		# big or small, and are accessed using the <tt>AllocationTable</tt>.
		#
		# The bbat allocation table's data is stored in the spare room in the header
		# block, and in extra blocks throughout the file as referenced by the meta
		# bat.  That chain is linear, as there is no higher level table.
		#
		class AllocationTable
			attr_reader :ole, :table
			def initialize ole
				@ole = ole
				@table = []
			end

			def self.load ole, chain
				at = AllocationTable.new ole
				at.table.replace ole.read_big_blocks(chain).unpack('L*')
				at
			end

			def chain start
				return [] if start >= (1 << 32) - 3
				raise "broken allocationtable chain" if start < 0 || start > @table.length
				[start] + chain(@table[start])
			end
		end

		#
		# A class which wraps an ole dir. Can be either a directory
		# (<tt>OleDir#dir?</tt>) or a file (<tt>OleDir#file?</tt>)
		#
		# Most interaction with <tt>Ole::Storage</tt> is through this class.
		# The 2 most important functions are <tt>OleDir#children</tt>, and
		# <tt>OleDir#data</tt>.
		# 
		class OleDir
			MEMBERS = [
				:name_utf16, :name_len, :type_id, :colour, :prev, :next, :child,
				:clsid, :flags, # dirs only
				:secs1, :days1, # create time
				:secs2, :days2, # modify time
				:first_block, :size, :reserved
			]
			PACK = 'a64 S C C L3 a16 L7 a4'
			SIZE = 128
			EPOCH = DateTime.parse '1601-01-01'
			TYPE_MAP = {
				1 => :dir,
				2 => :file,
				5 => :root
			}

			include Enumerable

			attr_accessor :idx, :ole
			# This returns all the children of this +OleDir+. It is filled in
			# when the tree structure is recreated.
			attr_accessor :children
			def initialize
				@values = []
			end

			def self.load str
				dir = OleDir.new
				dir.to_a.replace str.unpack(PACK)
				dir
			end

			def name
				UTF16_TO_UTF8[name_utf16[0...name_len].sub(/\x00\x00$/, '')]
			end

			def data
				# can now be implemented as
				# io.read
				return nil unless file? 
				bat = @ole.send(size > @ole.header.threshold ? :bbat : :sbat)
				chain = bat.chain(first_block)
				msg = size > @ole.header.threshold ? :read_big_blocks : :read_small_blocks
				@ole.send msg, chain, size
			end

			# provides io object to read the files contents from. no need to close, though
			# could provide a block form anyway i suppose
			def io
				return nil unless file?
				bat = @ole.send(size > @ole.header.threshold ? :bbat : :sbat)
				chain = bat.chain(first_block)
				if size > @ole.header.threshold
					# kind of takes over :read_big_blocks
					block_size = @ole.header.b_size
					ranges = chain[0, (size.to_f / block_size).ceil].map do |i|
						[block_size * (i + 1), block_size]
					end
					ranges.last[1] -= (ranges.length * block_size - size)
					RangesIO.new @ole.io, ranges
				else
					#:read_small_blocks
					# special case 0 size files. they have empty chains
					return RangesIO.new @io, [] if size == 0
					# more complicated
					block_size = @ole.header.s_size
					ranges = chain[0, (size.to_f / block_size).ceil].map do |block|
						# this tries to efficiently map a small block file to its position in the parent file.
						idx, pos = (block * block_size).divmod @ole.header.b_size
						[pos + @ole.header.b_size * (@ole.sb_blocks[idx] + 1), block_size]
					end
					ranges.last[1] -= (ranges.length * block_size - size)
					RangesIO.new @ole.io, ranges
				end
			end

			def type
				TYPE_MAP[type_id] or raise "unknown type #{type_id.inspect}"
			end

			def dir?
				# to count root as a dir.
				type != :file
			end

			def file?
				type == :file
			end

			def time
				# time is nil for streams, otherwise try to parse either of the time pairse (not
				# sure of their meaning - created / modified?)
				@time ||= file? ? nil : (OleDir.parse_time(secs1, days1) || OleDir.parse_time(secs2, days2))
			end

			def each(&block)
				@children.each(&block)
			end
			
			def [] idx
				if String === idx
					# path style look up. maybe take another arg which should
					# allow creation later on, like pole does.
					# this should maybe allow paths to be 'asdf/asdf/asdf', and
					# automatically split and recurse. is '/' invalid in an ole
					# dir name?
					# what about warning about multiple hits for the same name?
					children.find { |child| child.name == idx }
				else
					children[idx]
				end
			end

			# FIXME: doesn't belong here
			# Parse two 32 bit time values into a DateTime
			# Time is stored as a high and low 32 bit value, comprising the
			# 100's of nanoseconds since 1st january 1601 (Epoch).
			# struct FILETIME. see eg http://msdn2.microsoft.com/en-us/library/ms724284.aspx
			def self.parse_time low, high
				time = EPOCH + (high * (1 << 32) + low) * 1e-7 / 86400 rescue return
				# extra sanity check...
				unless (1800...2100) === time.year
					Log.warn "ignoring unlikely time value #{time.to_s}"
					return nil
				end
				time
			end

			def to_tree
				if children and !children.empty?
					str = "- #{inspect}\n"
					children.each_with_index do |child, i|
						last = i == children.length - 1
						child.to_tree.split(/\n/).each_with_index do |line, j|
							str << "  #{last ? (j == 0 ? "\\" : ' ') : '|'}#{line}\n"
						end
					end
					str
				else "- #{inspect}\n"
				end
			end

			MEMBERS.each_with_index do |sym, i|
				define_method(sym) { @values[i] }
			end

			def to_a
				@values
			end

			# perhaps i should remove the data snippet. its not that useful anyway.
			def inspect
				data = if file?
					tmp = io.read(9)
					tmp.length == 9 ? tmp[0, 5] + '...' : tmp
				end
				"#<OleDir:#{name.inspect} size=#{size}" +
					"#{time ? ' time=' + time.to_s.inspect : nil}" +
					"#{data ? ' data=' + data.inspect : nil}" +
					">"
			end
		end
	end
end

if $0 == __FILE__
	puts Ole::Storage.load(open(ARGV[0])).root.to_tree
end

