#! /usr/bin/ruby -w

$: << File.dirname(__FILE__) + '/..'

require 'iconv'
require 'date'
require 'support'

# move to support?
class IO
	def self.copy src, dst
		until src.eof?
			buf = src.read(4096)
			dst.write buf
		end
	end
end

#
# = Introduction
#
# +RangesIO+ is a basic class for wrapping another IO object allowing you to arbitrarily reorder
# slices of the input file by providing a list of ranges. Intended as an initial measure to curb
# inefficiencies in the Dirent#data method just reading all of a file's data in one hit, with
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
# * No buffering. by design at the moment. Intended for large reads
# 
# = TODO
# 
# On further reflection, this class is something of a joining/optimization of
# two separate IO classes. a SubfileIO, for providing access to a range within
# a File as a separate IO object, and a ConcatIO, allowing the presentation of
# a bunch of io objects as a single unified whole.
# 
# I will need such a ConcatIO if I'm to provide Mime#to_io, a method that will
# convert a whole mime message into an IO stream, that can be read from.
# It will just be the concatenation of a series of IO objects, corresponding to
# headers and boundaries, as StringIO's, and SubfileIO objects, coming from the
# original message proper, or RangesIO as provided by the Attachment#data, that
# will then get wrapped by Mime in a Base64IO or similar, to get encoded on-the-
# fly. Thus the attachment, in its plain or encoded form, and the message as a
# whole never exists as a single string in memory, as it does now. This is a
# fair bit of work to achieve, but generally useful I believe.
# 
# This class isn't ole specific, maybe move it to my general ruby stream project.
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

	def pos= pos, whence=IO::SEEK_SET
		# FIXME support other whence values
		raise NotImplementedError, "#{whence.inspect} not supported" unless whence == IO::SEEK_SET
		# just a simple pos calculation. invalidate buffers if we had them
		@pos = pos
	end

	alias seek :pos=
	alias tell :pos

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
		[r, off]
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

	# you may override this call to update @ranges and @size, if applicable. then write
	# support can grow below
	def truncate size
		raise NotImplementedError, 'truncate not supported'
	end
	# why not? :)
	alias size= :truncate

	def write data
		data_pos = 0
		# if we don't have room, we can use the truncate hook to make more space.
		if data.length > @size - @pos
			begin
				truncate @pos + data.length
			rescue NotImplementedError
				# FIXME maybe warn instead, then just truncate the data?
				raise "unable to satisfy write of #{data.length} bytes" 
			end
		end
		r, off = range_and_offset @pos
		i = ranges.index r
		([[r[0] + off, r[1] - off]] + ranges[i+1..-1]).each do |pos, len|
			@io.seek pos
			if data_pos + len > data.length
				chunk = data[data_pos..-1]
				@io.write chunk
				@pos += chunk.length
				data_pos = data.length
				break
			end
			@io.write data[data_pos, len]
			@pos += len
			data_pos += len
		end
		data_pos
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
		"#<#{self.class} io=#{io.inspect} size=#@size pos=#@pos "\
			"current_range=#{range_str}>"
	end
end

module Ole # :nodoc:
	Log = Logger.new_with_callstack

	# FIXME
	module Types
		# Parse two 32 bit time values into a DateTime
		# Time is stored as a high and low 32 bit value, comprising the
		# 100's of nanoseconds since 1st january 1601 (Epoch).
		# struct FILETIME. see eg http://msdn2.microsoft.com/en-us/library/ms724284.aspx
		def self.load_time str
			low, high = str.unpack 'L2'
			time = EPOCH + (high * (1 << 32) + low) * 1e-7 / 86400 rescue return
			# extra sanity check...
			unless (1800...2100) === time.year
				Log.warn "ignoring unlikely time value #{time.to_s}"
				return nil
			end
			time
		end

		# turn a binary guid into something displayable.
		# this will probably become a proper class later
		def self.load_guid str
			"{%08x-%04x-%04x-%02x%02x-#{'%02x' * 6}}" % str.unpack('L S S CC C6')
		end
	end


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
	#   ole = Ole::Storage.open 'myfile.msg'
	#   # => #<Ole::Storage io=#<File:myfile.msg> root=#<Dirent:"Root Entry"
	#          size=2816>>
	#   # get the top level root object and output a tree structure for
	#   # debugging
	#   puts ole.root.to_tree
	#   # =>
	#   - #<Dirent:"Root Entry" size=3840 time="2006-11-03T00:52:53Z">
	#     |- #<Dirent:"__nameid_version1.0" size=0 time="2006-11-03T00:52:53Z">
	#     |  |- #<Dirent:"__substg1.0_00020102" size=16 data="CCAGAAAAAADAAA...">
	#     ...
	#     |- #<Dirent:"__substg1.0_8002001E" size=4 data="MTEuMA==">
	#     |- #<Dirent:"__properties_version1.0" size=800 data="AAAAAAAAAAABAA...">
	#     \- #<Dirent:"__recip_version1.0_#00000000" size=0 time="2006-11-03T00:52:53Z">
	#        |- #<Dirent:"__substg1.0_0FF60102" size=4 data="AAAAAA==">
	#   	 ...
	#
	# = TODO
	#
	# 1. tests. lock down how things work at the moment - mostly good.
	# 2. lots of tidying up
	# 3. need to fix META_BAT support in #flush.
	# 4. maybe move io stuff to separate file
	#

	class Storage
		VERSION = '1.0.10'
		# All +Dirent+ names are in UTF16, which we convert
		UTF16_TO_UTF8 = Iconv.new('utf-8', 'utf-16le').method :iconv

		# The top of the ole tree structure
		attr_reader :root
		# The tree structure in its original flattened form. only valid after #load, or #flush.
		attr_reader :dirents
		# The underlying io object to/from which the ole object is serialized, whether we
		# should close it, and whether it is writeable
		attr_reader :io, :close_parent, :writeable
		# Low level internals, you probably shouldn't need to mess with these
		attr_reader :header, :bbat, :sbat, :sb_blocks

		# maybe include an option hash, and allow :close_parent => true, to be more general.
		# +arg+ should be either a file, or an +IO+ object, and needs to be seekable.
		def initialize arg, mode=nil
			# get the io object
			@close_parent, @io = if String === arg
				[true, open(arg, mode || 'rb')]
			else
				raise 'unable to specify mode string with io object' if mode
				[false, arg]
			end
			# do we have this file opened for writing? don't know of a better way to tell
			@writeable = begin
				@io.flush
				true
			rescue IOError
				false
			end
			# if the io object has data, we should load it, otherwise start afresh
			if !@io.size
				# first step though is to support modifying pre-existing and saving, then this
				# missing gap will be fairly straight forward - essentially initialize to
				# equivalent of loading an empty ole document.
				raise NotImplementedError, 'unable to create new ole objects from scratch as yet'
			else load
			end
		end

		def self.new arg, mode=nil
			ole = super
			if block_given?
				begin   yield ole
				ensure; ole.close
				end
			else ole
			end
		end

		class << self
			# encouraged
			alias open :new
			# deprecated
			alias load :new
		end

		# load document from file.
		def load
			# we always read 512 for the header block. if the block size ends up being different,
			# what happens to the 109 fat entries. are there more/less entries?
			@io.rewind
			header_block = @io.read 512
			@header = Header.load header_block

			bbat_chain_data =
				header_block[Header::SIZE..-1] +
				big_block_ranges((0...@header.num_mbat).map { |i| i + @header.mbat_start }).to_io.read
			@bbat = AllocationTable.load self, :big_block_ranges,
				bbat_chain_data.unpack('L*')[0, @header.num_bat]
			# FIXME i don't currently use @header.num_sbat which i should
			@sbat = AllocationTable.load self, :small_block_ranges,
				@bbat.chain(@header.sbat_start)
	
			# get block chain for directories, read it, then split it into chunks and load the
			# directory entries. semantics changed - used to cut at first dir where dir.type == 0
			@dirents = @bbat.ranges(@header.dirent_start).to_io.read.scan(/.{#{Dirent::SIZE}}/mo).
				map { |str| Dirent.load self, str }.reject { |d| d.type_id == 0 }

			# now reorder from flat into a tree
			# links are stored in some kind of balanced binary tree
			# check that everything is visited at least, and at most once
			# similarly with the blocks of the file.
			class << @dirents
				def to_tree idx=0
					return [] if idx == Dirent::EOT
					d = self[idx]
					d.children = to_tree d.child
					raise "directory #{d.inspect} used twice" if d.idx
					d.idx = idx
					to_tree(d.prev) + [d] + to_tree(d.next)
				end
			end

			@root = @dirents.to_tree.first
			Log.warn "root name was #{@root.name.inspect}" unless @root.name == 'Root Entry'
			@sb_blocks = @bbat.chain @root.first_block
			# sb_blocks could now be handled like this:
			#@sb_blocks = RangesIOResizable.new @io, @bbat, @root.first_block
			# # then redirect its first_block
			#root = @root
			#class << @sb_blocks
			#  def first_block
			#    root.first_block
			#  end
			#  def first_block= val
			#    root.first_block = val
			#  end
			#end
			# then @sb_blocks is read/writeable/resizeable not migrateable, and updates the @root
			# automagically. probably better is just to copy first_block at flush time. for less
			# integration.

			unused = @dirents.reject(&:idx).length
			Log.warn "* #{unused} unused directories" if unused > 0
		end

		def close
			flush if @writeable
			@io.close if @close_parent
		end

		# should have a #open_dirent i think. and use it in load and flush. neater.
=begin
thoughts on fixes:
1. reterminate any chain not ending in EOC. 
2. pass through all chain heads looking for collisions, and making sure nothing points to them
   (ie they are really heads).
3. we know the locations of the bbat data, and mbat data. ensure that there are placeholder blocks
   in the bat for them.
this stuff will ensure reliability of input better. otherwise, its actually worth doing a repack
directly after read, to ensure the above is probably acounted for, before subsequent writes possibly
destroy things.
=end
		def flush
			# recreate dirs from our tree, split into dirs and big and small files
			@root.type = :root
			# for now.
			@root.name = 'Root Entry'
			@dirents = @root.flatten
			#dirs, files = @dirents.partition(&:dir?)
			#big_files, small_files = files.partition { |file| file.size > @header.threshold }

			# maybe i should move the block form up to RangesIO, and get it for free at all levels.
			# Dirent#open gets block form for free then
			io = RangesIOResizeable.new @io, @bbat, @header.dirent_start
			io.truncate 0
			@dirents.each { |dirent| io.write dirent.save }
			padding = (io.size / @bbat.block_size.to_f).ceil * @bbat.block_size - io.size
			#p [:padding, padding]
			io.write 0.chr * padding
			@header.dirent_start = io.first_block
			io.close

			# similarly for the sbat data.
			io = RangesIOResizeable.new @io, @bbat, @header.sbat_start
			io.truncate 0
			io.write @sbat.save
			@header.sbat_start = io.first_block
			io.close

			# what follows will be slightly more complex for the bat fiddling.

			# now lets write out the bbat. the bbat's chain is not part of the bbat. but maybe i
			# should add blocks to the bbat to hold it.
			# firstly, create the bbat chain's actual chain data:
			# the size of bbat_data is equal to the
			#   bbat_data_size = ((number_of_normal_blocks + number_of_extra_bat_blocks) * 4 /
			#			block_size.to_f).ceil * block_size
			# saving it will require
			#   num_bbat_blocks = (bbat_data_size / block_size.to_f).ceil
			#   numer_of_extra_bat_blocks = num_bbat_blocks + num_mbat_blocks
			# which will then get added to number of blocks in the above, until it stabilises.
			# note that any existing free blocks can be used. this is the way to go, in order to
			# have the BAT properly appearing in the allocation table.
			# i just thought of the easiest way to do this:
			# create RangesIOResizeable hooked up to the bbat. use that to claim bbat blocks using
			# truncate. then when its time to write, convert that chain and some chunk of blocks at
			# the end, into META_BAT blocks. write out the chain, and those meta bat blocks, and its
			# done.

			#p @bbat
			@bbat.table.map! do |b|
				b == AllocationTable::BAT || b == AllocationTable::META_BAT ?
					AllocationTable::AVAIL : b
			end
			io = RangesIOResizeable.new @io, @bbat, AllocationTable::EOC
			# use crappy loop for now:
			while true
				bbat_data = @bbat.save
				#mbat_data = bbat_data.length / @bbat.block_size * 4
				mbat_chain = @bbat.chain io.first_block
				raise NotImplementedError, "don't handle writing out extra META_BAT blocks yet" if mbat_chain.length > 109
				# so we can ignore meta blocks in this calculation:
				break if io.size >= bbat_data.length # it shouldn't be bigger right?
				# this may grow the bbat, depending on existing available blocks
				io.truncate bbat_data.length
			end

			# now extract the info we want:
			ranges = io.ranges
			mbat_chain = @bbat.chain io.first_block
			io.close
			mbat_chain.each { |b| @bbat.table[b] = AllocationTable::BAT }
			#p @bbat.truncated_table
			#p ranges
			#p mbat_chain
			# not resizeable!
			io = RangesIO.new @io, ranges
			io.write @bbat.save
			io.close
			mbat_chain += [AllocationTable::AVAIL] * (109 - mbat_chain.length)

=begin
			bbat_data = new_bbat.save
			# must exist as linear chain stored in header.
			@header.num_bat = (bbat_data.length / new_bbat.block_size.to_f).ceil
			base = io.pos / new_bbat.block_size - 1
			io.write bbat_data
			# now that spanned a number of blocks:
			mbat = (0...@header.num_bat).map { |i| i + base }
			mbat += [AllocationTable::AVAIL] * (109 - mbat.length) if mbat.length < 109
			header_mbat = mbat[0...109]
			other_mbat_data = mbat[109..-1].pack 'L*'
			@header.mbat_start = base + @header.num_bat
			@header.num_mbat = (other_mbat_data.length / new_bbat.block_size.to_f).ceil
			io.write other_mbat_data
=end
			# now seek back and write the header out
			@io.seek 0
			@io.write @header.save + mbat_chain.pack('L*')
			@io.flush
		end

		# Turn a chain (an array given by +chain+) of big blocks, optionally
		# truncated to +size+, into an array of arrays describing the stretches of
		# bytes in the file that it belongs to.
		#
		# Big blocks are of size Ole::Storage::Header#b_size, and are stored
		# directly in the parent file.
		def big_block_ranges chain, size=nil
			block_size = @header.b_size
			# truncate the chain if required
			chain = chain[0...(size.to_f / block_size).ceil] if size
			# convert chain to ranges of the block size
			ranges = chain.map { |i| [block_size * (i + 1), block_size] }
			# truncate final range if required
			ranges.last[1] -= (ranges.length * block_size - size) if ranges.last and size
			io, ole = @io, self
			class << ranges; self; end.send(:define_method, :to_io) { RangesIO.new io, self }
			ranges
		end

		# As above, for +big_block_ranges+, but for small blocks.
		#
		# Small blocks are of size Ole::Storage::Header#s_size, and are stored
		# as a single file, serialized using big blocks. Single blocks are
		# mapped to big blocks using Ole::Storage#sb_blocks
		def small_block_ranges chain, size=nil
			block_size = @header.s_size
			chain = chain[0...(size.to_f / block_size).ceil] if size
			ranges = chain.map do |i|
				# this tries to efficiently map a small block file to its position in the parent file.
				idx, pos = (i * block_size).divmod @header.b_size
				[pos + @header.b_size * (@sb_blocks[idx] + 1), block_size]
			end
			ranges.last[1] -= (ranges.length * block_size - size) if ranges.last and size
			io = @io
			class << ranges; self; end.send(:define_method, :to_io) { RangesIO.new io, self }
			ranges
		end

		def bat_for_size size
			size > @header.threshold ? @bbat : @sbat
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
			PACK = 'a8 a16 S2 a2 S2 a6 L3 a4 L5'
			SIZE = 0x4c
			# i have seen it pointed out that the first 4 bytes of hex,
			# 0xd0cf11e0, is supposed to spell out docfile. hmmm :)
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

			def save
				@values.pack PACK
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
			# a free block (I don't currently leave any blocks free), although I do pad out
			# the allocation table with AVAIL to the block size.
			AVAIL		 = 0xffffffff
			EOC			 = 0xfffffffe # end of a chain
			# these blocks correspond to the bat, and aren't part of a file, nor available.
			# (I don't currently output these)
			BAT			 = 0xfffffffd
			META_BAT = 0xfffffffc

			attr_reader :ole, :table
			def initialize ole, range_conv
				@ole = ole
				@table = []
				@range_conv = range_conv
			end

			def self.load ole, range_conv, chain
				at = AllocationTable.new ole, range_conv
				at.table.replace ole.big_block_ranges(chain).to_io.read.unpack('L*')
				at
			end

			def truncated_table
				# this strips trailing AVAILs. come to think of it, this has the potential to break
				# bogus ole. if you terminate using AVAIL instead of EOC, like I did before. but that is
				# very broken. however, if a chain ends with AVAIL, it should probably be fixed to EOC
				# at load time.
				temp = @table.reverse
				not_avail = temp.find { |b| b != AVAIL } and temp = temp[temp.index(not_avail)..-1]
				temp.reverse
			end

			def save
				table = truncated_table #@table
				# pad it out some
				num = @ole.header.b_size / 4
				# do you really use AVAIL? they probably extend past end of file, and may shortly
				# be used for the bat. not really good.
				table += [AVAIL] * (num - (table.length % num)) if (table.length % num) != 0
				table.pack 'L*'
			end

			# rewriting this to be non-recursive. it broke on a large attachment
			# building up the chain, causing a stack error. need tail-call elimination...
			def chain start
				a = []
				idx = start
				until idx >= META_BAT
					raise "broken allocationtable chain" if idx < 0 || idx > @table.length
					a << idx
					idx = @table[idx]
				end
				Log.warn "invalid chain terminator #{idx}" unless idx == EOC
				a
			end
			
			def ranges chain, size=nil
				chain = self.chain(chain) unless Array === chain
				@ole.send @range_conv, chain, size
			end

			# ----------------------

			# up till now allocationtable's didn't know their block size. not anymore...
			# this should replace @range_conv
			# maybe cleaner to have a SmallAllocationTable, and a BigAllocationTable??
			def type
				@range_conv.to_s[/^[^_]+/].to_sym
			end

			# the plus 2 is,
			# 1 to get to the end of the block that starts at max_b, and the other due to the
			# `+ 1' in both range conversions (blocks are essentially indexed with -1 base, -1
			# corresponding to header bytes).
			# returns the size of the underlying object necessary to store all of our normal blocks
			def data_size
				#(@table.reject { |b| b >= META_BAT }.max + 2) * block_size
				(truncated_table.length + 1) * block_size
			end

			def get_free_block
				@table.each_index { |i| return i if @table[i] == AVAIL }
				@table.push AVAIL
				@table.length - 1
			end

			# must return first_block
			def resize_chain first_block, size
				new_num_blocks = (size / block_size.to_f).ceil
				blocks = chain first_block
				old_num_blocks = blocks.length
				if new_num_blocks < old_num_blocks
					# de-allocate some of our old blocks. TODO maybe zero them out in the file???
					(new_num_blocks...old_num_blocks).each { |i| @table[blocks[i]] = AVAIL }
					# if we have a chain, terminate it and return head, otherwise return EOC
					if new_num_blocks > 0
						@table[blocks[new_num_blocks-1]] = EOC
						first_block
					else EOC
					end
				elsif new_num_blocks > old_num_blocks
					# need some more blocks.
					last_block = blocks.last
					(new_num_blocks - old_num_blocks).times do
						block = get_free_block
						# connect the chain. handle corner case of blocks being [] initially
						if last_block
							@table[last_block] = block 
						else
							first_block = block
						end
						last_block = block
						# this is just to inhibit the problem where it gets picked as being a free block
						# again next time around.
						@table[last_block] = EOC
					end
					# for the growth case, we may now be bigger than the underlying file. that case
					# is currently handled manually for big blocks elsewhere, but what about small blocks:
					# need to re-size sbblocks
					if type == :small
						# the size of the small block file should be at least this big.
						# its really a bit of a hack, to stand in for a truncate call on the underlying not
						# doing the right thing as yet, because we optimized away the proper underlying
						sb_blocks_size = (@table.reject { |b| b >= META_BAT }.max + 2) * block_size
						root_first_block = @ole.bbat.resize_chain @ole.root.first_block, sb_blocks_size
						@ole.sb_blocks.replace @ole.bbat.chain(root_first_block)
						@ole.root.first_block = root_first_block
					end
					first_block
				else first_block
				end
			end

			def block_size
				@ole.header.send type == :big ? :b_size : :s_size
			end
		end

		# like normal RangesIO, but Ole::Storage specific. the ranges are backed by an
		# AllocationTable, and can be resized. used for read/write to 2 streams:
		# 1. serialized dirent data
		# 2. sbat table data
		# 3. all dirents but through RangesIOMigrateable below
		#
		# Note that all internal access to first_block is through accessors, as it is sometimes
		# useful to redirect it.
		class RangesIOResizeable < RangesIO
			attr_reader   :bat
			attr_accessor :first_block
			def initialize io, bat, first_block, size=nil
				@bat = bat
				self.first_block = first_block
				super(io, @bat.ranges(first_block, size))
			end

			def truncate size
				# note that old_blocks is != @ranges.length necessarily. i'm planning to write a
				# merge_ranges function that merges sequential ranges into one as an optimization.
				self.first_block = @bat.resize_chain first_block, size
				@ranges = @bat.ranges first_block, size
				@pos = @size if @pos > size

				# don't know if this is required, but we explicitly request our @io to grow if necessary
				# we never shrink it though. maybe this belongs in allocationtable, where smarter decisions
				# can be made.
				# maybe its ok to just seek out there later??
				max = @ranges.map { |pos, len| pos + len }.max || 0
				@io.truncate max if max > @io.size

				@size = size
			end
		end

		# like RangesIOResizeable, but Ole::Storage::Dirent specific. provides for migration
		# between bats based on size, and updating the dirent, instead of the ole copy back
		# on close.
		class RangesIOMigrateable < RangesIOResizeable
			attr_reader :dirent
			def initialize io, dirent
				@dirent = dirent
				super(io, @dirent.ole.bat_for_size(@dirent.size), @dirent.first_block, @dirent.size)
			end

			def truncate size
				bat = @dirent.ole.bat_for_size size
				if bat != @bat
					# bat migration needed! we need to backup some data. the amount of data
					# should be <= @ole.header.threshold, so we can just hold it all in one buffer.
					# backup this
					pos = @pos
					@pos = 0
					keep = read [@size, size].min
					# this does a normal truncate to 0, removing our presence from the old bat, and
					# rewrite the dirent's first_block
					super(0)
					@bat = bat
					# important to do this now, before the write. as the below write will always
					# migrate us back to sbat! this will now allocate us +size+ in the new bat.
					super
					@pos = 0
					write keep
					@pos = pos
				else
					super
				end
				# now just update the file
				@dirent.size = size
			end

			# forward this to the dirent
			def first_block
				@dirent.first_block
			end

			def first_block= val
				@dirent.first_block = val
			end
		end

		#
		# A class which wraps an ole directory entry. Can be either a directory
		# (<tt>Dirent#dir?</tt>) or a file (<tt>Dirent#file?</tt>)
		#
		# Most interaction with <tt>Ole::Storage</tt> is through this class.
		# The 2 most important functions are <tt>Dirent#children</tt>, and
		# <tt>Dirent#data</tt>.
		# 
		# was considering separate classes for dirs and files. some methods/attrs only
		# applicable to one or the other.
		class Dirent
			MEMBERS = [
				:name_utf16, :name_len, :type_id, :colour, :prev, :next, :child,
				:clsid, :flags, # dirs only
				:create_time_str, :modify_time_str, # files only
				:first_block, :size, :reserved
			]
			PACK = 'a64 S C C L3 a16 L a8 a8 L2 a4'
			SIZE = 128
			EPOCH = DateTime.parse '1601-01-01'
			TYPE_MAP = {
				# this is temporary
				0 => :empty,
				1 => :dir,
				2 => :file,
				5 => :root
			}
			COLOUR_MAP = {
				0 => :red,
				1 => :black
			}
			# used in the next / prev / child stuff to show that the tree ends here.
			# also used for first_block for directory.
			EOT = 0xffffffff

			include Enumerable

			# Dirent's should be created in 1 of 2 ways, either Dirent.new ole, [:dir/:file/:root],
			# or Dirent.load '... dirent data ...'
			# its a bit clunky, but thats how it is at the moment. you can assign to type, but
			# shouldn't.

			attr_accessor :idx
			# This returns all the children of this +Dirent+. It is filled in
			# when the tree structure is recreated.
			attr_accessor :children
			attr_reader :ole, :type, :create_time, :modify_time, :name
			def initialize ole, type
				@ole = ole
				# this isn't really good enough. need default values put in there.
				@values = []
				# maybe check types here. 
				@type = type
				@create_time = @modify_time = nil
				if file?
					@create_time = Time.now
					@modify_time = Time.now
				end
			end

			def self.load ole, str
				# load should function without the need for the initializer.
				dirent = Dirent.allocate
				dirent.load ole, str
				dirent
			end

			def load ole, str
				@ole = ole
				@values = str.unpack PACK
				@name = UTF16_TO_UTF8[name_utf16[0...name_len].sub(/\x00\x00$/, '')]
				@type = TYPE_MAP[type_id] or raise "unknown type #{type_id.inspect}"
				if file?
					@create_time = Types.load_time create_time_str
					@modify_time = Types.load_time modify_time_str
				end
			end

			# only defined for files really. and the above children stuff is only for children.
			# maybe i should have some sort of File and Dir class, that subclass Dirents? a dirent
			# is just a data holder. 
			# this can be used for write support if the underlying io object was opened for writing.
			# maybe take a mode string argument, and do truncation, append etc stuff.
			def open
				return nil unless file?
				bat = size > @ole.header.threshold ? @ole.bbat : @ole.sbat
				io = RangesIOMigrateable.new @ole.io, self
				if block_given?
					begin   yield io
					ensure; io.close
					end
				else io
				end
			end

			def read limit=nil
				open { |io| io.read limit }
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
				#@time ||= file? ? nil : (Dirent.parse_time(secs1, days1) || Dirent.parse_time(secs2, days2))
				create_time || modify_time
			end

			def each(&block)
				@children.each(&block)
			end
			
			def [] idx
				if Integer === idx
					children[idx]
				else
					# path style look up.
					# maybe take another arg to allow creation? or leave that to the filesystem
					# add on. 
					# not sure if '/' is a valid char in an Dirent#name, so no splitting etc at
					# this level.
					# also what about warning about multiple hits for the same name?
					children.find { |child| idx === child.name }
				end
			end

			# solution for the above '/' thing for now.
			def / path
				self[path]
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
				define_method(sym.to_s + '=') { |val| @values[i] = val }
			end

			def to_a
				@values
			end

			# flattens the tree starting from here into +dirents+. note it modifies its argument.
			def flatten dirents=[]
				@idx = dirents.length
				dirents << self
				children.each { |child| child.flatten dirents }
				self.child = Dirent.flatten_helper children
				dirents
			end

			# i think making the tree structure optimized is actually more complex than this, and
			# requires some intelligent ordering of the children based on names, but as long as
			# it is valid its ok.
			# actually, i think its ok. gsf for example only outputs a singly-linked-list, where
			# prev is always EOT.
			def self.flatten_helper children
				return EOT if children.empty?
				i = children.length / 2
				this = children[i]
				this.prev, this.next = [(0...i), (i+1..-1)].map { |r| flatten_helper children[r] }
				this.idx
			end

			attr_accessor :name, :type
			def save
				tmp = Iconv.new('utf-16le', 'utf-8').iconv(name) + 0.chr * 2
				tmp = tmp[0, 64] if tmp.length > 64
				self.name_len = tmp.length
				self.name_utf16 = tmp + 0.chr * (64 - tmp.length)
				begin
					self.type_id = TYPE_MAP.to_a.find { |id, name| @type == name }.first
				rescue
					raise "unknown type #{type.inspect}"
				end
				# for the case of files, it is assumed that that was handled already
				# note not dir?, so as not to override root's first_block
				self.first_block = Dirent::EOT if type == :dir
				if 0 #file?
					#self.create_time_str = ?? #Types.load_time create_time_str
					#self.modify_time_str = ?? #Types.load_time modify_time_str
				else
					self.create_time_str = 0.chr * 8
					self.modify_time_str = 0.chr * 8
				end
				@values.pack PACK
			end

			def inspect
				# perhaps i should remove the data snippet. its not that useful anymore.
				data = if file?
					tmp = read 9
					tmp.length == 9 ? tmp[0, 5] + '...' : tmp
				end
				"#<Dirent:#{name.inspect} size=#{size}" +
					"#{time ? ' time=' + time.to_s.inspect : nil}" +
					"#{data ? ' data=' + data.inspect : nil}" +
					">"
			end

			# --------
			# and for creation of a dirent. don't like the name. is it a file or a directory?
			# assign to type later? io will be empty.
			def new_child type
				child = Dirent.new ole, type
				children << child
				yield child if block_given?
				child
			end

			def delete child
				# remove from our child array, so that on reflatten and re-creation of @dirents, it will be gone
				raise "#{child.inspect} not a child of #{self.inspect}" unless @children.delete child
				# free our blocks
				child.open { |io| io.truncate 0 }
			end

			def self.copy src, dst
				# copies the contents of src to dst. must be the same type. this will throw an
				# error on copying to root. 
				raise unless src.type == dst.type
				src.name = dst.name
				if src.dir?
					src.children.each do |src_child|
						dst.new_child src_child.type { |dst_child| Dirent.copy src_child, dst_child }
					end
				else
					src.open do |src_io|
						dst.open dst_child.type { |dst_io| IO.copy src_io, dst_io }
					end
				end
			end
		end
	end
end

if $0 == __FILE__
	puts Ole::Storage.open(ARGV[0]) { |ole| ole.root.to_tree }
end

