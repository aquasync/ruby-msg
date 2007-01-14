#! /usr/bin/ruby -w

require 'iconv'
require 'date'

module Ole
	# class to provide access to OLE2 structured storage files, such as those produced by
	# microsoft office, eg *.doc, *.msg etc.
	# based on chicago's libole, source available at
	# http://prdownloads.sf.net/chicago/ole.tgz
	class Storage
		VERSION = '1.0.3'
		UTF16_TO_UTF8 = Iconv.new('utf-8', 'utf-16le').method :iconv

		def each_block bat, *blocks
			blocks.each do |block|
				@io.seek bat.block_size * (block + 1)
				yield @io.read(bat.block_size)
			end
		end

		def each_big_block *blocks, &block
			each_block @bbat, *blocks, &block
		end

		def each_small_block *blocks, &block
			each_block @sbat, *blocks, &block
		end

		attr_reader :io, :blocks, :header, :fat, :dirs, :root
		# +io+ needs to be seekable.
		def initialize io
			@io = io
			@blocks = Blocks.new self
			header_block = blocks[-1]
			@header = Header.load header_block
			@bbat = AllocationTable.new self, @header.b_shift
			@sbat = AllocationTable.new self, @header.s_shift
			# bbat chain is linear, as there is no other table.
			# the metabat is a small group of big blocks, made up of longs which are the big block
			# indices that point to blocks which are the big block allocation table chain.
			# when loading a stream, you either grab it from the big or small chain depending on its size...
			bbat_chain = header_block[FAT_START..-1].unpack('L*')
			@header.num_mbat.times do |i|
				each_big_block(@header.mbat_start + i) { |s| bbat_chain << s.unpack('L*') }
			end
			# trunacate if needed
			bbat_chain = bbat_chain[0...@header.num_bat] if bbat_chain.length > @header.num_bat
			@bbat.load_chain bbat_chain
			p @bbat

			# get block chain for sbat and load it
			@sbat.load_chain @bbat.get_chain(@header.sbat_start)
			p @sbat

			# get block chain for directories and load them
			@names = []
			each_big_block *@bbat.get_chain(@header.dirent_start) do |str|
				@names += str.scan(/.{128}/m).map { |s| dir = OleDir.from_str(s); dir.name if dir.type != 0 }.compact
			end
			p @names
			raise "not valid OLE2 structured storage file" unless @header.magic == MAGIC

			# load the fat. i think this is not exactly fat, but fat block numbers.
			# the closing if seems redundant but existed in sample code
			@fat = header_block[FAT_START..-1].unpack('L*')
			(0...@header.num_extra_fat_blocks).inject @header.fat_next_block do |i|
				ints = @blocks[i].unpack('L*')
				i = ints.shift
				@fat << ints
				i
			end if @header.fat_next_block > 0

			# get directories
			blknum = @header.root_start_block
			@dirs = @blocks[blknum].dirs
			while true
				fat = @blocks.get_fat_block blknum
				# is this % 128 just for array boundary reasons?. should it really be invalid?
				blknum = fat[blknum % 128]
				break if blknum == (1 << 32) - 2
				dirs = @blocks[blknum].dirs
				@dirs += dirs
				break unless dirs.length == 4
			end

			# now reorder from flat into a tree
			# links are stored in some kind of balanced binary tree
			# could maybe check that everything is visited only once, and that everything is covered.
			# similarly with the blocks of the file.
			class << @dirs
				def to_tree idx=0
					return [] if idx == (1 << 32) - 1
					dir = self[idx]
					dir.children = to_tree dir.dir_dirent
					raise "directory #{dir.inspect} used twice" if dir.idx
					dir.idx = idx
					to_tree(dir.prev_dirent) + [dir] + to_tree(dir.next_dirent)
				end
			end
			@root = @dirs.to_tree.first

			# extra check
			unused = @dirs.reject { |dir| dir.idx }.length
			warn "* #{unused} unused directories after to_tree" if unused > 0
		end

		def inspect
			"#<#{self.class} @io=#{@io.inspect} @root=#{@root.inspect}>"
		end

		# class to provide access to the file as a pseudo array of blocks
		# blocks[-1] is the header block with some fat pointers, blocks[0] is typically the first
		# fat block, and blocks[1] is typically the first directory block
		class Blocks
			BLOCK_SIZE = 512
			include Enumerable
			attr_reader :length

			def initialize ole
				@ole = ole
				@length = (@ole.io.stat.size + BLOCK_SIZE - 1) / BLOCK_SIZE - 1
			end

			def [] idx
				raise "block index #{idx.inspect} out of range" unless (-1...@length) === idx
				@ole.io.seek BLOCK_SIZE * (idx + 1)
				b = Block.new @ole.io.read(BLOCK_SIZE)
				b.ole = @ole
				b
			end

			def each
				# needless seeking
				length.times { |i| yield self[i] }
			end

			def get_fat_block idx
				# find the block that corresponds to the given block. each fat block has 128 entries.
				# further we divide by 128 here. still, it doesn't make sense yet.
				fat_idx = @ole.fat[idx / (BLOCK_SIZE / 4)]
				raise "invalid fat block for #{idx.inspect}" if fat_idx == (1 << 32) - 1
				self[fat_idx].unpack('L*')
			end

			class Block < String
				attr_accessor :ole
				def dirs
					dirs = self.scan(/.{#{OLE_DIR_SIZE}}/mo).
						map { |str| OleDir.from_str str }
					dirs.each { |dir| dir.ole = @ole }
					# sample code cuts at the first NO_ENTRY
					i = dirs.index dirs.find { |dir| dir.type == 0 }
					i ? dirs[0...i] : dirs
				end
			end
		end

		# class which wraps the ole header
		class Header
			LEGACY_MEMBERS = [
				:magic, :unk1, :unk1a, :unk1b, :unk1c, :num_fat_blocks, :root_start_block, :unk2,
				:unk3, :dir_flag, :unk4, :fat_next_block, :num_extra_fat_blocks
			]
			NEW_MEMBERS = [
				:magic, :unk1, :b_shift, :s_shift, :unk1d, :num_bat, :dirent_start, :unk2, :threshold,
				:sbat_start, :num_sbat, :mbat_start, :num_mbat
			]

			# 2 basic initializations, from scratch, or from a data string.
			def initialize
				@values = []
			end

			def self.load str
				h = Header.new
				h.to_a.replace str.unpack('a8 a22 S S a10 L8')
				h
			end

			[LEGACY_MEMBERS, NEW_MEMBERS].each do |members|
				members.each_with_index do |sym, i|
					define_method(sym) { @values[i] }
				end
			end

			def to_a
				@values
			end

			def inspect
				"#<#{self.class} " +
					NEW_MEMBERS.zip(@values).map { |k, v| "#{k}=#{v.inspect}" }.join(" ") +
					">"
			end
		end

		class AllocationTable
			attr_reader :ole, :block_size, :table
			def initialize ole, shift
				@ole = ole
				@block_size = 1 << shift
				@table = []
			end

			def load_chain chain
				@ole.each_big_block *chain do |s|
					@table += s.unpack('L*')
				end
			end

			def get_chain start
				return [] if start >= (1 << 32) - 3
				raise "dodgy chain #{start}" if start < 0 || start > @table.length
				[start] + get_chain(@table[start])
			end
		end

		MAGIC = "\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1"  # expected value of Header#magic
		FAT_START = 0x4c # should be equal to used size of the above HEADER_UNPACK string

		# class which wraps an ole dir
		OleDir = Struct.new :name_utf16, :name_len, :type, :filler1, :prev_dirent, :next_dirent, :dir_dirent,
			:unk1, :secs1, :days1, :secs2, :days2, :start_block, :size, :unk2
		OLE_DIR_UNPACK = 'a64 S C C L3 a20 L7'
		OLE_DIR_SIZE = 128
		EPOCH = DateTime.parse '1601-01-01'
		OleDir.class_eval do
			attr_accessor :idx, :children, :ole
			def self.from_str str
				OleDir.new(*str.unpack(OLE_DIR_UNPACK))
			end

			def name
				UTF16_TO_UTF8[name_utf16[0...name_len].sub(/\x00\x00$/, '')]
			end

			def data
				bat = @ole.instance_variable_get(size > @ole.header.threshold ? :@bbat : :@sbat)
				chain = bat.get_chain(start_block)
				p [start_block, bat]
				p chain
				data = ''
				@ole.each_block(bat, *chain) { |s| data << s }
				data[0, size]
			end

			def time
				# time is nil for streams, otherwise try to parse either of the time pairse (not
				# sure of their meaning - created / modified?)
				@time ||= type == 2 ? nil : (parse_time(secs1, days1) || parse_time(secs2, days2))
			end

			# time is made of a high and low 32 bit value, comprising of the 100's of nanoseconds
			# since 1st january 1601.
			# struct FILETIME. see eg http://msdn2.microsoft.com/en-us/library/ms724284.aspx
			def parse_time low, high
				time = EPOCH + (high * (1 << 32) + low) * 1e-7 / 86400 rescue nil
				# extra sanity check...
				time if time and (1800...2100) === time.year
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

			def inspect
#				"#<OleDir:#{name.inspect} size=#{size} #{[secs1, days1, secs2, days2].inspect}>"
				"#<OleDir:#{name.inspect} size=#{size}#{time ? ' time=' + time.to_s.inspect : nil}>"
			end
		end
	end
end

if $0 == __FILE__
	p Ole::Storage.new(open(ARGV[0])).header
	#puts Ole::Storage.new(open(ARGV[0])).root.to_tree
end

