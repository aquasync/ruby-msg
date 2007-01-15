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

		attr_reader :io, :header, :bbat, :sbat, :dirs, :sb_blocks, :root
		def initialize
		end

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

		def read_small_blocks blocks, size=nil
			data = ''
			blocks.each do |block|
				# interesting... small blocks are mapped to big blocks in a peculiar way they can be shifted
				# around arbitrarily, and you just update the sb_blocks map. without adjusting the allocation
				# table. it is an extra layer of indirection.
				idx, pos = (block * (1 << @header.s_shift)).divmod 1 << @header.b_shift
				pos += (1 << @header.b_shift) * (@sb_blocks[idx] + 1)
				@io.seek pos
				data << @io.read(1 << @header.s_shift)
			end
			data = data[0, size] if size and size < data.length
			data
		end

		def self.load io
			ole = Storage.new
			ole.load io
			ole
		end

		# +io+ needs to be seekable.
		def load io
			@io = io

			# we always read 512 for the header block. if the block size ends up being different,
			# what happens to the 109 fat entries. are there more/less entries?
			@io.seek 0
			header_block = @io.read 512

			@header = Header.load header_block
			raise "not valid OLE2 structured storage file" unless @header.magic == MAGIC

			# bbat chain is linear, as there is no other table.
			# the metabat is a small group of big blocks, made up of longs which are the big block
			# indices that point to blocks which are the big block allocation table chain.
			# when loading a stream, you either grab it from the big or small chain depending on its size...
			# note also some of the data coming from header block.
			# that provides an array of indices, which are loaded by the bbat.
			bbat_chain_data =
				header_block[FAT_START..-1] +
				read_big_blocks((0...@header.num_mbat).map { |i| @header.mbat_start + i })

			@bbat = AllocationTable.load self, bbat_chain_data.unpack('L*')[0, @header.num_bat]
			@sbat = AllocationTable.load self, @bbat.chain(@header.sbat_start)

			# get block chain for directories and load them
			#dirs = read_big_blocks(@bbat.get_chain(@header.dirent_start)).scan(/.{128}/m).
			@dirs = read_big_blocks(@bbat.chain(@header.dirent_start)).scan(/.{128}/m).
			# semantics mayn't be quite right. used to cut at first dir where dir.type == 0
				map { |str| OleDir.from_str str }.reject { |dir| dir.type == 0 }
			@dirs.each { |dir| dir.ole = self }

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
			@sb_blocks = @bbat.chain @root.start_block

			# extra check
			unused = @dirs.reject { |dir| dir.idx }.length
			warn "* #{unused} unused directories" if unused > 0
		end

		def inspect
			"#<#{self.class} io=#{@io.inspect} root=#{@root.inspect}>"
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
			attr_reader :ole, :table
			def initialize ole
				@ole = ole
				@table = []
			end

			def self.load ole, chain
				at = AllocationTable.new ole
				at.load chain
				at
			end

			def load chain
				@table = @ole.read_big_blocks(chain).unpack 'L*'
			end

			def chain start
				return [] if start >= (1 << 32) - 3
				raise "dodgy chain #{start}" if start < 0 || start > @table.length
				[start] + chain(@table[start])
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
				return nil if type != 2
				bat = @ole.instance_variable_get(size > @ole.header.threshold ? :@bbat : :@sbat)
				msg = size > @ole.header.threshold ? :read_big_blocks : :read_small_blocks
				chain = bat.chain(start_block)
				#p chain
				#p [start_block, bat]
				#p chain
				@ole.send msg, chain, size
			end

			def data64
				require 'base64'
				d = Base64.encode64(data[0..100]).delete("\n")
				d.length > 16 ? d[0..13] + '...' : d
				rescue
				nil
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
				data = self.data64
				"#<OleDir:#{name.inspect} size=#{size}" +
					"#{time ? ' time=' + time.to_s.inspect : nil}" +
					"#{data ? ' data=' + data.inspect : nil}" +
					">"
			end
		end
	end
end

if $0 == __FILE__
	#p Ole::Storage.new(open(ARGV[0])).header
	puts Ole::Storage.load(open(ARGV[0])).root.to_tree
end

