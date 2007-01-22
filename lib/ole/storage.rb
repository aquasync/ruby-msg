#! /usr/bin/ruby -w

require 'iconv'
require 'date'
require 'support'

module Ole
	Log = Logger.new_with_callstack

	# basic class to provide access to OLE2 structured storage files, such as those produced by
	# microsoft office, eg *.doc, *.msg etc.
	# based on chicago's libole, source available at
	# http://prdownloads.sf.net/chicago/ole.tgz
	# augmented later by pole, and a bit from gsf.
	class Storage
		VERSION = '1.0.8'
		UTF16_TO_UTF8 = Iconv.new('utf-8', 'utf-16le').method :iconv

		attr_reader :io, :header, :bbat, :sbat, :dirs, :sb_blocks, :root
		def initialize
			# creation of new ole objects not properly supported as yet
		end

		def self.load io
			ole = Storage.new
			ole.load io
			ole
		end

		# +io+ needs to be seekable.
		def load io
			# we always read 512 for the header block. if the block size ends up being different,
			# what happens to the 109 fat entries. are there more/less entries?
			@io = io
			@io.seek 0
			header_block = @io.read 512
			@header = Header.load header_block

			# bbat chain is linear, as there is no other table.
			# the metabat is a small group of big blocks, made up of longs which are the big block
			# indices that point to blocks which are the big block allocation table chain.
			# when loading a stream, you either grab it from the big or small chain depending on its size...
			# note also some of the data coming from header block.
			# that provides an array of indices, which are loaded by the bbat.
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
			unused = @dirs.reject { |dir| dir.idx }.length
			Log.warn "* #{unused} unused directories" if unused > 0

			@sb_blocks = @bbat.chain @root.first_block

			# this warn belongs in Ole::Storage.load, as nested msgs won't have this as a name
			Log.warn "root name was #{@root.name.inspect}" unless @root.name == 'Root Entry'
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
				# small blocks are essentially files within a a small block file.
				# this does an efficient map, of a small block file to its position in the parent file.
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

		# class which wraps the ole header
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

		# class which wraps an ole dir
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

			attr_accessor :idx, :children, :ole
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
				return nil unless file? 
				bat = @ole.send(size > @ole.header.threshold ? :bbat : :sbat)
				msg = size > @ole.header.threshold ? :read_big_blocks : :read_small_blocks
				chain = bat.chain(first_block)
				@ole.send msg, chain, size
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
				@time ||= file? ? nil : (OleDir.parse_time(secs1, days1) || OleDir.parse_time(secs2, days2))
			end

			# time is made of a high and low 32 bit value, comprising of the 100's of nanoseconds
			# since 1st january 1601.
			# struct FILETIME. see eg http://msdn2.microsoft.com/en-us/library/ms724284.aspx
			def self.parse_time low, high
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

			MEMBERS.each_with_index do |sym, i|
				define_method(sym) { @values[i] }
			end

			def to_a
				@values
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
	puts Ole::Storage.load(open(ARGV[0])).root.to_tree
end

