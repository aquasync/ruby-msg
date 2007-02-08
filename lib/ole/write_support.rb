=begin

plan is to begin with writing out a tree, as opposed to the api for the creation of that
tree in the first place.
essentially simple - reverse the stuff done in the load function.

that means, first flatten the tree structure:

TODO

* the core stuff is mostly here. openoffice will accept my output (and indeed
  i am converting an ole file made with openoffice), but word won't yet, so its
  not conforming in some way. i think my own library sees output as identical
  to input.
  need to get it working.
* lock it down with tests.
* refactor. first within this file, moving things from the monolithic load to
  correct classes. neaten, try to leave the ole object in a sane state after
  saving, etc, etc.
* then maybe merge this stuff into the core, and attempt to fill out api for
  creating an ole document from scratch.
* finally create further tests, and refactor until it feels nice and clean
  again.

=end

module Ole
	class Storage
		class AllocationTable
			# up till now allocationtable's didn't know their block size. not anymore...
			# this should replace @range_conv
			def type
				@range_conv.to_s[/^[^_]+/].to_sym
			end

			def clear
				@table = []
			end

			def block_size
				@ole.header.send type == :big ? :b_size : :s_size
			end

			# just allocate space for a chain
			def new_chain size
				# now need to worry about empty blocks etc, because we just create from scratch
				chain_head = @table.length
				(size / block_size.to_f).ceil.times { |i| @table << table.length + 1 }
				@table[-1] = (1 << 32) - 1
				chain_head
			end
		end

		class Dirent
			MEMBERS.each_with_index do |sym, i|
				define_method(sym.to_s + '=') { |val| @values[i] = val }
			end

			def flatten dirents=[]
				# add us to the dirs
				@idx = dirents.length
				dirents << self
				children.each { |child| child.flatten dirents }
				self.child = Dirent.flatten_helper children
				dirents
			end

			# i think making the tree structure optimized is actually more complex than this, and
			# requires some intelligent ordering of the children based on names, but as long as
			# it is valid its ok.
			def self.flatten_helper children
				return (1 << 32) - 1 if children.empty?
				point = children.length / 2
				child = children[point]
				# obviously i need to defined setters for these...
				child.prev = flatten_helper children[0...point]
				child.next = flatten_helper children[point+1..-1]
				child.idx
			end

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
				if 0 #file?
					#self.create_time_str = ?? #Types.load_time create_time_str
					#self.modify_time_str = ?? #Types.load_time modify_time_str
				else
					self.create_time_str = 0.chr * 8
					self.modify_time_str = 0.chr * 8
				end
				@values.pack PACK
			end
		end

		def save io
			io.rewind
			io.truncate 0
			# empty header to be filled in later.
			io.write 0.chr * @header.b_size
			# recreate dirs from our tree
			@dirents = @root.flatten
			# we assume that the header values are mostly all ok.
			# well that was the easy bit. what follows is to actually create all the file streams.
			# now serialize all the dirents first. that should be easy. actually the dir chain has
			# nothing to do with whether its a dir or not. then do big files, then do small files
			# first step then is to reverse this:
			#@dirs = read_big_blocks(@bbat.chain(@header.dirent_start)).scan(/.{#{OleDir::SIZE}}/mo).
			# semantics mayn't be quite right. used to cut at first dir where dir.type == 0
			#	map { |str| OleDir.load str }.reject { |dir| dir.type_id == 0 }
			# ie, we need to save all the dirents back to 512 blobs, store them in a sequential chain
			# starting at block whatever (1?), and then fill out our allocation table with the
			# correct dir chain, and update the header dirent_start.
			
			# we need a temporary bbat, not to clear the real one, as the file io later requires it
			# as we don't eagerly load all io at once.
			new_bbat = AllocationTable.new self, :big_block_ranges
			#@bbat.clear
			# this creates a new chain, fills it in with block values as written, and serializes
			# all the dirents. i need to pad it out with some dirs too.
			#@header.dirent_start = @bbat.new_chain do |io|
			#	@dirs.each do |dir|
			#		io.write dir.to_s
			#	end
			#end

			# then exactly the same for the big blocks, making them into chains.
			# for the small blocks, thats a bit more confusing, would need to think about it.
			# need to get sbblocks out of it somehow.
			# finally, we write the various allocation tables, and layers of indirection to find
			# everything to disk. has to be done after writing everything else. then fill out
			dirs, files = @dirents.partition &:dir?
			
			# oh, this stuff isn't known until later. so we either claim the relevant space above,
			# or we write it out afterwards.
			dirs.each { |dir| dir.first_block = (1 << 32) - 1 }
			big_files, small_files = files.partition { |file| file.size > @header.threshold }
			puts "#{big_files.length} big files"
			puts "#{small_files.length} small files"
			big_files.each do |file|
				# we need to get the io before we rewrite its first block etc
				file_io = file.io
				file.size = file_io.size
				file.first_block = new_bbat.new_chain file.size
				# note that this may seek over padding bytes that we didn't write. will these be 0
				# its probably better if we explicitly pad blocks, especially last block in the file.
				io.seek new_bbat.ranges(file.first_block).first.first
				file_io.seek 0
				# this sucks, i'm lazy atm
				io.write file_io.read
			end

			# now tackle small files, and the sbblocks
			# then put sbblocks in root first_block
			# first lets see how big all small files are:
			@root.first_block = new_bbat.new_chain small_files.map(&:size).sum
			io.seek new_bbat.ranges(@root.first_block).first.first
			new_sbat = AllocationTable.new self, :small_block_ranges
			small_files.each do |file|
				# we need to get the io before we rewrite its first block etc
				file_io = file.io
				file.size = file_io.size
				file.first_block = new_sbat.new_chain file.size
				# we won't seek
				io.write file_io.read
				# now make up the gap. i think we only pad to the sblock boundary
				pad = new_sbat.block_size - (file.size % new_sbat.block_size)
				pad = 0 if pad == new_sbat.block_size
				io.write 0.chr * pad
			end

			# write out dirents.
			@header.dirent_start = new_bbat.new_chain @dirents.length * Dirent::SIZE

			# now if i add write support to rangedio, its as simple as turning that chain into an
			# io, and writing to it. given my current approach though, i know my chains are created
			# sequentially. so i can just do
			io.seek new_bbat.ranges(@header.dirent_start).first.first
			# this 4 stuff is hardcoded bad. FIXME
			pad = 4 - (@dirents.length % 4)
			pad = 0 if pad == 4
			@dirents.each do |dirent|
				# write out the dirent to io
				io.write dirent.save
			end
			pad.times do
				# write out a padding dirent i think this is ok:
				io.write 0.chr * Dirent::SIZE
			end

			# now write out the various meta blocks, fill out header a bit
			# ok. how to do this. first thing i suppose is to make room for the size of the sbat
			sbat_data = new_sbat.table.pack('L*')
			# this idiom is getting old and ugly quickly... whats the better way to do it?
			pad = new_bbat.block_size - (sbat_data.length % new_bbat.block_size)
			pad = 0 if pad == new_bbat.block_size
			# add the right number of (1 << 32) - 1 to the table.
			sbat_data << 255.chr * pad
			@header.sbat_start = new_bbat.new_chain sbat_data.length
			io.seek new_bbat.ranges(@header.sbat_start).first.first
			io.write sbat_data
			# now lets write out the bbat. the bbat's chain is not part of the bbat. but maybe i
			# should add blocks to the bbat to hold it.
			# firstly, create the bbat chain's actual chain data:
			bbat_data = new_bbat.table.pack('L*')
			pad = new_bbat.block_size - (bbat_data.length % new_bbat.block_size)
			pad = 0 if pad == new_bbat.block_size
			bbat_data << 255.chr * pad
			# now we just append it as the final blocks in the file.
			# how many blocks will we need?
			@header.num_bat = bbat_data.length / new_bbat.block_size
			base = io.pos / new_bbat.block_size - 1
			io.write bbat_data
			# now that spanned a number of blocks:
			mbat = (0...@header.num_bat).map { |i| i + base }
			mbat += [(1 << 32) - 1] * (109 - mbat.length) if mbat.length < 109
			header_mbat = mbat[0...109]
			other_mbat_data = mbat[109..-1].pack 'L*'
			@header.mbat_start = base + @header.num_bat
			@header.num_mbat = (other_mbat_data.length / new_bbat.block_size.to_f).ceil
			io.write other_mbat_data

			# now seek back and write the header out
			header = @header.instance_variable_get(:@values).pack(Header::PACK) + header_mbat.pack('L*')
			io.seek 0
			io.write header

			# now finally migrate to the new allocation tables, flush, and we're done.
			@bbat = new_bbat
			@sbat = new_sbat
			# for this to fully work, we then need to migrate to it, like:
			# @io = io
			# we also need to update things like the sb_blocks
			@sb_blocks = new_bbat.chain @header.sbat_start
		end
	end
end

