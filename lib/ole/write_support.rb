=begin

plan is to begin with writing out a tree, as opposed to the api for the creation of that
tree in the first place.
essentially simple - reverse the stuff done in the load function.

TODO

* lock it down with tests.
* refactor. first within this file, moving things from the monolithic save to
  correct classes. neaten, try to leave the ole object in a sane state after
  saving, etc, etc.
	its currently incredibly UGLY
* then maybe merge this stuff into the core, and attempt to fill out api for
  creating an ole document from scratch.
* finally create further tests, and refactor until it feels nice and clean
  again.

* need to rethink what happens here. i now have write support for RangesIO, so
  save doesn't have to be monolithic like it is. it could be incremental.
	what about dealing with the inconsistencies in state.

its probably time to think about the api.
two things. i want to support loading IO or filename. i will provide opens:
Ole::Storage.open 'wordfile.doc'
io = open('wordfile.doc')
Ole::Storage.load io
to open a new file
Ole::Storage.open 'wordfile.doc', 'w'

i think the basics however, is:
ole = Ole::Storage.new
completely plain. where do file writes go?
ole files are completely loaded, not lazily. except for the underlying file data.
this is backed by a file:

Ole::Storage.open do |ole|
	# we have opened an ole storage object. as it wasn't otherwise specified, it is backed
	# by a stringio.
	ole.file.open("\001CompObj") { |f| p f.read }
	ole.dir.open('/').entries
end

=end

class IO
	def self.copy src, dst
		until src.eof?
			buf = src.read(4096)
			dst.write buf
		end
	end
end

class File
	def size
		stat.size
	end
end

class RangesIO
	def write data
		data_pos = 0
		# we have a certain amount of available room, and don't currently provide a facility to
		# expand it. i will later provide a hook to allow you to provide additional ranges.
		# then the Ole::Storage code can find an available block in the allocation table and add
		# its range, or expand the allocation table etc. but then i'd probably also want to provide
		# for a truncate call. in the case where i start with an empty allocation table, the effect
		# should be identical.
		raise "unable to satisfy write of #{data.length} bytes" if data.length > @size - @pos
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
end

module Ole
	class Storage
		class Header
			def save
				@values.pack PACK
			end
		end

		class AllocationTable
			AVAIL		 = 0xffffffff # a free block (I don't currently leave any blocks free)
			EOC			 = 0xfffffffe # end of a chain
			# these blocks correspond to the bat, and aren't part of a file, nor available.
			# (I don't currently output these)
			BAT			 = 0xfffffffd
			META_BAT = 0xfffffffc

			# up till now allocationtable's didn't know their block size. not anymore...
			# this should replace @range_conv
			def type
				@range_conv.to_s[/^[^_]+/].to_sym
			end

			def save
				table = @table
				# pad it out some
				num = @ole.header.b_size / 4
				# do you really use AVAIL? they probably extend past end of file, and may shortly
				# be used for the bat. not really good.
				table += [AVAIL] * (num - (table.length % num)) if (table.length % num) != 0
				table.pack 'L*'
			end

			def block_size
				@ole.header.send type == :big ? :b_size : :s_size
			end

			# just allocate space for a chain
			def new_chain size
				# no need to worry about empty blocks etc, because we just create from scratch
				chain_head = @table.length
				(size / block_size.to_f).ceil.times { |i| @table << table.length + 1 }
				# turns out this was what the problem was with word not accepting my file. i was
				# using AVAIL.
				@table[-1] = EOC
				chain_head
			end

			def new_chain2 size, io=nil, padding=nil
				# no need to worry about empty blocks etc, because we just create from scratch
				chain_head = @table.length
				num_blocks = (size / block_size.to_f).ceil
				num_blocks.times { |i| @table << table.length + 1 }
				@table[-1] = EOC
				padding_bytes = num_blocks * block_size - size
				if block_given?
					io.seek ranges(chain_head).first.first
					yield
					(padding_bytes / padding.length).times { io.write padding } if padding
				end
				chain_head
			end

			def new_chain3 size, io
				# no need to worry about empty blocks etc, because we just create from scratch
				chain_head = @table.length
				num_blocks = (size / block_size.to_f).ceil
				num_blocks.times { |i| @table << table.length + 1 }
				@table[-1] = EOC
				padding_bytes = num_blocks * block_size - size
				if block_given?
					# we can't use the #to_io method, because it uses @io, not io
					# rangesio will seek all over the joint. do those places have to exist first?
					# can we tell the file to be at least that big? like IO#truncate?
					sub_io = RangesIO.new(io, ranges(chain_head))
					max = sub_io.ranges.map { |pos, len| pos + len }.max
					# maybe its ok to just seek out there later??
					io.truncate max if max < io.size
					yield sub_io
					# further, if there are some unused are in the block, we fill it with zeros:
					sub_io.write 0.chr * (sub_io.size - sub_io.pos)
					#(padding_bytes / padding.length).times { io.write padding } if padding
				end
				chain_head
			end
		end

# maybe cleaner to have a SmallAllocationTable, and a BigAllocationTable.

		class Dirent
			# used in the next / prev / child stuff to show that the tree ends here.
			EOT = 0xffffffff

			MEMBERS.each_with_index do |sym, i|
				define_method(sym.to_s + '=') { |val| @values[i] = val }
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
		end

		def save io
			io.rewind
			io.truncate 0
			# empty header to be filled in later.
			#io.write 0.chr * @header.b_size

			# recreate dirs from our tree, split into dirs and big and small files
			@dirents = @root.flatten
			dirs, files = @dirents.partition &:dir?
			big_files, small_files = files.partition { |file| file.size > @header.threshold }

			# maybe later it'd be neater to do @dirents.each { |d| d.save io } ?
			# they have an @ole, which means they know the bat....

			# do big files.
			# we need a temporary bbat, not to clear the real one, as the file io later requires it
			# as we don't eagerly load all io at once.
			new_bbat = AllocationTable.new self, :big_block_ranges
			big_files.each do |file|
				# we need to get the io before we rewrite its first block etc
				file_io = file.io
				file_io.seek 0
				file.size = file_io.size
				file.first_block = new_bbat.new_chain3(file.size, io) { |sub_io| IO.copy file_io, sub_io }
			end

			# now tackle small files, and the sbblocks
			# then put sbblocks in root first_block
			# first lets see how big all small files are:
			new_sbat = AllocationTable.new self, :small_block_ranges
			@root.first_block = new_bbat.new_chain3 small_files.map(&:size).sum, io do |sub_io|
				# the big blocks that the sb_blocks file live in needs padding, and so does the small
				# block files too. but new_chain2's seeking stuff won't work for sbat
				small_files.each do |file|
					# we need to get the io before we rewrite its first block etc
					file_io = file.io
					file.size = file_io.size
					file.first_block = new_sbat.new_chain file.size
					file_io.seek 0
					IO.copy file_io, sub_io
					# now make up the gap. i think we only pad to the sblock boundary
					pad = new_sbat.block_size - (file.size % new_sbat.block_size)
					pad = 0 if pad == new_sbat.block_size
					sub_io.write 0.chr * pad
				end
			end

			# now if i add write support to rangedio, its as simple as turning that chain into an
			# io, and writing to it. given my current approach though, i know my chains are created
			# sequentially.

			# now that we've done all the files, we know enough to write out the dirents
			@header.dirent_start = new_bbat.new_chain3 @dirents.length * Dirent::SIZE, io do |sub_io|
				@dirents.each { |dirent| sub_io.write dirent.save }
			end

			# now write out the various meta blocks, fill out header a bit
			# ok. how to do this. first thing i suppose is to make room for the size of the sbat
			sbat_data = new_sbat.save
			@header.sbat_start = new_bbat.new_chain3 sbat_data.length, io do |sub_io|
				sub_io.write sbat_data
			end

			# now lets write out the bbat. the bbat's chain is not part of the bbat. but maybe i
			# should add blocks to the bbat to hold it.
			# firstly, create the bbat chain's actual chain data:
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

			# now seek back and write the header out
			io.seek 0
			io.write @header.save + header_mbat.pack('L*')

			# now finally migrate to the new allocation tables, flush, and we're done.
			@bbat = new_bbat
			@sbat = new_sbat
			# for this to fully work, we then need to migrate to it, like:
			@io = io
			# we also need to update things like the sb_blocks
			@sb_blocks = new_bbat.chain @header.sbat_start
		end
	end
end

