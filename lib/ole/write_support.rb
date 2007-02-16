=begin

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

but then, i lose 2 main features. the ability to easily assign to io, as just another property.
ie, i can currently do something like:
dirent.name = 'new name' # <= this is equivalent to "renaming"
dirent.io = open 'some file'
that doesn't seem like a big loss. i also lose the tree re-assignment:

ole = Ole::Storage.new
ole.instance_eval { @root = some_other_dirent }

but that seems like a supreme hack anyway. the equivalent to that would involve some sort
of tree copying implementation, like:

def Ole::Storage.copy dirent_from, path_to
	if dirent_from.file?
		open(path_to, 'w') { |f| ... }
	else
		dirent_from.each do |de|
			copy de, path_to + '/' + de.name
		end
	end
end

ole = Ole::Storage.new ...
ole.file.read("\001CompObj")

i'll try packing that on top for now.

ole = Ole::Stroage.open ...

# planned new interface for dirent manipulation:

Dirent#open, for read / write to existing dirents. ie not by-name access.
Dirent#read shortcut. ie:
ole.root["\001CompObj"].read
dirent = ole.root.last
p dirent.name
dirent.open do |io|
	io.read 5
	io.seek 0
	io.truncate
	io.write 'hello there'
end

# need to other things, deletion of existing dirents, and creation of new. do
# both through parent as we don't have access to our parent otherwise.

dirent.delete child

close will only stuff with things if you open with 'w'. or maybe 'w+', otherwise you'll
be re-writing dirents of a file you just meant to read from. the full file_system module
will be available and recommended usage, allowing Ole::Storage, Dir, and Zip::ZipFile to be
used pretty exchangably down the track. should be possible to write a recursive copy using
the plain api, such that you can copy dirs/files agnostically between any of ole docs, dirs,
and zip files.

# finally, planned interface for files:
ole = Ole::Storage.new/open filename/io object
ole.root. .....
# this finalises in the way our save currently does.
ole.close

# then we just need to be able to make ole in memory:
# eg, for a future conversion of embedded excel file into a genuine ole file
# attachment, in Attachment#to_mime, i will handle ole files by saving them:
stringio = StringIO.new
Ole::Storage.new stringio do |ole|
	Dirent.copy ole.root, msg.attachments[0].data
end
stringio.string
# then base64 encode and away we go...
=end

module Ole
	class Storage
		class Dirent
			# and for creation of a dirent. don't like the name. is it a file or a directory?
			# assign to type later? io will be empty.
			def new_child
				child = Dirent.new ole
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

			def / path
				self[path]
			end

			def self.copy src, dst
				src.name = dst.name
				if src.dir?
					raise unless dst.dir?
					src.children.each do |src_child|
						dst.new_child { |dst_child| Dirent.copy src_child, dst_child }
					end
				else
					raise unless dst.file?
					src.open do |src_io|
						dst.open { |dst_io| IO.copy src_io, dst_io }
					end
				end
			end
		end
	end
end

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
	attr_reader :first_block
	# temporarily putting this here. requires @bat, and @first_block. maybe a @dirent would
	# serve better.
	def truncate size
		# note that old_blocks is != @ranges.length necessarily. i'm planning to write a
		# merge_ranges function that merges sequential ranges into one as an optimization.
		old_num_blocks, new_num_blocks = [@size, size].map { |i| (i / @bat.block_size.to_f).ceil }
		if old_num_blocks == new_num_blocks
			@ranges = @bat.ranges @first_block, size
			@size = size
			return
		end
		blocks = @bat.chain @first_block
		at = Ole::Storage::AllocationTable
		if new_num_blocks < old_num_blocks
		# old_num_blocks should be blocks.length.
			# de-allocate some of our old blocks. TODO maybe zero them out in the file???
#			sample chain     [2, 1, 3]
#			sample bat table [AVAIL, 3, 1, EOC]
			(new_num_blocks...old_num_blocks).each { |i| @bat.table[blocks[i]] = at::AVAIL }
			# put EOC, but handle corner case of truncate 0
			if new_num_blocks > 0
				@bat.table[blocks[new_num_blocks-1]] = at::EOC 
			else
				@first_block = at::EOC
			end
		else # new_num_blocks > old_num_blocks
			# need some more blocks.
			last_block = blocks.last
			(new_num_blocks - old_num_blocks).times do
				block = @bat.get_free_block
				# connect the chain. handle corner case of blocks being [] initially
				if last_block
					@bat.table[last_block] = block 
				else
					@first_block = block
				end
				last_block = block
				# this is just to inhibit the problem where it gets picked as being a free block
				# again next time around.
				@bat.table[last_block] = at::EOC
			end
		end
		
		@ranges = @bat.ranges @first_block, size
		raise "something bogus happened: #{ranges.inspect} doesn't have #{new_num_blocks} blocks" unless @ranges.length == new_num_blocks

		# don't know if this is required, but we explicitly request our @io to grow if necessary
		# we never shrink it though. maybe this belongs in allocationtable, where smarter decisions
		# can be made.
		max = @ranges.map { |pos, len| pos + len }.max || 0
		# maybe its ok to just seek out there later??
		@io.truncate max if max > @io.size
		@size = size
	end

	def write data
		data_pos = 0
		# we have a certain amount of available room, and don't currently provide a facility to
		# expand it. i will later provide a hook to allow you to provide additional ranges.
		# then the Ole::Storage code can find an available block in the allocation table and add
		# its range, or expand the allocation table etc. but then i'd probably also want to provide
		# for a truncate call. in the case where i start with an empty allocation table, the effect
		# should be identical.
		if data.length > @size - @pos
			# need to get more bytes
			unless respond_to? :truncate
				raise "unable to satisfy write of #{data.length} bytes" 
				# FIXME maybe warn instead, then just truncate the data?
			else
				truncate @pos + data.length
				#p "made space by truncation. resized to #{@size} bytes"
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
end

module Ole
	class Storage
		class Header
			def save
				@values.pack PACK
			end
		end

		class AllocationTable
			# up till now allocationtable's didn't know their block size. not anymore...
			# this should replace @range_conv
			def type
				@range_conv.to_s[/^[^_]+/].to_sym
			end

			def get_free_block
				@table.each_index { |i| return i if @table[i] == AVAIL }
				@table.push AVAIL
				@table.length - 1
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

			# just allocate space for a chain. only used by sbat at the moment
			def new_chain size
				# no need to worry about empty blocks etc, because we just create from scratch
				chain_head = @table.length
				(size / block_size.to_f).ceil.times { |i| @table << table.length + 1 }
				# turns out this was what the problem was with word not accepting my file. i was
				# using AVAIL.
				@table[-1] = EOC
				chain_head
			end

			# create a non-allocated chain, and provide an io to write to it, updating the
			# allocation table as needed. only used by bbat at the moment, sbat needs migrating.
			# for sbat to use it properly, sbat ranges would need to be moved to RangesIO on top of
			# the sb_blocks implied file.
			def new_chain4 io
				sub_io = RangesIO.new io, []
				bat = self
				sub_io.instance_eval { @bat = bat; @first_block = EOC }
				if block_given?
					yield sub_io
					# further, if there are some unused are in the block, we fill it with zeros:
					sub_io.write 0.chr * (sub_io.size - sub_io.pos)
					nil
				else
					sub_io
				end
			end
		end

# maybe cleaner to have a SmallAllocationTable, and a BigAllocationTable.

		class Dirent
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
			dirs, files = @dirents.partition(&:dir?)
			big_files, small_files = files.partition { |file| file.size > @header.threshold }

			# maybe later it'd be neater to do @dirents.each { |d| d.save io } ?
			# they have an @ole, which means they know the bat....

			# do big files.
			# we need a temporary bbat, not to clear the real one, as the file io later requires it
			# as we don't eagerly load all io at once.
			new_bbat = AllocationTable.new self, :big_block_ranges
			big_files.each do |file|
				new_bbat.new_chain4 io do |sub_io|
					# we need to get the io before we rewrite its first block etc
					file_io = file.io
					file_io.seek 0
					IO.copy file_io, sub_io

					# now rewrite the file's dirent
					file.size = sub_io.size
					file.first_block = sub_io.first_block
					# when this block closes, the sub_io is automatically padded
				end
			end

			# now tackle small files, and the sbblocks
			# then put sbblocks in root first_block
			new_sbat = AllocationTable.new self, :small_block_ranges
			new_bbat.new_chain4 io do |sub_io|
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
				@root.first_block = sub_io.first_block
			end

			# it may be better, rather than using a new chain, to use the existing dirent chain,
			# truncate it to 0, and re-write it.
			new_bbat.new_chain4 io do |sub_io|
				@dirents.each { |dirent| sub_io.write dirent.save }
				@header.dirent_start = sub_io.first_block
			end

			# now write out the various meta blocks, fill out header a bit
			new_bbat.new_chain4 io do |sub_io|
				sub_io.write new_sbat.save
				@header.sbat_start = sub_io.first_block
			end

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

