require 'stringio'
require 'strscan'

class StringIO # :nodoc:
	begin
		instance_method :getbyte
	rescue NameError
		alias getbyte getc
	end
end

module Mapi
	#
	# = Introduction
	#
	# The +RTF+ module contains a few helper functions for dealing with rtf
	# in mapi messages: +rtfdecompr+, and <tt>rtf2html</tt>.
	#
	# Both were ported from their original C versions for simplicity's sake.
	#
	module RTF
		class Tokenizer
			def self.process io
				while true do
					case c = io.getc
					when ?{; yield :open_group
					when ?}; yield :close_group
					when ?\\
						case c = io.getc
						when ?{, ?}, ?\\; yield :text, c.chr
						when ?'; yield :text, [io.read(2)].pack('H*')
						when ?a..?z, ?A..?Z
							# read control word
							str = c.chr
							str << c while c = io.read(1) and c =~ /[a-zA-Z]/
							neg = 1
							neg = -1 and c = io.read(1) if c == '-'
							num = if c =~ /[0-9]/
								num = c
								num << c while c = io.read(1) and c =~ /[0-9]/
								num.to_i * neg
							end
							raise "invalid rtf stream" if neg == -1 and !num # ???? \blahblah- some text
							io.seek(-1, IO::SEEK_CUR) if c != ' '
							yield :control_word, str, num
						when nil
							raise "invalid rtf stream" # \EOF
						else
							# other kind of control symbol
							yield :control_symbol, c.chr
						end
					when nil
						return
					when ?\r, ?\n
						# ignore
					else yield :text, c.chr
					end
				end
			end
		end

		class Converter
			# this is pretty crap, its just to ensure there is always something readable if
			# there is an rtf only body, with no html encapsulation.
			def self.rtf2text str, format=:text
				group = 0
				text = ''
				text << "<html>\n<body>" if format == :html
				group_type = []
				group_tags = []
				RTF::Tokenizer.process(StringIO.new(str)) do |a, b, c|
					add_text = ''
					case a
					when :open_group; group += 1; group_type[group] = nil; group_tags[group] = []
					when :close_group; group_tags[group].reverse.each { |t| text << "</#{t}>" }; group -= 1;
					when :control_word; # ignore
						group_type[group] ||= b
						# maybe change this to use utf8 where possible
						add_text = if b == 'par' || b == 'line' || b == 'page'; "\n"
						elsif b == 'tab' || b == 'cell'; "\t"
						elsif b == 'endash' || b == 'emdash'; "-"
						elsif b == 'emspace' || b == 'enspace' || b == 'qmspace'; " "
						elsif b == 'ldblquote'; '"'
						else ''
						end
						if b == 'b' || b == 'i' and format == :html
							close = c == 0 ? '/' : ''
							text << "<#{close}#{b}>"
							if c == 0
								group_tags[group].delete b
							else
								group_tags[group] << b
							end
						end
						# lot of other ones belong in here.\
=begin
\bullet 	Bullet character.
\lquote 	Left single quotation mark.
\rquote 	Right single quotation mark.
\ldblquote 	Left double quotation mark.
\rdblquote
=end
					when :control_symbol; # ignore
						 group_type[group] ||= b
						add_text = ' ' if b == '~' # non-breakable space
						add_text = '-' if b == '_' # non-breakable hypen
					when :text
						add_text = b if group <= 1 or group_type[group] == 'rtlch' && !group_type[0...group].include?('*')
					end
					if format == :html
						text << add_text.gsub(/([<>&"'])/) do
							ent = { '<' => 'lt', '>' => 'gt', '&' => 'amp', '"' => 'quot', "'" => 'apos' }[$1]
							"&#{ent};"
						end
						text << '<br>' if add_text == "\n"
					else
						text << add_text
					end
				end
				text << "</body>\n</html>\n" if format == :html
				text
			end
		end

		RTF_PREBUF =
			"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}" \
			"{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript " \
			"\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier" \
			"{\\colortbl\\red0\\green0\\blue0\n\r\\par " \
			"\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"

		# Decompresses compressed rtf +data+, as found in the mapi property
		# +PR_RTF_COMPRESSED+. Code converted from my C version, which in turn
		# I wrote from a Java source, in JTNEF I believe.
		#
		# C version was modified to use circular buffer for back references,
		# instead of the optimization of the Java version to index directly into
		# output buffer. This was in preparation to support streaming in a
		# read/write neutral fashion.
		def rtfdecompr data
			io  = StringIO.new data
			buf = RTF_PREBUF + "\x00" * (4096 - RTF_PREBUF.length)
			wp  = RTF_PREBUF.length
			rtf = ''

			# get header fields (as defined in RTFLIB.H)
			compr_size, uncompr_size, magic, crc32 = io.read(16).unpack 'V*'
			#warn "compressed-RTF data size mismatch" unless io.size == data.compr_size + 4

			# process the data
			case magic
			when 0x414c454d # "MELA" magic number that identifies the stream as a uncompressed stream
				rtf = io.read uncompr_size
			when 0x75465a4c # "LZFu" magic number that identifies the stream as a compressed stream
				flag_count = -1
				flags = nil
				while rtf.length < uncompr_size and !io.eof?
					# each flag byte flags 8 literals/references, 1 per bit
					flags = ((flag_count += 1) % 8 == 0) ? io.getbyte : flags >> 1
					if 1 == (flags & 1) # each flag bit is 1 for reference, 0 for literal
						rp, l = io.getbyte, io.getbyte
						# offset is a 12 byte number. 2^12 is 4096, so thats fine
						rp = (rp << 4) | (l >> 4) # the offset relative to block start
						l = (l & 0xf) + 2 # the number of bytes to copy
						l.times do
							rtf << buf[wp] = buf[rp]
							wp = (wp + 1) % 4096
							rp = (rp + 1) % 4096
						end
					else
						rtf << buf[wp] = io.getbyte.chr
						wp = (wp + 1) % 4096
					end
				end
			else # unknown magic number
				raise "Unknown compression type (magic number 0x%08x)" % magic
			end

			# not sure if its due to a bug in the above code. doesn't seem to be
			# in my tests, but sometimes there's a trailing null. we chomp it here,
			# which actually makes the resultant rtf smaller than its advertised
			# size (+uncompr_size+).
			rtf.chomp! 0.chr
			rtf
		end

		# Note, this is a conversion of the original C code. Not great - needs tests and
		# some refactoring, and an attempt to correct some inaccuracies. Hacky but works.
		#
		# Returns +nil+ if it doesn't look like an rtf encapsulated rtf.
		#
		# Some cases that the original didn't deal with have been patched up, eg from
		# this chunk, where there are tags outside of the htmlrtf ignore block.
		#
		# "{\\*\\htmltag116 <br />}\\htmlrtf \\line \\htmlrtf0 \\line {\\*\\htmltag84 <a href..."
		#
		# We take the approach of ignoring all rtf tags not explicitly handled. A proper
		# parse tree would be nicer to work with. will need to look for ruby rtf library
		#
		# Some of the original comment to the c code is excerpted here:
		#
		# Sometimes in MAPI, the PR_BODY_HTML property contains the HTML of a message.
		# But more usually, the HTML is encoded inside the RTF body (which you get in the
		# PR_RTF_COMPRESSED property). These routines concern the decoding of the HTML
		# from this RTF body.
		#
		# An encoded htmlrtf file is a valid RTF document, but which contains additional
		# html markup information in its comments, and sometimes contains the equivalent
		# rtf markup outside the comments. Therefore, when it is displayed by a plain
		# simple RTF reader, the html comments are ignored and only the rtf markup has
		# effect. Typically, this rtf markup is not as rich as the html markup would have been.
		# But for an html-aware reader (such as the code below), we can ignore all the
		# rtf markup, and extract the html markup out of the comments, and get a valid
		# html document.
		#
		# There are actually two kinds of html markup in comments. Most of them are
		# prefixed by "\*\htmltagNNN", for some number NNN. But sometimes there's one
		# prefixed by "\*\mhtmltagNNN" followed by "\*\htmltagNNN". In this case,
		# the two are equivalent, but the m-tag is for a MIME Multipart/Mixed Message
		# and contains tags that refer to content-ids (e.g. img src="cid:072344a7")
		# while the normal tag just refers to a name (e.g. img src="fred.jpg")
		# The code below keeps the m-tag and discards the normal tag.
		# If there are any m-tags like this, then the message also contains an
		# attachment with a PR_CONTENT_ID property e.g. "072344a7". Actually,
		# sometimes the m-tag is e.g. img src="http://outlook/welcome.html" and the
		# attachment has a PR_CONTENT_LOCATION "http://outlook/welcome.html" instead
		# of a PR_CONTENT_ID.
		#
		def rtf2html rtf
			scan = StringScanner.new rtf
			# require \fromhtml. is this worth keeping? apparently you see \\fromtext if it
			# was converted from plain text.
			return nil unless rtf["\\fromhtml"]
      if scan.scan_until(/\\ansicpg/)
        code_page = "cp" + scan.scan(/\d+/)
        scan.pos = 0
      else
        code_page = 'ascii'
      end
			html = ''
			ignore_tag = nil
			# skip up to the first htmltag. return nil if we don't ever find one
			return nil unless scan.scan_until /(?=\{\\\*\\htmltag)/
			until scan.empty?
				if scan.scan /\{/
				elsif scan.scan /\}/
				elsif scan.scan /\\\*\\htmltag(\d+) ?/
					#p scan[1]
					if ignore_tag == scan[1]
						scan.scan_until /\}/
						ignore_tag = nil
					end
				elsif scan.scan /\\\*\\mhtmltag(\d+) ?/
						ignore_tag = scan[1]
				elsif scan.scan /\\par ?/
					html << "\r\n"
				elsif scan.scan /\\tab ?/
					html << "\t"
				elsif scan.scan /\\'([0-9A-Za-z]{2})/
					html << scan[1].hex.chr
				elsif scan.scan /\\pntext/
					scan.scan_until /\}/
				elsif scan.scan /\\htmlrtf/
					scan.scan_until /\\htmlrtf0 ?/
				# a generic throw away unknown tags thing.
				# the above 2 however, are handled specially
				elsif scan.scan /\\[a-z-]+(\d+)? ?/
				#elsif scan.scan /\\li(\d+) ?/
				#elsif scan.scan /\\fi-(\d+) ?/
				elsif scan.scan /[\r\n]/
				elsif scan.scan /\\([{}\\])/
					html << scan[1]
				elsif scan.scan /(.)/
					html << scan[1]
				else
					p :wtf
				end
			end
      html.strip.empty? ? nil : html.encode('utf-8', code_page)
		end

		module_function :rtf2html, :rtfdecompr
	end
end

