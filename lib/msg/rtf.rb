require 'stringio'

class Msg
	#
	# = Introduction
	#
	# The +RTF+ module contains a few helper functions for dealing with rtf
	# in msgs: +rtfdecompr+, and <tt>rtf2html</tt>.
	#
	# Both were ported from their original C versions for simplicity's sake.
	#
	module RTF
		RTF_PREBUF = 
			"{\\rtf1\\ansi\\mac\\deff0\\deftab720{\\fonttbl;}" \
			"{\\f0\\fnil \\froman \\fswiss \\fmodern \\fscript " \
			"\\fdecor MS Sans SerifSymbolArialTimes New RomanCourier" \
			"{\\colortbl\\red0\\green0\\blue0\n\r\\par " \
			"\\pard\\plain\\f0\\fs20\\b\\i\\u\\tab\\tx"

		# Decompresses compressed rtf +data+, as found in the mapi property
		# +PR_RTF_COMPRESSED+. Code converted from my C version, which in turn
		# was ported from Java source, in JTNEF I believe.
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
			compr_size, uncompr_size, magic, crc32 = io.read(16).unpack 'L*'
			#warn "compressed-RTF data size mismatch" unless io.size == data.compr_size + 4

			# process the data
			case magic
			when 0x414c454d # magic number that identifies the stream as a uncompressed stream
				rtf = io.read uncompr_size
			when 0x75465a4c # magic number that identifies the stream as a compressed stream
				flag_count = -1
				flags = nil
				while rtf.length < uncompr_size and !io.eof?
					#p [rtf.length, uncompr_size]
					# each flag byte flags 8 literals/references, 1 per bit
					flags = ((flag_count += 1) % 8 == 0) ? io.getc : flags >> 1
					if 1 == (flags & 1) # each flag bit is 1 for reference, 0 for literal
						rp, l = io.getc, io.getc
						# offset is a 12 byte number. 2^12 is 4096, so thats fine
						rp = (rp << 4) | (l >> 4) # the offset relative to block start
						l = (l & 0xf) + 2 # the number of bytes to copy
						l.times do
							rtf << (buf[wp] = buf[rp])
							wp = (wp + 1) % 4096
							rp = (rp + 1) % 4096
						end
					else
						rtf << (buf[wp] = io.getc)
						wp = (wp + 1) % 4096
					end
				end
			else # unknown magic number
				raise "Unknown compression type (magic number 0x%08x)" % magic
			end
			rtf
		end

=begin
# = RTF/HTML functions
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
# This code is experimental. It works on my own message archive, of about
# a thousand html-encoded messages, received in Outlook97 and Outlook2000
# and OutlookXP. But I can't guarantee that it will work on all rtf-encoded
# messages. Indeed, it used to be the case that people would simply stick
# {\fromhtml at the start of an html document, and } at the end, and send
# this as RTF. If someone did this, then it will almost work in my function
# but not quite. (Because I ignore \r and \n, and respect only \par. Thus,
# any linefeeds in the erroneous encoded-html will be ignored.)

# ISRTFHTML -- Given an uncompressed RTF body of the message, this
# function tells you whether it encodes some html.
# [in] (buf,*len) indicate the start and length of the uncompressed RTF body.
# [return-value] true or false, for whether it really does encode some html
bool isrtfhtml(const char *buf,unsigned int len)
{ // We look for the words "\fromhtml" somewhere in the file.
  // If the rtf encodes text rather than html, then instead
  // it will only find "\fromtext".
  const char *c;
  for (c=buf; c<buf+len; c++)
  { if (strncmp(c,"\\from",5)==0) return strncmp(c,"\\fromhtml",9)==0;
  }
  return false;
}


# DECODERTFHTML -- Given an uncompressed RTF body of the message,
# and assuming that it contains encoded-html, this function
# turns it onto regular html.
# [in] (buf,*len) indicate the start and length of the uncompressed RTF body.
# [out] the buffer is overwritten with the HTML version, null-terminated,
# and *len indicates the length of this HTML.
#
# Notes: (1) because of how the encoding works, the HTML version is necessarily
# shorter than the encoded version. That's why it's safe for the function to
# place the decoded html in the same buffer that formerly held the encoded stuff.
# (2) Some messages include characters \'XX, where XX is a hexedecimal number.
# This function simply converts this into ASCII. The conversion will only make
# sense if the right code-page is being used. I don't know how rtf specifies which
# code page it wants.
# (3) By experiment, I discovered that \pntext{..} and \liN and \fi-N are RTF
# markup that should be removed. There might be other RTF markup that should
# also be removed. But I don't know what else.
#
void decodertfhtml(char *buf,unsigned int *len)
{ // c -- pointer to where we're reading from
  // d -- pointer to where we're writing to. Invariant: d<c
  // max -- how far we can read from (i.e. to the end of the original rtf)
  // ignore_tag -- stores 'N': after \mhtmlN, we will ignore the subsequent \htmlN.
  char *c=buf, *max=buf+*len, *d=buf; int ignore_tag=-1;
  // First, we skip forwards to the first \htmltag.
  while (c<max && strncmp(c,"{\\*\\htmltag",11)!=0) c++;
  //
  // Now work through the document. Our plan is as follows:
  // * Ignore { and }. These are part of RTF markup.
  // * Ignore \htmlrtf...\htmlrtf0. This is how RTF keeps its equivalent markup separate from the html.
  // * Ignore \r and \n. The real carriage returns are stored in \par tags.
  // * Ignore \pntext{..} and \liN and \fi-N. These are RTF junk.
  // * Convert \par and \tab into \r\n and \t
  // * Convert \'XX into the ascii character indicated by the hex number XX
  // * Convert \{ and \} into { and }. This is how RTF escapes its curly braces.
  // * When we get \*\mhtmltagN, keep the tag, but ignore the subsequent \*\htmltagN
  // * When we get \*\htmltagN, keep the tag as long as it isn't subsequent to a \*\mhtmltagN
  // * All other text should be kept as it is.
=end


		# html encoded in rtf comments.
		# {\*\htmltag84 &quot;}\htmlrtf "\htmlrtf0

		# already generates better output that the c predecessor. eg from this chunk, where
		# there are tags outside of the htmlrtf ignore block. 
		# "{\\*\\htmltag116 <br />}\\htmlrtf \\line \\htmlrtf0 \\line {\\*\\htmltag84 <a href..."
		# we take the approach of ignoring
		# all rtf tags not explicitly handled. a proper parse tree would be nicer to work with.
		# ruby rtf library?
		# check http://homepage.ntlworld.com/peterhi/rtf_tools.html
		# and
		# http://rubyforge.org/projects/ruby-rtf/

		# Substandard conversion of the original C code.
		# Test and refactor, and try to correct some inaccuracies.
		# Returns +nil+ if it doesn't look like an rtf encapsulated rtf.
		#
		# Code is a hack, but it works.
		def rtf2html rtf
			scan = StringScanner.new rtf
			# require \fromhtml. is this worth keeping?
			return nil unless rtf["\\fromhtml"]
			html = ''
			ignore_tag = nil
			# skip up to the first htmltag. return nil if we don't ever find one
			return nil unless scan.scan_until /(?=\{\\\*\\htmltag)/
			until scan.empty?
				if scan.scan /\{/
				elsif scan.scan /\}/
				elsif scan.scan /\\\*\\htmltag(\d+) ?/
					p scan[1]
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
			html
		end

		module_function :rtf2html, :rtfdecompr
	end
end

