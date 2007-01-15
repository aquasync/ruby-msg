require 'ostruct'
require 'iconv'

module Ole
	#
	# Ole class
	# ------------------------------------------------------------------------------
	# simple class that is intended to abstract away the details of the access to
	# the ole objects. currently wraps a simple perl dumper based on
	# OLE::Storage_Lite, with base64 encoding to get around some YAML incompatibilities
	#
	# not used any more, as the replacement ruby version is ready (ole/storage)
	# moved here from main msg.rb
	#
	class Storage1 < OpenStruct
		UTF16_TO_UTF8 = Iconv.new('utf-8', 'utf-16le').method :iconv

		TYPE_MAP = {
			1 => :dir,
			2 => :file,
			5 => :root
		}

		def self.load filename
			data = File.popen "./ole2yaml.pl '#{filename.gsub /'/, %{'"'"'}}'" do |pipe|
				YAML.load(pipe)
			end
			Ole.decode data
		end
		
		def self.decode data
			data['name'] = UTF16_TO_UTF8[data['name']]
			data['data'] = Base64.decode64(data['data']) if data['data']
			data['kind'] = TYPE_MAP[data['kind']] or raise "unknown ole type #{data['kind']}"

			if data['children']
				data['children'].map! { |child| Ole.decode child }
			end

			Ole.new data
		end

		def dir?
			# treat root as a dir
			kind != :file
		end

		def file?
			kind == :file
		end

		def inspect
			"#<Ole::Storage @name=#{name.inspect}>"
		end
	end

	# would be interesting to try a simple poledump powered version.
	# just parses the plain output, for directory structure, then can
	# extract specific streams when #data is requested
end


