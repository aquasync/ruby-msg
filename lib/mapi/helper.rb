module Mapi
	class Helper
		# @return [String, nil]
		attr_reader :ansi_encoding
		# @return [Boolean]
		attr_reader :to_unicode

		# @param ansi_encoding [String]
		# @param to_unicode [Boolean]
		def initialize ansi_encoding=nil, to_unicode=false
			@ansi_encoding = ansi_encoding || "BINARY"
			@to_unicode = to_unicode
		end

		# Convert to requested format:
		#
		# - non Unicode string property
		# - body (rtf, text)
		#
		# @param str [String]
		# @return [Object]
		def convert_ansi_str str
			if @ansi_encoding
				if @to_unicode
					str.force_encoding(@ansi_encoding).encode("UTF-8")
				else
					str.force_encoding(@ansi_encoding)
				end
			else
				str
			end
		end
	end
end