class Symbol
	def to_proc
		proc { |a| a.send self }
	end
end

module Enumerable
	def group_by
		hash = Hash.new { |hash, key| hash[key] = [] }
		each { |item| hash[yield(item)] << item }
		hash
	end
end

