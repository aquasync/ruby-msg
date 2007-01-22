require 'logger'

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

class Logger
	def self.new_with_callstack logdev=STDERR
		log = Logger.new logdev
		log.level = WARN
		log.formatter = proc do |severity, time, progname, msg|
			# find where we were called from, in our code
			callstack = caller.dup
			callstack.shift while callstack.first =~ /\/logger\.rb:\d+:in/
			from = callstack.first.sub(/:in `(.*?)'/, ":\\1")
			"[%s %s]\n%-7s%s\n" % [time.strftime('%H:%M:%S'), from, severity, msg.to_s]
		end
		log
	end
end

