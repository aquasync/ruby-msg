#!/usr/bin/perl -w

use strict;
use OLE::Storage_Lite;
use locale;
#use YAML;
use YAML::Syck;
use POSIX qw(strftime);
use MIME::Base64;

sub item2yaml;
sub item2yaml {
	my $pps = shift;
	my $type = $pps->{Type};
	my $datefmt = '%Y-%m-%d %H:%M:%S';

	# no longer use this character conversion
	# OLE::Storage_Lite::Ucs2Asc
	my $item = {
		name => $pps->{Name},
		kind => $type,
		number => $pps->{No},
	};

#	$item->{name} =~ s/\x00//g;

	if ($type == 2) {
		$item->{size} = $pps->{Size};
		$item->{data} = encode_base64 $pps->{Data};
#		my $x = $pps->{Data};
#		$x =~ s/\x00//g;
#		$x =~ s/\x{00}//g;
#		$x =~ s/\z//g;
#		$item->{data} = $x;
	} else {
		my $time = $pps->{Time2nd} || $pps->{Time1st};
		$item->{time} = strftime($datefmt, @{$time}) . " +00:00";
	}

   	if ($pps->{Child} and length(@{$pps->{Child}})) {
		my @children = map { item2yaml $_ } @{$pps->{Child}};
		$item->{children} = [@children]; #\@children;
	}

	return $item;
}

for (@ARGV) {
	my $ole = OLE::Storage_Lite->new($_);
	my $pps = $ole->getPpsTree(1);
	die "fatal: `$_' not a valid OLE file.\n" unless $pps;
	print Dump item2yaml $pps;
}

