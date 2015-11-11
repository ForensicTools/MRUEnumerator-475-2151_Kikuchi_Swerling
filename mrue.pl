#!/usr/bin/perl

use strict;
use warnings;
use File::Basename;

#ascii art
print "███╗   ███╗██████╗ ██╗   ██╗███████╗\n";
print "████╗ ████║██╔══██╗██║   ██║██╔════╝\n";
print "██╔████╔██║██████╔╝██║   ██║█████╗  \n";
print "██║╚██╔╝██║██╔══██╗██║   ██║██╔══╝  \n";
print "██║ ╚═╝ ██║██║  ██║╚██████╔╝███████╗\n";
print "╚═╝     ╚═╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝ v1.0\n\n";

#expanded acronym
print "Most\n";
print "Recently\n";
print "Used\n";
print "Enumerator\n\n";

#banner
print "Hello!\n";
print "This tool is a modified version of RegRipper that will locate\n"; 
print "and enumerate common MRU lists that have been created by both\n";
print "Windows and other applications; it will then take all results\n";
print "and visualize them using D3.js for easier consumption.\n\n";

#prerequisites
print "Please be sure to have your registry files in the EVIDENCE folder!\n\n";

#wait for user to read intro and continue
print "Press any key to continue...\n\n";
<STDIN>;

#clear the terminal
system( "clear" );

################
#officedocs2010#
################

#call rip.pl, use the "officedocs" plugin, and put the output in a temporary file
system( qq{ perl rip.pl -r ./EVIDENCE/NTUSER.DAT -p officedocs2010 > "./TMP/output.tmp" });

#open the output file or die
my $fh;
open( $fh, "<", "./TMP/output.tmp" )
	or die "Fatal Error - Cannot open output.tmp.\n";

#read the contents of the file into an array
my @array;
while( <$fh> ){
	chomp;
	push @array, $_;
}

#close the file handle
close $fh;

#arrays to hold output
my @word;
my @excel;
my @access;
my @powerpoint;

#section keeper
my $section = 0;

#sort data
foreach ( @array ){
	#if the line is the beginning of a new section, increment $newsect
	if ( $_ =~ m/^Software./ ){
		$section++;

	#else add value to corrseponding array based on $newsect
	} else {
		if ( $section == 1 ){
			if ( $_ =~ m/^  Item./ ){
				push @word, $_;
			}
		}
		elsif ( $section == 2 ){
			if ( $_ =~ m/^  Item./ ){
				push @excel, $_;
			}
		}
		elsif ( $section == 3 ){
			if ( $_ =~ m/^  Item./ ){
				push @access, $_;
			}
		}
		elsif ( $section == 4 ){
			if ( $_ =~ m/^  Item./ ){
				push @powerpoint, $_;
			}
		}
	}
}

#remove "Item X -> " from the beginning of each element from all arrays
#replace all "\" with "/"
#remove trailing day, date, time, and year
#grab the filename from the full path
foreach ( @word ){
	$_ =~ s/^[^>]*> //;
	$_ =~ s/\\/\//g;
	$_ = substr( $_, 0, -26 );
	$_ = fileparse( $_ );
}

foreach ( @excel ){
	$_ =~ s/^[^>]*> //;
}

foreach ( @access ){
	$_ =~ s/^[^>]*> //;
}

foreach ( @powerpoint ){
	$_ =~ s/^[^>]*> //;
}

#open file handle in order to make output json
open( $fh, ">", "./RESULTS/results.json" )
	or die "Fatal Error - Cannot create output json.\n";

#print data to json
print $fh "{\n";
print $fh " \"name\": \"Main\",\n";
print $fh " \"children\": [\n";
print $fh "  {\n";
print $fh "   \"name\": \"MS Office\",\n";
print $fh "   \"children\": [\n";
print $fh "    {\n";
print $fh "     \"name\": \"Word\",\n";
print $fh "     \"children\": [\n";

#get size of array
my $arraysize = scalar( @word);

#subtract one to account for index 0, another to omit the last element
$arraysize = $arraysize - 2;

#print all but the last element
foreach my $x ( 0..$arraysize ){
	print $fh "      {\"name\": \"@word[ $x ]\", \"size\": 5000},\n";
}

#print the last element
my $lastelement = pop @word;
print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";

print $fh "     ]\n";
print $fh "    }\n";
print $fh "   ]\n";
print $fh "  }\n";
print $fh " ]\n";
print $fh "}\n";























