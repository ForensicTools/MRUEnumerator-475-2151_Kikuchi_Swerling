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
	#if the line is the beginning of a new section, increment $section
	if ( $_ =~ m/^Software./ ){
		$section++;

	#else add value to corrseponding array based on $section
	} else {
		if ( $section == 1 ){
			if ( $_ =~ m/^  Item./ ){
				push @word, $_;
			}
		} elsif ( $section == 2 ){
			if ( $_ =~ m/^  Item./ ){
				push @excel, $_;
			}
		} elsif ( $section == 3 ){
			if ( $_ =~ m/^  Item./ ){
				push @access, $_;
			}
		} elsif ( $section == 4 ){
			if ( $_ =~ m/^  Item./ ){
				push @powerpoint, $_;
			}
		}
	}
}

#find the size of all arrays
my $wordsize = scalar( @word );
my $excelsize = scalar( @excel );
my $accesssize = scalar( @access );
my $powerpointsize = scalar( @powerpoint );

#calculate process & print flag
my $ppflag = $wordsize . $excelsize . $accesssize . $powerpointsize;

#did office run flag
my $office2010rf = -1;

#if larger than 0000, call office2010pp sub
if ( $ppflag > 0 )
{
	$office2010rf = 1;
	&office2010pp;

#else print user message
} else {
	$office2010rf = 0;
	print "No Office 2010 data found.\n";
}

##########
#adoberdr#
##########

#call rip.pl, use the "officedocs" plugin, and put the output in a temporary file
system( qq{ perl rip.pl -r ./EVIDENCE/NTUSER.DAT -p adoberdr > "./TMP/output.tmp" });

#open the output file or die
open( $fh, "<", "./TMP/output.tmp" )
	or die "Fatal Error - Cannot open output.tmp.\n";

#read the contents of the file into an array
while( <$fh> ){
	chomp;
	push @array, $_;
}

#close the file handle
close $fh;

#array to hold output
my @adoberdr;

#sort data
foreach ( @array ){
	#if the line is a key, push it into the main data array
	if ( $_ =~ m/^c[0-9]*[0-9]*[0-9],./ ){
		push @adoberdr, $_;
	}
}

#pull out just the path
my @temp;

#split each element into its three parts
#replace the current element with just the full path
#remove the "/" at the beginning of each element
#replace all "\" with "/"
#grab the filename from the full path and throw away everything else
#number all elements
#increment counter
#reset counter after each loop
my $counter = 1;
foreach( @adoberdr ){
	my @temp = split( ',', $_ );
	$_ = $temp[ 1 ];
	$_ = substr( $_, 1, -1 );
	$_ =~ s/\\/\//g;
	$_ = fileparse( $_ );
	$_ = "$counter $_";
	$counter++;
}

#if office2010 did not produce results create a new file
if( $office2010rf == 0 ){

	#open file handle in order to make output json
	open( $fh, ">", "./RESULTS/data.json" )
		or die "Fatal Error - Cannot create output json.\n";

	#print header data to json
	print $fh "{\n";
	print $fh " \"name\": \"Main\",\n";
	print $fh " \"children\": [\n";
	print $fh "  {\n";
	print $fh "   \"name\": \"Adobe Reader\",\n";
	print $fh "   \"children\": [\n";
	print $fh "    {\n";

	

#else open and append to the existing file
} else {
	#open file handle in order to make output json
	open( $fh, ">>", "./RESULTS/data.json" )
		or die "Fatal Error - Cannot open output json.\n";

	print $fh "  {\n";
	print $fh "   \"name\": \"Adobe Reader\",\n";
	print $fh "   \"children\": [\n";
}

#get size of word array
my $arraysize = scalar( @adoberdr );

#just print nothing
if( $arraysize == 0 ){
	print $fh "     {\"name\": \"No Adobe Reader Data.\", \"size\": 5000}\n";
	
#print the only element
} elsif ( $arraysize == 1 ){
	#print the last element
	my $lastelement = pop @adoberdr;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
#print element 0 & 1
} elsif ( $arraysize == 2 ){
	$arraysize = 1;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$adoberdr[ $x ]\", \"size\": 5000},\n";
	}
	
#all other array sizes (2+)
} else {
	#subtract one to account for index 0, another to omit the last element
	$arraysize = $arraysize - 2;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$adoberdr[ $x ]\", \"size\": 5000},\n";
	}

	#print the last element
	my $lastelement = pop @adoberdr;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
}

#print closing braces
print $fh "   ]\n";
print $fh "  }\n";
print $fh " ]\n";
print $fh "}\n";

##################
#officedocs2010pp#
##################

sub office2010pp
{
	#remove "Item X -> " from the beginning of each element from all arrays
	#replace all "\" with "/"
	#remove trailing day, date, time, and year
	#grab the filename from the full path and throw away everything else
	#number all elements
	#increment counter
	#reset counter after each loop
	my $counter = 1;
	foreach ( @word ){
		$_ =~ s/^[^>]*> //;
		$_ =~ s/\\/\//g;
		$_ = substr( $_, 0, -26 );
		$_ = fileparse( $_ );
		$_ = "$counter $_";
		$counter++; 
	}
	$counter = 1;
	foreach ( @excel ){
		$_ =~ s/^[^>]*> //;
		$_ =~ s/\\/\//g;
		$_ = substr( $_, 0, -26 );
		$_ = fileparse( $_ );
		$_ = "$counter $_";
		$counter++; 
	}
	$counter = 1;
	foreach ( @access ){
		$_ =~ s/^[^>]*> //;
		$_ =~ s/\\/\//g;
		$_ = substr( $_, 0, -26 );
		$_ = fileparse( $_ );
		$_ = "$counter $_";
		$counter++; 
	}
	$counter = 1;
	foreach ( @powerpoint ){
		$_ =~ s/^[^>]*> //;
		$_ =~ s/\\/\//g;
		$_ = substr( $_, 0, -26 );
		$_ = fileparse( $_ );
		$_ = "$counter $_";
		$counter++; 
	}

	#open file handle in order to make output json
	open( $fh, ">", "./RESULTS/data.json" )
		or die "Fatal Error - Cannot create output json.\n";

	#print data to json
	print $fh "{\n";
	print $fh " \"name\": \"Main\",\n";
	print $fh " \"children\": [\n";
	print $fh "  {\n";
	print $fh "   \"name\": \"MS Office\",\n";
	print $fh "   \"children\": [\n";
	print $fh "    {\n";

	#WORD section
	print $fh "     \"name\": \"Word\",\n";
	print $fh "     \"children\": [\n";

	#get size of word array
	my $arraysize = scalar( @word);

	#print "no data" child
	if( $arraysize == 0 ){
		print $fh "       {\"name\": \"No Word Data.\", \"size\": 5000}\n";

	#print the only element	
	} elsif ( $arraysize == 1 ){
		#print the last element
		my $lastelement = pop @word;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
	#print element 0 & 1
	} elsif ( $arraysize == 2 ){
		$arraysize = 1;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$word[ $x ]\", \"size\": 5000},\n";
		}

	#all other array sizes (2+)
	} else {
		#subtract one to account for index 0, another to omit the last element
		$arraysize = $arraysize - 2;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$word[ $x ]\", \"size\": 5000},\n";
		}

		#print the last element
		my $lastelement = pop @word;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";
	}

	#print closing child bracket and new parent brackets
	print $fh "     ]\n";
	print $fh "    },\n";
	print $fh "    {\n";

	#EXCEL section
	print $fh "     \"name\": \"Excel\",\n";
	print $fh "     \"children\": [\n";

	#get size of word array
	$arraysize = scalar( @excel);

	#print "no data" child
	if( $arraysize == 0 ){
		print $fh "       {\"name\": \"No Excel Data.\", \"size\": 5000}\n";
	
	#print the only element
	} elsif ( $arraysize == 1 ){
		#print the last element
		my $lastelement = pop @excel;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";

	#print element 0 & 1	
	} elsif ( $arraysize == 2 ){
		$arraysize = 1;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$powerpoint[ $x ]\", \"size\": 5000},\n";
		}

	#all other array sizes (2+)
	} else {
		#subtract one to account for index 0, another to omit the last element
		$arraysize = $arraysize - 2;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$excel[ $x ]\", \"size\": 5000},\n";
		}

		#print the last element
		my $lastelement = pop @excel;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";
	}

	#print closing child bracket and new parent brackets
	print $fh "     ]\n";
	print $fh "    },\n";
	print $fh "    {\n";

	#ACCESS section
	print $fh "     \"name\": \"Access\",\n";
	print $fh "     \"children\": [\n";

	#get size of word array
	$arraysize = scalar( @access);

	#print "no data" child
	if( $arraysize == 0 ){
		print $fh "       {\"name\": \"No Access Data.\", \"size\": 5000}\n";
	
	#print the only element
	} elsif ( $arraysize == 1 ){
		#print the last element
		my $lastelement = pop @access;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
	#print element 0 & 1	
	} elsif ( $arraysize == 2 ){
		$arraysize = 1;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$access[ $x ]\", \"size\": 5000},\n";
		}
	
	#all other array sizes (2+)
	} else {
		#subtract one to account for index 0, another to omit the last element
		$arraysize = $arraysize - 2;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$access[ $x ]\", \"size\": 5000},\n";
		}

		#print the last element
		my $lastelement = pop @access;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";
	}

	#print closing child bracket and new parent brackets
	print $fh "     ]\n";
	print $fh "    },\n";
	print $fh "    {\n";

	#POWERPOINT section
	print $fh "     \"name\": \"Powerpoint\",\n";
	print $fh "     \"children\": [\n";

	#get size of word array
	$arraysize = scalar( @powerpoint);

	#print "no data" child
	if( $arraysize == 0 ){
		print $fh "       {\"name\": \"No PowerPoint Data.\", \"size\": 5000}\n";
	
	#print the only element
	} elsif ( $arraysize == 1 ){
		#print the last element
		my $lastelement = pop @powerpoint;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
	#print element 0 & 1
	} elsif ( $arraysize == 2 ){
		$arraysize = 1;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$powerpoint[ $x ]\", \"size\": 5000},\n";
		}
	
	#all other array sizes (2+)
	} else {
		#subtract one to account for index 0, another to omit the last element
		$arraysize = $arraysize - 2;

		#print all but the last element
		foreach my $x ( 0..$arraysize ){
			print $fh "      {\"name\": \"$powerpoint[ $x ]\", \"size\": 5000},\n";
		}

		#print the last element
		my $lastelement = pop @powerpoint;
		print $fh "      {\"name\": \"$lastelement\", \"size\": 5000}\n";
	}

	#print closing braces
	print $fh "     ]\n";
	print $fh "    }\n";
	print $fh "   ]\n";
	print $fh "  },\n";
}























