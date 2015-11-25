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
print "╚═╝     ╚═╝╚═╝  ╚═╝ ╚═════╝ ╚══════╝ v1.5\n\n";

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
print "Please be sure to have NTUSER.DAT in the EVIDENCE folder!\n\n";

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

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/WORD.txt" )
	or die "Fatal Error - Cannot open WORD.txt.\n";

#export sorted word data
#if array is size 0, export no data
if( scalar( @word ) == 0 ){
	print $fh "No Word data found.";

#else export all data
} else{
	foreach( @word ){
		my $temp = $_;
		$temp = substr( $_, 2 );
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/EXCEL.txt" )
	or die "Fatal Error - Cannot open EXCEL.txt.\n";

#export sorted excel data
#if array is size 0, export no data
if( scalar( @excel ) == 0 ){
	print $fh "No Excel data found.";

#else export all data
} else{
	foreach( @excel ){
		my $temp = $_;
		$temp = substr( $_, 2 );
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/ACCESS.txt" )
	or die "Fatal Error - Cannot open ACCESS.txt.\n";

#export sorted access data
#if array is size 0, export no data
if( scalar( @access ) == 0 ){
	print $fh "No Access data found.";

#else export all data
} else{
	foreach( @access ){
		my $temp = $_;
		$temp = substr( $_, 2 );
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/POWERPOINT.txt" )
	or die "Fatal Error - Cannot open POWERPOINT.txt.\n";

#export sorted powerpoint data
#if array is size 0, export no data
if( scalar( @powerpoint ) == 0 ){
	print $fh "No PowerPoint data found.";

#else export all data
} else{
	foreach( @powerpoint ){
		my $temp = $_;
		$temp = substr( $_, 2 );
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#call office2010pp sub
&office2010pp;

##########
#adoberdr#
##########

#call rip.pl, use the "adoberdr" plugin, and put the output in a temporary file
system( qq{ perl rip.pl -r ./EVIDENCE/NTUSER.DAT -p adoberdr > "./TMP/output.tmp" });

#open the output file or die
open( $fh, "<", "./TMP/output.tmp" )
	or die "Fatal Error - Cannot open output.tmp.\n";

#read the contents of the file into an array
@array = ();
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

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/ADOBE READER.txt" )
	or die "Fatal Error - Cannot open ADOBE READER.txt.\n";

#print header
print $fh "File name\n";
print $fh "sDate\n";
print $fh "uFileSize\n";
print $fh "uPageCount\n\n";

#export sorted adobe reader data
#if array is size 0, export no data
my @temp;
if( scalar( @adoberdr ) == 0 ){
	print $fh "No Adobe Reader data found.";

#else export all data
} else{
	foreach( @adoberdr ){
		@temp = split( ',', $_ );
		#remove "/" from the beginning of line and "\00" from end of line
		$temp[ 1 ] = substr( $temp[ 1 ], 1, -1 );
		print $fh "$temp[ 1 ]\n";
		#remove "/" from the beginning of line and "\00" from end of line
		$temp[ 2 ] = substr( $temp[ 2 ], 1, -1 );
		print $fh "$temp[ 2 ]\n";
		#remove "/" from the beginning of line and "\00" from end of line
		$temp[ 3 ] = substr( $temp[ 3 ], 1, -1 );
		print $fh "$temp[ 3 ]\n";
		#remove "/" from the beginning of line and "\00" from end of line
		$temp[ 4 ] = substr( $temp[ 4 ], 1 );
		print $fh "$temp[ 4 ]\n\n";
	}
}

#close file handle
close $fh;

#split each element into its three parts
#replace the current element with just the full path
#remove the "/" at the beginning of each element
#replace all "\" with "/"
#grab the filename from the full path and throw away everything else
#number all elements
#increment counter
my $counter = 1;
foreach( @adoberdr ){
	@temp = ();
	@temp = split( ',', $_ );
	$_ = $temp[ 1 ];
	$_ = substr( $_, 1, -1 );
	$_ =~ s/\\/\//g;
	$_ = fileparse( $_ );
	$_ = "$counter $_";
	$counter++;
}

#open file handle in order to make output json
open( $fh, ">>", "./RESULTS/data.json" )
	or die "Fatal Error - Cannot open output json.\n";

print $fh "  {\n";
print $fh "   \"name\": \"Adobe Reader\",\n";
print $fh "   \"children\": [\n";

#get size of adoberdr array
my $arraysize = scalar( @adoberdr );

#print "no data" child
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

#print closing brackets & braces
print $fh "   ]\n";
print $fh "  },\n";

#close file handle
close $fh;

################
#wordwheelquery#
################

#call rip.pl, use the "wordwheelquery" plugin, and put the output in a temporary file
system( qq{ perl rip.pl -r ./EVIDENCE/NTUSER.DAT -p wordwheelquery > "./TMP/output.tmp" });

#open the output file or die
open( $fh, "<", "./TMP/output.tmp" )
	or die "Fatal Error - Cannot open output.tmp.\n";

#read the contents of the file into an array
@array = ();
while( <$fh> ){
	chomp;
	push @array, $_;
}

#close the file handle
close $fh;

#array to hold output
my @wwq;

#sort data
foreach ( @array ){
	#if the line is a key, push it into the main data array
	if ( $_ =~ m/^[0-9]*[0-9]*[0-9]./ ){
		push @wwq, $_;
	}
}

#reset counter
#empty temp array
#split all elements using whitespace
#replace the current element with just the query
#number all elements
#increment counter
$counter = 1;
foreach( @wwq ){
	@temp = ();
	@temp = split( /\s+\s*\s/, $_ );
	$_ = pop( @temp );
	$_ = "$counter $_";
	$counter++;
}

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/WORDWHEELQUERY.txt" )
	or die "Fatal Error - Cannot open WORDWHEELQUERY.txt.\n";

#export sorted wordwheelquery data
#if array is size 0, export no data
if( scalar( @wwq ) == 0 ){
	print $fh "No WordWheelQuery data found.";

#else export all data
} else{
	foreach( @wwq ){
		my $temp = $_;
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#open file handle in order to make output json
open( $fh, ">>", "./RESULTS/data.json" )
	or die "Fatal Error - Cannot open output json.\n";

print $fh "  {\n";
print $fh "   \"name\": \"WordWheelQuery\",\n";
print $fh "   \"children\": [\n";

#get size of wwq array
$arraysize = scalar( @wwq );

#print "no data" child
if( $arraysize == 0 ){
	print $fh "     {\"name\": \"No WordWheelQuery Data.\", \"size\": 5000}\n";
	
#print the only element
} elsif ( $arraysize == 1 ){
	#print the last element
	my $lastelement = pop @wwq;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
#print element 0 & 1
} elsif ( $arraysize == 2 ){
	$arraysize = 1;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$wwq[ $x ]\", \"size\": 5000},\n";
	}
	
#all other array sizes (2+)
} else {
	#subtract one to account for index 0, another to omit the last element
	$arraysize = $arraysize - 2;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$wwq[ $x ]\", \"size\": 5000},\n";
	}

	#print the last element
	my $lastelement = pop @wwq;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
}

#print closing brackets & braces
print $fh "   ]\n";
print $fh "  },\n";

#close file handle
close $fh;

########
#runmru#
########

#call rip.pl, use the "runmru" plugin, and put the output in a temporary file
system( qq{ perl rip.pl -r ./EVIDENCE/NTUSER.DAT -p runmru > "./TMP/output.tmp" });

#open the output file or die
open( $fh, "<", "./TMP/output.tmp" )
	or die "Fatal Error - Cannot open output.tmp.\n";

#read the contents of the file into an array
@array = ();
while( <$fh> ){
	chomp;
	push @array, $_;
}

#close the file handle
close $fh;

#array to hold output
my @runmru;

#sort data
foreach ( @array ){
	#if the line is a key, push it into the main data array
	if ( $_ =~ m/^[a-z]\s/ ){
		push @runmru, $_;
	}
}

#reset counter
#empty temp array
#split all elements using whitespace
#replace the current element with just the query
#remove trailing "\1" at the end of each element
#number all elements
#increment counter
$counter = 1;
foreach( @runmru ){
	@temp = ();
	@temp = split( /  /, $_ );
	$_ = pop( @temp );
	$_ = substr( $_, 0, -2 );
	$_ = "$counter $_";
	$counter++;
}

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/RUNMRU.txt" )
	or die "Fatal Error - Cannot open RUNMRU.txt.\n";

#export sorted runmru data
#if array is size 0, export no data
if( scalar( @runmru ) == 0 ){
	print $fh "No RunMRU data found.";

#else export all data
} else{
	foreach( @runmru ){
		my $temp = $_;
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#open file handle in order to make output json
open( $fh, ">>", "./RESULTS/data.json" )
	or die "Fatal Error - Cannot open output json.\n";

print $fh "  {\n";
print $fh "   \"name\": \"RunMRU\",\n";
print $fh "   \"children\": [\n";

#get size of runmru array
$arraysize = scalar( @runmru );

#print "no data" child
if( $arraysize == 0 ){
	print $fh "     {\"name\": \"No RunMRU Data.\", \"size\": 5000}\n";
	
#print the only element
} elsif ( $arraysize == 1 ){
	#print the last element
	my $lastelement = pop @runmru;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
#print element 0 & 1
} elsif ( $arraysize == 2 ){
	$arraysize = 1;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$runmru[ $x ]\", \"size\": 5000},\n";
	}
	
#all other array sizes (2+)
} else {
	#subtract one to account for index 0, another to omit the last element
	$arraysize = $arraysize - 2;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$runmru[ $x ]\", \"size\": 5000},\n";
	}

	#print the last element
	my $lastelement = pop @runmru;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
}

#print closing brackets & braces
print $fh "   ]\n";
print $fh "  },\n";

#close file handle
close $fh;

############
#recentdocs#
############

#call rip.pl, use the "recentdocs" plugin, and put the output in a temporary file
system( qq{ perl rip.pl -r ./EVIDENCE/NTUSER.DAT -p recentdocs > "./TMP/output.tmp" });

#open the output file or die
open( $fh, "<", "./TMP/output.tmp" )
	or die "Fatal Error - Cannot open output.tmp.\n";

#read the contents of the file into an array
@array = ();
while( <$fh> ){
	chomp;
	push @array, $_;
}

#close the file handle
close $fh;

#array to hold output
my @recentdocs;

#reset section keeper
$section = 0;

#sort data
foreach ( @array ){
	#if the line is the beginning of a new section, increment $section
	if ( $_ =~ m/^Software./ ){
		$section++;

	#else add value to corrseponding array based on $section
	} else {
		if ( $section == 1 ){
			if ( $_ =~ m/^  [0-9]*[0-9]*[0-9]\s[=]{1}./ ){
				push @recentdocs, $_;
			}
		}
	}
}

#reset counter
#empty temp array
#split all elements based on regex
#replace the current element with just the query
#number all elements
#increment counter
$counter = 1;
foreach( @recentdocs ){
	@temp = ();
	@temp = split( /^\s\s[0-9]*[0-9]*[0-9]\s[=]\s/, $_ );
	$_ = pop( @temp );
	$_ = "$counter $_";
	$counter++;
}

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/RECENTDOCS.txt" )
	or die "Fatal Error - Cannot open RECENTDOCS.txt.\n";

#export sorted recentdocs data
#if array is size 0, export no data
if( scalar( @recentdocs ) == 0 ){
	print $fh "No RecentDocs data found.";

#else export all data
} else{
	foreach( @recentdocs ){
		my $temp = $_;
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#open file handle in order to make output json
open( $fh, ">>", "./RESULTS/data.json" )
	or die "Fatal Error - Cannot open output json.\n";

print $fh "  {\n";
print $fh "   \"name\": \"Recent Docs\",\n";
print $fh "   \"children\": [\n";

#get size of recentdocs array
$arraysize = scalar( @recentdocs );

#print "no data" child
if( $arraysize == 0 ){
	print $fh "     {\"name\": \"No Recent Docs Data.\", \"size\": 5000}\n";
	
#print the only element
} elsif ( $arraysize == 1 ){
	#print the last element
	my $lastelement = pop @recentdocs;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
#print element 0 & 1
} elsif ( $arraysize == 2 ){
	$arraysize = 1;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$recentdocs[ $x ]\", \"size\": 5000},\n";
	}
	
#all other array sizes (2+)
} else {
	#subtract one to account for index 0, another to omit the last element
	$arraysize = $arraysize - 2;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$recentdocs[ $x ]\", \"size\": 5000},\n";
	}

	#print the last element
	my $lastelement = pop @recentdocs;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
}

#print closing brackets & braces
print $fh "   ]\n";
print $fh "  },\n";

#close file handle
close $fh;

#####
#mmc#
#####

#call rip.pl, use the "mmc" plugin, and put the output in a temporary file
system( qq{ perl rip.pl -r ./EVIDENCE/NTUSER.DAT -p mmc > "./TMP/output.tmp" });

#open the output file or die
open( $fh, "<", "./TMP/output.tmp" )
	or die "Fatal Error - Cannot open output.tmp.\n";

#read the contents of the file into an array
@array = ();
while( <$fh> ){
	chomp;
	push @array, $_;
}

#close the file handle
close $fh;

#array to hold output
my @mmc;

#reset section keeper
$section = 0;

#sort data
foreach ( @array ){
	#if the line is a key, push it into the main data array
	if ( $_ =~ m/^\s\sFile[0-9]*[0-9]*[0-9]\s[-][>]/ ){
		push @mmc, $_;
	}
}

#open the export file or die
open( $fh, ">", "./RESULTS/DATA/MMC.txt" )
	or die "Fatal Error - Cannot open MMC.txt.\n";

#export sorted mmc data
#if array is size 0, export no data
if( scalar( @mmc ) == 0 ){
	print $fh "No MMC data found.";

#else export all data
} else{
	foreach( @mmc ){
		my $temp = $_;
		$temp = substr( $temp, 2 );
		print $fh "$temp\n";
	}
}

#close file handle
close $fh;

#reset counter
#empty temp array
#split all elements based on regex
#replace the current element with just the query
#replace all "\" with "/"
#grab the filename from the full path and throw away everything else
#number all elements
#increment counter
$counter = 1;
foreach( @mmc ){
	@temp = ();
	@temp = split( /^\s\sFile[0-9]*[0-9]*[0-9]\s[-][>]\s/, $_ );
	$_ = pop( @temp );
	$_ =~ s/\\/\//g;
	$_ = fileparse( $_ );
	$_ = "$counter $_";
	$counter++;
}

#open file handle in order to make output json
open( $fh, ">>", "./RESULTS/data.json" )
	or die "Fatal Error - Cannot open output json.\n";

print $fh "  {\n";
print $fh "   \"name\": \"MMC\",\n";
print $fh "   \"children\": [\n";

#get size of mmc array
$arraysize = scalar( @mmc );

#print "no data" child
if( $arraysize == 0 ){
	print $fh "     {\"name\": \"No MMC Data.\", \"size\": 5000}\n";
	
#print the only element
} elsif ( $arraysize == 1 ){
	#print the last element
	my $lastelement = pop @mmc;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
	
#print element 0 & 1
} elsif ( $arraysize == 2 ){
	$arraysize = 1;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$mmc[ $x ]\", \"size\": 5000},\n";
	}
	
#all other array sizes (2+)
} else {
	#subtract one to account for index 0, another to omit the last element
	$arraysize = $arraysize - 2;

	#print all but the last element
	foreach my $x ( 0..$arraysize ){
		print $fh "    {\"name\": \"$mmc[ $x ]\", \"size\": 5000},\n";
	}

	#print the last element
	my $lastelement = pop @mmc;
	print $fh "    {\"name\": \"$lastelement\", \"size\": 5000}\n";
}

#print END B&B
print $fh "   ]\n";
print $fh "  }\n";
print $fh " ]\n";
print $fh "}\n";

#close file handle
close $fh;

#notify user that everything's all set and done
print "\nAll Done! - Go to the RESULTS folder and open RESULTS.html in your browser.\n";
print "Full details can be found in the DATA folder.\n";

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

	#print closing brackets & braces
	print $fh "     ]\n";
	print $fh "    }\n";
	print $fh "   ]\n";
	print $fh "  },\n";
}























