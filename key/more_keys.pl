#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Slurp qw( read_file );
use DDP;
use Spreadsheet::Read;
use Excel::Writer::XLSX;
use List::Util;
 
# 读取google.xlsx A列的关键次数据到数组a
my $book = Spreadsheet::Read->new ("google_keys.xlsx");
my $sheet = $book->sheet (1);
my @google;
my %h;
my $r;


for(1..50000) {
	last unless ($sheet->cell("A$_"));
	push @google, $sheet->cell("A$_");
}

# 去除google中重复元素
my %saw;
@saw{ @google } = ( );
@google = sort keys %saw;

############################################################
# 读取pinpai.xlsx A列的关键次数据到数组b
my $ppbook = Spreadsheet::Read->new ("pinpai.xlsx");
my $ppsheet = $ppbook->sheet (1);
my @pinpai;
#1475
for(1..10000) {
	last unless ($ppsheet->cell("A$_"));
	push @pinpai, $ppsheet->cell("A$_");
}
############################################################

chomp(@google, @pinpai);
my %fake;

foreach(@google) {
	my $google_key = $_;
	for(@pinpai){
		if($google_key =~ /\b$_\b/i) {
			#say "$_";
			$fake{$google_key} = $_;
		}
	}
}

#p %fake;


# Create a new Excel workbook
my $workbook = Excel::Writer::XLSX->new( 'fake.xlsx' );
# Add a worksheet
my $worksheet = $workbook->add_worksheet();

my $num = 1;
while(my($key, $value) = each %fake) {
	#进行过滤google品牌词提取出好的
	@google = grep { $_ ne "$key" } @google;
	$worksheet->write( "A$num", "$key" );
	$worksheet->write( "B$num", "$value" );
	$num++;
}

sub gen_search_key {
	#随机打乱数组
	@google=List::Util::shuffle @google;

	my $len = 0;
	my $line_data = "";
	for(@google){
		chomp;
		$len = $len + length("$_");
		if($len < 940) {
			$line_data = "$_" . " " . "$line_data";
		} elsif($len >= 940) {
			#$line_data = $line_data . " " . "$_";
			$worksheet->write( "C$num", "$line_data" );
			$num++;
			$len = 0;
			$line_data = "";
		}
	}
}

$num = 1;
for(1..10) {
	&gen_search_key;
}

$workbook->close();