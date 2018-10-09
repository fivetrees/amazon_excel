#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Slurp qw( read_file );
use DDP;
use Spreadsheet::Read;
use Excel::Writer::XLSX;
use List::Util qw(shuffle);

# 读取google.xlsx A列的关键次数据到数组a
my $book = Spreadsheet::Read->new ("google_keys.xlsx");
my $sheet = $book->sheet (1);
my @google;
for(1..1695) {
	push @google, $sheet->cell("A$_");
	last unless ($sheet->cell("A$_"));
}
############################################################

# 读取pinpai.xlsx A列的关键次数据到数组b
my $ppbook = Spreadsheet::Read->new ("pinpai.xlsx");
my $ppsheet = $ppbook->sheet (1);
my @pinpai;
for(1..897) {
	push @pinpai, $ppsheet->cell("A$_");
	last unless ($ppsheet->cell("A$_"));
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


@google = shuffle(@google);

$num = 1;
for(@google){
	chomp;
	say "$_ ...";
	$worksheet->write( "C$num", "$_" );
	$num++;
}


$workbook->close();