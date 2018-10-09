#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use File::Slurp qw( read_file );
use DDP;
use Spreadsheet::Read;
use Excel::Writer::XLSX;

# 读取google.xlsx A列的关键次数据到数组a
my $book = Spreadsheet::Read->new ("google_keys.xlsx");
my $sheet = $book->sheet (1);
my @google;
for(1..1000) {
	push @google, $sheet->cell("A$_");
	last unless ($sheet->cell("A$_"));
}
############################################################

# 读取pinpai.xlsx A列的关键次数据到数组b
my $ppbook = Spreadsheet::Read->new ("pinpai.xlsx");
my $ppsheet = $ppbook->sheet (1);
my @pinpai;
for(1..361) {
	push @pinpai, $ppsheet->cell("A$_");
	last unless ($ppsheet->cell("A$_"));
}
############################################################

chomp(@google, @pinpai);
my %fake;

foreach(@google) {
	my $google_key = $_;
	for(@pinpai){
		say "$_ :: $google_key";
		if("$google_key" =~ /\b"$_"\b/) {
			#say "$_ :: $google_key";
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

=pod
$num = 1;
for(@google){
	chomp;
	$worksheet->write( "C$num", "$_" );
	$num++;
}
=cut

$num = 1;
my $len = 0;
my $line_data = "";
for(@google){
	chomp;
	$len = $len + length("$_");
	if($len < 400) {
		$line_data = $line_data . " " . "$_";
	} elsif($len >= 400) {
		$line_data = $line_data . " " . "$_";
		$worksheet->write( "C$num", "$line_data" );
		$num++;
		$len = 0;
		$line_data = "";
	}
}


$workbook->close();