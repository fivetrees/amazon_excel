#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;

exit if(@ARGV != 1);
my $filename = "$ARGV[0]";

# 删除wish-*.xlsx
say "";
copy("${filename}", "bak-${filename}") or die "Copy ${filename} failed: $!";


# 读取表格文件
my $book = Spreadsheet::Read->new("$filename");

# 读取亚马逊的模板数据，在excel的第1个工作区
my $sheet = $book->sheet(1);

my @nums;
my $null_num = 0;

for(2..5000) {

	my $sku = $sheet->cell("A$_");
	if($sku ne ""){
		push @nums, $_;
	} elsif($sku eq ""){
		$null_num++;
		last if($null_num == 2);
	} else{
		say "error !!!!";
		last;
	}
}

my $Excel;
# use existing instance if Excel is already running
eval {$Excel = Win32::OLE->GetActiveObject('Excel.Application')};
die "Excel not installed" if $@;
unless (defined $Excel) {
    $Excel = Win32::OLE->new('Excel.Application', sub {$_[0]->Quit;})
          or die "Oops, cannot start Excel";
}
# 关掉Excel的提示，比如是否保存之类的。
$Excel->{DisplayAlerts} = 'False'; 
 
sub update_price {
	# open existing excel document
	my $Book = $Excel->Workbooks->Open("E:\\tool\\update_price\\$filename") //  die "Can not open $filename book\n" ;	
	# 使用该Excel文档中名为"template"的Sheet
	my $Sheet = $Book->Worksheets(1);

	#写数据到表格里面
	for my $num (@nums){
		
		my $cur_price = $Sheet->Cells($num,3)->{Value};
		my $upd_price = int($cur_price * 1.7);
		$Sheet->Cells($num,3)->{Value} = $upd_price;
		
	}

	$Book->Save;
	$Book->Close;
	undef $Book;
}

&update_price;

undef $Excel;
