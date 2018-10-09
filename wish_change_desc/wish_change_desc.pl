#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;
use LWP::Simple;


unlink "ERROR.txt", if(-f "ERROR.txt");

#定义当前目录
my $mydir = "E:/tool/wish_change_desc/imgdir";

chdir "$mydir";
my @files = <*>;


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
 
# 新描述
my $wish_desc = "Welcome to our store! We will provide you high quality products.
Due to the light and screen, our color have little difference, please forgive us.
If you have any question about our products, please contact us in time.
Please check your order information before you finsh your order.
Please check your size again before you make order.";

for my $newxlsx (@files) {

	# open existing excel document
	my $Book = $Excel->Workbooks->Open("E:\\tool\\wish_change_desc\\imgdir\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
	my $Sheet = $Book->Worksheets(1);

	for my $num (2..200) {

		my $desc = $Sheet->range("H$num")->{Value};
		last unless($desc);
		say "$newxlsx H$num";

		if($desc ne ""){
			$Sheet->range("H$num")->{Value} = "$wish_desc";
		}
	}

	$Book->Save;
	$Book->Close;
	undef $Book;

}


undef $Excel;


