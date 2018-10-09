#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Win32::OLE;
use File::Copy;


#先删除ERROR.txt
unlink "ERROR.txt", if(-f "ERROR.txt");


# amazon xlsl 转换成wish xlsl
#my $filename = "$ARGV[0]";
my $filename = "a.xlsx";


#初始化excel文件
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

# 打开us.xlsx文件
my $Book11 = $Excel->Workbooks->Open("E:\\tool\\amazon_convert_wish_shoes\\$filename") //  die "Can not open $filename book\n" ;	
my $Sheet11 = $Book11->Worksheets(4);


#读取新文件名
my $psku	= $Sheet11->range("A4")->{Value};
my $ptitle 	= $Sheet11->range("B4")->{Value};
my @huohaos = split '-', $psku;
my $huohao 	= $huohaos[-1];
my @ptitles = split ' ', $ptitle;
my $brand   = $ptitles[0];
my $bullet_point1 = $Sheet11->range("AX5")->{Value};
my $bullet_point2 = $Sheet11->range("AY5")->{Value};
my $bullet_point3 = $Sheet11->range("AZ5")->{Value};
my $bullet_point4 = $Sheet11->range("BA5")->{Value};
my $bullet_point5 = $Sheet11->range("BB5")->{Value};
my $desc = "$bullet_point1" . "\n" . "\n" . "$bullet_point2" . "\n" . "\n" . "$bullet_point3" . "\n" . "\n" . "$bullet_point4" . "\n" . "\n" . "$bullet_point5"; 

#拷贝ca excel文件
my $newxlsx = "wish-${brand}-${huohao}.xlsx";
copy("wish_original.xlsx", "$newxlsx") or die "Copy wish_original.xlsx failed: $!";

# 打开新文档 写数据
my $Book22 = $Excel->Workbooks->Open("E:\\tool\\amazon_convert_wish_shoes\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
my $Sheet22 = $Book22->Worksheets(1);



my $line = 0;
for(5..200) {
	say "No Defined A$_" and last unless(defined $Sheet11->range("A$_")->{Value});
	$line = $_;
}
say "line is $line";

my $wishline = $line - 3;
sub copy_paste_multi {
	my $t1 = shift;
	my $t2 = shift;
	
	$Sheet11->range("${t1}5:${t1}${line}")->copy();
	$Sheet22->range("${t2}2:${t2}${wishline}")->Select();
	$Sheet22->paste();
}

#AMAZON左列 WISH右列
my %cphash = qw(
A B
B C
CH D
CI E
C F
J J
K K
BG N
BH O
BI P
BJ Q
BK R
BL S
);

$Sheet22->Range("A2:A$wishline") ->{Value} = "$psku"; #父SKU
$Sheet22->Range("G2:G$wishline") ->{Value} = "100"; #库存
$Sheet22->Range("L2:L$wishline") ->{Value} = "4"; #运费
$Sheet22->Range("I2:I$wishline") ->{Value} = "$desc"; #描述
#$Sheet22->Range("M2:M$wishline") ->{Value} = "7-15"; #


while(my($key, $value) = each %cphash) {
	&copy_paste_multi("$key", "$value");
}


$Book11->Close;
undef $Book11;

$Book22->Save;
$Book22->Close;
undef $Book22;

undef $Excel;


=pod
Amazon Wish clothes


=cut