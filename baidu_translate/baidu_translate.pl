#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use LWP::UserAgent;
use Digest::MD5;
use Encode;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;


sub baidu_translate {
	my $q = shift;
	my $appid = '20170504000046317';
	my $secert = 'oRFnDOfTz0Tm5a32tzTN';
	my $fanyiurl = "http://api.fanyi.baidu.com/api/trans/vip/translate?";

	#my $q = "hello world";
	my $salt = int(rand(20000)) + 100;

	my $string = "$appid" . "$q" . "$salt" . "$secert";
	my $sign = Digest::MD5->new->add($string)->hexdigest;
	my $url = "$fanyiurl" . "q=" . "$q" . "&from=en&to=jp&appid=" . "$appid" . "&salt=" . "$salt" . "&sign=" . "$sign";

	my $ua = new LWP::UserAgent;
	$ua->timeout(10);
	my $response = $ua->get( $url );
  
	if ($response->is_success) {
		my $tmp = $response->decoded_content;
		$tmp =~ s/\\u([0-9a-fA-F]{4})/pack("U",hex($1))/eg;
		$tmp = encode("gb2312",$tmp);
		#say "$tmp";
		return $tmp;
	} else {
		die $response->status_line;
	}
}


exit if(@ARGV != 1);
my $filename = "$ARGV[0]";
my $newxlsx = "amazonjp-${filename}";

# 删除amazonjp-*.xlsx
say "";
unlink glob "amazonjp-*.xlsx";
copy("original.xlsx", "$newxlsx") or die "Copy original.xlsx failed: $!";

my ($sku, $title);

# 读取excel文件
my $book = Spreadsheet::Read->new("$filename");

# 读取亚马逊的模板数据，在excel的第四个工作区
my $sheet = $book->sheet(4);

# 读取sku
$sku 			= $sheet->cell("A4");
$title			= $sheet->cell("B4");

my ($pnum, $psku, $ptitle, @nums);
my $null_num = 0;

for(4..80) {
	my $color = $sheet->cell("CL$_");
	my $size = $sheet->cell("DK$_");
	my $sku = $sheet->cell("A$_");
	if($sku ne "" and $color eq "" and $size eq "") {
		$pnum = $_;
		$psku = $sku;
		$ptitle = $sheet->cell("B$_");
	} elsif($sku ne "" and $color ne "" and $size ne ""){
		push @nums, $_;
	} elsif($sku eq "" and $color eq "" and $size eq ""){
		#say "line $_ ## null number $null_num ## null data";
		#say "sku is $sku ## size is $size ## color is $color";
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


sub amazonjp {
	# open existing excel document
	my $Book = $Excel->Workbooks->Open("E:\\tool\\baidu_translate\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
	# 使用该Excel文档中名为"Upload Template"的Sheet
	my $Sheet = $Book->Worksheets('Template');

	#写数据到表格里面
	for my $num (@nums){

		#标题B4 及马来西亚标题B4
		#$Sheet->Cells($num,78)->{Value} = $sheet->cell("B$num");
		my $temp = $sheet->cell("B$num");
		say "temp is $temp";
		$Sheet->Cells($num,78)->{Value} = &baidu_translate($sheet->cell("B$num"));	
		
	}
	
	$Book->Save;
	#$Book->SaveAs("C:\\Users\\senlin\\linux\\$newxlsx") or die "Save $newxlsx failer.";
	$Book->Close;
	undef $Book;

}


&amazonjp;

undef $Excel;
