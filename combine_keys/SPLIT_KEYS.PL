#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Win32::OLE;
use File::Copy;
use File::Slurp;
use List::Util;
use Array::Split qw( split_by split_into );
use POSIX;

my @pinpai = read_file('pinpai.txt');
unlink "Amazonkeys.xlsx", if(-f "Amazonkeys.xlsx");
my $filename = "Amazonkeys.xlsx";
copy("NOT_DEL_AMAZONKEYS.xlsx", "$filename") or die "Copy NOT_DEL_AMAZONKEYS.xlsx failed: $!";

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

# 打开start.xlsx文件
my $Book11 = $Excel->Workbooks->Open("E:\\tool\\combine_keys\\$filename") //  die "Can not open $filename book\n" ;	
my $Sheet11 = $Book11->Worksheets(1);

#将newkeys.txt保存为文件句柄
open(NEWKEYS, "newkeys.txt") or die"";
open(TITLEBULLET, "Our_title_Top10_title_bullet.txt") or die"";

#命名以下将会用到的数组array和标量scalar
my (@newkeys, @allkeys, @allspkeys, @amazonkeys, @ourtitles, @toptitlebullets, @onlytitles, @onlytitlebullets, @specialkeys, @sppointers);
my ($newkey, $allnewkey, $onlynewkey, $onlynewkeys_num, $split_size, $title, $ourtitle, $toptitlebullet, $specialkey, $row_num);


#将newkeys提取出来变成数组
while (<NEWKEYS>) {
	chomp;
    if (/^\s*$/) {
		push @newkeys, $newkey if($newkey ne "" or $newkey ne " ");
		$newkey = "";
		#say "null line";
		next;
	} else {
		$newkey .= " $_";
	}
}

my $count_num = 0;
while (<TITLEBULLET>) {
	$count_num++;
	chomp;
	if($count_num == 1) {
		push @ourtitles, "$_";
	} else {
		next if (/^\s*$/);
		push @toptitlebullets, "$_";
	}
}

#say "toptitlebullets :: @toptitlebullets";

#参数是标量 去除重复的单词
sub removeduplicate {
	my $word = shift;
	my @words = split / /, $word;
	my %count;
	my @newwords = grep { ++$count{ $_ } < 2; } @words;
	my $newword = join ' ', @newwords;
	return $newword;
}


push @ourtitles, "women's womens women woman men's mens men man girl's girls girl boy's boys boy kid's kids kid";
$ourtitle = join ' ', @ourtitles;
$toptitlebullet = join ' ', @toptitlebullets;

#$ourtitle = &removeduplicate($ourtitle);
#$toptitlebullet = &removeduplicate($toptitlebullet);

#大写转换成小写
$ourtitle = lc($ourtitle);
$toptitlebullet = lc($toptitlebullet);

$ourtitle =~ s/\+/ /g;
$ourtitle =~ s/:/ /g;
$ourtitle =~ s/;/ /g;
$ourtitle =~ s/,/ /g;
$ourtitle =~ s/\(/ /g;
$ourtitle =~ s/\)/ /g;
$ourtitle =~ s/\// /g;
$ourtitle =~ s/\./ /g;
$ourtitle =~ s/\?/ /g;
$ourtitle =~ s/\`/ /g;
$ourtitle =~ s/\!/ /g;
$ourtitle = &removeduplicate($ourtitle);
$ourtitle =~ s/^\s+//g;
$ourtitle =~ s/\s+$//g;
$ourtitle =~ s/\ \ \ /\ /g;
$ourtitle =~ s/\ \ /\ /g;

$toptitlebullet =~ s/\+/ /g;
$toptitlebullet =~ s/:/ /g;
$toptitlebullet =~ s/;/ /g;
$toptitlebullet =~ s/,/ /g;
$toptitlebullet =~ s/\(/ /g;
$toptitlebullet =~ s/\)/ /g;
$toptitlebullet =~ s/\// /g;
$toptitlebullet =~ s/\./ /g;
$toptitlebullet =~ s/\?/ /g;
$toptitlebullet =~ s/\`/ /g;
$toptitlebullet =~ s/\!/ /g;
$toptitlebullet = &removeduplicate($toptitlebullet);
$toptitlebullet =~ s/^\s+//g;
$toptitlebullet =~ s/\s+$//g;
$toptitlebullet =~ s/\ \ \ /\ /g;
$toptitlebullet =~ s/\ \ /\ /g;

chomp(@pinpai);
for(@pinpai) {
	if($toptitlebullet =~ /\b$_\b/i) {
		$toptitlebullet =~ s/$_//i;
	}
}

#say "ourtitle::#$ourtitle#";
#say "toptitlebullet::#$toptitlebullet#";


@onlytitles = split / /, $ourtitle;
@onlytitlebullets = split / /, $toptitlebullet;


my $excel_num = 0;
my $five_num = 0;
for(1..40) {
	#将数组打乱重新组合 但不破坏单个元素
	@newkeys=List::Util::shuffle @newkeys;
	#将数组组成标量
	$allnewkey = join ' ', @newkeys;
	#去除重复单词
	$onlynewkey = &removeduplicate($allnewkey);
	#say "onlynewkey  :: $onlynewkey";
	
	#大写转换成小写
	$onlynewkey = lc($onlynewkey);
	
	#去除自己标题的单词
	for(@onlytitles) {
		chomp;
		#s/\b$_\b//ig; 加g是错误的
		$onlynewkey =~ s/\b$_\b//i;
	}
	
	#去除前面及后面的空格
	$onlynewkey =~ s/\+/ /g;
	$onlynewkey =~ s/:/ /g;
	$onlynewkey =~ s/;/ /g;
	$onlynewkey =~ s/,/ /g;
	$onlynewkey =~ s/^\s+//g;
	$onlynewkey =~ s/\s+$//g;
	$onlynewkey =~ s/\ \ \ /\ /g;
	$onlynewkey =~ s/\ \ /\ /g;
	
	for(@onlytitlebullets) {
		chomp;
		unless($onlynewkey =~ /$_/i) {
			unless($ourtitle =~ /$_/i) {
				push @specialkeys, "$_";
			}
		}
	}
	
	$specialkey = join ' ', @specialkeys;
	
	#say "specialkey::#$specialkey#";
	
	#将标量转成数组
	my @onlynewkeys = split / /, $onlynewkey;
	#say "onlynewkeys :: @onlynewkeys";
	#计算数组个数
	$onlynewkeys_num = @onlynewkeys;
	#say "onlynewkeys_num::#$onlynewkeys_num#";
	#分割数组的size进行求余向上取整
	if($onlynewkeys_num < 450) {
		$row_num = 5;
	} elsif(450 < $onlynewkeys_num and $onlynewkeys_num <= 900) {
		$row_num = 10;
	} elsif(900 < $onlynewkeys_num and $onlynewkeys_num <= 1350) {
		$row_num = 15;
	} elsif(1350 < $onlynewkeys_num and $onlynewkeys_num <= 1800) {
		$row_num = 20;
	}
	$split_size = POSIX::ceil($onlynewkeys_num / $row_num);
	#say "split_size = $split_size :: onlynewkeys_num = $onlynewkeys_num :: row_num = $row_num";
	#分割出多个数组
	@allkeys = split_by ( $split_size, @onlynewkeys );
	#say "------------------------------------------------------";
	#拆分出多个数组引用
	
	@allspkeys = split_by ( 30, @specialkeys );
	
	for my $sppointer (@allspkeys) {
		my $i = scalar(@$sppointer);
		$i = 0;
		
		my $sppointer_str = "";
		for(@$sppointer){
			$sppointer_str .= " $$sppointer[$i++]";
		}
		
		$sppointer_str =~ s/^\s+//g;
		$sppointer_str =~ s/\s+$//g;

		push @sppointers, "$sppointer_str";
		#say "sppointer_str::#$sppointer_str#";
	}
	
	@sppointers=List::Util::shuffle @sppointers;
	
	for my $pointer (@allkeys) {
		#say "pointer :: $pointer";
		my $i = scalar(@$pointer);
		#say "i :: $i";
		$i = 0;
		
		my $pointer_str = "";
		for(@$pointer){
			$pointer_str .= " $$pointer[$i++]";
		}
		
		#写到A列
		$excel_num++;
		
		if($excel_num == 1 or $excel_num == 6 or $excel_num == 11 or $excel_num == 16) {
			my $spkey = shift @sppointers;
			say "spkey :: $spkey";
			$pointer_str .= " $spkey";
		}
		
		$five_num++;
		$pointer_str =~ s/\s\"\s//g;
		$pointer_str =~ s/\s\-\s//g;
		$pointer_str =~ s/\s\.\s//g;
		$pointer_str =~ s/^\s+//g;
		$pointer_str =~ s/\s+$//g;
		#say "pointer_str :: $pointer_str";
		

		$Sheet11->Cells($excel_num,1)->{Value} = "$pointer_str";
		push @amazonkeys, "$pointer_str\n";
		#say "";
	}
	
	
}


$Book11->Save;
$Book11->Close;
undef $Book11;

undef $Excel;
