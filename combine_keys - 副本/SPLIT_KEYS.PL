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
my (@newkeys, @allkeys, @amazonkeys, @ourtitles, @toptitlebullets, @onlytitles, @onlytitlebullets, @specialkeys);
my ($newkey, $allnewkey, $onlynewkey, $onlynewkeys_num, $split_size, $title, $ourtitle, $toptitlebullet, $specialkey);


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

#say "ourtitle::#$ourtitle#";
#say "toptitlebullet::#$toptitlebullet#";


@onlytitles = split / /, $ourtitle;
@onlytitlebullets = split / /, $toptitlebullet;



for(1..1) {
	#将数组打乱重新组合 但不破坏单个元素
	@newkeys=List::Util::shuffle @newkeys;
	#将数组组成标量
	$allnewkey = join ' ', @newkeys;
	#去除重复单词
	$onlynewkey = &removeduplicate($allnewkey);
	#say "onlynewkey  :: $onlynewkey";
	
	#去除自己标题的单词
	for(@onlytitles) {
		chomp;
		#s/\b$_\b//ig; 加g是错误的
		$onlynewkey =~ s/\b$_\b//i;
	}
	
	#去除前面及后面的空格
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
	#分割数组的size进行求余向上取整
	$split_size = POSIX::ceil($onlynewkeys_num / 10);
	#say "split_size = $split_size :: onlynewkeys_num = $onlynewkeys_num";
	#分割出多个数组
	@allkeys = split_by ( $split_size, @onlynewkeys );
	say "------------------------------------------------------";
	#拆分出多个数组引用
	my $excel_num = 0;
	for my $pointer (@allkeys) {
		#say "pointer :: $pointer";
		my $i = scalar(@$pointer);
		#say "i :: $i";
		$i = 0;
		
		my $pointer_str = "";
		for(@$pointer){
			$pointer_str .= " $$pointer[$i++]";
		}
		
		$pointer_str =~ s/^\ //g;
		$pointer_str =~ s/\ $//g;
		#say "pointer_str :: $pointer_str";
		
		#写到B列
		$excel_num++;
		$Sheet11->Cells($excel_num,1)->{Value} = "$pointer_str";
		push @amazonkeys, "$pointer_str\n";
		#say "";
	}
}


$Book11->Save;
$Book11->Close;
undef $Book11;

undef $Excel;
