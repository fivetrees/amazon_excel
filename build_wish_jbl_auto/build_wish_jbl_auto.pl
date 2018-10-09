#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;
use LWP::Simple;

my @products = qw(
YXT901
YXT902
);

unlink "ERROR.txt", if(-f "ERROR.txt");

#Excel ###############################################################################
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

# wish关键词文件 wish_tags_shoes.xlsx
my $wish_book = Spreadsheet::Read->new("NOT_DEL_wish_tags_shoes.xlsx");
my $wish_sheet = $wish_book->sheet(1);

# 读取title.xlsx的模板数据
my $Tbook = Spreadsheet::Read->new("NOT_DEL_title.xlsx");
my $Tsheet = $Tbook->sheet(1);
#Excel ###############################################################################




# 产生wish随机数#############################################################
sub wishtag { 

my %h;
my (@h, @tags);
my $tag;
while (@h < 10) {
	my $r = int rand(702);
	push @h, $r if (!exists $h{$r} and $r > 0);
	$h{$r} = 1;
}

push @tags, $wish_sheet->cell("A$_") for(@h);

$tag=join(", ",@tags);
return $tag;

}
# 产生wish随机数#############################################################


#产生随机标题#################################################################################
sub title {
	my $title = "Threelove Men's";

	#材质
	my $a1 = int rand(11) + 1;
	my $ta1 = $Tsheet->cell("A$a1");
	$title = $title . " " . "$ta1";
	#say "a1 $a1 :: $ta1";
	
	#形容词
	my $b1 = int rand(163) + 1;
	my $b2 = int rand(163) + 1;
	if($b1 != $b2){
		my $tb1 = $Tsheet->cell("B$b1");
		my $tb2 = $Tsheet->cell("B$b2");
		$title = $title . " " . "$tb1" . " " . "$tb2";
		#say "b1 $b1 :: $tb1";
		#say "b2 $b2 :: $tb2";
	} else {
		$b2 = int rand(163) + 1;
		my $tb1 = $Tsheet->cell("B$b1");
		my $tb2 = $Tsheet->cell("B$b2");
		$title = $title . " " . "$tb1" . " " . "$tb2";
		#say "b1 $b1 :: $tb1";
		#say "b2 $b2 :: $tb2";
	}

	
	#场所
	my $c1 = int rand(18) + 1;
	my $tc1 = $Tsheet->cell("C$c1");
	$title = $title . " " . "$tc1";
	#say "c1 $c1 :: $tc1";
	
	#定位
	my $d1 = int rand(18) + 1;
	my $td1 = $Tsheet->cell("D$d1");
	#say "d1 $d1 :: $td1";
	
	$title = $title . " " . "$td1";
	
	$title =~ s/\"//g;
	$title = $title . " " . "shoes";
	$title =~ s/\s+/ /g;
	#say "$d1 :: $td1";
	return "$title";
}
#产生随机标题#################################################################################


for my $pro (@products) {


	#读取新文件名
	my $psku   = "zsl-jbl-wm-$pro";
	my @huohaos = split '-', $psku;
	my $huohao = $huohaos[-1];
	my $diysku = "$huohaos[0]" . "-" . "$huohaos[1]" . "-" . "$huohaos[2]" . "-";





	my $ptitle = &title;
	my @sizes	= split ' ', '6 6.5 7 8 8.5 9';
	my $price 	= int rand(15) + 30;
	my $msrp  	= int rand(30) + 80;
	my $shipping = 6;
	my $quantity = 90;
	my (%colorsize, @urls);
	my $wish_num = 2;
	my $store_name = "";
	my $url_head = "http://img.hejiegm.cn/img/wh/jbl/";

	my $wish_desc = "Dear customer, welcome to our store! We will provide you high quality products. 
	Please choose the right size before you buy our items. 

	Shoes features:
	flyknit screen cloth  
	rubber big bottom running shoe  
	anti-skidding durable anti-friction
	siper light and quick dry  
	flyknit screen side 
	big hole design
	material is air cushion

	please bother us if you have questions!";


	#判断标题是否包含feisen
	if($ptitle =~ /threelove/i) {
		$store_name = "hejie";
		#判断标题里面店铺名称是否和url匹配
		unless($url_head =~ /$store_name/i) {
			open ERROR_FH, ">ERROR.txt";
			say ERROR_FH "url $url_head and store $store_name is not match!!";
			say "url $url_head and store $store_name is not match!!";
			close ERROR_FH;
			$Book->Close;
			undef $Book;
			undef $Excel;
			exit;
		}
	}

	#拷贝新的文件名
	say "chdir E:/tool/build_wish_jbl_auto/ failed!!" and exit unless(chdir "E:/tool/build_wish_jbl_auto/");
	my $newxlsx = "${store_name}-wish-${huohao}.xlsx";
	copy("NOT_DEL_wish_original.xlsx", "$newxlsx") or die "Copy NOT_DEL_wish_original.xlsx failed: $!";


	# open existing excel document
	my $Book = $Excel->Workbooks->Open("E:\\tool\\build_wish_jbl_auto\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
	my $Sheet = $Book->Worksheets(1);
	
	
	my $mydir = "E:/tool/build_wish_jbl_auto/imgdir/$huohao";
	say "chdir $mydir failed!!" and exit unless(chdir "$mydir");
	my @colors = <*>;

	for my $color (@colors) {
		chdir "$mydir";
		if(-d "$color") {
			for my $size (@sizes){
				my $img_num = 12;
				chdir "$color";
				my @pics = <*>;
				for(@pics){
					my $url = "$url_head" . "$huohao/" . "$color/" . "$_";
					push @urls, $url;
					$Sheet->Cells($wish_num,$img_num)->{Value} = "$url";
					#say "$wish_num $color :: $_ :: $size :: $img_num";
					$img_num++;
				}
				
				#sku 标题数据写入
				#say "$wish_num :: $color :: $size";
				$Sheet->Cells($wish_num,1)->{Value} = "$psku";
				$Sheet->Cells($wish_num,2)->{Value} = "$psku" . "-" . "$color" . "-" . "$size";
				$Sheet->Cells($wish_num,3)->{Value} = "$ptitle" . " " . "$color";
				
				#say "$wish_num :: $color :: $size";
				#颜色CL列 以及CM列
				$Sheet->Cells($wish_num,4)->{Value} = "$color";

				#尺寸列 DK DL
				$Sheet->Cells($wish_num,5)->{Value} = "$size";
				
				#库存
				$Sheet->Cells($wish_num,6)->{Value} = "$quantity";
				
				#tags关键词
				my $tag = &wishtag;
				$Sheet->Cells($wish_num,7)->{Value} = "$tag";
				
				#描述
				$Sheet->Cells($wish_num,8)->{Value} = "$wish_desc";
				
				#价格
				$Sheet->Cells($wish_num,9)->{Value} = "$price";
				
				#出厂价
				$Sheet->Cells($wish_num,10)->{Value} = "$msrp";
				
				#运费
				$Sheet->Cells($wish_num,11)->{Value} = "$shipping";
				
				
				$wish_num++;
			}
		}
	}

	#去除重复链接
	my %count;
	my @onlyurls = grep { ++$count{ $_ } < 2; } @urls;

	#检测图片url是否存在
	for my $url (@onlyurls) {
		my $response = head( $url );
		unless($response) {
			#如果图片链接不存在则退出
			#IO::Socket::INET听说这个检测更快
			chdir "E:/tool/build_wish_jbl_auto";
			open ERROR_FH, ">ERROR.txt";
			say "can't get url $url";
			say ERROR_FH "can't get url $url";
			close ERROR_FH;
			$Book->Close;
			undef $Book;
			undef $Excel;
			exit;
		}
	}

	$Book->Save;
	$Book->Close;
	undef $Book;
	sleep 5;
	
}

undef $Excel;



=pod

表格字母列对应数字
A # 1
B # 2
C # 3
D # 4
E # 5
F # 6
G # 7
H # 8
I # 9
J # 10
K # 11
L # 12
M # 13
N # 14
O # 15
P # 16
Q # 17
R # 18
S # 19
T # 20
U # 21
V # 22
W # 23
X # 24
Y # 25
Z # 26
=cut