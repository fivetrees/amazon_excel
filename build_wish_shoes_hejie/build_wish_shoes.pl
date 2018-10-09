#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;
use LWP::Simple;
use File::Slurp;

unlink "ERROR.txt", if(-f "ERROR.txt");
my $filename = "input.xlsx";


# 产生wish随机数
# wish关键词文件 NOT_DEL_wish_tags_shoes.xlsx
my $wish_book = Spreadsheet::Read->new("NOT_DEL_wish_tags_shoes.xlsx");
my $wish_sheet = $wish_book->sheet(1);
my %h;
my (@h, @tags);
my $tag;
while (@h < 10) {
    my $r = int rand(100);
    push @h, $r if (!exists $h{$r} and $r > 0);
    $h{$r} = 1;
}

push @tags, $wish_sheet->cell("A$_") for(@h);
$tag=join(", ",@tags);

my @upcs = read_file( 'wish_upc.txt' );

# 读取表格文件
my $book = Spreadsheet::Read->new("$filename");
# 读取input.xlsx的模板数据，在excel的第1个工作区
my $sheet = $book->sheet(1);

#读取新文件名
my $psku   = $sheet->cell("A2");
my $ptitle = $sheet->cell("B2");
my $store_name = "";
my @huohaos = split '-', $psku;
my $huohao = $huohaos[-1];
my $diysku = "$huohaos[0]" . "-" . "$huohaos[1]" . "-" . "$huohaos[2]" . "-";


#判断标题是否包含Threelove
if($ptitle =~ /Threelove/i) {
	$store_name = "hejiegm";
}

#拷贝新的文件名
my $newxlsx = "${store_name}-wish-${huohao}.xlsx";
copy("NOT_DEL_wish_original.xlsx", "$newxlsx") or die "Copy NOT_DEL_wish_original.xlsx failed: $!";


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

# open existing excel document
my $Book = $Excel->Workbooks->Open("E:\\tool\\build_wish_shoes_hejie\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
my $Sheet = $Book->Worksheets(1);


my @sizes	= split ' ', $sheet->cell("C2");
my $price 	= $sheet->cell("D2");
my $msrp  	= $sheet->cell("E2");
my $shipping = $sheet->cell("J2");
my $quantity = $sheet->cell("G2");
my (%colorsize);
my $wish_num = 2;
my $mydir = "E:/tool/build_wish_shoes_hejie/imgdir/$huohao";
my $url_head = "http://img.hejiegm.cn/img/am/";

#鞋子描述
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

#判断标题是否包含Threelove
if($ptitle =~ /Threelove/i) {
	$store_name = "hejiegm";
	#判断标题里面店铺名称是否和url匹配
	unless($url_head =~ /$store_name/i) {
		say "$mydir not exist" and exit unless(chdir "$mydir");
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

say "$mydir not exist" and exit unless(chdir "$mydir");
my @colors = <*>;

for my $color (@colors) {
	chdir "$mydir";
	if(-d "$color") {
	
		chdir "$color";
		my @pics = <*>;
		
		#检测图片url是否存在
		for(@pics) {
			next if($_ eq "00");
			my $url = "$url_head" . "$huohao/" . "$color/" . "$_";
			my $response = head( $url );
			unless($response) {
				#如果图片链接不存在则退出
				#IO::Socket::INET听说这个检测更快
				chdir "$mydir";
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
		
		
		for my $size (@sizes){
			my $img_num = 12;
			#chdir "$color";
			#my @pics = <*>;
			for(@pics){
			
				my $url = "$url_head" . "$huohao/" . "$color/" . "$_";
				$Sheet->Cells($wish_num,$img_num)->{Value} = "$url";
				#say "$wish_num $color :: $_ :: $size :: $img_num";
				$img_num++;
			}
			
			#sku 标题数据写入
			#say "$wish_num :: $color :: $size";
			$Sheet->Cells($wish_num,1)->{Value} = "$psku";
			$Sheet->Cells($wish_num,2)->{Value} = "$psku" . "-" . "$color" . "-" . "$size";
			$Sheet->Cells($wish_num,3)->{Value} = "$ptitle" . " " . "$color";
			
			#写入UPC
			my $upc = shift @upcs;
			$Sheet->Cells($wish_num,24)->{Value} = "$upc";
			
			#say "$wish_num :: $color :: $size";
			#颜色CL列 以及CM列
			$Sheet->Cells($wish_num,4)->{Value} = "$color";

			#尺寸列 DK DL
			$Sheet->Cells($wish_num,5)->{Value} = "$size";
			
			#库存
			$Sheet->Cells($wish_num,6)->{Value} = "$quantity";
			
			#tags关键词
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

write_file( 'wish_upc.txt', @upcs );

$Book->Save;
$Book->Close;
undef $Book;
undef $Excel;

chdir "E:/tool/build_wish_shoes_hejie";
rename "$filename", "${huohao}_input.xlsx";
copy("NOT_DEL_input_original.xlsx", "input.xlsx") or die "Copy NOT_DEL_input_original.xlsx failed: $!";


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
AA # 27
AB # 28
AC # 29
AD # 30
=cut