#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;
use LWP::Simple;


unlink "ERROR.txt", if(-f "ERROR.txt");

my $filename = "input.xlsx";

# 删除wish-*.xlsx
say "";
unlink glob "wish-*.xlsx";


# 产生wish随机数
# wish关键词文件 wish_tags.xlsx
my $wish_book = Spreadsheet::Read->new("wish_tags.xlsx");
my $wish_sheet = $wish_book->sheet(1);
my %h;
my (@h, @tags);
my $tag;
while (@h < 9) {
    my $r = int rand(100);
    push @h, $r if (!exists $h{$r} and $r > 0);
    $h{$r} = 1;
}
push @tags, "fashion";
push @tags, $wish_sheet->cell("A$_") for(@h);
$tag=join(", ",@tags);


# 读取表格文件
my $book = Spreadsheet::Read->new("$filename");
# 读取input.xlsx的模板数据，在excel的第1个工作区
my $sheet = $book->sheet(1);

#读取新文件名
my $psku	= $sheet->cell("A2");
my @huohaos = split '-', $psku;
my $huohao = $huohaos[-1];
my $diysku = "$huohaos[0]" . "-" . "$huohaos[1]" . "-" . "$huohaos[2]" . "-";


#拷贝新的文件名
my $newxlsx = "wish-${huohao}.xlsx";
copy("wish_original.xlsx", "$newxlsx") or die "Copy wish_original.xlsx failed: $!";


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
my $Book = $Excel->Workbooks->Open("E:\\tool\\build_wish\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
my $Sheet = $Book->Worksheets(1);

#男生休闲短裤
my $wish_desc = "Style: fashionable leisure
Category: youth popular
Pants type: straight type
Pants long: pants / pants
Waist type: waist

Fashion adjustment belt
Will breathe the fabric
Creative pocket
Breathable
Green cotton
High quality zipper";


# line代表产品sku数
my (%colorsize, @image_dirs);
my ($null_num, $num, $line) = 0;
my $wish_num = 2;
my $ptitle = $sheet->cell("B2");
my $price = $sheet->cell("G2");
my $msrp = $sheet->cell("H2");
my $shipping = $sheet->cell("M2");
my $store_name = "";
my $url_head = "http://img.feisenkj.com/img/wh/";

#判断标题是否包含feisen
if($ptitle =~ /feisen/i) {
	$store_name = "feisen";
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


for my $num (2..40) {

	my $color = $sheet->cell("C$num");
	my $size = $sheet->cell("D$num");
	my $image_dir = $sheet->cell("E$num");
	my $image = $sheet->cell("F$num");

	#去除重复image_dir
	unless ( grep { $_ eq $image_dir } @image_dirs ){
		push @image_dirs, $image_dir;
	}

	if($color ne "" and $size ne ""){
		$colorsize{$color} = "$size" . "#" . "$image_dir" . "#" . "$image";

		#总行数
		$line += split ' ', $size;
		#say "line is $line";

		my @sizes = split ' ', $size;
		my @images = split ' ', $image;
		
		#检测图片url是否存在
		for(@images) {
			next if($_ eq "00");
			my $url = "$url_head" . "$image_dir/" . "$_.jpg";
			my $response = head( $url );
			unless($response) {
				#如果图片链接不存在则退出
				#IO::Socket::INET听说这个检测更快
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
		
		for my $wh_size (@sizes) {
			my $img_num = 12;
			for(@images) {
				next if($_ eq "00");
				#say "$wish_num :: $color :: $image_dir :: $wh_size :: $_.jpg";
				my $url = "$url_head" . "$image_dir/" . "$_.jpg";
				$Sheet->Cells($wish_num,$img_num)->{Value} = "$url";
				$img_num++;
			}

			#sku 标题数据写入
			#say "$wish_num :: $color :: $wh_size";
			$Sheet->Cells($wish_num,1)->{Value} = "$psku";
			$Sheet->Cells($wish_num,2)->{Value} = "$psku" .  "-" . "$color" . "-" . "$wh_size";
			$Sheet->Cells($wish_num,3)->{Value} = "$ptitle" . " " . "$color";
			
			#say "$wish_num :: $color :: $wh_size";
			#颜色CL列 以及CM列
			$Sheet->Cells($wish_num,4)->{Value} = "$color";

			#尺寸列 DK DL
			$Sheet->Cells($wish_num,5)->{Value} = "$wh_size";
			
			#库存
			$Sheet->Cells($wish_num,6)->{Value} = "99";
			
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

	} elsif($color eq "" and $size eq ""){
		$null_num++;
		last if($null_num == 2);
	} else{
		say "error !!!!";
		last;
	}
}



$Book->Save;
$Book->Close;
undef $Book;
undef $Excel;

rename "$filename", "${huohao}_input.xlsx";
copy("input_original.xlsx", "input.xlsx") or die "Copy input.xlsx failed: $!";

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