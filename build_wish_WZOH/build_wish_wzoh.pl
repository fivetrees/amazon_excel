#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;


my $filename = "input.xlsx";

# 删除wish-*.xlsx
say "";
unlink glob "wish-*.xlsx";


# 产生wish随机数
# wish关键词文件 wish_tags_shoes.xlsx
my $wish_book = Spreadsheet::Read->new("wish_tags_shoes.xlsx");
my $wish_sheet = $wish_book->sheet(1);
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


# 读取表格文件
my $book = Spreadsheet::Read->new("$filename");
# 读取input.xlsx的模板数据，在excel的第1个工作区
my $sheet = $book->sheet(1);

#读取新文件名
my $psku   = $sheet->cell("A2");
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
my $Book = $Excel->Workbooks->Open("E:\\tool\\build_wish_WZOH\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
my $Sheet = $Book->Worksheets(1);


my $ptitle	= $sheet->cell("B2");
my @sizes	= split ' ', $sheet->cell("C2");
my $price 	= $sheet->cell("D2");
my $msrp  	= $sheet->cell("E2");
my $shipping = $sheet->cell("J2");
my $quantity = $sheet->cell("G2");
my $wish_desc = $sheet->cell("I2");
my (%colorsize);
my $wish_num = 2;

my $mydir = "C:/Users/senlin/linux/build_wish_WZOH/imgdir/$huohao";

chdir "$mydir";
my @colors = <*>;

for my $color (@colors) {
	chdir "$mydir";
	if(-d "$color") {
		for my $size (@sizes){
			my $img_num = 12;
			chdir "$color";
			my @pics = <*>;
			for(@pics){
			
				my $url = "http://img.hejiegm.cn/img/wh/" . "$huohao/" . "$color/" . "$_";
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



$Book->Save;
$Book->Close;
undef $Book;
undef $Excel;

chdir "C:/Users/senlin/linux/build_wish_WZOH";
rename "$filename", "${huohao}_input.xlsx";
copy("input_original.xlsx", "input.xlsx") or die "Copy input.xlsx failed: $!";



=pod

凉鞋
产地：温州 中国
上市时间：2017
鞋面材质：超纤	
内里材质：超纤皮
鞋底材质：TPR
风格：休闲
侧帮款式：侧空
鞋帮高度：低帮
后帮款式：前后绊带
功能：透气 轻便舒适柔软
穿着方式：套筒/套鞋
款式：露趾
Origin: Wenzhou China
Time to market: 2017
Upper material: super fiber
Inside material: ultra-filament skin
Sole material: TPR
style: Casual
Side style: side empty
Upper height: low help
After the style: before and after the trip with
Function: breathable light and comfortable
Wear Style: Sleeve
Style: open toe




皮鞋
产地：温州 中国
上市时间：2017
内里材质：仿皮	
风格：休闲
鞋帮高度：低帮	
鞋跟形状：马蹄跟	
鞋底工艺：粘胶鞋
鞋跟高度：低跟（1-3CM）	
穿着方式：前系带	
适用运动：通用	
鞋垫材质：PU	
款式：商务鞋	
适用场景：办公室
功能：透气 轻便舒适柔软
Origin: Wenzhou China
Time to market: 2017
Inside material: imitation leather
style: Casual
Upper height: low help
Heel shape: horseshoe heel
Soles craft: viscose shoes
Heel height: low with (1-3CM)
Wearing style: front lace
Applicable Sports: General
Insole Material: PU
Style: Business shoes
Applicable to the scene: office
Function: breathable light and comfortable and soft





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