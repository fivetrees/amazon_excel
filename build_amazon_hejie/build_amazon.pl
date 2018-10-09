#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;
use File::Slurp;

my $filename = "start.xlsx";
my $newxlsx = "amazon-copy.xlsx";

# 删除amazon_*.xlsx
say "";
unlink glob "amazon_*.xlsx";
copy("amazon.xlsx", "$newxlsx") or die "Copy $newxlsx failed: $!";


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
my $Book = $Excel->Workbooks->Open("E:\\tool\\build_amazon_hejie\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
# 使用该Excel文档中名为"Template"的Sheet
my $Sheet = $Book->Worksheets('Template');

# 读取excel文件
my $book = Spreadsheet::Read->new("$filename");
# 读取输入的数据，在excel的第1个工作区
my $sheet = $book->sheet(1);

# line代表产品sku数
my (@main_urls, %colorsize, @image_dirs);
my ($null_num, $num, $line) = 0;
my $amazon_num = 5;
my $main_img_num = 4;
my $psku = $sheet->cell("A2");
my $ptitle = $sheet->cell("B2");

my @huohaos = split '-', $psku;
my $huohao = $huohaos[-1];
my $zdysku = "$huohaos[0]" . "-" . "$huohaos[1]" . "-" . "$huohaos[2]" . "-";


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
		
		for my $am_size (@sizes) {
			my $img_num = 58;
			for(@images) {
				#say "$amazon_num :: $color :: $image_dir :: $am_size :: $_.jpg";
				my $url = "http://img.hejiegm.cn/img/am/" . "$image_dir/" . "$_.jpg";
				$Sheet->Cells($amazon_num,$img_num)->{Value} = "$url";
				#主图url
				my $main_url = "http://img.hejiegm.cn/img/am/" . "$image_dir/" . "$images[0].jpg";
				push @main_urls, $main_url;
				
				$img_num++;
			}

			#sku 标题数据写入
			#say "$amazon_num :: $color :: $am_size";
			$Sheet->Cells($amazon_num,1)->{Value} = "$zdysku" . "$image_dir" . "-" . "$color" . "-" . "$am_size";
			$Sheet->Cells($amazon_num,2)->{Value} = "$ptitle" . " " . "$color" . " " . "$am_size";
			
			#主图数据写入
			my %count;
			my $first_main_img = 58;
			my @main_urls = grep { ++$count{ $_ } < 2; } @main_urls;
			for(@main_urls) {
				$Sheet->Cells($main_img_num,$first_main_img)->{Value} = "$_";
				$first_main_img++;
			}

			#say "$amazon_num :: $color :: $am_size";
			#颜色CL列 以及CM列
			$Sheet->Cells($amazon_num,90)->{Value} = "$color";
			$Sheet->Cells($amazon_num,91)->{Value} = "$color";

			#尺寸列 DK DL
			$Sheet->Cells($amazon_num,115)->{Value} = "$am_size";
			if($am_size eq "XXS") {
				$Sheet->Cells($amazon_num,116)->{Value} = "XX-Small";
			} elsif($am_size eq "XS") {
				$Sheet->Cells($amazon_num,116)->{Value} = "X-Small";
			} elsif($am_size eq "S") {
				$Sheet->Cells($amazon_num,116)->{Value} = "Small";
			} elsif($am_size eq "M") {
				$Sheet->Cells($amazon_num,116)->{Value} = "Medium";
			} elsif($am_size eq "L") {
				$Sheet->Cells($amazon_num,116)->{Value} = "Large";
			} elsif($am_size eq "XL") {
				$Sheet->Cells($amazon_num,116)->{Value} = "X-Large";
			} elsif($am_size eq "XXL") {
				$Sheet->Cells($amazon_num,116)->{Value} = "XX-Large";
			} elsif($am_size eq "3XL") {
				$Sheet->Cells($amazon_num,116)->{Value} = "XXX-Large";
			} elsif($am_size eq "4XL") {
				$Sheet->Cells($amazon_num,116)->{Value} = "XXXX-Large";
			} elsif($am_size eq "5XL") {
				$Sheet->Cells($amazon_num,116)->{Value} = "XXXXX-Large";
			} elsif($am_size eq "6XL") {
				$Sheet->Cells($amazon_num,116)->{Value} = "XXXXXX-Large";
			} else {
				$Sheet->Cells($amazon_num,116)->{Value} = "$am_size";
			}

			$amazon_num++;
		}

	} elsif($color eq "" and $size eq ""){
		$null_num++;
		last if($null_num == 2);
	} else{
		say "error !!!!";
		last;
	}
}

my $fnum = 4;
my $snum = 5;

##################
my @upcs = read_file( 'upc.txt' ) ;
#chomp(@upcs);
#say "upcs num is $upcs";
##################

#主标题 主sku写入
$Sheet->Cells(4,1)->{Value} = "$psku";
$Sheet->Cells(4,2)->{Value} = "$ptitle";
	
for my $ln (1..$line){

	#写入upc
	my $upc = shift @upcs;
	$Sheet->Cells($snum,3)->{Value} = "$upc";
	say "upc $upc";
	
	#J K M Q AZ BA BB BC BD BE 列写入
	$Sheet->Cells($snum,10)->{Value} = $sheet->cell("G2");
	$Sheet->Cells($snum,11)->{Value} = $sheet->cell("H2");
	$Sheet->Cells($snum,13)->{Value} = $sheet->cell("U3");
	$Sheet->Cells($snum,17)->{Value} = $sheet->cell("V3");
	
	#BT BU BV BW CN CP CR DF DG DP DS DT EA 列写入
	$Sheet->Cells(4,72)->{Value} = $sheet->cell("W3");
	$Sheet->Cells($snum,72)->{Value} = $sheet->cell("W4");
	$Sheet->Cells($snum,73)->{Value} = "$psku";
	$Sheet->Cells($snum,74)->{Value} = $sheet->cell("X3");
	
	$snum++;
}


$line += 1;
for my $ln (1..$line){

	#写入upc
	my $upc = shift @upcs;
	$Sheet->Cells($fnum,8)->{Value} = "$upc";
	
	$Sheet->Cells($fnum,52)->{Value} = $sheet->cell("J2");
	$Sheet->Cells($fnum,53)->{Value} = $sheet->cell("K2");
	$Sheet->Cells($fnum,54)->{Value} = $sheet->cell("L2");
	$Sheet->Cells($fnum,55)->{Value} = $sheet->cell("M2");
	$Sheet->Cells($fnum,56)->{Value} = $sheet->cell("N2");
	$Sheet->Cells($fnum,57)->{Value} = $sheet->cell("O2");

	### 固定数据写入
	#D E F G列写入
	$Sheet->Cells($fnum,4)->{Value} = $sheet->cell("Q3");
	$Sheet->Cells($fnum,5)->{Value} = $sheet->cell("R3");
	$Sheet->Cells($fnum,6)->{Value} = $sheet->cell("S3");
	$Sheet->Cells($fnum,7)->{Value} = $sheet->cell("T3");
	
	#BT BU BV BW CN CP CR DF DG DP DS DT EA 列写入
	$Sheet->Cells($fnum,75)->{Value} = $sheet->cell("Y3");
	$Sheet->Cells($fnum,92)->{Value} = $sheet->cell("Z3");
	$Sheet->Cells($fnum,94)->{Value} = $sheet->cell("AA3");
	$Sheet->Cells($fnum,96)->{Value} = $sheet->cell("AB3");
	
	$Sheet->Cells($fnum,110)->{Value} = $sheet->cell("AC3");
	$Sheet->Cells($fnum,111)->{Value} = $sheet->cell("AD3");
	$Sheet->Cells($fnum,120)->{Value} = $sheet->cell("AE3");
	$Sheet->Cells($fnum,123)->{Value} = $sheet->cell("AF3");
	$Sheet->Cells($fnum,124)->{Value} = $sheet->cell("AG3");
	$Sheet->Cells($fnum,131)->{Value} = $sheet->cell("AH3");

	$fnum++;	
}

#say "upcs num is $upcs";
write_file( 'upc.txt', @upcs ) ;

$Book->Save;
$Book->Close;
undef $Book;
undef $Excel;


my $newfile = join '_', @image_dirs;
$newfile =~ s/_$//;

rename "$newxlsx", "$newfile.xlsx";


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
AE # 31
AF # 32
AG # 33
AH # 34
AI # 35
AJ # 36
AK # 37
AL # 38
AM # 39
AN # 40
AO # 41
AP # 42
AQ # 43
AR # 44
AS # 45
AT # 46
AU # 47
AV # 48
AW # 49
AX # 50
AY # 51
AZ # 52
BA # 53
BB # 54
BC # 55
BD # 56
BE # 57
BF # 58
BG # 59
BH # 60
BI # 61
BJ # 62
BK # 63
BL # 64
BM # 65
BN # 66
BO # 67
BP # 68
BQ # 69
BR # 70
BS # 71
BT # 72
BU # 73
BV # 74
BW # 75
BX # 76
BY # 77
BZ # 78
CA # 79
CB # 80
CC # 81
CD # 82
CE # 83
CF # 84
CG # 85
CH # 86
CI # 87
CJ # 88
CK # 89
CL # 90
CM # 91
CN # 92
CO # 93
CP # 94
CQ # 95
CR # 96
CS # 97
CT # 98
CU # 99
CV # 100
CW # 101
CX # 102
CY # 103
CZ # 104
DA # 105
DB # 106
DC # 107
DD # 108
DE # 109
DF # 110
DG # 111
DH # 112
DI # 113
DJ # 114
DK # 115
DL # 116
DM # 117
DN # 118
DO # 119
DP # 120
DQ # 121
DR # 122
DS # 123
DT # 124
DU # 125
DV # 126
DW # 127
DX # 128
DY # 129
DZ # 130
EA # 131
EB # 132
EC # 133
ED # 134
EE # 135
EF # 136
EG # 137
EH # 138
EI # 139
EJ # 140
EK # 141
EL # 142
EM # 143
EN # 144
EO # 145
EP # 146
EQ # 147
ER # 148
ES # 149
ET # 150
EU # 151
EV # 152
EW # 153
EX # 154
EY # 155
EZ # 156
FA # 157
FB # 158
FC # 159
FD # 160
FE # 161
FF # 162
FG # 163
FH # 164
FI # 165
FJ # 166
FK # 167
FL # 168
FM # 169
FN # 170
FO # 171
FP # 172
FQ # 173
FR # 174
FS # 175
FT # 176
FU # 177
FV # 178
FW # 179
FX # 180
FY # 181
FZ # 182
GA # 183
GB # 184
GC # 185
GD # 186
GE # 187
GF # 188
GG # 189
GH # 190
GI # 191
GJ # 192
GK # 193
GL # 194
GM # 195
GN # 196
GO # 197
GP # 198
GQ # 199
GR # 200
GS # 201
GT # 202
GU # 203
GV # 204
GW # 205
GX # 206
GY # 207
GZ # 208
=cut