#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;

# SKU A4
# 标题 B4
# 主图 BF5
# 其他图片 BG5 BH5 BI5 BJ5 BK5 BL5
# 颜色 CL5 ... CL50
# 尺寸 DK5 ... DK50

exit if(@ARGV != 1);
my $filename = "$ARGV[0]";
my $newxlsx = "lazada-${filename}";

# 删除lazada-*.xlsx
say "";
unlink glob "lazada-*.xlsx";
copy("lazada_original.xlsx", "$newxlsx") or die "Copy lazada_original.xlsx failed: $!";

my ($sku, $title, $main_image, $color, @sizes);

# 读取excel文件
my $book = Spreadsheet::Read->new("$filename");

# 读取亚马逊的模板数据，在excel的第四个工作区
my $sheet = $book->sheet(4);

# 读取sku
$sku 			= $sheet->cell("A4");
$title			= $sheet->cell("B4");
#$main_image 	= $sheet->cell("BF5");
#$color			= $sheet->cell("CL5");
#@extra_image	= qw(BF5 BG5 BH5 BI5 BJ5 BK5 BL5);

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

#say "pnum is $pnum";
#say "psku is $psku";
#say "ptitle is $ptitle";

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

my $lazada_Highlights = "
Feisen fashion one piece swimsuits with zipper is of nice quality and Good elasticity., making it comfortable to wear. 
If you are confused with our size, don't hesitate to contact us.    
1. Material: Polyester,  Spandex,  Net Yarn. 
2. Nice quality 
3. Hand Wash .Suit for Swimming,Beach. 
4. It is well designed,it fits and shape your body.
";
sub lazada {
	# open existing excel document
	my $Book = $Excel->Workbooks->Open("E:\\tool\\amazon_convert_lazada\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
	# 使用该Excel文档中名为"Upload Template"的Sheet
	my $Sheet = $Book->Worksheets('Upload Template');

	#写数据到表格里面
	for my $num (@nums){
		my $lz_num = $num - 1;
		#lazada 颜色D4
		$Sheet->Cells($lz_num,4)->{Value} = $sheet->cell("CL$num");
		#lazada 标题B4 及马来西亚标题B4
		$Sheet->Cells($lz_num,78)->{Value} = $sheet->cell("B$num");
		$Sheet->Cells($lz_num,79)->{Value} = $sheet->cell("B$num");
		
		#$Sheet->Cells($lz_num,78)->{Value} = "$ptitle";
		#$Sheet->Cells($lz_num,79)->{Value} = "$ptitle";
		#lazada Highlights CD4
		$Sheet->Cells($lz_num,82)->{Value} = "$lazada_Highlights";
		#lazada size DK4
		my $temp = $sheet->cell("DK$num");
		if($temp eq "XXXL" or $temp eq "xxxl") {
			$Sheet->Cells($lz_num,115)->{Value} = 'Int:' . "3XL";
		} elsif($temp eq "XXXXL" or $temp eq "xxxxl") {
			$Sheet->Cells($lz_num,115)->{Value} = 'Int:' . "4XL";
		} elsif($temp eq "XXXXXL" or $temp eq "xxxxxl") {
			$Sheet->Cells($lz_num,115)->{Value} = 'Int:' . "5XL";
		} else {
			$Sheet->Cells($lz_num,115)->{Value} = 'Int:' . "$temp";
		}
		#lazada price CK4
		my $price = $sheet->cell("J$num");
		my $ship_price = $sheet->cell("M$num");
		$Sheet->Cells($lz_num,90)->{Value} = int(($sheet->cell("J$num") + $sheet->cell("M$num") + 5) * 100 / 23.0335);
	
		#lazada出厂价 CL4
		$Sheet->Cells($lz_num,89)->{Value} = $Sheet->Cells($lz_num,90)->{Value} * 2;
		#lazada sku CO4
		$Sheet->Cells($lz_num,93)->{Value} = $sheet->cell("A$num");
		#lazada image CK4
		my @extra_image	= ("BF$num", "BG$num", "BH$num", "BI$num", "BJ$num", "BK$num", "BL$num");
		my $lz_img_num = 104;
		for(@extra_image){
			my $img = $sheet->cell("$_");
			#say "extra_image is $_";
			if($img ne "") {
				$Sheet->Cells($lz_num,$lz_img_num)->{Value} = $img;
				$lz_img_num++;
			}
		}
	}
	
	$Book->Save;
	#$Book->SaveAs("C:\\Users\\senlin\\linux\\$newxlsx") or die "Save $newxlsx failer.";
	$Book->Close;
	undef $Book;
}

&lazada;

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