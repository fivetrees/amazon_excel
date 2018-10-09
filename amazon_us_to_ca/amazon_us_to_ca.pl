#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;

# 美国站xlsl 转换成加拿大xlsl

my $filename = "us.xlsx";
my $newxlsx = "ca.xlsx";

# 删除amazon-ca-*.xlsx
say "";
unlink glob "amazon-ca-*.xlsx";
copy("amazon-ca_original.xlsx", "$newxlsx") or die "Copy amazon-ca_original.xlsx failed: $!";

# 删除错误文件
unlink "ERROR.txt", if(-f "ERROR.txt");

# 读取excel文件
my $book = Spreadsheet::Read->new("$filename");

# 读取亚马逊的模板数据，在excel的第四个工作区
my $sheet = $book->sheet(4);

# 读取父sku
my $psku = $sheet->cell("A4");
my $pproduct_id = $sheet->cell("C4");
my ($pnum, @nums);
my $null_num = 0;

#确定psku存在
say "psku is error!!" and exit if($psku eq "" or $pproduct_id ne "");

for(4..80) {
	my $sku = $sheet->cell("A$_");
	if($sku ne ""){
		push @nums, $_;
	} elsif($sku eq ""){
		#say "line $_ ## null number $null_num ## null data";
		#say "sku is $sku ## size is $size ## color is $color";
		$null_num++;
		last if($null_num == 2);
	} else{
		say "error !!!!";
		last;
	}
}


#创建新表格
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


sub us_to_ca {
	# open existing excel document
	my $Book = $Excel->Workbooks->Open("E:\\tool\\amazon_us_to_ca\\$newxlsx") //  die "Can not open $newxlsx book\n" ;

	# 使用该Excel文档中名为"Upload Template"的Sheet
	my $Sheet = $Book->Worksheets('Template');

	#写数据到表格里面
	for my $num (@nums){
		my $btitle = $sheet->cell("B$num");
		my $len = length("$btitle");
		if($len >= 80) {
			open ERROR_FH,">>ERROR.txt";
			say "ERROR: length is  $len  > 80 !!!";;
			say  ERROR_FH "ERROR: length is $len !!!";;
		}
		
		$Sheet->Cells($num,1)->{Value} = $sheet->cell("A$num");
		$Sheet->Cells($num,3)->{Value} = "Swimwear";
		$Sheet->Cells($num,4)->{Value} = $sheet->cell("F$num");
		$Sheet->Cells($num,5)->{Value} = "$psku";
		$Sheet->Cells($num,6)->{Value} = $sheet->cell("H$num");
		$Sheet->Cells($num,7)->{Value} = $sheet->cell("E$num");
		$Sheet->Cells($num,8)->{Value} = $sheet->cell("B$num");
		$Sheet->Cells($num,9)->{Value} = $sheet->cell("D$num");
		$Sheet->Cells($num,10)->{Value} = $sheet->cell("C$num");
		$Sheet->Cells($num,11)->{Value} = $sheet->cell("E$num");

		#有些数据从第5行开始写
		if($num > 4) {
			$Sheet->Cells($num,12)->{Value} = $sheet->cell("Q$num");
			$Sheet->Cells($num,13)->{Value} = $sheet->cell("J$num") + 3;
			$Sheet->Cells($num,26)->{Value} = $sheet->cell("K$num");
			$Sheet->Cells($num,29)->{Value} = $sheet->cell("M$num");
			$Sheet->Cells($num,69)->{Value} = $sheet->cell("CL$num");
			$Sheet->Cells($num,70)->{Value} = $sheet->cell("CM$num");
			$Sheet->Cells($num,84)->{Value} = $sheet->cell("DK$num");
			$Sheet->Cells($num,85)->{Value} = $sheet->cell("DL$num");
		}
		
		$Sheet->Cells($num,40)->{Value} = $sheet->cell("AZ$num");
		$Sheet->Cells($num,41)->{Value} = $sheet->cell("BA$num");
		$Sheet->Cells($num,42)->{Value} = $sheet->cell("BB$num");
		$Sheet->Cells($num,43)->{Value} = $sheet->cell("BC$num");
		$Sheet->Cells($num,44)->{Value} = $sheet->cell("BD$num");

		$Sheet->Cells($num,51)->{Value} = $sheet->cell("BF$num");
		$Sheet->Cells($num,59)->{Value} = $sheet->cell("BW$num");
		$Sheet->Cells($num,60)->{Value} = $sheet->cell("BV$num");
		$Sheet->Cells($num,61)->{Value} = $sheet->cell("BU$num");
		$Sheet->Cells($num,62)->{Value} = $sheet->cell("BT$num");
		$Sheet->Cells($num,71)->{Value} = $sheet->cell("CP$num");

		#图片处理
		my @extra_image	= ("BG$num", "BH$num", "BI$num", "BJ$num", "BK$num");
		my $lz_img_num = 46;
		for(@extra_image){
			my $img = $sheet->cell("$_");
			#say "extra_image is $_";
			if($img ne "") {
				$Sheet->Cells($num,$lz_img_num)->{Value} = $img;
				$lz_img_num++;
			}
		}
	}

	$Book->Save;
	$Book->Close;
	undef $Book;
}

&us_to_ca;

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