#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Win32::OLE;
use File::Copy;
use LWP::Simple;
use File::Slurp;

#先删除ERROR.txt
unlink "ERROR.txt", if(-f "ERROR.txt");
my $filename = "start.xlsx";

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
my $Book11 = $Excel->Workbooks->Open("E:\\tool\\build_amazon_us_shoes_hejie\\$filename") //  die "Can not open $filename book\n" ;	
my $Sheet11 = $Book11->Worksheets(2);


#读取新文件名
my $psku	= $Sheet11->range("A2")->{Value};
my $ptitle 	= $Sheet11->range("B2")->{Value};
my @huohaos = split '-', $psku;
my $huohao 	= $huohaos[-1];
my $diysku = "$huohaos[0]" . "-" . "$huohaos[1]" . "-" . "$huohaos[2]" . "-";
my $url_head = "http://img.hejiegm.cn/img/am/";
my $store_name = "Threelove";
my $urlcc = "img.hejiegm.cn";


sub check_title_url {
	my $t1 = shift;
	my $t2 = shift;
	unless($ptitle =~ /$t1/i and $url_head =~ /$t2/i) {
		open ERROR_FH, ">ERROR.txt";
		say ERROR_FH "Title Or Url is not match Store Name!!";
		say "Title Or Url is not match Store Name!!!!";
		close ERROR_FH;
		
		$Book11->Close;
		undef $Book11;

		undef $Excel;
		exit;
	}
}

&check_title_url("$store_name", "$urlcc");

#拷贝新的文件名
my $newxlsx = "${store_name}-amazon-${huohao}.xlsx";
copy("NOT_DEL_AMAZON.xlsx", "$newxlsx") or die "Copy NOT_DEL_AMAZON.xlsx failed: $!";

# 打开新文档 写数据
my $Book22 = $Excel->Workbooks->Open("E:\\tool\\build_amazon_us_shoes_hejie\\$newxlsx") //  die "Can not open $newxlsx book\n" ;	
my $Sheet22 = $Book22->Worksheets(4);

sub copy_paste {
	my $t1 = shift;
	my $t2 = shift;
	
	$Sheet11->range("$t1")->copy();
	$Sheet22->range("$t2")->Select();
	$Sheet22->paste();
}

sub copy_paste_multi {
	my $t1 = shift;
	my $t2 = shift;
	my $t3 = shift;
	
	$Sheet11->range("$t1")->copy();
	$Sheet22->range("$t2:$t3")->Select();
	$Sheet22->paste();
}

# line代表产品sku数
my (@main_urls, %colorsize, @image_dirs);
my ($num, $line) = 0;
my $amazon_num = 5;
my $main_img_num = 4;


for my $num (11..200) {

	my $color = $Sheet11->range("A$num")->{Value};
	my $size = $Sheet11->range("B$num")->{Value};
	my $image_dir = $Sheet11->range("C$num")->{Value};
	my $image = $Sheet11->range("D$num")->{Value};

	#如果A列和B列图像类为空则退出
	last unless($Sheet11->range("A$num")->{Value} or $Sheet11->range("B$num")->{Value});

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
		
		for(@images) {
			next if($_ eq "00");
			my $url = "$url_head" . "$image_dir/" . "$_.jpg";
			my $response = head( $url );
			unless($response) {
				sleep 3;
				$response = head( $url );
				unless($response) {
					#如果图片链接不存在则退出
					#IO::Socket::INET听说这个检测更快
					open ERROR_FH, ">ERROR.txt";
					say "can't get url $url";
					say ERROR_FH "can't get url $url";
					close ERROR_FH;
					$Book11->Close;
					undef $Book11;
					$Book22->Close;
					undef $Book22;
					undef $Excel;
					exit;
				}
			}
		}
		
		for my $am_size (@sizes) {
			my $img_num = 59;
			for(@images) {
				next if($_ eq "00");
				#say "$amazon_num :: $color :: $image_dir :: $am_size :: $_.jpg";
				my $url = "$url_head" . "$image_dir/" . "$_.jpg";
				$Sheet22->Cells($amazon_num,$img_num)->{Value} = "$url";
				#主图url
				my $main_url = "$url_head" . "$image_dir/" . "$images[0].jpg";
				push @main_urls, $main_url;

				$img_num++;
				last if($img_num == 67);
			}

			#sku 标题数据写入 # 尺寸转化
			my $title_size = "";
			if($am_size eq "34") {
				$title_size = "US 3.5 (EU 33)";
			} elsif($am_size eq "35") {
				$title_size = "US 4 (EU 34)";
			} elsif($am_size eq "36") {
				$title_size = "US 4.5 (EU 35)";
			} elsif($am_size eq "37") {
				$title_size = "US 5 (EU 36)";
			} elsif($am_size eq "38") {
				$title_size = "US 5.5 (EU 37)";
			} elsif($am_size eq "39") {
				$title_size = "US 6 (EU 38)";
			} elsif($am_size eq "40") {
				$title_size = "US 6.5 (EU 39)";
			} elsif($am_size eq "41") {
				$title_size = "US 7 (EU 40)";
			} elsif($am_size eq "42") {
				$title_size = "US 8 (EU 41)";
			} elsif($am_size eq "43") {
				$title_size = "US 8.5 (EU 42)";
			} elsif($am_size eq "44") {
				$title_size = "US 9.5 (EU 43)";
			} elsif($am_size eq "45") {
				$title_size = "US 10.5 (EU 44)";
			} elsif($am_size eq "46") {
				$title_size = "US 11 (EU 45)";
			} elsif($am_size eq "47") {
				$title_size = "US 12 (EU 46)";
			} elsif($am_size eq "48") {
				$title_size = "US 13 (EU 47)";
			} elsif($am_size eq "49") {
				$title_size = "US 14 (EU 48)";
			} elsif($am_size eq "50") {
				$title_size = "US 15 (EU 49)";
			} elsif($am_size eq "51") {
				$title_size = "US 16 (EU 50)";
			} else {
				$title_size = "$am_size";
			}
			
			#sku 标题数据写入
			#say "$amazon_num :: $color :: $am_size";
			$Sheet22->Cells($amazon_num,1)->{Value} = "$diysku" . "$image_dir" . "-" . "$color" . "-" . "$am_size";
			$Sheet22->Cells($amazon_num,2)->{Value} = "$ptitle" . " " . "$color" . " " . "$am_size";
			
			#主图数据写入
			my %count;
			my $first_main_img = 59;
			my @main_urls = grep { ++$count{ $_ } < 2; } @main_urls;
			for(@main_urls) {
				$Sheet22->Cells($main_img_num,$first_main_img)->{Value} = "$_";
				
				$first_main_img++;
				last if($first_main_img == 67);
			}

			#say "$amazon_num :: $color :: $am_size";
			#颜色CM列 以及CN列
			$Sheet22->Cells($amazon_num,87)->{Value} = "$color";
			$Sheet22->Cells($amazon_num,88)->{Value} = "$color";

			#尺寸列
			if($am_size eq "34") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 3.5 ( EU 33 )";
			} elsif($am_size eq "35") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 4   ( EU 34 )";
			} elsif($am_size eq "36") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 4.5 ( EU 35 )";
			} elsif($am_size eq "37") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 5   ( EU 36 )";
			} elsif($am_size eq "38") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 5.5 ( EU 37 )";
			} elsif($am_size eq "39") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 6   ( EU 38 )";
			} elsif($am_size eq "40") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 6.5 ( EU 39 )";
			} elsif($am_size eq "41") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 7   ( EU 40 )";
			} elsif($am_size eq "42") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 8   ( EU 41 )";
			} elsif($am_size eq "43") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 8.5 ( EU 42 )";
			} elsif($am_size eq "44") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 9.5 ( EU 43 )";
			} elsif($am_size eq "45") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 10.5( EU 44 )";
			} elsif($am_size eq "46") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 11  ( EU 45 )";
			} elsif($am_size eq "47") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 12  ( EU 46 )";
			} elsif($am_size eq "48") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 13  ( EU 47 )";
			} elsif($am_size eq "49") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 14  ( EU 48 )";
			} elsif($am_size eq "50") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 15  ( EU 49 )";
			} elsif($am_size eq "51") {
				$Sheet22->Cells($amazon_num,89)->{Value} = "US 16  ( EU 50 )";
			} else {
				$Sheet22->Cells($amazon_num,89)->{Value} = "$am_size";
			}

			$amazon_num++;
		}
		
	} else{
		say "error !!!!";
		last;
	}
}

my $fnum = 4;
my $snum = 5;
$line += 4;

#主标题 主sku写入
$Sheet22->Cells(4,1)->{Value} = "$psku";
$Sheet22->Cells(4,2)->{Value} = "$ptitle";

#写入UPC########################################
#my @upcs = read_file( 'upc.txt' ) ;
#for my $ln (1..$line){
#	my $upc = shift @upcs;
#	$Sheet22->Cells($snum,3)->{Value} = "$upc";
#	say "upc $upc";	
#	$snum++;
#}
#write_file( 'upc.txt', @upcs ) ;
################################################



#gen copy4 array
#use Array::Utils ":all";
#my @all = qw(A B C D E F G H I J K L M N O P Q R S T U V W X Y Z AA AB AC AD AE AF AG AH AI AJ AK AL AM AN AO AP AQ AR AS AT AU AV AW AX AY AZ BA BB BC BD BE BF BG BH BI BJ BK BL BM BN BO BP BQ BR BS BT BU BV BW BX BY BZ CA CB CC CD CE CF CG CH CI CJ CK CL CM CN CO CP CQ CR CS CT CU CV CW CX CY CZ DA DB DC DD DE DF DG DH DI DJ DK DL DM DN DO DP DQ DR DS DT DU DV DW DX DY DZ EA EB EC ED EE EF EG EH EI);
#my @ignore = qw(A B C D J K M Q BC BG BH BI BJ BK BL BM BN BV BW BX CI CJ CK CS CT CU);
#my @c = array_diff(@all,@ignore);
#print "$_ " for(@c);

my @copy4 = qw(E F G H I L N O P R S T U V W X Y Z AA AB AC AD AE AF AG AH AI AJ AK AL AM AN AO AP AQ AR AS AT AU AV AW AX AY AZ BA BB BD BE BF BO BP BQ BR BS BT BU BY BZ CA CB CC CD CE CF CG CH CL CM CN CO CP CQ CR CV CW CX CY CZ DA DB DC DD DE DF DG DH DI DJ DK DL DM DN DO DP DQ DR DS DT DU DV DW DX DY DZ EA EB EC ED EE EF EG EH EI);


my @copy5 = qw(C D J K M Q BX);


#BU BV行特别处理
&copy_paste("BV2", "BV4");
&copy_paste_multi("BV3", "BV5", "BV$line");
&copy_paste_multi("A2", "BW5", "BW$line");

#从第五行开始复制
for(@copy5) {
	&copy_paste_multi("${_}2", "$_$snum", "$_$line");
	say "55555  ${_}2, ${_}${snum}, ${_}${line}";
}

#从第四行开始复制
for(@copy4) {
	say "No Defined" and next unless(defined $Sheet11->range("${_}2")->{Value});
	my $avalue = $Sheet11->range("${_}2")->{Value};
	say "${_}2 NULL" and next if("$avalue" eq " " or "$avalue" eq "");
	&copy_paste_multi("${_}2", "${_}${fnum}", "${_}${line}");
	say "44444 ${_}2, $_${fnum}, ${_}${line}";
	say "avalue##$avalue##44444 ${_}2, $_${fnum}, ${_}${line}";
}


$Book11->Save;
$Book11->Close;
undef $Book11;

$Book22->Save;
$Book22->Close;
undef $Book22;

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