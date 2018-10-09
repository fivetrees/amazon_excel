#!/usr/bin/perl

use 5.010;
use strict;
use warnings;
use Spreadsheet::Read;
use Win32::OLE;
use File::Copy;
use LWP::Simple;

my @products = qw(
1108
1116
1122
1132
1133
1158
1165
1302
1520
1522
1559
1569
1713
1718
1733
1737
1759
27182728
5201
557
7597
7760
858
8800
8809
8866
8899
B256
C911
CA29
CJ03
D002
D0057
D225
D226
FF956
FF957
FF958
JBL2712
WM229
YD006
YX8901
YXT90
);

unlink "ERROR.txt", if(-f "ERROR.txt");


for my $pro (@products) {

	#读取新文件名
	my $psku   = "zsl-jbl-wm-$pro";
	my @huohaos = split '-', $psku;
	my $huohao = $huohaos[-1];
	my $diysku = "$huohaos[0]" . "-" . "$huohaos[1]" . "-" . "$huohaos[2]" . "-";

	#拷贝新的文件名
	say "chdir failed!!" and exit unless(chdir "E:/tool/build_wish_jbl_auto/");

	my @sizes	= split ' ', '6 6.5 7 8 8.5 9';
	my (%colorsize, @urls);
	my $store_name = "";
	my $url_head = "http://img.hejiegm.cn/img/wh/jbl/";

	my $mydir = "E:/tool/build_wish_jbl_auto/imgdir/$huohao";
	say "chdir failed!!" and exit unless(chdir "$mydir");
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
					$img_num++;
				}

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
			exit;
		}
	}

	
}




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